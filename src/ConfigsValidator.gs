/**
 * グローバルに定義されている設定値を「勝手に拾って」検証し、
 * Logger に警告/エラーを出すだけのバリデータ。データ記入はしない。
 *
 * 使い方:
 *   new ConfigsValidator().run();
 *
 * 方針:
 * - throw はしない（止めない）
 * - エラー/警告は Logger.log で出力
 * - 可能な限り続行して全部の問題を列挙する
 */
class ConfigsValidator {
  constructor() {
    /** @type {string[]} */
    this.errors = [];
    /** @type {string[]} */
    this.warnings = [];
    /** @type {string[]} */
    this.info = [];
  }

  /** 実行エントリ */
  run() {
    this.errors = [];
    this.warnings = [];
    this.info = [];

    this.validateScalars_();
    this.validateSheetNames_();
    this.validateWillNotUpdateBooksUrl_();
    this.validateDataMoveConfig_();

    this.flushLogs_();
  }

  // -------------------------
  // Validators
  // -------------------------

  validateScalars_() {
    // spreadsheetFirstRowNumber
    if (typeof spreadsheetFirstRowNumber !== "number" || !Number.isInteger(spreadsheetFirstRowNumber) || spreadsheetFirstRowNumber < 1) {
      this.errors.push(`spreadsheetFirstRowNumber は 1以上の整数が必要です（現在: ${String(spreadsheetFirstRowNumber)}）`);
    }

    // spreadsheetColNumber
    if (typeof spreadsheetColNumber !== "number" || !Number.isInteger(spreadsheetColNumber) || spreadsheetColNumber < 1) {
      this.errors.push(`spreadsheetColNumber は 1以上の整数が必要です（現在: ${String(spreadsheetColNumber)}）`);
    }

    // macroProcessMsTimeLimit
    if (typeof macroProcessMsTimeLimit !== "number" || !Number.isFinite(macroProcessMsTimeLimit) || macroProcessMsTimeLimit <= 0) {
      this.errors.push(`macroProcessMsTimeLimit は 0より大きい数値が必要です（現在: ${String(macroProcessMsTimeLimit)}）`);
    }
  }

  validateSheetNames_() {
    if (!Array.isArray(sheetNames)) {
      this.errors.push("sheetNames は配列である必要があります。");
      return;
    }

    if (sheetNames.length === 0) {
      this.warnings.push("sheetNames が空です。更新対象がありません。");
      return;
    }

    const bad = sheetNames.filter(s => typeof s !== "string" || !s.trim());
    if (bad.length) {
      this.errors.push(`sheetNames に空/非文字列が含まれています: ${JSON.stringify(bad)}`);
    }

    const trimmed = sheetNames
      .filter(s => typeof s === "string")
      .map(s => s.trim());

    const dups = this.findDuplicates_(trimmed);
    if (dups.length) {
      this.warnings.push(`sheetNames に重複があります: ${dups.join(", ")}`);
    }
  }

  validateWillNotUpdateBooksUrl_() {
    if (!Array.isArray(willNotUpdateBooksUrl)) {
      this.errors.push("willNotUpdateBooksUrl は配列である必要があります。");
      return;
    }

    const bad = willNotUpdateBooksUrl.filter(u => typeof u !== "string" || !u.trim());
    if (bad.length) {
      this.errors.push(`willNotUpdateBooksUrl に空/非文字列が含まれています: ${JSON.stringify(bad)}`);
    }

    const notUrlLike = willNotUpdateBooksUrl
      .filter(u => typeof u === "string")
      .map(u => u.trim())
      .filter(u => u.length > 0)
      .filter(u => !/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9-_]+/.test(u));

    if (notUrlLike.length) {
      this.warnings.push(`willNotUpdateBooksUrl にスプレッドシートURLっぽくないものがあります: ${JSON.stringify(notUrlLike)}`);
    }

    const dups = this.findDuplicates_(willNotUpdateBooksUrl.map(String));
    if (dups.length) {
      this.warnings.push(`willNotUpdateBooksUrl に重複があります: ${dups.join(", ")}`);
    }
  }

  validateDataMoveConfig_() {
    if (!dataMoveConfig || typeof dataMoveConfig !== "object" || Array.isArray(dataMoveConfig)) {
      this.errors.push("dataMoveConfig は { [sheetName]: [{src,dest}, ...] } 形式のオブジェクトである必要があります。");
      return;
    }

    const keys = Object.keys(dataMoveConfig);
    if (keys.length === 0) {
      this.warnings.push("dataMoveConfig が空です。データ移行設定がありません。");
      return;
    }

    // sheetNames との整合（存在しないキー）
    if (Array.isArray(sheetNames)) {
      const missing = sheetNames
        .filter(s => typeof s === "string")
        .map(s => s.trim())
        .filter(s => s.length > 0)
        .filter(s => !(s in dataMoveConfig));

      if (missing.length) {
        // 「移行なし更新」もありうるので warning
        this.warnings.push(`sheetNames にあるが dataMoveConfig に無いシート（移行なし更新ならOK）: ${missing.join(", ")}`);
      }

      const extra = keys.filter(k => !sheetNames.includes(k));
      if (extra.length) {
        this.info.push(`dataMoveConfig にあるが sheetNames に無いシート: ${extra.join(", ")}`);
      }
    }

    // 各シート設定の検証
    keys.forEach(sheetName => {
      const rules = dataMoveConfig[sheetName];

      if (!Array.isArray(rules)) {
        this.errors.push(`dataMoveConfig["${sheetName}"] は配列である必要があります。`);
        return;
      }

      if (rules.length === 0) {
        this.warnings.push(`dataMoveConfig["${sheetName}"] が空です（移行ルールなし）。`);
        return;
      }

      // dest 重複チェック
      const dests = [];

      rules.forEach((r, idx) => {
        if (!r || typeof r !== "object" || Array.isArray(r)) {
          this.errors.push(`dataMoveConfig["${sheetName}"][${idx}] は {src,dest} オブジェクトである必要があります。`);
          return;
        }

        if (typeof r.src !== "string" || typeof r.dest !== "string") {
          this.errors.push(`dataMoveConfig["${sheetName}"][${idx}] の src/dest は文字列である必要があります。`);
          return;
        }

        const src = r.src.trim();
        const dest = r.dest.trim();
        dests.push(dest);

        const srcParsed = this.parseA1Range_(src);
        const destParsed = this.parseA1Range_(dest);

        if (!srcParsed.ok) {
          this.errors.push(`dataMoveConfig["${sheetName}"][${idx}].src が不正: "${src}"（${srcParsed.reason}）`);
        }
        if (!destParsed.ok) {
          this.errors.push(`dataMoveConfig["${sheetName}"][${idx}].dest が不正: "${dest}"（${destParsed.reason}）`);
        }

        if (srcParsed.ok && destParsed.ok) {
          if (srcParsed.height !== destParsed.height || srcParsed.width !== destParsed.width) {
            this.errors.push(
              `dataMoveConfig["${sheetName}"][${idx}] の範囲サイズ不一致: ` +
              `src="${src}"(${srcParsed.height}x${srcParsed.width}) / dest="${dest}"(${destParsed.height}x${destParsed.width})`
            );
          }
        }
      });

      const dupDest = this.findDuplicates_(dests);
      if (dupDest.length) {
        this.warnings.push(`dataMoveConfig["${sheetName}"] で dest 範囲が重複しています: ${dupDest.join(", ")}`);
      }
    });
  }

  // -------------------------
  // Helpers
  // -------------------------

  /**
   * A1範囲の簡易パース（"A1", "A1:B2", "$A$1:$B$2" を許容）
   * シート名付き（'Sheet'!A1:B2）は想定外としてエラー扱い
   */
  parseA1Range_(a1) {
    const s = String(a1 || "").trim();

    if (!s) return { ok: false, reason: "空です" };

    if (s.includes("!")) {
      return { ok: false, reason: "シート名付きA1表記（!）は想定外です" };
    }

    const re = /^\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/;
    const m = s.match(re);
    if (!m) return { ok: false, reason: "A1表記として解析できません（例: A1 / A1:B2）" };

    const c1 = this.colToNumber_(m[1]);
    const r1 = Number(m[2]);
    const c2 = m[3] ? this.colToNumber_(m[3]) : c1;
    const r2 = m[4] ? Number(m[4]) : r1;

    if (!Number.isFinite(c1) || !Number.isFinite(c2) || c1 < 1 || c2 < 1) {
      return { ok: false, reason: "列が不正です" };
    }
    if (!Number.isInteger(r1) || !Number.isInteger(r2) || r1 < 1 || r2 < 1) {
      return { ok: false, reason: "行が不正です" };
    }

    const minC = Math.min(c1, c2);
    const maxC = Math.max(c1, c2);
    const minR = Math.min(r1, r2);
    const maxR = Math.max(r1, r2);

    return {
      ok: true,
      width: (maxC - minC + 1),
      height: (maxR - minR + 1),
      startCol: minC,
      endCol: maxC,
      startRow: minR,
      endRow: maxR,
    };
  }

  colToNumber_(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      const code = letters.charCodeAt(i);
      if (code < 65 || code > 90) return NaN; // A-Z
      n = n * 26 + (code - 64);
    }
    return n;
  }

  findDuplicates_(arr) {
    const seen = new Set();
    const dup = new Set();
    arr.forEach(v => {
      const key = String(v);
      if (seen.has(key)) dup.add(key);
      else seen.add(key);
    });
    return Array.from(dup);
  }

  flushLogs_() {
    Logger.log("===== ConfigsValidator Report =====");
    if (this.errors.length === 0 && this.warnings.length === 0) {
      Logger.log("OK: エラー/警告は見つかりませんでした。");
    } else {
      if (this.errors.length) {
        Logger.log("ERRORS:");
        this.errors.forEach(e => Logger.log("  - " + e));
      }
      if (this.warnings.length) {
        Logger.log("WARNINGS:");
        this.warnings.forEach(w => Logger.log("  - " + w));
      }
    }
    if (this.info.length) {
      Logger.log("INFO:");
      this.info.forEach(i => Logger.log("  - " + i));
    }
    Logger.log("===================================");
  }
}

/** 便利：単発で走らせる用（任意） */
function validateConfigs() {
  new ConfigsValidator().run();
}
