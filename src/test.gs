function myFunction() {
  
  const srcBookId = PropertiesService.getScriptProperties().getProperty('TemplateSsId');
  const srcBook = SpreadsheetApp.openById(srcBookId); // template 

  Logger.log(srcBook.getName())
}
