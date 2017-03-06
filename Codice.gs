function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ABIT')
      .addItem('Nuovo mese', 'nuovoMeseSidebar')
      .addToUi();
}

function nuovoMeseSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('NuovoMeseSidebar')
      .setTitle('Crea un nuovo mese')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function creaFoglio(mese, anno) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("template");
  
  var nomiMesi = new Array();
  nomiMesi[0] = "Gennaio";
  nomiMesi[1] = "Febbraio";
  nomiMesi[2] = "Marzo";
  nomiMesi[3] = "Aprile";
  nomiMesi[4] = "Maggio";
  nomiMesi[5] = "Giugno";
  nomiMesi[6] = "Luglio";
  nomiMesi[7] = "Agosto";
  nomiMesi[8] = "Settembre";
  nomiMesi[9] = "Ottobre";
  nomiMesi[10] = "Novembre";
  nomiMesi[11] = "Dicembre";
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  if (sheets.length > 1) {
    for (var i=0; i<sheets.length; i++){
      if (sheets[i].getName() == nomiMesi[Number(mese)-1] + " " + anno){
        var error = true
        
        var ui = SpreadsheetApp.getUi();

        var result = ui.alert(
          'Il foglio \"' + nomiMesi[Number(mese)-1] + " " + anno + '\" esiste già',
          'Scegliere un altro mese',
          ui.ButtonSet.OK);

        } else {
          error = false;
        }
    }
  }

  if (!error){
    var newSheet = sourceSheet.copyTo(ss).setName(nomiMesi[Number(mese)-1] + " " + anno);
    newSheet.activate()
      .getRange("A1").setValue("01/"+mese+"/"+anno);
  
    var lastDay = new Date(Number(anno), Number(mese), 0);
    nrGiorni = lastDay.getDate();
    if (nrGiorni < 31){
      newSheet.deleteColumns(2+nrGiorni, 31-nrGiorni);
    }
  }
  
}
