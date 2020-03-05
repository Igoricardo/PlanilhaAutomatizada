function myFunction() {
  // recupera a planilha/aba corrente
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var sh = planilha.getActiveSheet();
  // recupera os valores das células
  var calName = sh.getRange("A1").getValue();
  var date_deb = sh.getRange("B1").getValue();
  var date_fin = sh.getRange("C1").getValue();
  // acessa a API do Google Calendar utilizando o nome do calendário
  var cal = CalendarApp.openByName(calName);
  // procura os eventos do calendário utilizando as datas recuperadas da planilha
  var events = cal.getEvents(new Date(date_deb), new Date(date_fin));

  // percorre o array events
  for (var i = 1 ; i < events.length ; ++i) {
    // para cada evento
    var event = events[i];
    var row = sh.getLastRow()+1;
    // escreve o título na coluna A
    sh.getRange("A"+row).setValue(event.getTitle());
    // a data de início na coluna B
    sh.getRange("B"+row).setValue(event.getStartTime());
    // a data do final do evento na Coluna C
    sh.getRange("C"+row).setValue(event.getEndTime());
  }
}
