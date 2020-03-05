/*
   # Função para envio de E-mail ao clicar no Check List
   # Linguagem: JavaScript 
*/

function EnviarEmail3() {          // Nome da função
  
 var start_line = 1;
 var shipping_columm = 2;
 var status = 2;
 var condition = "VERDADEIRO3";          // Condição para o envio do E-mail
 var txt_sent = "FALSO3"                  // Condição de comparação e confirmação de envio 
  
 var guide = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2020');    // Aba ou Guia da planilha referenciada
 var interval = guide.getRange(start_line,1,guide.getLastRow()-start_line+1,status);
 var information = interval.getValues();
 var you_sent = false;
 var email,subject_matter,subject_matter2,subject_matter3, message;         // Variaveis de E-mails, Assunto e Menssagem do E-mail
 
  for(var x=0; x<information.length; ++x)
  
  {
    if((information[x][shipping_columm-1]==condition) && (information[x][shipping_columm-1]!=txt_sent)){    // Condição para enviar o E-mail
      
   var email = "@dankia.com.br";
   subject_matter = "Externo - Reteste Final Finalizado";          // Dados do 1° Destino de E-mail
   var message = "E-mail automático, por favor não responder! - Processo Finalizado!"
   
   /*
   var email2 = "exemplo2@gmail.com"
   subject_matter2 = "Externo - Reteste Final Finalizado"          // Dados do 2° Destino de E-mail
   var message2 = "Processo Finalizado!"
    
   var email3 = "exemplo3@hotmail.com"
   subject_matter3 = "Externo - Reteste Final Finalizado"         // Dados do 3° Destino de E-mail
   var message3 = "Processo Finalizado!"
   */
   
    MailApp.sendEmail(email,subject_matter,message,{htmlBody:message});            // 1° Destino de E-mail
      
    /*
    MailApp.sendEmail(email2,subject_matter2,message2,{htmlBody:message2});       // 2° Destino de E-mail
    MailApp.sendEmail(email3,subject_matter3,message3,{htmlBody:message3});       // 3° Destino de E-mail
    */
      
    guide.getRange(start_line+x,status).setValue(txt_sent);
    you_sent = true;
    SpreadsheetApp.flush(); 
    }
    
    else{}    // Esta condição vazia foi colocada pois ao desmarcar a caixa de seleção, era enviado um email novamente. Com isto o problema foi solucionado.
    
  } 
}