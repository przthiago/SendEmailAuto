function sendEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2; // A partir da segunda linha
  const emailRange = sheet.getRange('A' + startRow + ':A');
  const emailList = emailRange.getValues();
  
  const subject = "Soluções Logísticas Personalizadas";
  const message = "Olá, boa tarde!\n\nMeu nome é Thiago e sou representante da Zion Logística Transporte e Armazenamento. Estou entrando em contato para apresentar nossos serviços especializados em transporte, com cobertura completa na região de São Paulo e Grande São Paulo e Regiões Metropolitanas.\n\nEm anexo, envio uma apresentação detalhada da nossa empresa.\n\nAgradeço se puder enviar ao destinatário correto!\n\nAtenciosamente,\n\n(11) 99266-1479";

  const attachments = [
    DriveApp.getFileById('18JbR0vGEPqi4wHE6F6XuAhsUqTR6KeiK'), // ID do arquivo
    DriveApp.getFileById('1RiyKH3N8tFJUWp6ssmSscoLJ4oD2u_Vs'), // ID do arquivo
    DriveApp.getFileById('1kbRkitiPPLWCoAO3j5pOl54aI511J89W')  // ID do arquivo
  ];

  for (let i = 0; i < emailList.length; i++) {
    const email = emailList[i][0];
    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: message,
          attachments: attachments
        });
        Logger.log("E-mail enviado para: " + email);
        Utilities.sleep(5000); // Espera 5 segundos
      } catch (e) {
        Logger.log("Erro ao enviar e-mail para " + email + ": " + e.message);
      }
    } else {
      Logger.log("E-mail vazio na linha: " + (i + startRow));
    }
  }
}
