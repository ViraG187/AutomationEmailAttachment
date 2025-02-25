function emailAutomation() {
  var count = GmailApp.getInboxUnreadCount();
  const thread = GmailApp.getInboxThreads(0,count+7);
  const messages = GmailApp.getMessagesForThreads(thread);

  var autofolderIterator = DriveApp.getFoldersByName('test_inputs');
  var autofolder = autofolderIterator.hasNext() ? autofolderIterator.next() : DriveApp.createFolder('test_inputs');
  // Logger.log("autofolder found")

  for (let i = 0; i < messages.length; i++) {
    for (let j = Math.max(0,messages[i].length-3); j < messages[i].length; j++) {
      var sender = messages[i][j].getFrom();

     if ( sender == "pe@ideametrics.co.in"){
      Logger.log(`email skipped from ${sender}`)
      continue;
     }
      var attachment = messages[i][j].getAttachments();
      var projectName0 = messages[i][j].getSubject().trim();
      var projectName = projectName0.replace(/^(Fwd:|Fw:|RE:|Re:|Technical Discussions :|Technical Discussions :|)\s*/gi, "");
         Logger.log(`foldername: ${projectName}`)
      var content0 = messages[i][j].getBody()
      var content = refine_content(content0)
      var date = messages[i][j].getDate();
      var fdate = Utilities.formatDate(date,Session.getScriptTimeZone(), "dd/MMM/yy@HH-MM")
      var folderDate = Utilities.formatDate(date,Session.getScriptTimeZone(),"dd/MMM/yy")
      var fname = "Email_"+ fdate + ".html"

      var dateFolder = getOrcreateDateFolder(folderDate,autofolder);
      var projectFolder = getOrcreateProjectFolder(projectName,dateFolder)
  
      var folder_iterator = dateFolder.getFilesByName(fname);

      if(folder_iterator.hasNext()){
       Logger.log("file already exists")
      } else{ projectFolder.createFile(fname,content)
          Logger.log(`${fname} is saved in ${projectFolder}`)
      }
          
      for(let k=0; k < attachment.length; k++){
        if (attachment.length > 0){
          var attachment_name = attachment[k].getName()
          var rename = fdate + attachment_name
          var attachment_iterator = projectFolder.getFilesByName(rename)

          if (!attachment_iterator.hasNext()){
           var file = projectFolder.createFile(attachment[k].copyBlob())
           
           var new_name = file.setName(rename)
           Logger.log("Attachment saved :" + new_name)

          }else {
            Logger.log("attachment already exists")
          }

        } else { Logger.log("No attachment")}

      }
   }
  } 
}  
function getOrcreateDateFolder(folderDate,auto_inputs){

  var folderIterator = auto_inputs.getFoldersByName(folderDate);

  if (folderIterator.hasNext()){
    Logger.log("Datewise folder Already exists")
    return folderIterator.next()
  } else {
    return auto_inputs.createFolder(folderDate)
  }  

}
function getOrcreateProjectFolder(projectName,dateFolder){

  var folder = dateFolder.getFoldersByName(projectName)

  if(folder.hasNext()){
    Logger.log( projectName + "Folder already exist : ")
    return folder.next()
  } else {
     Logger.log(projectName + "New folder created")
     return dateFolder.createFolder(projectName);
  } 
}
function refine_content(content) {
  content = content.replace(/From:.*|To:.*|Sent:.*|Subject:.*/gi, ""); // Remove "From:", "To:", "Sent:", "Subject:" lines
  content = content.replace(/[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}/g, ""); // Remove email addresses (e.g., john.doe@example.com)
  content = content.replace(/On .* wrote:/gi, ""); // Remove quoted email history (e.g., "On Monday, John wrote:")
  
  var signatureRegex = /(Thanks[\s\S]*|Best[\s\S]*|Regards[\s\S]*|Warm Regards[\s\S]*|Kind Regards[\s\S]*|Sincerely[\s\S]*|Cheers[\s\S]*|Sent from my[\s\S]*)/i;
  content = content.replace(signatureRegex, "").trim();
  return content.trim();
}
