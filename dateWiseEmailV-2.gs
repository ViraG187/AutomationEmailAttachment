function emailAutomationV2() {
  var count = GmailApp.getInboxUnreadCount();
  if ( count === 0){ // Gaurd clause, === strict compaire as js auto corrects type
    return;
  }
  Logger.log(`unread : ${count}`)
  const thread = GmailApp.getInboxThreads(0,count);
  const messages = GmailApp.getMessagesForThreads(thread);

  var autofolderIterator = DriveApp.getFoldersByName('test_inputs');
  var autofolder = autofolderIterator.hasNext() ? autofolderIterator.next() : DriveApp.createFolder('test_inputs');
  // Logger.log("autofolder found")

  for (let i = 0; i < messages.length; i++) {
    for (let j = Math.max(0,messages[i].length-2); j < messages[i].length; j++) {
      var sender = messages[i][j].getFrom();
      Logger.log(sender)
     if ( sender == "Project Engineering <pe@ideametrics.co.in>"){
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
      var folderDate = Utilities.formatDate(date,Session.getScriptTimeZone(),"MMM/dd/yy")
      var fname = "Email_"+ fdate + ".html"

      var dateFolder = getOrcreateDateFolder(folderDate,autofolder);
      var projectFolder0 = getOrcreateProjectFolder(projectName,dateFolder)
      var projectFolder = getOrcreateDateFolder(folderDate,projectFolder0)
  
      var folder_iterator = projectFolder.getFilesByName(fname);

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

        dc_doc(projectFolder,fdate,attachment);

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
     return dateFolder.createFolder(projectName)
  } 
}
function refine_content(content) {
  content = content.replace(/From:.*|To:.*|Sent:.*|Subject:.*/gi, ""); // Remove "From:", "To:", "Sent:", "Subject:" lines
  content = content.replace(/[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}/g, ""); // Remove email addresses (e.g., john.doe@example.com)
  content = content.replace(/On .* at .* wrote:/gi, ""); // Remove quoted email history (e.g., "On Monday, John wrote:")
  
  var signatureRegex = /(Thanks[\s\S]*|Best[\s\S]*|Regards[\s\S]*|Warm Regards[\s\S]*|Kind Regards[\s\S]*|Sincerely[\s\S]*|Cheers[\s\S]*|Sent from my[\s\S]*)/i;
  content = content.replace(signatureRegex, "").trim();
  return content.trim();
}
function dc_doc(dateFolder,fdate,attachment){
  //create csv file
  var csv_name = fdate+".csv";
  var csvIterator = dateFolder.getFilesByName(csv_name);
  var content = "Sr.,Doc name,Doc Type,Modeler Review,FEA Review\n"; 


  if(!csvIterator.hasNext()){
    for(let i=0;i < attachment.length; i++){
      var docnametype = attachment[i].getName();
      var docnamestr = docnametype.split(/\./);
      var docname = docnamestr[0];
      var doctype = docnamestr[1];
      var sr = i+1;

      content += `${sr},${docname},${doctype},,\n`;
    }

    dateFolder.createFile(csv_name,content,MimeType.CSV);
  }

}
