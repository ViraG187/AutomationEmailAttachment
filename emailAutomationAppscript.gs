function emailAutomation() {
  var query = GmailApp.search('in:inbox newer_than:4d');  // Search inbox for last week's emails
  var threads = GmailApp.getMessagesForThreads(query);    // Get messages from each thread in 2D array
  
  Logger.log("Total threads: " + threads.length);  // Log total number of threads

  var autofolder = DriveApp.getFoldersByName('Auto_Inputs');
  autofolder = autofolder.next();

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i];           // Iterate through each thread
    var messagesCount = messages.length; // Get total number of messages in the thread
    
    
    Logger.log("Thread " + (i + 1) + ": " + messagesCount + " messages.");  // Log messages count for this thread
    //for (var j = Math.max(0,messagesCount - 2) ; j < messagesCount; j++)
    for (var j = Math.max(0,messagesCount-2) ; j < messagesCount; j++) {  // Iterate through the most recent 2 messages
      var message = messages[j];
      
      Logger.log(message);

      var subject = message.getSubject().trim(); // Get the email subject
      Logger.log("Processing email with subject: " + subject);
      //var folder_name = subject.trim("RE:");

      var emailDate = message.getDate();
      var formattedDate = Utilities.formatDate(emailDate, Session.getScriptTimeZone(), "dd-MM-yy_HH:mm");
      Logger.log(formattedDate)

      // Step 3: Create or find the folder in Google Drive for the project
      var projectFolder = getOrCreateProjectFolder(subject,autofolder);
      Logger.log("Created or found folder for subject: " + subject);

      // Step 4: Save email content to a text file
      var emailContent = message.getPlainBody(); // Get the email body (plain text)
      var emailContent = recentplainmessage(message); // with use of custom functionality
      projectFolder.createFile('Email_Content_' + formattedDate + '.txt', emailContent); // Save it to the folder

      Logger.log("Saved email content for: " + subject);

      // Step 5: Save attachments if any exist
      var attachments = message.getAttachments();
      Logger.log(attachments) // Debugging
      if (attachments.length > 0) {
        var attachmentsFolder = getOrCreateAttachmentsFolder(projectFolder); // Create subfolder
        for (var k = 0; k < attachments.length; k++) {  // Use 'k' for the attachment loop
          var attachment = attachments[k];

          try {
          var file = attachmentsFolder.createFile(attachment);
          Logger.log(attachment.getName());

          var newname = attachment.getName() + "_" + formattedDate ;
          file.setName(newname);
          Logger.log(attachment.getName());

          Logger.log(newname);
        } catch(e) { Logger.log("Attchment not supports MIME")} // Save each attachment
        }
        Logger.log("Saved attachments for: " + subject);
      }
    }
  }
}
//triming for clear inputs
function recentplainmessage(message){
 var body = message.getPlainBody();
 var seperator = /(?:On.*|From:.*|To:.*|Cc.*wrote:)/g;

  var input = body.split(seperator);
 
  return input[0].trim();
}

// Helper function to create or find the folder based on project name (email subject)
function getOrCreateProjectFolder(projectName,autofolder) {
  Logger.log("Looking for folder: " + projectName);
  
  var folders = DriveApp.getFoldersByName(projectName);
  if (folders.hasNext()) {
    Logger.log("Folder already exists: " + projectName);
    return folders.next(); // Return existing folder if found
  } else {
    Logger.log("Creating new folder: " + projectName);
    return autofolder.createFolder(projectName)
    // DriveApp.createFolder(projectName); // Create new folder if not found
  }
}

// Helper function to create an "Attachments" subfolder
function getOrCreateAttachmentsFolder(projectFolder) {
  Logger.log("Looking for attachments folder in: " + projectFolder.getName());
  
  var folders = projectFolder.getFoldersByName('Attachments');
  if (folders.hasNext()) {
    Logger.log("Attachments folder already exists in: " + projectFolder.getName());
    return folders.next(); // Return existing subfolder if found
  } else {      
    Logger.log("Creating attachments folder in: " + projectFolder.getName());
    return projectFolder.createFolder('Attachments'); // Create subfolder
  }
}
