function emailAutomation() {
  var count = GmailApp.getInboxUnreadCount(); 
  Logger.log("new unread email :" + count)
    const threads = GmailApp.getInboxThreads(0, 50);
    const messages = GmailApp.getMessagesForThreads(threads);

    var autofolderIterator = DriveApp.getFoldersByName('Auto_Inputs');
    var autofolder = autofolderIterator.hasNext() ? autofolderIterator.next() : DriveApp.createFolder('Auto_Inputs');

    for (let i = 0; i < messages.length; i++) {
        for (let j = Math.max(0, messages[i].length - 3); j < messages[i].length; j++) {
            var sender = messages[i][j].getFrom();

            if (sender !== "pe@ideametrics.co.in") {
                var attachments = messages[i][j].getAttachments();
                var projectName = messages[i][j].getSubject().trim();
                var content = messages[i][j].getBody();
                var date = messages[i][j].getDate();
                var fdate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MMM/yy@HH");
                var fname = "Email_" + fdate + ".html";

                var projectFolder = getOrCreateProjectFolder(projectName, autofolder);
                var fileIterator = projectFolder.getFilesByName(fname);

                if (fileIterator.hasNext()) {
                    Logger.log("File already exists: " + fname);
                } else {
                    projectFolder.createFile(fname, content);
                    Logger.log("Saved email as file: " + fname);
                }

                if (attachments.length > 0) {
                      var attachmentFolder = getOrCreateAttachmentFolder(projectFolder);
                      for (let k = 0; k < attachments.length; k++) {
                        var attachmentName = attachments[k].getName();
                        var newName = fdate + "_" + attachmentName;
                        var attachmentIterator = attachmentFolder.getFilesByName(newName);

                        if (!attachmentIterator.hasNext()) {
                          var file = attachmentFolder.createFile(attachments[k].copyBlob());
                          file.setName(newName);
                          Logger.log("Attachment saved: " + newName);
                          
                        } else {
                          Logger.log("Attachment already exists: " + attachmentName);
                        }
                      }
                    } else {
                      Logger.log("No attachments found for this email");
                }
            }
        }
    }
}

function getOrCreateProjectFolder(projectName, parentFolder) {
    var folderIterator = DriveApp.getFoldersByName(projectName);

    if (folderIterator.hasNext()) {
        Logger.log("Folder already exists: " + projectName);
        return folderIterator.next();
    } else {
        Logger.log("Creating new project folder: " + projectName);
        return parentFolder.createFolder(projectName);
    }
}

function getOrCreateAttachmentFolder(projectFolder) {
    var folderIterator = projectFolder.getFoldersByName("Attachments");

    if (folderIterator.hasNext()) {
        Logger.log("Attachment folder already exists.");
        return folderIterator.next();
    } else {
        Logger.log("Creating new attachment folder.");
        return projectFolder.createFolder("Attachments");
    }
}
