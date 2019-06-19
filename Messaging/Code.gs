/**
  * Creates a custom menu in Google Sheets when the spreadsheet opens.
  */
function onOpen() {
    SpreadsheetApp.getUi().createMenu('MessagingApp')//Create custom menu
        .addItem('Choose File', 'showPicker')//Create item in the custom menu
        .addItem('Send Message', 'main')//Create item in the custom menu
        .addToUi();
}

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();//get the current spreadsheet

//getting all the sheets by their names:
var outboxSheet = spreadsheet.getSheetByName("Outbox");
var sentSheet = spreadsheet.getSheetByName("Sent");
var recipientsSheet = spreadsheet.getSheetByName("Recipients");
var settingsSheet = spreadsheet.getSheetByName("Settings");

//data member that keep the outbox data and the current row in the outbox data
var data = outboxSheet.getDataRange().getValues();

var indexOfList;//gives the flexibility to change the lists's names

//Enum that represent the message types
var MessageType = { sms: "SMS", email: "Email", emailAndSMS: "Email and SMS" };

function main() {
    for (var i = 1, d = new Date(); i < data.length; i++)//for each row in the outbox sheet, starts in row 2:
    {
        var currentRowData = getData(i);//get all the data from the current row

        if (currentRowData.dateTime.getTime() <= d.getTime())//if the date and the time fits
        {
            for (var j = 0; j < currentRowData.list.length; j++)//for each recipient in the list
            {
                var tags = {};
                //a tag can be added like this: tags.tagName = tagValue
                tags.firstName = currentRowData.list[j].firstName;

                if (!isNullOrEmpty(currentRowData.email))//if email is chosen
                {
                    if (currentRowData.list[j].emailSMS == MessageType.email || currentRowData.list[j].emailSMS == MessageType.emailAndSMS)   //if the current recipient want to get email
                    {
                        try
                        {
                            sendMail(currentRowData.list[j], currentRowData, tags);
                        }
                        catch (e)
                        {
                            Logger.log(JSON.stringify(e));
                        }
                    }
                }
                if (!isNullOrEmpty(currentRowData.text))   //if text (sms) is chosen
                {
                    if (currentRowData.list[j].emailSMS == MessageType.sms || currentRowData.list[j].emailSMS == MessageType.emailAndSMS) //if the current recipient want to get sms
                    {
                        sendSMS(currentRowData.list[j], currentRowData, tags);
                    }
                }
            }

            var sentFolder = DriveApp.getFolderById(setting.sentFolderId);  //get the sent folder

            if (!isNullOrEmpty(currentRowData.attachment)) //if an attachment was sent
            {
                moveFile(currentRowData.attachment, sentFolder);   //move the attachment file to the sent folder
            }
            if (!isNullOrEmpty(currentRowData.htmlEmail))   //if the email include DocId
            {
                moveFile(currentRowData.docId, sentFolder);    //move the DocId file to the sent folder
            }

            sentSheet.appendRow(data[i]);   //adding current row to the send sheet

            outboxSheet.deleteRow(i + 1);   //removing current row from the outbox sheet
            data.splice(i, 1);
            i--;
        }
    }
}

function getSetting()//get the setting from the settingsSheet
{
    var settingsSheetValues = settingsSheet.getDataRange().getValues();
    //to get the value: settingsSheetValues[row -1][column -1]
    var setting = {
        defaultList: settingsSheetValues[1][1],
        defaultSubject: settingsSheetValues[2][1],
        sentFolderId: settingsSheetValues[3][1],
//      defaultTitle: settingsSheetValues[4][1]
    };
    return setting;
}
var setting = getSetting();

function getData(i)//set all the data from the current row, if columns are blank get the defaults from the setting sheet
{
    var currentRowData = {
        email: data[i][0], //set the email from the email column
        text: data[i][1],  //set the text from the text column
        listName: data[i][2],  //set the listName from the list column
        subject: data[i][3],   //set the subject from the subject column
        dateTime: data[i][4],  //set the dateTime from the date column
        time: data[i][5],  //get the time from the time column
        attachment: data[i][6],//set the attachment from the attachment column
//      title: data[i][7],//set the title from the title column
//      subTitle: data[i][8]//set the subTitle from the subTitle column
    }
    if (currentRowData.email.indexOf("DocId:") == 0)//if using a Google Doc set htmlEmail
    {
        currentRowData.docId = currentRowData.email.slice(6); //set docId
        currentRowData.htmlEmail = getGoogleDocumentAsHTML(currentRowData.docId); //set htmlEmail
    }

    if (isNullOrEmpty(currentRowData.listName))//if list column is blank get the default list
    {
        currentRowData.listName = setting.defaultList;//set the listName from the default list
    }
    currentRowData.list = getList(currentRowData.listName);//set the actual recipients list

    if (isNullOrEmpty(currentRowData.subject))//if subject column is blank get the default subject
    {
        currentRowData.subject = setting.defaultSubject;//set the subject from the default subject
    }

    if (isNullOrEmpty(currentRowData.dateTime))//if date column is blank get the current date-1
    {
        currentRowData.dateTime = new Date(0);//set the dateTime to a old date
    }

    if (isNullOrEmpty(currentRowData.time))//if time column is blank set to 8:00 AM
    {
        currentRowData.dateTime.setHours(8,0,0,0);
    }

    else//set the dateTime from the time column
    {
        var hours = Math.floor(getValueAsHours(currentRowData.time));
        var minutes = getValueAsMinutes(currentRowData.time) - hours * 60;
        var seconds = getValueAsSeconds(currentRowData.time) - hours * 60 * 60 - minutes * 60;

        currentRowData.dateTime.setHours(hours,minutes,seconds,0);
    }

//    if (isNullOrEmpty(currentRowData.title))//if title column is blank get the default title
//    {
//        currentRowData.title = setting.defaultTitle;//set the title from the default title
//    }

    return currentRowData;
}

function sendSMS(recipient, currentRowData, tags) {

    var text = currentRowData.text;//get the text data
    text = replaceTags(tags, text);//replace the tags
    sendSMS_(recipient.cellphone, text);//send sms
}

function sendMail(recipient, currentRowData, tags)//sending email to the recipients list
{
    var template = "Hi " + recipient.firstName + ", \n\n";

    var htmlBody;

    if (isNullOrEmpty(currentRowData.htmlEmail))//if the email include simple text
    {
        htmlBody = template.replace(/\n/g, '<br>') + currentRowData.email.replace(/\n/g, '<br>');//the body will include the template + the email column content
        if (!isNullOrEmpty(currentRowData.attachment))//if the email include simple text and attachment
        {
            var file = DriveApp.getFileById(currentRowData.attachment);
            MailApp.sendEmail(recipient.emailAddress, currentRowData.subject, template + currentRowData.email, { htmlBody: htmlBody, attachments: [file.getBlob()] });
        }
        else//if the email include simple text and no attachment
        {
            MailApp.sendEmail(recipient.emailAddress, currentRowData.subject, template + currentRowData.email, { htmlBody: htmlBody });
        }
    }
    else if(currentRowData.htmlEmail == "DocId:the file does not exist")//if the email include DocId and the doc does not exist
    {
        //error message
    }
    else//if the email include DocId
    {
        htmlBody = replaceTags(tags, currentRowData.htmlEmail);//the body will include the the document html content and the tags will be replaced

        if (!isNullOrEmpty(currentRowData.attachment))//if the email include DocId and attachment
        {
            var file = DriveApp.getFileById(currentRowData.attachment);
            MailApp.sendEmail(recipient.emailAddress, currentRowData.subject, currentRowData.email, { htmlBody: htmlBody, attachments: [file.getBlob()] });
        }
        else//if the email include DocId and no attachment
        {
            MailApp.sendEmail(recipient.emailAddress, currentRowData.subject, currentRowData.email, { htmlBody: htmlBody });
        }
    }
}

function getList(listName)//the function get a list name and return a recipients array which in that list
{
    var recipientsInLIst = [];

    var recipients = recipientsSheet.getDataRange().getValues();
    var firstRow = recipients.shift();//remove the first row

    indexOfList = firstRow.indexOf(listName);
    if (!(indexOfList < 0))//if the name of the list exists
    {
        for (var i = 0; i < recipients.length; i++) {
            if (!isNullOrEmpty(recipients[i][indexOfList]))//if the current recipient is in the list
            {
                recipientsInLIst.push(recipients[i]);
            }
        }
    }
    recipientsInLIst = recipientsInLIst
    // Create an descriptive object
        .map(function (recipient) {
            return {
                firstName: recipient[0],
                lastName: recipient[1],
                emailAddress: recipient[2],
                cellphone: recipient[3],
                emailSMS: recipient[indexOfList]
            };
        });
    return recipientsInLIst;
}

function attachFile(fileId)//get file id from the file picker and attach the file to the active cell
{
    var cell = spreadsheet.getActiveCell();//cell - the active cell
    if (cell.getColumn() == 1)//if using a Google Doc in the email column
    {
        cell.setValue("DocId:" + fileId);
    }
    if (cell.getColumn() == 7)//if adding a file to attachment column
    {
        cell.setValue(fileId);
    }
}

function getGoogleDocumentAsHTML(id)//get google doc id and return html content
{
    var html;
    try
    {
        DriveApp.getFileById(id);
        var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
        var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + id + "&exportFormat=html";
        var param = {
            method: "get",
            headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true,
        };
        html = UrlFetchApp.fetch(url, param).getContentText();
    }
    catch (e)
    {
        html = "DocId:the file does not exist";
        Logger.log(JSON.stringify(e));
        //in this place show the error message
    }
    finally
    {
        return html;
    }
}

function moveFile(fileId, targetFolder) {
    var file = DriveApp.getFileById(fileId);//get the file
    var fileParents = file.getParents();//get the file Parents
    if (fileParents.hasNext())//if there is a sourceFolder
    {
        var sourceFolder = fileParents.next();//get the file source folder
        sourceFolder.removeFile(file);//remove the file from its original folder
    }
    targetFolder.addFile(file);//add the attachment file to the sent folder
}

function replaceTags(tags, text)//get tags array and text and return the the replaced text
{
    for (var tag in tags)//each tag in the text will be replaced by its value
    {
        if (tags.hasOwnProperty(tag)) {
            text = text.replace('{' + tag + '}', tags[tag]);
        }
    }
    return text;
}

function isNullOrEmpty(str) {
    return (str == null || str == "");
}


///
//From 'https://stackoverflow.com/questions/17715841/how-to-read-the-correct-time-values-from-google-spreadsheet/17727300#17727300'
//(with a little changes)
///
function getValueAsSeconds(value) {

    // Get the date value in the spreadsheet's timezone.
    var spreadsheetTimezone = spreadsheet.getSpreadsheetTimeZone();
    var dateString = Utilities.formatDate(value, spreadsheetTimezone,
        'EEE, d MMM yyyy HH:mm:ss');
    var date = new Date(dateString);

    // Initialize the date of the epoch.
    var epoch = new Date('Dec 30, 1899 00:00:00');

    // Calculate the number of milliseconds between the epoch and the value.
    var diff = date.getTime() - epoch.getTime();

    // Convert the milliseconds to seconds and return.
    return Math.round(diff / 1000);
}

function getValueAsMinutes(value) {
    return getValueAsSeconds(value) / 60;
}

function getValueAsHours(value) {
    return getValueAsMinutes(value) / 60;
}
///
///
