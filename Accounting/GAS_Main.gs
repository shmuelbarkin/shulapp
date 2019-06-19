var
  // Edit URL handle names here
  URL_HANDLES = [
    'first-name',      // First Name
    'last-name',       // Last Name
    'email',           // Email
    'address-1',       // Address Line 1
    'address-2',       // Address Line 2
    'city',            // City
    'state',           // State
    'zip',             // Zip
    'item',            // Item
    'amount'           // Amount
  ],
  URL_PAYMENT_SITE = '%s' + '?g=1'
    // Add URL parameters
    + URL_HANDLES
      // Include separators per parameter
      .map( function(handle) { return ('&' + handle + '=%s'); })
      // Combine all parameters into a single string
      .reduce(function(prevValue, currValue) { return prevValue + currValue; }),

  SHEETNAME_MEMBERS = 'Members',
  SHEETNAME_SETTINGS = 'Settings',

  ROW_OFFSET_SETTINGS = 2,

  DATEFORMAT_RECEIPT = 'MM/d/yyyy',
  DAYS_TO_MILLISECONDS = 86400000;


function onOpen()
{
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Utilities')
    .addSubMenu(ui.createMenu('Pledges')
      .addItem('Process Bills', 'processBills')
      .addItem('Process Receipts', 'processReceiptsAndSend')
      .addItem('Process Payments', 'processReceiptsAndDontSend')
    )
    .addSubMenu(ui.createMenu('Membership')
      .addItem('Process Increment Date and Decrement Numleft', 'incrementDateAndDecrementNumLeft')
      .addItem('Process Membership Messages', 'sendMembershipReminders')
      .addItem('Process Signup Campaign', 'sendSignUpCampaign')
    )
    .addSubMenu(ui.createMenu('Simcha')
      .addItem('Process Invitation Message', 'sendSimchaInvitations')
      .addItem('Process Simcha Reminder', 'sendSimchaReminder2_A')
      .addItem('Process Bris Reminder', 'sendSimchaReminder2_B')
      .addItem('Process Shalom Zachor, Kiddush Reminder', 'sendSimchaReminder2_C')
    )
  .addToUi();
} // onOpen()



function getTemplateTokens_(ss)
{
  // Get template tokens
  var
    sheetSettings = ss.getSheetByName(SHEETNAME_SETTINGS),
    sheetSettingsLastRow = sheetSettings.getLastRow();

  // Premature exit if there are no template tokens
  if (sheetSettingsLastRow < ROW_OFFSET_SETTINGS) return this;

  // Get A(ROW_OFFSET_SETTINGS):B(sheetSettingsLastRow)
  return sheetSettings.getRange(ROW_OFFSET_SETTINGS, 1, sheetSettingsLastRow - ROW_OFFSET_SETTINGS + 1, 2).getValues();
} // getTemplateTokens_()



// Replace template tokens with the values from the Settings sheet
String.prototype.replaceTemplateTokens = function(templateTokens) {
  var newString = this;

  // Go through all template tokens
  for (var i = 0, iLen = templateTokens.length; i < iLen; i++)
  {
    // Replace token
    newString = newString.replace(new RegExp('\{\{' + escapeRegExp(templateTokens[i][0]) + '\}\}', 'g'), templateTokens[i][1]);
  } // for all template tokens i

  return newString;
} // replaceTemplateTokens()



// http://stackoverflow.com/a/6969486
function escapeRegExp(str)
{
  return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
} // escapeRegExp()
