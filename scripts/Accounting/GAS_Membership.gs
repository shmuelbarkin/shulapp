var
  SHEETNAME_MEMBERSHIP = 'Membership',
  ROW_OFFSET_MEMBERS = 2,
  ROW_OFFSET_MEMBERSHIPS = 2,

  NUMBER_DAYS_EXPIRED_CC   =  5,
  NUMBER_DAYS_REMINDERSENT = 30,
  NUMBER_DAYS_EXPIRING_CC  = 30,
  NUMBER_DAYS_SIGNUP       = 90,

  COL_MEMBERSHIP_FULLNAME      = 0,
  COL_MEMBERSHIP_DATE          = 1,
  COL_MEMBERSHIP_AMOUNT        = 2,
  COL_MEMBERSHIP_ITEM          = 3,
  COL_MEMBERSHIP_NUMLEFT       = 4,
  COL_MEMBERSHIP_CC_EXPIRATION = 5,
  COL_MEMBERSHIP_CC_LAST_4     = 6,
  COL_MEMBERSHIP_DECLINED      = 7,
  COL_MEMBERSHIP_REMINDER_SENT = 8,

  REGEXP_SEARCH_URL_PAYMENT_NO_AMOUNT = /^(.+)\&xAmount=.*$/
  REGEXP_REPLACE_URL_PAYMENT_NO_AMOUNT = '$1',

  EMAIL_SUBJECT_MEMBERSHIP_EXPIRED     = 'Shul Membership Expired',
  EMAIL_SUBJECT_MEMBERSHIP_EXPIRING    = 'Shul Membership Expiring',
  EMAIL_SUBJECT_MEMBERSHIP_CC_EXPIRED  = 'Membership CC Expired',
  EMAIL_SUBJECT_MEMBERSHIP_CC_EXPIRING = 'Membership CC Expiring',
  EMAIL_SUBJECT_MEMBERSHIP_CC_DECLINED = 'Membership CC Declined',
  EMAIL_SUBJECT_MEMBERSHIP_SIGNUP      = 'Shul Membership',

  SMS_BODY_MEMBERSHIP_EXPIRED     = 'Your Shul membership has expired.\nTo renew your membership please contact us, or click on the link below.\n{{Website}}',
  SMS_BODY_MEMBERSHIP_EXPIRING    = 'Your Shul membership is expiring.\nTo extend your membership please contact us, or respond with the number of months you would like to add.\n{{Website}}',
  SMS_BODY_MEMBERSHIP_CC_EXPIRED  = 'Your Shul membership credit card has expired.\nTo update your credit card information please contact us, or click on the link below.\n{{Website}}',
  SMS_BODY_MEMBERSHIP_CC_EXPIRING = 'Your Shul membership credit card will expire in one month.\nTo update your credit card information please contact us.\n{{Website}}',
  SMS_BODY_MEMBERSHIP_CC_DECLINED = 'Your Shul membership credit card has been declined.\nTo update your credit card information please contact us, or click on the link below.\n{{Website}}',
  SMS_BODY_MEMBERSHIP_SIGNUP      = 'Help us help you! Sign up for Shul membership today.\nTo sign up please contact us, or click on the link below.\nThanks,\b{{Website}}';


// Nightly triggered now
function incrementDateAndDecrementNumLeft()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getSheetByName(SHEETNAME_MEMBERSHIP),
    sheetLastRow = sheet.getLastRow();

  // Premature exit if list of data rows is empty
  if (sheetLastRow < ROW_OFFSET_MEMBERSHIPS) return;

  var
    // Get B(ROW_OFFSET_MEMBERSHIPS):E(last row)
    valsDateAndNumLeft = sheet.getRange(ROW_OFFSET_MEMBERSHIPS, 2, sheetLastRow - ROW_OFFSET_MEMBERSHIPS + 1, 4).getValues(),
    valsDate = [],
    valsNumLeft = [],

    dateToday = new Date(),
    timeToday = dateToday.getTime();


  // Go through all data rows
  for (var i = 0, iLen = valsDateAndNumLeft.length; i < iLen; i++)
  {
    var
      dateCurrent = valsDateAndNumLeft[i][0],
      numLeftCurrent = valsDateAndNumLeft[i][3];

    // Check if date is before today
    if ((isDate_(dateCurrent)) && (dateCurrent.getTime() < timeToday))
    {
      // Increment date by one month
      dateCurrent.setMonth(dateCurrent.getMonth() + 1);

      // Check if numLeft is numeric
      if (isNumeric_(numLeftCurrent))
      {
        // Once a cell is = 0 don't decrease to a negative number
        numLeftCurrent = Math.max(Number(numLeftCurrent) - 1, 0);
      } // if numLeft is numeric
    } // if date is before today

    // Push to arrays to write
    valsDate.push([dateCurrent]);
    valsNumLeft.push([numLeftCurrent]);
  } // for all data rows i

  // Write to Date column
  sheet.getRange(ROW_OFFSET_MEMBERSHIPS, 2, sheetLastRow - ROW_OFFSET_MEMBERSHIPS + 1, 1).setValues(valsDate);
  // Write to Numleft column
  sheet.getRange(ROW_OFFSET_MEMBERSHIPS, 5, sheetLastRow - ROW_OFFSET_MEMBERSHIPS + 1, 1).setValues(valsNumLeft);

} // incrementDateAndDecrementNumLeft()



function sendMembershipReminders()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetMemberships = ss.getSheetByName(SHEETNAME_MEMBERSHIP),
    sheetMembershipsLastRow = sheetMemberships.getLastRow();

  // Premature exit if list of data rows is empty
  if (sheetMembershipsLastRow < ROW_OFFSET_MEMBERSHIPS) return;

  var
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    sheetMembersLastRow = sheetMembers.getLastRow();

  // Premature exit if there are no rows to process
  if (sheetMembersLastRow < 2) return;

  var
    // Get A2:O(last row)
    members = sheetMembers.getRange(2, 1, sheetMembersLastRow - 1, 15).getValues()
      // Don't include empty rows (considering first and last name, not full name)
      .filter(function (member) { return ((!!member[1]) && (!!member[2])); })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName:  member[0],
          firstName: member[1],
          lastName:  member[2],
          address: {
            street:  member[3],
            city:    member[4],
            state:   member[5],
            zip:     member[6]
          },
          billing: {
            email: ((member[14].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[14].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    // Get A(ROW_OFFSET_MEMBERSHIPS):I(last row)
    memberships = sheetMemberships.getRange(ROW_OFFSET_MEMBERSHIPS, 1, sheetMembershipsLastRow - ROW_OFFSET_MEMBERSHIPS + 1, 9).getValues(),
    // Template tokens
    templateTokens = getTemplateTokens_(ss),
    dateToday = new Date();

  // Remove time component of date today
  dateToday.setHours(0, 0, 0, 0);

  // Go through all membership rows
  for (var i = 0, iLen = memberships.length; i < iLen; i++)
  {
    // Don't send any message if the the ReminderSent value in Column I is on or later than 30 days before today
    if (
      (isDate_(memberships[i][COL_MEMBERSHIP_REMINDER_SENT]))
      && (memberships[i][COL_MEMBERSHIP_REMINDER_SENT].getTime() >= (dateToday.getTime() - (NUMBER_DAYS_REMINDERSENT * DAYS_TO_MILLISECONDS)))
    ) continue;

    // Look for the current row's full name in the members list
    var searchMember = findFirstMatch_(members, 'fullName', memberships[i][COL_MEMBERSHIP_FULLNAME]);

    // Check if a member was found
    // Column E == 0
    if (Object.keys(searchMember).length)
    {
      var
        detailsMessage = {
          email: {
            subject:  '',
            message:  ''
          },
          smsMessage: ''
        },
        urlPaymentSite = Utilities.formatString(URL_PAYMENT_SITE,
          '{{Membership URL}}',
          // First Name
          encodeURIComponent(searchMember.firstName),
          // Last Name
          encodeURIComponent(searchMember.lastName),
          // Email
          encodeURIComponent(searchMember.billing.email),
          // Address Line 1
          encodeURIComponent(searchMember.address.street),
          // Address Line 2
          '',
          // City
          encodeURIComponent(searchMember.address.city),
          // State
          encodeURIComponent(searchMember.address.state),
          // Zip
          encodeURIComponent(searchMember.address.zip),
          // Item
          encodeURIComponent(memberships[i][COL_MEMBERSHIP_ITEM]),
          // Amount
          encodeURIComponent(memberships[i][COL_MEMBERSHIP_AMOUNT])
        ).replaceTemplateTokens(templateTokens);

      // Check different conditions for reminders
      // 1. Expired Membership
      if (
        (isNumeric_(memberships[i][COL_MEMBERSHIP_NUMLEFT]))
        && (memberships[i][COL_MEMBERSHIP_NUMLEFT] == 0)
      )
      {
        detailsMessage.email.subject = EMAIL_SUBJECT_MEMBERSHIP_EXPIRED.replaceTemplateTokens(templateTokens);
        detailsMessage.email.message = Utilities.formatString(
          HTML_CODE_EMAIL_BODY_MEMBERSHIP_EXPIRED.replaceTemplateTokens(templateTokens),
          searchMember.firstName,
          urlPaymentSite
        );
        detailsMessage.smsMessage    = SMS_BODY_MEMBERSHIP_EXPIRED.replaceTemplateTokens(templateTokens);
      }
      // 2. Expired Credit Card
      else if (
        // Column F is earlier than 5 days after today
        (
          (isDate_(memberships[i][COL_MEMBERSHIP_CC_EXPIRATION]))
          && (memberships[i][COL_MEMBERSHIP_CC_EXPIRATION].getTime() < (dateToday.getTime() + (NUMBER_DAYS_EXPIRED_CC * DAYS_TO_MILLISECONDS)))
        )
        // Column E > 0
        && (
          (isNumeric_(memberships[i][COL_MEMBERSHIP_NUMLEFT]))
          && (memberships[i][COL_MEMBERSHIP_NUMLEFT] > 0)
        )
      )
      {
        detailsMessage.email.subject = EMAIL_SUBJECT_MEMBERSHIP_CC_EXPIRED.replaceTemplateTokens(templateTokens);
        detailsMessage.email.message = Utilities.formatString(
          HTML_CODE_EMAIL_BODY_MEMBERSHIP_CC_EXPIRED.replaceTemplateTokens(templateTokens),
          searchMember.firstName,
          memberships[i][COL_MEMBERSHIP_CC_LAST_4],
          urlPaymentSite
        );
        detailsMessage.smsMessage    = SMS_BODY_MEMBERSHIP_CC_EXPIRED.replaceTemplateTokens(templateTokens);
      }
      // 3. Credit Card Declined
      else if (memberships[i][COL_MEMBERSHIP_DECLINED] == 'Yes')
      {
        detailsMessage.email.subject = EMAIL_SUBJECT_MEMBERSHIP_CC_DECLINED.replaceTemplateTokens(templateTokens);
        detailsMessage.email.message = Utilities.formatString(
          HTML_CODE_EMAIL_BODY_MEMBERSHIP_CC_DECLINED.replaceTemplateTokens(templateTokens),
          searchMember.firstName,
          memberships[i][COL_MEMBERSHIP_CC_LAST_4],
          urlPaymentSite
        );
        detailsMessage.smsMessage    = SMS_BODY_MEMBERSHIP_CC_DECLINED.replaceTemplateTokens(templateTokens);
      }
      // 4. Expiring Membership
      else if (
        (isNumeric_(memberships[i][COL_MEMBERSHIP_NUMLEFT]))
        && (memberships[i][COL_MEMBERSHIP_NUMLEFT] == 1)
      )
      {
        detailsMessage.email.subject = EMAIL_SUBJECT_MEMBERSHIP_EXPIRING.replaceTemplateTokens(templateTokens);
        detailsMessage.email.message = Utilities.formatString(
          HTML_CODE_EMAIL_BODY_MEMBERSHIP_EXPIRING.replaceTemplateTokens(templateTokens),
          searchMember.firstName
        );
        detailsMessage.smsMessage    = SMS_BODY_MEMBERSHIP_EXPIRING.replaceTemplateTokens(templateTokens);
      }
      // 5. Expiring Credit Card
      else if (
        // Column F is earlier than 30 days after today
        (
          (isDate_(memberships[i][COL_MEMBERSHIP_CC_EXPIRATION]))
          && (memberships[i][COL_MEMBERSHIP_CC_EXPIRATION].getTime() < (dateToday.getTime() + (NUMBER_DAYS_EXPIRING_CC * DAYS_TO_MILLISECONDS)))
        )
        // Column E is > 0
        && (
          (isNumeric_(memberships[i][COL_MEMBERSHIP_NUMLEFT]))
          && (memberships[i][COL_MEMBERSHIP_NUMLEFT] > 0)
        )
      )
      {
        detailsMessage.email.subject = EMAIL_SUBJECT_MEMBERSHIP_CC_EXPIRING.replaceTemplateTokens(templateTokens);
        detailsMessage.email.message = Utilities.formatString(
          HTML_CODE_EMAIL_BODY_MEMBERSHIP_CC_EXPIRING.replaceTemplateTokens(templateTokens),
          searchMember.firstName,
          memberships[i][COL_MEMBERSHIP_CC_LAST_4]
        );
        detailsMessage.smsMessage    = SMS_BODY_MEMBERSHIP_CC_EXPIRING.replaceTemplateTokens(templateTokens);
      }
      else
      {
        // Move to next row (nothing to do here!)
        continue;
      } // if reminder conditions

      // Send message(s)

      // Send E-Mail
      // Check if billing email address exists
      if (!!searchMember.billing.email)
      {
        // Send email receipt
        GmailApp.sendEmail(
          searchMember.billing.email,
          detailsMessage.email.subject,
          '',
          {
            htmlBody: detailsMessage.email.message
          }
        )
      } // if billing email address exists

      // Send SMS
      // Check if SMS # exists
      if (!!searchMember.billing.phone)
      {
        sendSMS_(
          searchMember.billing.phone,
          detailsMessage.smsMessage
        );
      } // if SMS # exists

      // Date Stamp ReminderSent in Column I
      // Get I(current row)
      sheetMemberships.getRange(i + ROW_OFFSET_MEMBERSHIPS, COL_MEMBERSHIP_REMINDER_SENT + 1).setValue(dateToday);

    } // if member was found
  } // for all membership rows i
} // sendMembershipReminders()



function sendSignUpCampaign()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    sheetMembersLastRow = sheetMembers.getLastRow();

  // Premature exit if list of data rows is empty
  if (sheetMembersLastRow < ROW_OFFSET_MEMBERS) return;

  var
    sheetMemberships = ss.getSheetByName(SHEETNAME_MEMBERSHIP),
    sheetMembershipsLastRow = sheetMemberships.getLastRow(),
    membershipsFullNames = [];

  // Check if there are data rows in the memberships tab
  if (sheetMembershipsLastRow >= ROW_OFFSET_MEMBERSHIPS)
  {
    // Get A2:A(last row)
    membershipsFullNames = sheetMemberships.getRange(2, 1, sheetMembershipsLastRow - ROW_OFFSET_MEMBERSHIPS + 1, 1)
      .getValues()
      // Flatten to 1D array
      .map(function(fullName) { return fullName[0]; });
  } // if sheetMembershipsLastRow

  var
    // Get A2:O(last row)
    members = sheetMembers.getRange(2, 1, sheetMembersLastRow - 1, 15).getValues()
      // Include only unlisted members with a type of "Regular"
      .filter(function (member) {
        return (
          // Members Col L is "Regular"
          (member[11] == 'Regular')
          // Not listed in Col A of Membership
          && (membershipsFullNames.indexOf(member[0]) == -1)
        );
      })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName:  member[0],
          firstName: member[1],
          lastName:  member[2],
          address: {
            street:  member[3],
            city:    member[4],
            state:   member[5],
            zip:     member[6]
          },
          billing: {
            email: ((member[14].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[14].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    // Template tokens
    templateTokens = getTemplateTokens_(ss);

  // Go through remaining members to send to
  for (var i = 0, iLen = members.length; i < iLen; i++)
  {
    // Send message(s)

    // Send E-Mail
    // Check if billing email address exists
    if (!!members[i].billing.email)
    {
      var
        urlPaymentSite = Utilities.formatString(URL_PAYMENT_SITE,
          '{{Membership URL}}',
          // First Name
          encodeURIComponent(members[i].firstName),
          // Last Name
          encodeURIComponent(members[i].lastName),
          // Email
          encodeURIComponent(members[i].billing.email),
          // Address Line 1
          encodeURIComponent(members[i].address.street),
          // Address Line 2
          '',
          // City
          encodeURIComponent(members[i].address.city),
          // State
          encodeURIComponent(members[i].address.state),
          // Zip
          encodeURIComponent(members[i].address.zip),
          // Item
          encodeURIComponent('Monthly Membership'),
          // Amount
          ''
        ).replaceTemplateTokens(templateTokens);

      // Send email receipt
      GmailApp.sendEmail(
        members[i].billing.email,
        EMAIL_SUBJECT_MEMBERSHIP_SIGNUP.replaceTemplateTokens(templateTokens),
        '',
        {
          htmlBody: Utilities.formatString(
            HTML_CODE_EMAIL_BODY_MEMBERSHIP_SIGNUP.replaceTemplateTokens(templateTokens),
            members[i].firstName,
            urlPaymentSite.replace(REGEXP_SEARCH_URL_PAYMENT_NO_AMOUNT, REGEXP_REPLACE_URL_PAYMENT_NO_AMOUNT)
          )
        }
      )
    } // if billing email address exists

    // Send SMS
    // Check if SMS # exists
    if (!!members[i].billing.phone)
    {
      sendSMS_(
        members[i].billing.phone,
        SMS_BODY_MEMBERSHIP_SIGNUP.replaceTemplateTokens(templateTokens)
      );
    } // if SMS # exists
  } // for all remaining members to send to i
} // sendSignUpCampaign()
