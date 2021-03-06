var
  SHEETNAME_SIMCHA = 'Simcha List',
  ROW_OFFSET_SIMCHA = 3,

  COL_SIMCHA_NAME                 =  0,
  COL_SIMCHA_EVENT                =  1,
  COL_SIMCHA_EVENT_INVITE_TEXT    =  2,
  COL_SIMCHA_EVENT_DATE           =  3,
  COL_SIMCHA_EVENT_TIME           =  4,
  COL_SIMCHA_EVENT_PLACE          =  5,
  COL_SIMCHA_EVENT_ADDRESS        =  6,
  COL_SIMCHA_EANDW_CHOSSON        =  7,
  COL_SIMCHA_EANDW_KALLAH         =  8,
  COL_SIMCHA_EANDW_KABOLAS        =  9,
  COL_SIMCHA_EANDW_CHUPA          = 10,
  COL_SIMCHA_EANDW_SIMCHAS        = 11,
  COL_SIMCHA_TIMESTAMP_INVITATION = 12,
  COL_SIMCHA_TIMESTAMP_REMINDER   = 13,

  TIMEFORMAT_SIMCHA = 'hh:mm aa',
  ORDER_LISTITEMS_SIMCHA = [
    {index: COL_SIMCHA_EANDW_KABOLAS, header: 'Kabolas Ponim'},
    {index: COL_SIMCHA_EANDW_CHUPA,   header: 'Chupa'},
    {index: COL_SIMCHA_EANDW_SIMCHAS, header: 'Simchas Chosson v\'Kallah'},
    {index: COL_SIMCHA_EANDW_CHOSSON, header: 'Chosson'},
    {index: COL_SIMCHA_EANDW_KALLAH,  header: 'Kallah'},
    {index: COL_SIMCHA_EVENT_PLACE,   header: 'Place'},
    {index: COL_SIMCHA_EVENT_ADDRESS, header: 'Address'}
  ],
  NUMBER_DAYS_REMINDER_2C = 2,

  SMS_BODY_SIMCHA_INVITATION = '%s %s\n%s %s\n%s, %s',
  SMS_BODY_SIMCHA_REMINDER   = '%s - %s %s\n%s %s\n%s, %s';


function sendSimchaInvitations()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetSimcha = ss.getSheetByName(SHEETNAME_SIMCHA),
    sheetSimchaLastRow = sheetSimcha.getLastRow();

  // Premature exit if there are no data rows to process
  if (sheetSimchaLastRow < ROW_OFFSET_SIMCHA) return;

  var
    // Get A(ROW_OFFSET_SIMCHA):M(last row)
    rowsSimcha = sheetSimcha.getRange(ROW_OFFSET_SIMCHA, 1, sheetSimchaLastRow - ROW_OFFSET_SIMCHA + 1, 13).getValues(),
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    // Get A2:O(last row)
    members = sheetMembers.getRange(2, 1, sheetMembers.getLastRow() - 1, 16).getValues()
      // Don't include empty rows (considering first and last name, not full name)
      .filter(function (member) { return ((!!member[1]) && (!!member[2])); })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName: member[0],
          firstName: member[1],
          lastName: member[2],
          address: {
            street: member[3],
            zip: member[6]
          },
          billing: {
            email: ((member[15].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[15].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    // Template tokens
    templateTokens = getTemplateTokens_(ss),
    dateToday = new Date();

  // Remove time component of date today
  dateToday.setHours(0, 0, 0, 0);

  // Go through all Simcha list rows
  for (var i = rowsSimcha.length - 1; i >= 0; i--)
  {
    // Skip if Col D: Event Date is empty or not a date
    if ((!rowsSimcha[i][COL_SIMCHA_EVENT_DATE]) || (!isDate_(rowsSimcha[i][COL_SIMCHA_EVENT_DATE]))) continue;

    // Check if Col D: Event Date is before today
    if (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getTime() < dateToday.getTime())
    {
      // Delete this row
      sheetSimcha.deleteRow(i + ROW_OFFSET_SIMCHA);
      // Go to next row
      continue;
    } // if event date is before today

    // Skip if the required parameters do not exist
    if (
      // Col A: Member Name is empty
      (!rowsSimcha[i][COL_SIMCHA_NAME])
      // Col B: Event is empty
      || (!rowsSimcha[i][COL_SIMCHA_EVENT])
      // Col M: Invitation Timestamp is not empty
      || (!!rowsSimcha[i][COL_SIMCHA_TIMESTAMP_INVITATION])
    ) continue;

    // Look for the current row's full name in the members list
    var searchMember = findFirstMatch_(members, 'fullName', rowsSimcha[i][COL_SIMCHA_NAME]);

    // Check if name is in the member list
    if (Object.keys(searchMember).length)
    {
      var
        strEventDate = Utilities.formatDate(rowsSimcha[i][COL_SIMCHA_EVENT_DATE], Session.getScriptTimeZone(), DATEFORMAT_RECEIPT),
        strEventTime = ((isDate_(rowsSimcha[i][COL_SIMCHA_EVENT_TIME]))
          ? Utilities.formatDate(rowsSimcha[i][COL_SIMCHA_EVENT_TIME], Session.getScriptTimeZone(), TIMEFORMAT_SIMCHA)
          : rowsSimcha[i][COL_SIMCHA_EVENT_TIME]
        ),
        strListEmailRecipients = members
          // Remove blanks and duplicates
          .filter(function(member, index, existingMembers) {
            // Premature exit if no email
            if (!member.billing.email) return false;

            // Go through all existing members in the list
            for (var j = 0, jLen = index; j < jLen; j++)
            {
              // Check if existing before
              if (existingMembers[j].billing.email == member.billing.email)
              {
                // Premature exit
                return false;
              } // if existing before
            } // for all existing members in the list j

            return true;
          })
          .map(function(member) {
            return {
              firstName: member.firstName,
              email: member.billing.email
            };
          }),
        listNumbers = members.filter(function(member) { return (!!member.billing.phone); })
          .map(function(member) { return member.billing.phone; })
          .unique();

      // Send message(s)
      // Check if billing email addresses exist
      if (strListEmailRecipients.length)
      {
        // Send E-Mail
        var
          titleEmail = searchMember.lastName + ' ' + rowsSimcha[i][COL_SIMCHA_EVENT],
          htmlListItems =
            Utilities.formatString(HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens), 'Date: ' + strEventDate)
            + ((!!strEventTime)
              ? Utilities.formatString(
                HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                'Time: ' + strEventTime
              )
              : ''
            );

        // Go through order of list items to add
        for (var j = 0, jLen = ORDER_LISTITEMS_SIMCHA.length; j < jLen; j++)
        {
          // Check if field is not empty
          if (!!(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[j].index]))
          {
            // Add to HTML code of list items
            htmlListItems += Utilities.formatString(
              HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
              ORDER_LISTITEMS_SIMCHA[j].header + ': '
                + ((isDate_(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[j].index]))
                  ? Utilities.formatDate(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[j].index], Session.getScriptTimeZone(), TIMEFORMAT_SIMCHA)
                  : rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[j].index]
                )
            );
          } // if field is not empty
        } // for all list items to add j

        // Go through list of email recipients to send to
        for (var j = 0, jLen = strListEmailRecipients.length; j < jLen; j++)
        {
          // Send email to one recipient
          GmailApp.sendEmail(
            strListEmailRecipients[j].email,
            titleEmail,
            '',
            {
              htmlBody: Utilities.formatString(
                HTML_CODE_EMAIL_BODY_SIMCHA.replaceTemplateTokens(templateTokens),
                // Title
                titleEmail,
                // First Name
                strListEmailRecipients[j].firstName,
                // Event
                rowsSimcha[i][COL_SIMCHA_EVENT],
                // Event Invite Text
                rowsSimcha[i][COL_SIMCHA_EVENT_INVITE_TEXT],
                // List Items
                htmlListItems,
                // Name
                searchMember.fullName
              )
            }
          );
        } // for all email recipients j
      } // if list of email recipients is not empty

      // Send SMS
      // Go through all numbers to send to
      for (var k = 0, kLen = listNumbers.length; k < kLen; k++)
      {
        sendSMS_(
          listNumbers[k],
          Utilities.formatString(
            SMS_BODY_SIMCHA_INVITATION.replaceTemplateTokens(templateTokens),
            // Name
            searchMember.fullName,
            // Event
            rowsSimcha[i][COL_SIMCHA_EVENT],
            // Date
            strEventDate,
            // Time
            strEventTime,
            // Place
            rowsSimcha[i][COL_SIMCHA_EVENT_PLACE],
            // Address
            rowsSimcha[i][COL_SIMCHA_EVENT_ADDRESS]
          )
        );
      } // for all numbers to send to k

      // Write timestamp to Col M: Invitation Timestamp
      // Get M(current row)
      sheetSimcha.getRange(i + ROW_OFFSET_SIMCHA, COL_SIMCHA_TIMESTAMP_INVITATION + 1).setValue(dateToday);

    } // if name is in the member list
  } // for all Simcha list rows i
} // sendSimchaInvitations()



function sendSimchaReminder2_A() { sendSimchaReminder2_Today_(['Chasuna', 'Sheva Brochos', 'Vort']); } // sendSimchaReminder2_A()
function sendSimchaReminder2_B() { sendSimchaReminder2_Today_(['Bris']); }                             // sendSimchaReminder2_B()


function sendSimchaReminder2_Today_(eventTypesValid)
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetSimcha = ss.getSheetByName(SHEETNAME_SIMCHA),
    sheetSimchaLastRow = sheetSimcha.getLastRow();

  // Premature exit if there are no data rows to process
  if (sheetSimchaLastRow < ROW_OFFSET_SIMCHA) return;

  var
    // Get A(ROW_OFFSET_SIMCHA):N(last row)
    rowsSimcha = sheetSimcha.getRange(ROW_OFFSET_SIMCHA, 1, sheetSimchaLastRow - ROW_OFFSET_SIMCHA + 1, 14).getValues(),
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    // Get A2:O(last row)
    members = sheetMembers.getRange(2, 1, sheetMembers.getLastRow() - 1, 16).getValues()
      // Don't include empty rows (considering first and last name, not full name)
      .filter(function (member) { return ((!!member[1]) && (!!member[2])); })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName: member[0],
          firstName: member[1],
          lastName: member[2],
          address: {
            street: member[3],
            zip: member[6]
          },
          billing: {
            email: ((member[15].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[15].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    // Template tokens
    templateTokens = getTemplateTokens_(ss),
    dateToday = new Date();

  // Remove time component of date today
  dateToday.setHours(0, 0, 0, 0);

  // Go through all Simcha list rows
  for (var i = rowsSimcha.length - 1; i >= 0; i--)
  {
    // Skip if Col D: Event Date is empty or not a date
    if ((!rowsSimcha[i][COL_SIMCHA_EVENT_DATE]) || (!isDate_(rowsSimcha[i][COL_SIMCHA_EVENT_DATE]))) continue;

    // Check if Col D: Event Date is before today
    if (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getTime() < dateToday.getTime())
    {
      // Delete this row
      sheetSimcha.deleteRow(i + ROW_OFFSET_SIMCHA);
      // Go to next row
      continue;
    } // if event date is before today

    // Skip if the required parameters do not exist
    if (
      // Col A: Member Name is empty
      (!rowsSimcha[i][COL_SIMCHA_NAME])
      // Col B: Event is empty
      || (!rowsSimcha[i][COL_SIMCHA_EVENT])
      // Col D: Event Date is not today
      || (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getDate() != dateToday.getDate())
      || (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getMonth() != dateToday.getMonth())
      || (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getFullYear() != dateToday.getFullYear())
      // Col N: Reminder Timestamp is not empty
      || (!!rowsSimcha[i][COL_SIMCHA_TIMESTAMP_REMINDER])
    ) continue;

    // Go through list of possible event types
    for (var j = 0, jLen = eventTypesValid.length; j < jLen; j++)
    {
      // Check if event types match
      if (eventTypesValid[j] == rowsSimcha[i][COL_SIMCHA_EVENT])
      {
        // Look for the current row's full name in the members list
        var searchMember = findFirstMatch_(members, 'fullName', rowsSimcha[i][COL_SIMCHA_NAME]);

        // Check if name is in the member list
        if (Object.keys(searchMember).length)
        {
          var
            strEventDate = Utilities.formatDate(rowsSimcha[i][COL_SIMCHA_EVENT_DATE], Session.getScriptTimeZone(), DATEFORMAT_RECEIPT),
            strEventTime = ((isDate_(rowsSimcha[i][COL_SIMCHA_EVENT_TIME]))
              ? Utilities.formatDate(rowsSimcha[i][COL_SIMCHA_EVENT_TIME], Session.getScriptTimeZone(), TIMEFORMAT_SIMCHA)
              : rowsSimcha[i][COL_SIMCHA_EVENT_TIME]
            ),
            strListEmailRecipients = members
              // Remove blanks and duplicates
              .filter(function(member, index, existingMembers) {
                // Premature exit if no email
                if (!member.billing.email) return false;

                // Go through all existing members in the list
                for (var k = 0, kLen = index; k < kLen; k++)
                {
                  // Check if existing before
                  if (existingMembers[k].billing.email == member.billing.email)
                  {
                    // Premature exit
                    return false;
                  } // if existing before
                } // for all existing members in the list k

                return true;
              })
              .map(function(member) {
                return {
                  firstName: member.firstName,
                  email: member.billing.email
                };
              }),
            listNumbers = members.filter(function(member) { return (!!member.billing.phone); })
              .map(function(member) { return member.billing.phone; })
              .unique();

          // Send message(s)
          // Check there are email recipients to send to
          if (strListEmailRecipients.length)
          {
            // Send E-Mail
            var
              titleEmail = searchMember.lastName + ' ' + rowsSimcha[i][COL_SIMCHA_EVENT],
              htmlListItems =
                Utilities.formatString(
                  HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                  'Date: ' + strEventDate
                )
                + ((!!strEventTime)
                  ? Utilities.formatString(
                    HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                    'Time: ' + strEventTime
                  )
                  : ''
                );

            // Go through order of list items to add
            for (var k = 0, kLen = ORDER_LISTITEMS_SIMCHA.length; k < kLen; k++)
            {
              // Check if field not empty
              if (!!(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index]))
              {
                // Add to HTML code of list items
                htmlListItems += Utilities.formatString(
                  HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                  ORDER_LISTITEMS_SIMCHA[k].header + ': '
                    + ((isDate_(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index]))
                      ? Utilities.formatDate(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index], Session.getScriptTimeZone(), TIMEFORMAT_SIMCHA)
                      : rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index]
                    )
                );
              } // if field is not empty
            } // for all list items to add k

            // Go through list of email recipients to send to
            for (var k = 0, kLen = strListEmailRecipients.length; k < kLen; k++)
            {
              // Send email to one recipient
              GmailApp.sendEmail(
                strListEmailRecipients[k].email,
                'Today - ' + titleEmail,
                '',
                {
                  htmlBody: Utilities.formatString(
                    HTML_CODE_EMAIL_BODY_SIMCHA.replaceTemplateTokens(templateTokens),
                    // Title
                    titleEmail,
                    // First Name
                    strListEmailRecipients[k].firstName,
                    // Event
                    rowsSimcha[i][COL_SIMCHA_EVENT],
                    // Event Invite Text
                    rowsSimcha[i][COL_SIMCHA_EVENT_INVITE_TEXT],
                    // List Items
                    htmlListItems,
                    // Name
                    searchMember.fullName
                  )
                }
              );
            } // for all email recipients k
          } // if list of email recipients is not empty

          // Send SMS
          // Go through all numbers to send to
          for (var k = 0, kLen = listNumbers.length; k < kLen; k++)
          {
            sendSMS_(
              listNumbers[k],
              Utilities.formatString(
                SMS_BODY_SIMCHA_REMINDER.replaceTemplateTokens(templateTokens),
                // Reminder Type
                'Today',
                // Name
                searchMember.fullName,
                // Event
                rowsSimcha[i][COL_SIMCHA_EVENT],
                // Date
                strEventDate,
                // Time
                strEventTime,
                // Place
                rowsSimcha[i][COL_SIMCHA_EVENT_PLACE],
                // Address
                rowsSimcha[i][COL_SIMCHA_EVENT_ADDRESS]
              )
            );
          } // for all numbers to send to k

          // Write timestamp to Col M: Invitation Timestamp
          // Get N(current row)
          sheetSimcha.getRange(i + ROW_OFFSET_SIMCHA, COL_SIMCHA_TIMESTAMP_REMINDER + 1).setValue(dateToday);

        } // if name is in the member list

        break;
      } // if event types match
    } // for all possible event types j
  } // for all Simcha list rows i
} // sendSimchaReminder2_Today_()



function sendSimchaReminder2_C()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetSimcha = ss.getSheetByName(SHEETNAME_SIMCHA),
    sheetSimchaLastRow = sheetSimcha.getLastRow(),
    eventTypesValid = ['Shalom Zachor', 'Kiddush'];

  // Premature exit if there are no data rows to process
  if (sheetSimchaLastRow < ROW_OFFSET_SIMCHA) return;

  var
    // Get A(ROW_OFFSET_SIMCHA):N(last row)
    rowsSimcha = sheetSimcha.getRange(ROW_OFFSET_SIMCHA, 1, sheetSimchaLastRow - ROW_OFFSET_SIMCHA + 1, 14).getValues(),
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    // Get A2:O(last row)
    members = sheetMembers.getRange(2, 1, sheetMembers.getLastRow() - 1, 16).getValues()
      // Don't include empty rows (considering first and last name, not full name)
      .filter(function (member) { return ((!!member[1]) && (!!member[2])); })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName: member[0],
          firstName: member[1],
          lastName: member[2],
          address: {
            street: member[3],
            zip: member[6]
          },
          billing: {
            email: ((member[15].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[15].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    // Template tokens
    templateTokens = getTemplateTokens_(ss),
    dateToday = new Date(),
    dateTomorrow = {};

  // Remove time component of date today
  dateToday.setHours(0, 0, 0, 0);
  // Get date tomorrow
  dateTomorrow = dateToday.addDays(1);

  // Go through all Simcha list rows
  for (var i = rowsSimcha.length - 1; i >= 0; i--)
  {
    // Skip if Col D: Event Date is empty or not a date
    if ((!rowsSimcha[i][COL_SIMCHA_EVENT_DATE]) || (!isDate_(rowsSimcha[i][COL_SIMCHA_EVENT_DATE]))) continue;

    // Check if Col D: Event Date is before today
    if (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getTime() < dateToday.getTime())
    {
      // Delete this row
      sheetSimcha.deleteRow(i + ROW_OFFSET_SIMCHA);
      // Go to next row
      continue;
    } // if event date is before today

    // Skip if the required parameters do not exist
    if (
      // Col A: Member Name is empty
      (!rowsSimcha[i][COL_SIMCHA_NAME])
      // Col B: Event is empty
      || (!rowsSimcha[i][COL_SIMCHA_EVENT])
      // Col D: Event Date is not tomorrow
      || (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getDate() != dateTomorrow.getDate())
      || (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getMonth() != dateTomorrow.getMonth())
      || (rowsSimcha[i][COL_SIMCHA_EVENT_DATE].getFullYear() != dateTomorrow.getFullYear())
      // Col M: Invitation Timestamp is empty or not a date
      || (!rowsSimcha[i][COL_SIMCHA_TIMESTAMP_INVITATION]) || (!isDate_(rowsSimcha[i][COL_SIMCHA_TIMESTAMP_INVITATION]))
      // Col M: Invitation Timestamp is not more than NUMBER_DAYS_REMINDER_2C days ago
      || (((dateToday.getTime() - rowsSimcha[i][COL_SIMCHA_TIMESTAMP_INVITATION].getTime()) / DAYS_TO_MILLISECONDS) <= NUMBER_DAYS_REMINDER_2C)
      // Col N: Reminder Timestamp is not empty
      || (!!rowsSimcha[i][COL_SIMCHA_TIMESTAMP_REMINDER])
    ) continue;

    // Go through list of possible event types
    for (var j = 0, jLen = eventTypesValid.length; j < jLen; j++)
    {
      // Check if event types match
      if (eventTypesValid[j] == rowsSimcha[i][COL_SIMCHA_EVENT])
      {
        // Look for the current row's full name in the members list
        var searchMember = findFirstMatch_(members, 'fullName', rowsSimcha[i][COL_SIMCHA_NAME]);

        // Check if name is in the member list
        if (Object.keys(searchMember).length)
        {
          var
            strEventDate = Utilities.formatDate(rowsSimcha[i][COL_SIMCHA_EVENT_DATE], Session.getScriptTimeZone(), DATEFORMAT_RECEIPT),
            strEventTime = ((isDate_(rowsSimcha[i][COL_SIMCHA_EVENT_TIME]))
              ? Utilities.formatDate(rowsSimcha[i][COL_SIMCHA_EVENT_TIME], Session.getScriptTimeZone(), TIMEFORMAT_SIMCHA)
              : rowsSimcha[i][COL_SIMCHA_EVENT_TIME]
            ),
            strListEmailRecipients = members
              // Remove blanks and duplicates
              .filter(function(member, index, existingMembers) {
                // Premature exit if no email
                if (!member.billing.email) return false;

                // Go through all existing members in the list
                for (var k = 0, kLen = index; k < kLen; k++)
                {
                  // Check if existing before
                  if (existingMembers[k].billing.email == member.billing.email)
                  {
                    // Premature exit
                    return false;
                  } // if existing before
                } // for all existing members in the list k

                return true;
              })
              .map(function(member) {
                return {
                  firstName: member.firstName,
                  email: member.billing.email
                };
              }),
            listNumbers = members.filter(function(member) { return (!!member.billing.phone); })
              .map(function(member) { return member.billing.phone; })
              .unique();

          // Send message(s)
          // Check if there are email recipients to send to
          if (strListEmailRecipients.length)
          {
            // Send E-Mail
            var
              titleEmail = searchMember.lastName + ' ' + rowsSimcha[i][COL_SIMCHA_EVENT],
              htmlListItems =
                Utilities.formatString(
                  HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                  'Date: ' + strEventDate
                )
                + ((!!strEventTime)
                  ? Utilities.formatString(
                    HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                    'Time: ' + strEventTime
                  )
                  : ''
                );

            // Go through order of list items to add
            for (var k = 0, kLen = ORDER_LISTITEMS_SIMCHA.length; k < kLen; k++)
            {
              // Check if field not empty
              if (!!(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index]))
              {
                // Add to HTML code of list items
                htmlListItems +=
                  Utilities.formatString(
                    HTML_CODE_ITEMS.replaceTemplateTokens(templateTokens),
                    ORDER_LISTITEMS_SIMCHA[k].header + ': '
                      + ((isDate_(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index]))
                        ? Utilities.formatDate(rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index], Session.getScriptTimeZone(), TIMEFORMAT_SIMCHA)
                        : rowsSimcha[i][ORDER_LISTITEMS_SIMCHA[k].index]
                      )
                );
              } // if field is not empty
            } // for all list items to add k

            // Go through list of email recipients to send to
            for (var k = 0, kLen = strListEmailRecipients.length; k < kLen; k++)
            {
              // Send email to one recipient
              GmailApp.sendEmail(
                strListEmailRecipients[k].email,
                'Tomorrow - ' + titleEmail,
                '',
                {
                  htmlBody: Utilities.formatString(
                    HTML_CODE_EMAIL_BODY_SIMCHA.replaceTemplateTokens(templateTokens),
                    // Title
                    titleEmail,
                    // First Name
                    strListEmailRecipients[k].firstName,
                    // Event
                    rowsSimcha[i][COL_SIMCHA_EVENT],
                    // Event Invite Text
                    rowsSimcha[i][COL_SIMCHA_EVENT_INVITE_TEXT],
                    // List Items
                    htmlListItems,
                    // Name
                    searchMember.fullName
                  )
                }
              );
            } // for all email recipients k
          } // if billing email address exists

          // Send SMS
          // Go through all numbers to send to
          for (var k = 0, kLen = listNumbers.length; k < kLen; k++)
          {
            sendSMS_(
              listNumbers[k],
              Utilities.formatString(
                SMS_BODY_SIMCHA_REMINDER.replaceTemplateTokens(templateTokens),
                // Reminder Type
                'Tomorrow',
                // Name
                searchMember.fullName,
                // Event
                rowsSimcha[i][COL_SIMCHA_EVENT],
                // Date
                strEventDate,
                // Time
                strEventTime,
                // Place
                rowsSimcha[i][COL_SIMCHA_EVENT_PLACE],
                // Address
                rowsSimcha[i][COL_SIMCHA_EVENT_ADDRESS]
              )
            );
          } // for all numbers to send to k

          // Write timestamp to Col M: Invitation Timestamp
          // Get N(current row)
          sheetSimcha.getRange(i + ROW_OFFSET_SIMCHA, COL_SIMCHA_TIMESTAMP_REMINDER + 1).setValue(dateToday);

        } // if name is in the member list

        break;
      } // if event types match
    } // for all possible event types j
  } // for all Simcha list rows i
} // sendSimchaReminder2_C()
