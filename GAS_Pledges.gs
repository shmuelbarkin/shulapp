var  
  SHEETNAME_PLEDGES = 'Pledges',
  SHEETNAME_TRANSACTIONS = 'Transactions',
  ROW_OFFSET_PLEDGES = 3,
  
  NUMBER_PLEDGES_ROW = 4,
  COL_START_PLEDGE = 2,
    
  COL_OFFSET_PLEDGE = 4,
  COL_OFFSET_DATE =   0,
  COL_OFFSET_AMOUNT = 1,
  COL_OFFSET_ITEM =   2,
  COL_OFFSET_PAID =   3,
  
  COL_LAST_TOUCH = 17,
  COL_BILL_SENT =  18,
  
  NUMBER_DAYS_AGO_BILL = 29,
    
  SUBJECT_EMAIL_RECEIPT = 'Official Receipt - $%s',
  SUBJECT_EMAIL_BILL = 'Pledge Reminder - $%s',
  
  SMS_BODY_RECEIPT = '%s,\nThank you for your donation of $%s.\nWhiteStreetShul.org',
  SMS_BODY_BILL = '%s,\nPledge Reminder: You have $%s in outstanding pledges.\nWhiteStreetShul.org';



function onEdit(e)
{
  // If processable edit
  if (
    // must originate from Pledges sheet
    (e.source.getSheetName() == SHEETNAME_PLEDGES)    
    // must be a column change in the ff: B, C, D, F, G, H, J, K, L, N, O, P
    && (e.range.columnStart >= 2) && (e.range.columnEnd <= 16)
      // Exclude E
      && ((e.range.columnStart >  5) || (e.range.columnEnd <  5))
      // Exclude I
      && ((e.range.columnStart >  9) || (e.range.columnEnd <  9))
      // Exclude M
      && ((e.range.columnStart > 13) || (e.range.columnEnd < 13))
    // must be one row tall only, and starting from row 3
    && (e.range.rowStart == e.range.rowEnd) && (e.range.rowStart >= 3)
  )
  {
    // Get R(current row)
    e.source.getActiveSheet().getRange(e.range.rowStart, 18)
      .setValue(new Date());
  } // if processable edit
} // onEdit()



function processReceipts()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    // Get A2:N(last row)
    members = sheetMembers.getRange(2, 1, sheetMembers.getLastRow() - 1, 14).getValues()
      // Don't include empty rows (considering first and last name, not full name)
      .filter(function (member) { return ((!!member[1]) && (!!member[2])); })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName: member[0],
          firstName: member[1],
          address: {
            street: member[3],
            zip: member[6]
          },
          billing: {
            email: ((member[13].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[13].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    sheetPledges = ss.getSheetByName(SHEETNAME_PLEDGES),
    // Get A3:Q(last row)
    rowsPledges = sheetPledges.getRange(3, 1, sheetPledges.getLastRow() - ROW_OFFSET_PLEDGES + 1, 17).getValues(),
    sheetTransactions = ss.getSheetByName(SHEETNAME_TRANSACTIONS),
    dateToday = new Date(),
    strDateToday = Utilities.formatDate(dateToday, Session.getScriptTimeZone(), DATEFORMAT_RECEIPT);
    
  // Go through all the rows in the Pledges tab
  for (var i = rowsPledges.length - 1; i >= 0; i--)
  {
    // Look for the current row's full name in the members list
    var searchMember = findFirstMatch_(members, 'fullName', rowsPledges[i][0]);
    
    // Check if a member was found
    if (Object.keys(searchMember).length)
    {
      var
        // States of all NUMBER_PLEDGES_ROW in the current row
        pledgeStates = {
          // Filled state. true for filled, and false otherwise
          filled: Array.apply(null, Array(NUMBER_PLEDGES_ROW)).map(function () { return false; }),
          // Paid state. true for paid, and false otherwise
          paid: Array.apply(null, Array(NUMBER_PLEDGES_ROW)).map(function () { return false; }),
          areAllPaid: true,
        },
        htmlListItems = '',
        numberPaidItems = 0,
        currentPaidAmount = 0,
        totalPaidAmount = 0;
            
      // Check if the paid columns E, I, M, Q has any entries ($##.##)
      // Get pledge states
      // Go through all pledges in the row
      for (var j = 0; j < NUMBER_PLEDGES_ROW; j++)
      {
        // Starting index of the current pledge
        var startIndex = COL_START_PLEDGE + (j * COL_OFFSET_PLEDGE) - 1;
              
        pledgeStates.filled[j] = (
          (!!rowsPledges[i][startIndex + COL_OFFSET_DATE])
          && (!!rowsPledges[i][startIndex + COL_OFFSET_AMOUNT])
          && (!!rowsPledges[i][startIndex + COL_OFFSET_ITEM])
        );
        
        pledgeStates.paid[j] = (!!rowsPledges[i][startIndex + COL_OFFSET_PAID]);
        
        // Check if this row is filled
        if (pledgeStates.filled[j])
        {
          // Check if paid 
          if (pledgeStates.paid[j])
          {
            // Get current paid amount
            currentPaidAmount = ((isNumeric_(rowsPledges[i][startIndex + COL_OFFSET_PAID]))
              ? Number(rowsPledges[i][startIndex + COL_OFFSET_PAID])
              : 0
            );
            
            // Add to list of paid items
            // HTML
            htmlListItems += Utilities.formatString(HTML_CODE_SUBITEMS,
              Utilities.formatString('%s, %s, $%s',
                // Item Description
                rowsPledges[i][startIndex + COL_OFFSET_ITEM],                
                // Item Date
                ((isDate_(rowsPledges[i][startIndex + COL_OFFSET_DATE]))
                  ? Utilities.formatDate(rowsPledges[i][startIndex + COL_OFFSET_DATE], Session.getScriptTimeZone(), DATEFORMAT_RECEIPT)
                  : rowsPledges[i][startIndex + COL_OFFSET_DATE]
                ),
                // Item paid amount
                currentPaidAmount.toString()
              )
            );
            
            // Add to list of transactions
            sheetTransactions.appendRow([
              // Col A: Name
              rowsPledges[i][0],
              // Col B: Date
              strDateToday,
              // Col C: Item
              rowsPledges[i][startIndex + COL_OFFSET_ITEM],
              // Col D: Amount Paid
              rowsPledges[i][startIndex + COL_OFFSET_PAID]
            ]);
            
            // Add to total amount
            totalPaidAmount += currentPaidAmount;
            
            // Count up
            numberPaidItems += 1;
          }
          else if (pledgeStates.areAllPaid)
          {
            // Update all-paid status to false
            pledgeStates.areAllPaid = false;
          } // if paid
        } // if filled
      } // for all pledges in the row j
     
      // Generate receipt email/SMS for total amount and provide the list of items paid
      // Check if at least one item was paid
      if (numberPaidItems > 0)
      {
        // Check if billing email address exists
        if (!!searchMember.billing.email)
        {
          // Send E-mail receipt      
          GmailApp.sendEmail(
            searchMember.billing.email,
            Utilities.formatString(SUBJECT_EMAIL_RECEIPT, totalPaidAmount.toString()),
            '',
            {
              htmlBody: Utilities.formatString(HTML_CODE_EMAIL_BODY_RECEIPT,
                // First Name
                searchMember.firstName,
                // Date Today
                strDateToday,
                // Total Amount
                totalPaidAmount.toString(),
                // Items
                ((numberPaidItems != 1) ? 's' : '') + ':' + htmlListItems
              )
            }
          );
        } // if billing email address exists
        
        // Send SMS              
        // Check if SMS # exists
        if (!!searchMember.billing.phone)
        {
          sendSMS_(
            searchMember.billing.phone,
            Utilities.formatString(SMS_BODY_RECEIPT,
              // First Name
              searchMember.firstName,
              // Total Amount
              totalPaidAmount.toString()
            )
          );
        } // if SMS # exists
        
      } // if numberPaidItems > 0

      // Delete paid items [Set of cells, or row if all items are paid] 
      // Check if all items were paid
      if (pledgeStates.areAllPaid)
      {
        // Delete the whole row
        sheetPledges.deleteRow(i + ROW_OFFSET_PLEDGES);
      }
      else
      {
        // Go through all pledges in the row
        for (var j = 0; j < NUMBER_PLEDGES_ROW; j++)
        {
          // Check if pledge was paid
          if ((pledgeStates.filled[j]) && (pledgeStates.paid[j]))
          {
            // Remove pledge cells
            // Get (start index)(current row):(start index + 3)(current row)
            sheetPledges.getRange((i + ROW_OFFSET_PLEDGES), COL_START_PLEDGE + (j * COL_OFFSET_PLEDGE), 1, COL_OFFSET_PLEDGE)
              .setValue('');
          } // if pledge was paid
        } // for all pledges in the row j      
      } // if all are paid
    }
    else if (!!rowsPledges[i][0])
    {
      ss.toast('Skipping Row ' + (i + ROW_OFFSET_PLEDGES).toString() + '. Member Name "' + rowsPledges[i][0] + '" was not found in the list.');
    } // if member was found    
  } // for all rows in Pledges i
} // processReceipts()



function processBills()
{
  var
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetMembers = ss.getSheetByName(SHEETNAME_MEMBERS),
    // Get A2:N(last row)
    members = sheetMembers.getRange(2, 1, sheetMembers.getLastRow() - 1, 14).getValues()
      // Don't include empty rows (considering first and last name, not full name)
      .filter(function (member) { return ((!!member[1]) && (!!member[2])); })
      // Create an descriptive object
      .map(function(member) {
        return {
          fullName: member[0],
          firstName: member[1],
          address: {
            street: member[3],
            zip: member[6]
          },
          billing: {
            email: ((member[13].toString().contains('Email')) ? member[10] : ''),
            phone: ((member[13].toString().contains('SMS'))   ? member[ 9] : '')
          }
        };
      }),
    sheetPledges = ss.getSheetByName(SHEETNAME_PLEDGES),
    // Get A3:S(last row)
    rowsPledges = sheetPledges.getRange(3, 1, sheetPledges.getLastRow() - ROW_OFFSET_PLEDGES + 1, 19).getValues(),
    dateToday = new Date();
    
  // Go through all the rows in the Pledges tab
  for (var i = rowsPledges.length - 1; i >= 0; i--)
  {
    // Check bill sent timestamp
    if (
      // S is blank
      (!rowsPledges[i][COL_BILL_SENT])
      || (
        // S is of type date
        (isDate_(rowsPledges[i][COL_BILL_SENT]))          
        && (
          // S is more than NUMBER_DAYS_AGO_BILL days ago
          (((dateToday.getTime() - rowsPledges[i][COL_BILL_SENT].getTime()) / DAYS_TO_MILLISECONDS) > NUMBER_DAYS_AGO_BILL)          
          // R is greater than S
          || (
            // R is of type date
            (isDate_(rowsPledges[i][COL_LAST_TOUCH]))
            // S is less than or equal to R
            && (rowsPledges[i][COL_BILL_SENT].getTime() <= rowsPledges[i][COL_LAST_TOUCH].getTime())
          )
        )
      )
    )
    {
      // Look for the current row's full name in the members list
      var searchMember = findFirstMatch_(members, 'fullName', rowsPledges[i][0]);
      
      // Check if a member was found
      if (Object.keys(searchMember).length)
      {
        var
          urlParamDescription = '',
          htmlListItems = '',
          numberUnpaidItems = 0,
          currentPledgeAmount = 0,
          totalPledgeAmount = 0;
          
        // Check if the unpaid columns E, I, M, Q have any entries ($##.##)
        // Get pledge states
        // Go through all pledges in the row
        for (var j = 0; j < NUMBER_PLEDGES_ROW; j++)
        {
          // Starting index of the current pledge
          var startIndex = COL_START_PLEDGE + (j * COL_OFFSET_PLEDGE) - 1;
                
          // Check if this row is filled but unpaid
          if (
            // Filled
            (!!rowsPledges[i][startIndex + COL_OFFSET_DATE])
            && (!!rowsPledges[i][startIndex + COL_OFFSET_AMOUNT])
            && (!!rowsPledges[i][startIndex + COL_OFFSET_ITEM])            
            // Unpaid
            && (!rowsPledges[i][startIndex + COL_OFFSET_PAID])
          )
          {
            // Get current pledge amount
            currentPledgeAmount = ((isNumeric_(rowsPledges[i][startIndex + COL_OFFSET_AMOUNT]))
              ? Number(rowsPledges[i][startIndex + COL_OFFSET_AMOUNT])
              : 0
            );
            
            // Add to list of unpaid items
            // URL parameter
            urlParamDescription += Utilities.formatString('%s $%s,',
              rowsPledges[i][startIndex + COL_OFFSET_ITEM],
              currentPledgeAmount.toString()              
            );
            // HTML
            htmlListItems += Utilities.formatString(HTML_CODE_ITEMS,
              Utilities.formatString('%s, %s, $%s',
                // Item Description
                rowsPledges[i][startIndex + COL_OFFSET_ITEM],
                // Item Date
                ((isDate_(rowsPledges[i][startIndex + COL_OFFSET_DATE]))
                  ? Utilities.formatDate(rowsPledges[i][startIndex + COL_OFFSET_DATE], Session.getScriptTimeZone(), DATEFORMAT_RECEIPT)
                  : rowsPledges[i][startIndex + COL_OFFSET_DATE]
                ),
                // Item pledge amount
                currentPledgeAmount.toString()                
              )
            );
            
            // Add to total amount
            totalPledgeAmount += currentPledgeAmount;
            
            // Count up
            numberUnpaidItems += 1;
          } // if filled
        } // for all pledges in the row j
        
        // Remove extraneous comma in urlParamDescription if it exists
        urlParamDescription = ((urlParamDescription.endsWith(',')) ? urlParamDescription.slice(0, -1) : urlParamDescription);        
        
        // Generate receipt email/SMS for total amount and provide the list of items unpaid
        // Check if at least one item was unpaid
        if (numberUnpaidItems > 0)
        {
          // Check if billing email address exists
          if (!!searchMember.billing.email)
          {
            // Send E-mail bill
            GmailApp.sendEmail(
              searchMember.billing.email,
              Utilities.formatString(SUBJECT_EMAIL_BILL, totalPledgeAmount.toString()),
              '',
              {
                htmlBody: Utilities.formatString(HTML_CODE_EMAIL_BODY_BILL,
                  // First Name
                  searchMember.firstName,
                  // Total Amount
                  totalPledgeAmount.toString(),
                  // Items
                  htmlListItems,
                  // Payment Site URL
                  Utilities.formatString(URL_PAYMENT_SITE,
                    encodeURIComponent(searchMember.fullName),
                    encodeURIComponent(searchMember.address.street),
                    encodeURIComponent(searchMember.address.zip),
                    encodeURIComponent(urlParamDescription),
                    encodeURIComponent(totalPledgeAmount)
                  )                 
                )
              }
            );
          } // if billing email address exists
          
          // Send SMS              
          // Check if SMS # exists
          if (!!searchMember.billing.phone)
          {
            sendSMS_(
              searchMember.billing.phone,
              Utilities.formatString(SMS_BODY_BILL,
                // First Name
                searchMember.firstName,
                // Total Amount
                totalPledgeAmount.toString()
              )
            );
          } // if SMS # exists          
        } // if numberUnpaidItems > 0        
      
        // Enter date stamp in column S        
        // Get S(current row)
        sheetPledges.getRange((i + ROW_OFFSET_PLEDGES), COL_BILL_SENT + 1).setValue(dateToday);                
      }
      else if (!!rowsPledges[i][0])
      {
        ss.toast('Skipping Row ' + (i + ROW_OFFSET_PLEDGES).toString() + '. Member Name "' + rowsPledges[i][0] + '" was not found in the list.');
      } // if a member was found
    } // if bill sent timestamp 
  } // for all rows in Pledges i
} // processBills()