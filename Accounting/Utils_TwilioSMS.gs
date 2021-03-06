var
  TWILIO_ACCOUNT_SID = '',
  TWILIO_AUTH_TOKEN  = '',
  URL_TWILIO_SEND_SMS = 'https://api.twilio.com/2010-04-01/Accounts/' +  TWILIO_ACCOUNT_SID + '/Messages.json',
  TWILIO_FROM = '';


function sendSMS_(numberRecipient, message)
{
  try
  {
    var
      responseTwilioSMS = UrlFetchApp.fetch(
        URL_TWILIO_SEND_SMS,
        {
          headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(TWILIO_ACCOUNT_SID + ':' + TWILIO_AUTH_TOKEN) },
          method: 'post',
          payload: {
            From: TWILIO_FROM,
            To: numberRecipient,
            Body: message
          }
        }
      ).getContentText();
//      resultTwilioSMS = JSON.parse(responseTwilioSMS);
  }
  catch (e)
  {
    Logger.log(JSON.stringify(e))
  }
} // sendSMS()
