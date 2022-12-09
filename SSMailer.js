function myFunction() 
{
  // Constants for main sequence (mSeq)
  var mSeqVertStartLine       = 2;    // Index of line where parsing starts
  var mSeqVertLength          = 500;  // Count of lines that will be parsed
  var mSeqOrgNameLetter       = 'A';  // Letter of col that contains organization names
  var mSeqDomainLetter        = 'B';  // Letter of col that contains name of something (domain, program, hosting and etc) that should renew
  var mSeqDaysToRemindLetter  = 'C';  // Letter of col that contains the date when you'd wanna be alerted
  var mSeqExpireDateLetter    = 'D';  // Letter of col that contains the date when the program or domain will expire

  // Constants for organization and user sequence (ouSeq)
  var ouSeqVertStartLine      = 2;    // Index of line where parsing starts
  var ouSeqVertLength         = 15;   // Count of lines that will be parsed
  var ouSeqOrgsLetter         = 'G';  // Letter of col that contains organization names to bind them to emails
  var ouSeqEmailsLetter       = 'H';  // Letter of col that contains emails that should be binded to some organizations

  // Variables
  var daysToRemind = 30;
  var organizations = {};             // List that contains the binding between organizations and emails
  var dateArr = new Array(mSeqVertStartLine+mSeqVertLength);    // Array for temporary 
  var today;

  for (var l = mSeqVertStartLine; l < mSeqVertStartLine+mSeqVertLength; l++)
  {
    if (SpreadsheetApp.getActiveSheet().getRange(mSeqExpireDateLetter+l).getValue() != '')
    {
      SpreadsheetApp.getActiveSheet().getRange(mSeqDaysToRemindLetter+l).setValue("=" + mSeqExpireDateLetter + l + "-" + daysToRemind);
    }
  }

  //parsing OU sequence (ouSeq)
  for (var vpos = ouSeqVertStartLine; vpos < ouSeqVertStartLine+ouSeqVertLength; vpos++)
  {
    var org = SpreadsheetApp.getActiveSheet().getRange(ouSeqOrgsLetter+vpos).getValue();
    if (org != null && org != "")
    {
      var emailz = SpreadsheetApp.getActiveSheet().getRange(ouSeqEmailsLetter+vpos).getValue();
      if (emailz.toString().includes(' '))
      {
        var ems = emailz.toString().split(' ');
        organizations[org] = ems;
        continue;
      }
      else 
      {
        var ems0 = emailz.toString();
        organizations[org] = ems0;
      }
    }
  }

  // Parsing main sequence (mSeq)
  today = Utilities.formatDate(new Date(), "GMT+5", "dd/MM/yyyy");
  for (var vpos0 = mSeqVertStartLine; vpos0 < mSeqVertStartLine+mSeqVertLength; vpos0++)
  {
    var date = SpreadsheetApp.getActiveSheet().getRange(mSeqDaysToRemindLetter+vpos0).getValue();
    if (date != null && date != "")
    {
      dateArr[vpos0] = Utilities.formatDate(SpreadsheetApp.getActiveSheet().getRange(mSeqDaysToRemindLetter+vpos0).getValue(), "GMT+5", "dd/MM/yyyy");
    }
  }
  for (var ind = 0; ind < dateArr.length; ind++)
  {
    if (dateArr[ind] == today)
    {
      // Painting the line to yellow
      SpreadsheetApp.getActiveSheet().getRange(mSeqOrgNameLetter+ind).setBackground("yellow");
      SpreadsheetApp.getActiveSheet().getRange(mSeqDomainLetter+ind).setBackground("yellow");
      SpreadsheetApp.getActiveSheet().getRange(mSeqDaysToRemindLetter+ind).setBackground("yellow");
      SpreadsheetApp.getActiveSheet().getRange(mSeqExpireDateLetter+ind).setBackground("yellow");
      var lorg = SpreadsheetApp.getActiveSheet().getRange(mSeqOrgNameLetter+ind).getValue().toString();
      if (lorg != null && lorg != "")
      {
        var email = organizations[lorg].toString();
        if (email.includes(','))
        {
          var email0 = email.split(',');
          email0.forEach(
            function(elem)
            {
              MailApp.sendEmail(
                {
                  to: elem,
                  subject: "[SS] Организация: " + lorg + ". Необходимо продление!",
                  htmlBody: SpreadsheetApp.getActiveSheet().getRange(mSeqDomainLetter+ind).getValue().toString() + " скоро закончится.<br>Вы просили уведомить вас: " + dateArr[ind], 
                }
              );
            }
          );
        }
        else
        {
          MailApp.sendEmail(
            {
              to: email,
              subject: "[SS] Организация: " + lorg + ". Необходимо продление!",
              htmlBody: SpreadsheetApp.getActiveSheet().getRange(mSeqDomainLetter+ind).getValue().toString() + " скоро закончится.<br>Вы просили уведомить вас: " + dateArr[ind], 
            }
          );
        }
      }
    }
  }
}
