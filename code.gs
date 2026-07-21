const CONFIG = {
  GROUP_EMAIL: 'group-email@example.com', // הכתובת של הגרופס
  GROUP_NAME: 'chat name', // השם של הקבוצה (יופיע במייל האישור)
  GROUP_URL: 'chatURL', // קישור לצאט (הכפתור של כניסה לקבוצה במייל האישור)
  SENDER_NAME: 'admin', // השם שיופיע במייל האישור (לא ישפיע על הכתובת מייל, רק על השם תצוגה)
  BLACKLIST_SHEET_URL: '', // קישור לקובץ גוגל שיטס של רשימה שחורה - אופציונלי, ניתן להשאיר ריק
  SEND_SUCCESS_EMAIL: true, // האם לשלוח מייל אישור - true = כן, false = לא
  COLORS: {
    primary: '#4285F4',
    secondary: '#34A853',
    background: '#F8F9FA',
    text: '#202124'
  }
};

function onFormSubmit(e) {
  try {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    let userEmail = formResponse.getRespondentEmail();

    if (!userEmail) {
      for (let i = 0; i < itemResponses.length; i++) {
        const itemTitle = itemResponses[i].getItem().getTitle();
        if (itemTitle.includes('מייל') || itemTitle.toLowerCase().includes('email')) {
          userEmail = itemResponses[i].getResponse();
          break;
        }
      }
    }

    if (!userEmail) {
      Logger.log('לא נמצאה כתובת מייל בטופס');
      return;
    }

    userEmail = userEmail.trim();

    if (isBlacklisted(userEmail)) {
      return;
    }
    const addResult = addToGroup(userEmail);

    if (addResult.success) {
      sendConfirmationEmail(userEmail);
      Logger.log('✓ הרישום הושלם בהצלחה עבור: ' + userEmail);
    } else {
      Logger.log('✗ שגיאה בצירוף: ' + addResult.error);
    }
  } catch (error) {
    Logger.log('שגיאה כללית: ' + error.toString());
  }
}

function addToGroup(email) {
  try {
    const member = {
      email: email,
      role: 'MEMBER',
      delivery_settings: 'ALL_MAIL'
    };
    AdminDirectory.Members.insert(member, CONFIG.GROUP_EMAIL);
    return { success: true };
  } catch (error) {
    if (error.toString().includes('duplicate') || error.toString().includes('Member already exists')) {
      return { success: true, alreadyMember: true };
    }
    return { success: false, error: error.toString() };
  }
}

function sendConfirmationEmail(email) {
  if (!CONFIG.SEND_SUCCESS_EMAIL) return;

  const subject = CONFIG.EMAIL_SUBJECT || 'צורפת בהצלחה!';
  const htmlBody = (typeof CONFIG.htmlTemplate === 'function')
    ? CONFIG.htmlTemplate(CONFIG)
    : defaultHtmlTemplate(CONFIG);

  GmailApp.sendEmail(email, subject, 'אנא הצג מייל זה בתצוגת HTML', {
    htmlBody: htmlBody,
    name: CONFIG.SENDER_NAME || '.'
  });
}

function defaultHtmlTemplate(config) {
  const colors = config.COLORS || {
    primary: '#4285F4',
    secondary: '#34A853',
    background: '#F8F9FA',
    text: '#202124'
  };

  return `
  <!DOCTYPE html>
  <html dir="rtl" lang="he">
  <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #E6F7FF; direction: rtl;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #E6F7FF; padding: 30px 0;">
      <tr>
        <td align="center">
          <table width="650" cellpadding="0" cellspacing="0" style="background-color: white; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border: 1px solid #D1D5DB;">
            <tr>
              <td style="padding: 30px; direction: rtl;">
                <h2 style="margin: 0 0 20px 0; font-size: 18px; color: ${colors.primary}; font-weight: bold;">
                   צורפת בהצלחה ל${config.GROUP_NAME}!
                </h2>
                <a href="${config.GROUP_URL}" style="display: inline-block; background-color: ${colors.primary}; color: white; text-decoration: none; padding: 12px 30px; border-radius: 25px; font-size: 14px;"> כניסה לקבוצה </a>
                <p style="margin: 20px 0 0 0; font-size: 14px; color: #374151;"><strong>במידה והקישור לא עובד יש לרענן את הדף ולנסות שוב</strong><br> בברכה,<br> הנהלת הקבוצה </p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
  </html>`;
}


function isBlacklisted(email){if(!CONFIG.BLACKLIST_SHEET_URL){return false;}const values=SpreadsheetApp.openByUrl(CONFIG.BLACKLIST_SHEET_URL).getSheets()[0].getDataRange().getValues();email=email.toLowerCase();for(const row of values){for(const cell of row){if(String(cell).trim().toLowerCase()===email){return true;}}}return false;}
