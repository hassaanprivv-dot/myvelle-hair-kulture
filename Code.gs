// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//  Myvelle Hair Kulture — Booking Form Backend
//  Google Apps Script  |  Deploy as Web App
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── CONFIG: Fill these in before deploying ──────────
const SHEET_ID     = '1HOymeARy2nkqN6Mdokuc3IvoY071GRmZNWPYJPhQtMA';
const NOTIFY_EMAIL = 'myvelle.hair.kulture@gmail.com';
// ───────────────────────────────────────────────────


function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    writeToSheet(data);
    sendNotification(data);
    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// Allow CORS preflight
function doGet(e) {
  return jsonResponse({ status: 'Myvelle Booking API is live 💜' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── Write row to Google Sheet ───────────────────────
function writeToSheet(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Bookings') || ss.insertSheet('Bookings');

  // Auto-header on first run
  if (sheet.getLastRow() === 0) {
    const headers = ['Timestamp','Name','Instagram','Phone','Services','Date','Time','Notes','Status'];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
               .setBackground('#1e0c22')
               .setFontColor('#ffaedd')
               .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(5, 280);
    sheet.setColumnWidth(8, 300);
  }

  sheet.appendRow([
    new Date(),
    data.name        || '',
    data.instagram   || '',
    data.phone       || '',
    data.services    || '',
    data.date        || '',
    data.time        || '',
    data.notes       || '',
    'New ✨'
  ]);
}


// ── Email notification ──────────────────────────────
function sendNotification(data) {
  MailApp.sendEmail({
    to:       NOTIFY_EMAIL,
    subject:  `💜 New Booking — ${data.name || 'Someone'} wants an appointment`,
    htmlBody: buildEmailHTML(data),
    name:     'Myvelle Hair Kulture'
  });
}

function buildEmailHTML(data) {
  const rows = [
    ['Name',           data.name],
    ['Phone',          data.phone],
    ['Instagram',      data.instagram || '—'],
    ['Service(s)',     data.services],
    ['Preferred Date', data.date],
    ['Preferred Time', data.time],
    ['Notes',          data.notes || '—'],
  ];

  const rowsHTML = rows.map(([label, val]) => `
    <tr>
      <td style="padding:12px 16px;border-bottom:1px solid rgba(255,174,221,0.1);
                 color:rgba(247,217,248,0.5);font-size:12px;font-weight:700;
                 text-transform:uppercase;letter-spacing:0.08em;
                 width:36%;vertical-align:top;">
        ${label}
      </td>
      <td style="padding:12px 16px;border-bottom:1px solid rgba(255,174,221,0.1);
                 color:#f7d9f8;font-size:14px;font-weight:600;vertical-align:top;">
        ${val || '—'}
      </td>
    </tr>
  `).join('');

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
</head>
<body style="margin:0;padding:0;background:#0d0515;font-family:'Helvetica Neue',Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#0d0515;padding:40px 16px;">
  <tr><td align="center">
  <table width="580" cellpadding="0" cellspacing="0"
         style="max-width:580px;width:100%;">

    <!-- ── HEADER ── -->
    <tr>
      <td style="background:linear-gradient(135deg,#4a006a 0%,#7000ff 55%,#ff4dcc 100%);
                 border-radius:20px 20px 0 0;padding:40px 32px;text-align:center;">
        <p style="margin:0 0 10px;font-size:36px;line-height:1;">💜</p>
        <h1 style="margin:0 0 6px;color:#fff;font-size:26px;font-weight:800;
                   letter-spacing:-0.02em;">New Booking Request!</h1>
        <p style="margin:0;color:rgba(255,255,255,0.7);font-size:14px;
                  letter-spacing:0.05em;text-transform:uppercase;">
          Myvelle Hair Kulture
        </p>
      </td>
    </tr>

    <!-- ── BODY ── -->
    <tr>
      <td style="background:#1e0c22;border:1px solid rgba(255,174,221,0.12);
                 border-top:none;border-radius:0 0 20px 20px;padding:0 0 28px;">

        <!-- alert badge -->
        <div style="margin:24px 28px 20px;background:rgba(255,174,221,0.08);
                    border:1px solid rgba(255,174,221,0.2);border-radius:10px;
                    padding:12px 16px;color:#ffaedd;font-size:13px;text-align:center;">
          ✨ A new client is ready to get slay'd — respond within 24 hrs!
        </div>

        <!-- data table -->
        <table width="100%" cellpadding="0" cellspacing="0"
               style="border-collapse:collapse;">
          ${rowsHTML}
        </table>

        <!-- CTA button -->
        <div style="text-align:center;margin:28px 28px 0;">
          <a href="https://www.instagram.com/direct/inbox/"
             style="display:inline-block;
                    background:linear-gradient(135deg,#ff4dcc,#9b00ff);
                    color:#fff;text-decoration:none;
                    padding:14px 36px;border-radius:50px;
                    font-size:13px;font-weight:700;
                    letter-spacing:0.1em;text-transform:uppercase;">
            Reply on Instagram DM
          </a>
        </div>

        <!-- footer -->
        <p style="margin:28px 28px 0;padding-top:20px;
                  border-top:1px solid rgba(255,174,221,0.08);
                  color:rgba(247,217,248,0.25);font-size:11px;text-align:center;">
          Myvelle Hair Kulture &nbsp;·&nbsp; Dallas, TX
          <br/>This is an automated notification from your booking form.
        </p>

      </td>
    </tr>

  </table>
  </td></tr>
</table>

</body>
</html>`;
}
