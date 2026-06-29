// ─────────────────────────────────────────────────────────────────────────────
// MEGAN & SAM — Save the Date Emailer
// ─────────────────────────────────────────────────────────────────────────────
//
// SETUP:
//   1. Go to script.google.com → New project
//   2. Paste this entire file in
//   3. Edit the CONFIG section below
//   4. Create a Google Sheet with these columns in row 1:
//        A: Name   B: Email   C: Sent
//      Then fill in your guest list from row 2 onwards.
//   5. In the script editor: Run → sendSaveDates
//      (First run will ask for Gmail permissions — allow them)
//
// The script skips anyone already marked "YES" in column C,
// so you can re-run safely if it gets interrupted.
// ─────────────────────────────────────────────────────────────────────────────

const CONFIG = {
  SHEET_NAME:    "Guests",          // Name of the tab in your Google Sheet
  FROM_NAME:     "Megan & Sam",     // Display name emails are sent from
  SUBJECT:       "Save the Date 💚 Megan & Sam — 14 March 2027",
  WEBSITE_URL:   "https://sam9dc-del.github.io/wedding/",
  DELAY_MS:      300,               // Milliseconds between sends (avoids Gmail rate limits)
};

// ─────────────────────────────────────────────────────────────────────────────

function sendSaveDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found. Check the tab name.`);

  const data = sheet.getDataRange().getValues();
  let sent = 0, skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const [name, email, sentFlag] = data[i];

    if (!email || String(sentFlag).toUpperCase() === "YES") {
      skipped++;
      continue;
    }

    const firstName = String(name).trim().split(" ")[0];

    try {
      GmailApp.sendEmail(
        email,
        CONFIG.SUBJECT,
        stripHtml(buildEmailHtml(firstName)),   // plain text fallback
        {
          name:     CONFIG.FROM_NAME,
          htmlBody: buildEmailHtml(firstName),
          replyTo:  Session.getActiveUser().getEmail(),
        }
      );

      // Mark as sent so we don't double-send
      sheet.getRange(i + 1, 3).setValue("YES");
      sent++;

      Utilities.sleep(CONFIG.DELAY_MS);

    } catch (err) {
      Logger.log(`Failed for ${email}: ${err.message}`);
      sheet.getRange(i + 1, 3).setValue("ERROR: " + err.message);
    }
  }

  const msg = `Done! Sent: ${sent} · Skipped/already sent: ${skipped}`;
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// Send a single test email to yourself before blasting the full list
function sendTestEmail() {
  const me = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(
    me,
    "[TEST] " + CONFIG.SUBJECT,
    stripHtml(buildEmailHtml("Friend")),
    {
      name:     CONFIG.FROM_NAME,
      htmlBody: buildEmailHtml("Friend"),
    }
  );
  SpreadsheetApp.getUi().alert(`Test email sent to ${me} — check your inbox!`);
}

// ─────────────────────────────────────────────────────────────────────────────
// EMAIL TEMPLATE
// ─────────────────────────────────────────────────────────────────────────────

function buildEmailHtml(firstName) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Save the Date — Megan & Sam</title>
</head>
<body style="margin:0;padding:0;background:#f4f1ec;font-family:Georgia,'Times New Roman',serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f1ec;padding:40px 20px;">
    <tr><td align="center">

      <!-- Card -->
      <table width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#6c725b;border-radius:8px;overflow:hidden;">

        <!-- Top band -->
        <tr>
          <td style="background:#575e49;padding:12px 40px;text-align:center;">
            <p style="margin:0;font-family:Arial,sans-serif;font-size:11px;letter-spacing:0.25em;text-transform:uppercase;color:rgba(255,255,255,0.65);">Save the Date</p>
          </td>
        </tr>

        <!-- Main content -->
        <tr>
          <td style="padding:52px 48px 44px;text-align:center;">

            <!-- Names -->
            <h1 style="margin:0 0 6px;font-family:Georgia,'Times New Roman',serif;font-size:52px;font-weight:400;color:#ffffff;line-height:1;letter-spacing:0.01em;">
              Megan &amp; Sam
            </h1>

            <!-- Divider -->
            <table cellpadding="0" cellspacing="0" style="margin:24px auto;">
              <tr>
                <td style="width:40px;height:1px;background:rgba(255,255,255,0.3);"></td>
                <td style="padding:0 12px;color:rgba(255,255,255,0.45);font-size:16px;">◆</td>
                <td style="width:40px;height:1px;background:rgba(255,255,255,0.3);"></td>
              </tr>
            </table>

            <!-- Date -->
            <p style="margin:0 0 6px;font-family:Georgia,'Times New Roman',serif;font-size:28px;font-weight:400;color:#ffffff;letter-spacing:0.04em;">
              14 March 2027
            </p>

            <!-- Venue -->
            <p style="margin:0 0 32px;font-family:Arial,sans-serif;font-size:13px;letter-spacing:0.18em;text-transform:uppercase;color:rgba(255,255,255,0.72);">
              Ohakuri Rise &nbsp;·&nbsp; Taupo, New Zealand
            </p>

            <!-- Personalised note -->
            <p style="margin:0 0 36px;font-family:Arial,sans-serif;font-size:15px;line-height:1.7;color:rgba(255,255,255,0.85);">
              Dear ${firstName},<br/>
              We're getting married and we'd love you to be there.<br/>
              Formal invitations to follow — for now, please mark your calendar!
            </p>

            <!-- CTA button -->
            <a href="${CONFIG.WEBSITE_URL}"
               style="display:inline-block;padding:14px 44px;background:#ffffff;color:#575e49;font-family:Arial,sans-serif;font-size:12px;font-weight:700;letter-spacing:0.2em;text-transform:uppercase;text-decoration:none;border-radius:40px;">
              View Our Website
            </a>

          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="background:#575e49;padding:20px 40px;text-align:center;">
            <p style="margin:0;font-family:Arial,sans-serif;font-size:11px;color:rgba(255,255,255,0.5);letter-spacing:0.1em;">
              Megan &amp; Sam &nbsp;·&nbsp; 14 · 03 · 2027 &nbsp;·&nbsp; Taupo, New Zealand
            </p>
          </td>
        </tr>

      </table>
      <!-- /Card -->

    </td></tr>
  </table>

</body>
</html>`;
}

// Plain-text fallback (strips HTML tags)
function stripHtml(html) {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}
