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
  WEBSITE_URL:   "https://sam9dc-del.github.io/wedding/save-the-date.html",
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
<title>We're Getting Married!</title>
<link href="https://fonts.googleapis.com/css2?family=Jost:wght@300;400&display=swap" rel="stylesheet"/>
<style>@font-face{font-family:'Mozart Script';src:url('https://sam9dc-del.github.io/wedding/MozartScript-Regular.ttf') format('truetype');}</style>
</head>
<body style="margin:0;padding:0;background:#f4f1ec;font-family:Arial,sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f1ec;padding:40px 20px;">
    <tr><td align="center">

      <table width="100%" cellpadding="0" cellspacing="0" style="max-width:520px;background:#ffffff;border-radius:4px;overflow:hidden;">

        <!-- Sage top rule -->
        <tr>
          <td style="background:#6c725b;height:5px;font-size:0;line-height:0;">&nbsp;</td>
        </tr>

        <!-- Photo -->
        <tr>
          <td style="padding:0;line-height:0;">
            <img src="https://sam9dc-del.github.io/wedding/couple2.jpeg" alt="Megan and Sam" width="520"
                 style="width:100%;height:340px;object-fit:cover;object-position:center 55%;display:block;" />
          </td>
        </tr>

        <!-- Main -->
        <tr>
          <td style="padding:72px 48px 48px;text-align:center;">

            <!-- Script headline -->
            <p style="margin:0 0 24px;font-family:Georgia,'Times New Roman',serif;font-style:italic;font-size:44px;line-height:1;color:#1a1a1a;">
              Megan &amp; Sam
            </p>

            <!-- Names in spaced caps -->
            <p style="margin:0 0 16px;font-family:'Jost',Arial,sans-serif;font-weight:300;font-size:10px;letter-spacing:0.32em;text-transform:uppercase;color:#6c725b;">
              Megan Spring &amp; Sam Cook
            </p>

            <!-- Personal note -->
            <p style="margin:0 0 40px;font-family:Georgia,'Times New Roman',serif;font-style:italic;font-size:16px;line-height:1.75;color:#555;">
              Hello ${firstName},<br/><br/>
              We've set a date for our wedding and<br/>
              we can't wait to share it with you!
            </p>

            <!-- CTA -->
            <a href="${CONFIG.WEBSITE_URL}"
               style="display:inline-block;padding:14px 44px;background:#6c725b;color:#ffffff;font-family:'Jost',Arial,sans-serif;font-size:10px;font-weight:400;letter-spacing:0.28em;text-transform:uppercase;text-decoration:none;border-radius:40px;">
              View Save The Date
            </a>

          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="padding:20px 40px;text-align:center;border-top:1px solid #ede9e3;">
            <p style="margin:0;font-family:'Jost',Arial,sans-serif;font-size:9px;font-weight:300;letter-spacing:0.28em;text-transform:uppercase;color:#aaa;">
              Megan &amp; Sam &nbsp;·&nbsp; 14 March 2027 &nbsp;·&nbsp; Taupo
            </p>
          </td>
        </tr>

      </table>

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
