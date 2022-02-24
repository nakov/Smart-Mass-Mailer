# Smart-Mass-Mailer

This is a very simple bulk email sender (client), implemented in C# and .NET Core. It sends an email message (by given HTML template file) to a list of recipients (given in Excel spreadsheet, holding columns "email" + "name"), with simple mail-merge functionality.

This software is for developers only. It is not intended for end-users, no GUI is available.

# Goals

The goal of this software is to send reliably bulk emails (e.g. 10,000 emails) without using MailChimp or similar email marketing software. Sending thousands of emails will not work for most email providers (like Office 365 and GMail). It will say "stop, are you a spammer?". Registering own SMTP server or using an SMTP from sites like MailJet + mail client like Outlook / Thunderbird / Evolution + mail merge will do the job, but most emails will be marked as spam. This is because you send too agressivly, e.g. 10,000 emails for 5 minutes. Best results come when you send emails one by one with 5-30 seconds delay after each mail sent. This is what this software does.

# How to Use It?

1. Setup an SMTP server (be sure to configure correctly the SPF records for your domain + reverse DNS + others). Or purchase SMTP from MailGun / MailJet / other.
2. Setup your email template HTML file (see the sample).
3. Prepare your Excel database holding the target emails and person names (see the sample).
4. Configure the app settings: `appsettings.json` (SMTP server settings, email sender, mail subject, delay between emails, etc.)
5. Run the app and wait. It takes time (intentially). I run this in the night.
