using System;
using System.IO;

namespace TeamsBot.Server.Services
{
    public static class EmailTemplateRenderer
    {
        private const string TemplateFilePath = "/Users/ankithravindran/Development/Antigravity/prospectplus-application/mockups/pp_email_base_template.html";

        public static string RenderOtpEmail(string recipientEmail, string code)
        {
            string name = ExtractFirstName(recipientEmail);
            string emailBodyHtml = $@"
                <p style=""margin: 0 0 16px; font-size: 15px; color: #2d3748; font-family: 'Inter', system-ui, -apple-system, sans-serif; line-height: 1.6;"">
                    You have requested access to <strong>Minutes.Plus Security Portal</strong> (Teams Recording & AI Transcription Bot).
                </p>
                <p style=""margin: 0 0 12px; font-size: 14px; color: #4a5568; font-family: 'Inter', system-ui, -apple-system, sans-serif; font-weight: 600;"">
                    Your single-use 6-digit security verification code is:
                </p>
                <div style=""text-align: center; margin: 24px 0; padding: 20px; background-color: #f0fdf4; border: 1.5px solid #bbf7d0; border-radius: 12px; box-shadow: inset 0 2px 4px rgba(0,0,0,0.02);"">
                    <span style=""font-size: 32px; font-family: 'Courier New', Courier, monospace; font-weight: 700; color: #059669; letter-spacing: 8px; display: inline-block;"">{code}</span>
                </div>
                <p style=""margin: 0 0 16px; font-size: 13px; color: #718096; font-family: 'Inter', system-ui, -apple-system, sans-serif; line-height: 1.5;"">
                    This passcode is valid for <strong>10 minutes</strong>. Do not share this code with anyone. If you did not request this login code, please notify corporate IT security immediately.
                </p>
";

            try
            {
                if (File.Exists(TemplateFilePath))
                {
                    string templateHtml = File.ReadAllText(TemplateFilePath);
                    string html = templateHtml.Replace("Hi {{contact.firstName}},", $"Hi {name},");
                    html = html.Replace("<p style=\"margin: 0 0 16px; font-size: 15px; color: #2d3748; font-family: 'Inter', system-ui, -apple-system, sans-serif; line-height: 1.6;\">\n                Insert your email body text here.\n              </p>", emailBodyHtml);
                    html = html.Replace("Insert your email body text here.", emailBodyHtml);
                    return html;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[EmailTemplateRenderer] Warning: Failed to load external template ({ex.Message}). Using fallback embedded template.");
            }

            // Fallback inline template matching pp_email_base_template.html design
            return $@"<!DOCTYPE html>
<html lang=""en"">
<head>
  <meta charset=""utf-8"" />
  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />
  <title>Minutes.Plus Security Verification Code</title>
</head>
<body style=""margin: 0; padding: 0; width: 100% !important; background-color: #f4f7f8; font-family: 'Inter', system-ui, -apple-system, sans-serif;"">
  <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""background-color: #f4f7f8; padding: 20px 0; width: 100%;"">
    <tr>
      <td align=""center"">
        <table class=""email-container"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""600"" style=""max-width: 600px; width: 100%; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 24px rgba(9, 92, 123, 0.06);"">
          <tr>
            <td class=""content-cell"" style=""padding: 45px 35px; color: #2d3748; font-size: 15px; line-height: 1.6; font-family: 'Inter', system-ui, -apple-system, sans-serif;"">
              <div class=""greeting"" style=""font-size: 20px; color: #095c7b; font-weight: 700; margin-bottom: 16px; letter-spacing: -0.5px;"">
                Hi {name},
              </div>
              {emailBodyHtml}
              <p style=""margin: 16px 0 6px; font-size: 15px; color: #2d3748; line-height: 1.6;"">
                Kind regards,
              </p>
              <p style=""margin: 0; font-size: 15px; color: #2d3748; line-height: 1.6;"">
                <strong style=""font-weight: 700; color: #2d3748;"">MailPlus Corporate Security Team</strong>
              </p>
            </td>
          </tr>
          <tr>
            <td align=""center"" style=""background-color: #095c7b; padding: 25px 20px; text-align: center;"">
              <img src=""https://lh3.googleusercontent.com/d/1hhLMkl8NmyhkhDT9jDg9AYIhbIRsjQQD"" alt=""MailPlus Logo"" width=""135"" style=""display: inline-block; vertical-align: middle; border: 0; max-height: 42px; width: auto;"" />
            </td>
          </tr>
          <tr>
            <td align=""center"" style=""background-color: #f8fafb; padding: 30px 20px; text-align: center; border-top: 1px solid #edf2f7; font-size: 12px; color: #718096; line-height: 1.5;"">
              <p style=""margin: 0 0 6px; font-size: 12px;"">
                <strong style=""font-weight: 700; color: #4a5568;"">MailPlus</strong> | Business logistics, made simple.
              </p>
              <p style=""margin: 0 0 15px; font-size: 12px;"">
                Powered by MailPlus Australia
              </p>
              <p style=""margin: 0; font-size: 11px; color: #a0aec0; line-height: 1.5;"">
                &copy; 2026 MailPlus. All rights reserved.
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>";
        }

        private static string ExtractFirstName(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return "User";
            string prefix = email.Split('@')[0];
            string[] parts = prefix.Split('.');
            if (parts.Length > 0 && !string.IsNullOrWhiteSpace(parts[0]))
            {
                return char.ToUpper(parts[0][0]) + parts[0].Substring(1).ToLower();
            }
            return "User";
        }
    }
}
