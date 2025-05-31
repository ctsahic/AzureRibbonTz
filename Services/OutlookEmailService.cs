using System.Net;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1.Services
{
    public class OutlookEmailService : IEmailService
    {
        public string CleanDescription(string emailBody)
        {
            emailBody = Regex.Replace(emailBody, "<[^>]*>", string.Empty);
            emailBody = Regex.Replace(emailBody, @"\s+", " ");
            return emailBody.Trim();
        }

        public Outlook.MailItem GetSelectedEmail()
        {
            Outlook.Application app = new Outlook.Application();
            return app.ActiveExplorer().Selection[1] as Outlook.MailItem;
        }

        public Outlook.Attachments GetEmailAttachments(Outlook.MailItem mailItem)
        {
            return mailItem?.Attachments;
        }

        public string GetFormattedDescription(Outlook.MailItem mailItem)
        {
            // Get HTML body if available, otherwise use plain text
            string body = mailItem.HTMLBody ?? mailItem.Body;

            if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                // Azure DevOps supports a subset of HTML - let's ensure it's clean and compatible
                body = CleanHtmlForAzureDevOps(body);
            }
            else
            {
                // For plain text, preserve line breaks by converting them to <br/> tags
                body = WebUtility.HtmlEncode(body) // Replaced System.Web.HttpUtility with System.Net.WebUtility
                    .Replace("\r\n", "<br/>")
                    .Replace("\n", "<br/>");
            }

            return body;
        }

        private string CleanHtmlForAzureDevOps(string html)
        {
            // Remove potentially problematic scripts and styles
            html = Regex.Replace(html, @"<script.*?</script>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"<style.*?</style>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Ensure all tags are properly closed and HTML is well-formed
            // Keep only basic formatting tags that Azure DevOps supports
            html = Regex.Replace(html, @"</?(?!b|i|u|strong|em|strike|br|p|div|h[1-6]|ul|ol|li|code|pre)[^>]*>", "");

            // Clean up excessive whitespace while preserving structure
            html = Regex.Replace(html, @">\s+<", "><");

            return html.Trim();
        }
    }
}