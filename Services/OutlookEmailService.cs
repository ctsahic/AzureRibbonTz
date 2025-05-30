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
    }
}