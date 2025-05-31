using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1.Services
{
    public interface IEmailService
    {
        Outlook.MailItem GetSelectedEmail();
        string CleanDescription(string emailBody);
        Outlook.Attachments GetEmailAttachments(Outlook.MailItem mailItem);
    }
}