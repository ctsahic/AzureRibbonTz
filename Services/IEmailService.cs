using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1.Services
{
    public interface IEmailService
    {
        Outlook.MailItem GetSelectedEmail();
        string GetFormattedDescription(Outlook.MailItem mailItem);
        Outlook.Attachments GetEmailAttachments(Outlook.MailItem mailItem);
    }
}