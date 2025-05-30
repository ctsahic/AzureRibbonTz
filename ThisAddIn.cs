using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace AzureRibbonTz
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Set TLS 1.2 as the default security protocol
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                // Initialize any global event handlers or services here
                Application.ItemSend += Application_ItemSend;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during add-in startup: {ex.Message}", "Startup Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            // This is where you can add any pre-send email processing
            // For example, you could validate the email content or add custom headers
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Clean up any event handlers
                Application.ItemSend -= Application_ItemSend;
            }
            catch (Exception ex)
            {
                // Log the error but don't show UI since Outlook is shutting down
                System.Diagnostics.Debug.WriteLine($"Error during add-in shutdown: {ex.Message}");
            }
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
