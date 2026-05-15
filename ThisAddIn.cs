using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
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
            // Reserved for future email send event processing
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                Application.ItemSend -= Application_ItemSend;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during add-in shutdown: {ex.Message}");
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
