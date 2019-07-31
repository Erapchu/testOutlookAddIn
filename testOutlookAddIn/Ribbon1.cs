using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace testOutlookAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (e.Control.Context is Outlook.Inspector)
                {
                    var inspector = e.Control.Context as Outlook.Inspector;
                    var thisMail = inspector.CurrentItem as Outlook.MailItem;
                    FileProcess.SaveToFile(thisMail.Subject, thisMail.Body);

                    Marshal.ReleaseComObject(inspector);
                    Marshal.ReleaseComObject(thisMail);
                }
            }
            catch
            {
                MessageBox.Show("Не удалось инициировать сохранение сообщения", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
