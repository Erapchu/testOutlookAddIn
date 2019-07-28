using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace testOutlookAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Hello mf");
        }
    }
}
