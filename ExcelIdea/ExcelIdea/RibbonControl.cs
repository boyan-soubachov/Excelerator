using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelIdea
{
    public partial class RibbonControl
    {
        private void RibbonControl_Load(object sender, RibbonUIEventArgs e)
        {
            //load data source setting
            switch (Globals.ThisAddIn.dataSource)
            {
                case "file":
                    ddl_InteropMethod.SelectedItemIndex = 0;
                    break;
                case "ipc":
                    ddl_InteropMethod.SelectedItemIndex = 1;
                    break;
                default:
                    ddl_InteropMethod.SelectedItemIndex = 0;
                    break;
            }
        }

        private void btn_start_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Start();
        }

        private void btn_Tools_Settings_Click(object sender, RibbonControlEventArgs e)
        {
            FormSettings frm_settings = new FormSettings();
            frm_settings.ShowInTaskbar = false;
            frm_settings.ShowDialog();
        }

        private void ddl_InteropMethod_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.dataSource = ddl_InteropMethod.SelectedItem.Tag.ToString();
        }
    }
}
