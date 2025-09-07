using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordMarkdownAddIn
{
    public partial class MarkdownRibbon
    {
        private void MarkdownRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.PaneControl.SaveMarkdownFile(); 
        }

        private void btnPanel_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TogglePane();
        }

        private void btnOpen_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.PaneControl.OpenMarkdownFile();
        }

        private void bBold_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.PaneControl.InsertInline("**", "**");
        }
    }
}
