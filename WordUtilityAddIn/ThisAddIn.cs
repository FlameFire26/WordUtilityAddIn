using System;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using WordUtilityAddIn.Ribbon;
using WordUtilityAddIn.Services;

namespace WordUtilityAddIn
{
    public partial class ThisAddIn
    {
        private UtilityRibbon utilityRibbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            utilityRibbon = new UtilityRibbon(
                new WordDocumentService(() => Application));
            return utilityRibbon;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.DocumentChange += Application_DocumentChange;
            Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.DocumentChange -= Application_DocumentChange;
            Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        private void Application_DocumentChange()
        {
            utilityRibbon?.NotifyDocumentChanged();
        }

        private void Application_WindowSelectionChange(Word.Selection selection)
        {
            utilityRibbon?.NotifySelectionChanged();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
