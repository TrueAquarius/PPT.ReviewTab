using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace PPT.ReviewTab
{
    public partial class ThisAddIn
    {
       

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

#if true
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ReviewRibbonTab();
        }
#endif

#if false
        protected override Microsoft.Office.Tools.Ribbon.IRibbonExtension[] CreateRibbonObjects()
        {
            ReviewRibbonTab reviewRibbonTab = new ReviewRibbonTab();
            return new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new ReviewTab(), reviewRibbonTab };
        }
#endif


        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
