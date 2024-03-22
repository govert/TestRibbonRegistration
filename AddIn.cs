using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace TestRibbonRegistration
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // It's better if we don't load other add-ins during the AutoOpen method
            // So we schedule the ribbon load to run after the add-in is initialized, 
            // when Excel is idle again.
            ExcelAsyncUtil.QueueAsMacro(() => RegisterRibbon());
        }

        public void AutoClose()
        {
        }

        public void RegisterRibbon()
        {
            var app = (Application)ExcelDnaUtil.Application;
            var progId = ComConstants.RibbonProgId;
            // Enumerate the registered add-ins, and load the one with the right ProgId
            foreach (COMAddIn addIn in app.COMAddIns)
            {
                if (addIn.ProgId == progId)
                {
                    addIn.Connect = true;
                    break;
                }
            }
        }
    }
}
