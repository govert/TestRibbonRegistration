using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace TestRibbonRegistration
{
    internal class ComConstants
    {
        // Choose your own ProgId and a new Guid here
        public const string RibbonProgId = "TestRibbonRegistration.MyCustomRibbon";
        public const string RibbonGuid = "E595901B-903E-4FE7-8D06-6E7F3D5A2C4F";

        // This is the Guid for the IRibbonExtensibility interface - don't change this
        public const string IRibbonExtensibilityGuid = "000C0396-0000-0000-C000-000000000046";
    }

    [ComVisible(true)]
    [ComImport]
    [Guid(ComConstants.IRibbonExtensibilityGuid)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    [ProgId(ComConstants.RibbonProgId)]
    [Guid(ComConstants.RibbonGuid)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class MyCustomRibbon : ExcelComAddIn, IRibbonExtensibility
    {
        public string GetCustomUI(string RibbonID)
        {
            return 
                @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                  <ribbon>
                    <tabs>
  
                      <tab id='tab1' label='My Tab'  >
                        <group id='group1' label='My Group'>
                          <button id='button1' label='My Button' onAction='OnButtonPressed'  />
                        </group >
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            var app = (Application)ExcelDnaUtil.Application;
            app.StatusBar = "Button Pressed!";
        }
    }
}
