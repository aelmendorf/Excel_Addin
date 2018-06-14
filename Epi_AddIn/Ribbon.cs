using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Epi_AddIn.Worker_Code;
using System.Data;
using System.Windows.Forms;

//using 

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Epi_AddIn {
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;
        private DB_EpiWrapper epiWrapper=new DB_EpiWrapper();

        public Ribbon() {

        }



        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("Epi_AddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        public void GetSpectrum_Callback(Office.IRibbonControl control) {
          /*  Excel.Range sel = Globals.ThisAddIn.Application.Selection as Excel.Range;
            List<string> wafers = new List<string>();
            foreach(var cell in sel.Cells) {
                wafers.Add((string)(((Excel.Range)cell).Value2));
            }
            DataSet dSet = this.epiWrapper.GetWaferAll(wafers.ToArray());
            this.epiWrapper.OutToExcel(dSet);*/
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for(int i = 0; i < resourceNames.Length; ++i) {
                if(string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using(StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if(resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
