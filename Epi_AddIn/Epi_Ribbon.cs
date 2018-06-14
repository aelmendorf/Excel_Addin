using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Epi_AddIn.Worker_Code;
using System.Data;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using System.Threading;

namespace Epi_AddIn {
    public partial class Epi_Ribbon {
        private DB_EpiWrapper epiWrapper = new DB_EpiWrapper();
        private const string EWAT = "NewEWAT 2018.xlsb";

        private void Epi_Ribbon_Load(object sender, RibbonUIEventArgs e) {

        }

       /* private async void getSpectrum_Click(object sender, RibbonControlEventArgs e) {
            Excel.Range sel = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if(sel != null) {
                this.getSpectrum.Enabled = false;
                this.openEWAT.Enabled = false;
                this.importBurn.Enabled = false;
                List<string> wafers = new List<string>();
                if(Globals.ThisAddIn.Application.ActiveWorkbook.Name == EWAT) {
                    if(sel.Column == 2) {
                        foreach(Excel.Range cell in sel.Cells) {
                            if(cell.Value2 != null  && (string)cell.Value2!="") {
                                wafers.Add((string)cell.Value2);
                            }
                        }
                        if(wafers.Count > 0) {
                            MessageBox.Show("Collecting Data" + Environment.NewLine + "new workbook will open when done");

                            if(SynchronizationContext.Current == null)
                                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                            await this.GetSpectrumData(wafers.ToArray());
                        }
                    } else {
                        MessageBox.Show("Please select from RunID Column and try again ");
                    }//if EWAT file make sure that correct column is selected
                } else {
                    foreach(Excel.Range cell in sel.Cells) {
                        if(cell.Value2 != null && (string)cell.Value2 != "") {
                            wafers.Add((string)cell.Value2);
                        }
                    }
                    if(wafers.Count > 0) {
                        MessageBox.Show("Collecting Data" + Environment.NewLine + "new workbook will open when done");

                        if(SynchronizationContext.Current == null)
                            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                        await this.GetSpectrumData(wafers.ToArray());
                    }
                }//End check file
                this.getSpectrum.Enabled = true;
                this.openEWAT.Enabled = true;
                this.importBurn.Enabled = true;
            } else {
                MessageBox.Show("Invalid selection" + Environment.NewLine + "Please try selecting again");
            }
        }*/

        private void getSpectrum_Click(object sender, RibbonControlEventArgs e) {
            Excel.Range sel = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if(sel != null) {
                this.getSpectrum.Enabled = false;
                this.openEWAT.Enabled = false;
                this.importBurn.Enabled = false;
                List<string> wafers = new List<string>();
                if(Globals.ThisAddIn.Application.ActiveWorkbook.Name == EWAT) {
                    if(sel.Column == 2) {
                        foreach(Excel.Range cell in sel.Cells) {
                            if(cell.Value2 != null && (string)cell.Value2 != "") {
                                if(epiWrapper.Exist((string)cell.Value2, TEST_TYPE.INITIAL)==1) {
                                    wafers.Add((string)cell.Value2);
                                }
                            }
                        }//End loop
                        if(wafers.Count > 0) {
                            MessageBox.Show("Collecting Data" + Environment.NewLine + "new workbook will open when done");
                            this.epiWrapper.GetWaferAll(wafers.ToArray());
                        } else {
                            MessageBox.Show("No wafer entry found"+Environment.NewLine+"Please check selection and try again");
                        }//End check for verified wafers
                    } else {
                        MessageBox.Show("EWAT File "+Environment.NewLine+" Please check that you have"+Environment.NewLine+
                            "RunID Column selected and try again ");
                    }//if EWAT file make sure that correct column is selected
                } else {
                    foreach(Excel.Range cell in sel.Cells) {
                        if(cell.Value2 != null && (string)cell.Value2 != "") {
                            if(epiWrapper.Exist((string)cell.Value2, TEST_TYPE.INITIAL) == 1) {
                                wafers.Add((string)cell.Value2);
                            }
                        }
                    }//End loop through selection
                    if(wafers.Count > 0) {
                        MessageBox.Show("Collecting Data" + Environment.NewLine + "new workbook will open when done");
                        this.epiWrapper.GetWaferAll(wafers.ToArray());
                    } else {
                        MessageBox.Show("No wafer entry found" + Environment.NewLine + "Please check selection and try again");
                    }//End check if found
                }//End check file
                this.getSpectrum.Enabled = true;
                this.openEWAT.Enabled = true;
                this.importBurn.Enabled = true;
            } else {
                MessageBox.Show("Invalid selection" + Environment.NewLine + "Please try selecting again");
            }
        }

        private async Task GetSpectrumData(string[] wafers) {
            await Task.Run(() => this.epiWrapper.GetWaferAll(wafers));
        }

        private void openEWAT_Click(object sender, RibbonControlEventArgs e) {
            if(Globals.ThisAddIn.Application.ActiveWorkbook.Name != EWAT) {
                FileInfo file = new FileInfo(@"\\172.20.4.6\Data\Interdepartmental data\EpiData\NewEWAT 2018.xlsb");
                Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(file.FullName);
            } else {
                MessageBox.Show("You alread have file open");
            }//End check if file open

        }

        private void importBurn_Click(object sender, RibbonControlEventArgs e) {
            Excel.Range sel = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if(sel != null) {
                List<string> notFound = new List<string>();
                TEST_TYPE type = TEST_TYPE.AFTER;

                string start="", stop="",date="";
                if(this.testType.Text == "Initial") {
                    type = TEST_TYPE.INITIAL;
                    start = "M";
                    stop = "AJ";
                    date = "L";
                } else if(this.testType.Text == "After") {
                    type = TEST_TYPE.AFTER;
                    start = "BC";
                    stop = "BZ";
                    date = "BB";
                } else {
                    MessageBox.Show("Please select a Test Type" + Environment.NewLine + "and try again.");
                    return;
                }//End check for selection
                this.getSpectrum.Enabled = false;
                this.openEWAT.Enabled = false;
                this.importBurn.Enabled = false;
                Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                foreach(Excel.Range cell in sel.Cells) {
                    if(cell.Value2 != null && (string)cell.Value2 != "") {
                        if(epiWrapper.Exist((string)cell.Value2,type) == 1) {
                            Excel.Range output = ws.get_Range(start + cell.Row, stop+ cell.Row);
                            DataTable data = this.epiWrapper.GetPointData_BurnIn(cell.Value2,type);
                            object[,] Cells = new object[data.Rows.Count, data.Columns.Count];
                            for(int j = 0; j < data.Rows.Count; j++) {
                                for(int i = 0; i < data.Columns.Count; i++) {
                                    if((double)data.Rows[j][i] == 0.00) {
                                        Cells[j, i] = "";
                                    } else {
                                        Cells[j, i] = data.Rows[j][i];
                                    }//End check for 0, empty if 0
                                }
                            }//End loop through
                            output.Value = Cells;
                            ws.get_Range(date + cell.Row).Value = DateTime.Now.ToShortDateString();
                        } else {
                            notFound.Add((string)cell.Value2);
                        }//End check if exist
                    }//double check not null
                }//End loop through range
                if(notFound.Count > 0) {
                    string message = "";
                    foreach(string not in notFound) {
                        message += not + Environment.NewLine;
                    }
                    MessageBox.Show("Wafers not Found: " + message);
                }//End check for not found wafers and message
                MessageBox.Show("Import Done");
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                this.getSpectrum.Enabled = true;
                this.openEWAT.Enabled = true;
                this.importBurn.Enabled = true;
            }//End check for selection
        }//End importBurn
    }//End ribbon
}
