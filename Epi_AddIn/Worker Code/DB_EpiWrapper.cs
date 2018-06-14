using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using ExcelApp = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Data;
using Newtonsoft.Json;
using MySql.Data.MySqlClient;
using System.Collections;

namespace Epi_AddIn.Worker_Code {
    public enum TEST_TYPE { AFTER = 2, INITIAL = 1 };
    public enum TEST_AREA { CENTERA = 1, CENTERB = 2, CENTERC = 3, RIGHT = 4, TOP = 5, LEFT = 6 };

    public class DB_EpiWrapper {
        private string server, uid, password, database, connectionString;
        private const int ARRAY_SIZE = 2048;

        public DB_EpiWrapper() {
            server = "172.20.4.20";
            uid = "aelmendorf";
            password = "Drizzle123!";
            database = "epi";
            connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";" + "SslMode=none";
        }

        public void GetBurnIn(string[] wafers,ExcelApp.Range sel) {


        }

        public void GetWaferAll(string[] wafers) {
            DataSet dSet = new DataSet("Burn-in Data");
            DataTable temp = this.GetWaferPointData(wafers);
            if(temp != null) {
                dSet.Tables.Add(temp);
                this.GetWaferSpectData(wafers, ref dSet);
                this.OutToExcel(dSet);
            }//End check if null
        }//End get all

        public DataTable ConvertTransposeSpectrum(DataTable tbl) {
            Dictionary<string, double[]> values = new Dictionary<string, double[]>();
            List<double> wlArrs = new List<double>();
            List<double> spectArrs = new List<double>();

            DataTable tblNew = new DataTable();
            tblNew = tbl.Clone();

            foreach(DataRow row in tbl.Rows) {
                foreach(DataColumn col in tbl.Columns) {

                    if(!DBNull.Value.Equals(row[col])) {
                        values.Add(col.ColumnName, JsonConvert.DeserializeObject<double[]>((string)row[col]));
                        //Console.WriteLine(row[col]);
                    } else {
                        values.Add(col.ColumnName, new double[ARRAY_SIZE]);
                    }
                }
            }//End transpose/convert

            foreach(KeyValuePair<string, double[]> item in values) {
                for(int i = 0; i < item.Value.Length; i++) {
                    tblNew.Rows.Add();
                }
            }

            for(int y = 0; y < tblNew.Columns.Count; y++) {
                for(int i = 0; i < values[tblNew.Columns[y].ColumnName].Length; i++) {
                    tblNew.Rows[i][y] = values[tblNew.Columns[y].ColumnName][i];
                }
            }
            return tblNew;
        }//End ConvertTransposeSpectrum

        public DataTable GetWaferPointData(string[] wafers) {
            DataTable init = new DataTable();
            DataTable after = new DataTable();
            DataTable main = new DataTable();

            for(int i = 0; i < wafers.Length; i++) {
                DataTable temp = new DataTable();

                init = this.GetPoint(wafers[i], TEST_TYPE.INITIAL);
                after = this.GetPoint(wafers[i], TEST_TYPE.AFTER);
                if(init != null && after != null) {
                    temp = this.JoinDataTable(init, after, "WaferID");
                } else {
                    return null;
                }

                if(i == 0) {
                    main = temp.Clone();
                    foreach(DataRow row in temp.Rows) {
                        DataRow nRow = main.NewRow();
                        foreach(DataColumn col in temp.Columns) {
                            nRow[col.ColumnName] = row[col.ColumnName];
                        }
                        main.Rows.Add(nRow);
                    }
                } else {
                    foreach(DataRow row in temp.Rows) {
                        DataRow nRow = main.NewRow();
                        foreach(DataColumn col in temp.Columns) {
                            nRow[col.ColumnName] = row[col.ColumnName];
                        }
                        main.Rows.Add(nRow);
                    }
                }//End check not 0
            }//End loop
            main.TableName = "PointData";
            return main;
        }//End getWafer

        public DataTable GetPoint(string wafer, TEST_TYPE type) {
            DataTable tbl = new DataTable();
            try {
                using(MySqlConnection connect = new MySqlConnection(connectionString)) {
                    connect.Open();
                    string query = "getWaferData_All";
                    MySqlCommand cmd = new MySqlCommand(query, connect);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.Parameters.AddWithValue("@wafer", wafer);
                    cmd.Parameters.AddWithValue("@test", (int)type);
                    cmd.ExecuteNonQuery();

                    using(MySqlDataAdapter adp = new MySqlDataAdapter(cmd)) {
                        adp.Fill(tbl);
                    }
                    connect.Close();
                    return tbl;
                }
            } catch(MySqlException ex) {
                return null;
            }
        }

        public DataTable GetPointData_BurnIn(string wafer, TEST_TYPE type) {
            DataTable tbl = new DataTable();
            try {
                using(MySqlConnection connect = new MySqlConnection(connectionString)) {
                    connect.Open();
                    string query = "getWaferData";
                    MySqlCommand cmd = new MySqlCommand(query, connect);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.Parameters.AddWithValue("@wafer", wafer);
                    cmd.Parameters.AddWithValue("@test", (int)type);
                    cmd.ExecuteNonQuery();

                    using(MySqlDataAdapter adp = new MySqlDataAdapter(cmd)) {
                        adp.Fill(tbl);
                    }
                    connect.Close();
                    return tbl;
                }
            } catch(MySqlException ex) {
                return null;
            }
        }

        public DataTable GetSpect(string wafer, TEST_TYPE type) {
            DataTable tbl = new DataTable();
            try {
                using(MySqlConnection connect = new MySqlConnection(connectionString)) {
                    connect.Open();
                    string query = "get_spectrum";
                    MySqlCommand cmd = new MySqlCommand(query, connect);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.Parameters.AddWithValue("@wafer", wafer);
                    cmd.Parameters.AddWithValue("@test", (int)type);
                    cmd.ExecuteNonQuery();
                    using(MySqlDataAdapter adp = new MySqlDataAdapter(cmd)) {
                        adp.Fill(tbl);
                    }
                }
                return this.ConvertTransposeSpectrum(tbl);
            } catch(MySqlException ex) {
                return null;
            }
        }

        public void GetWaferSpectData(string[] wafers, ref DataSet main) {

            DataTable init = new DataTable();
            DataTable after = new DataTable();

            // DataTable main = new DataTable();

            for(int i = 0; i < wafers.Length; i++) {
                init = this.GetSpect(wafers[i], TEST_TYPE.INITIAL);
                after = this.GetSpect(wafers[i], TEST_TYPE.AFTER);
                if(init != null && after != null) {
                    main.Tables.Add(this.JoinDataTableSpectrum(init, after, wafers[i] + "_Spectrum"));
                }
            }

        }//End get Spectrum data

        public DataTable JoinDataTable(DataTable dt_1, DataTable dt_2, string joinField) {
            var dt = new DataTable();

            foreach(DataColumn col in dt_1.Columns) {
                dt.Columns.Add(col.ColumnName + "_Initial", typeof(string));
            }

            foreach(DataColumn col in dt_2.Columns) {
                dt.Columns.Add(col.ColumnName + "_After", typeof(string));
            }

            var nRow = dt.NewRow();
            foreach(DataColumn col in dt.Columns) {
                if(col.ColumnName.Contains("_Initial")) {
                    nRow[col.ColumnName] = dt_1.Rows[0][col.ColumnName.Substring(0, col.ColumnName.IndexOf("_Initial"))];
                } else {
                    nRow[col.ColumnName] = dt_2.Rows[0][col.ColumnName.Substring(0, col.ColumnName.IndexOf("_After"))];
                }
            }
            dt.Rows.Add(nRow);
            dt.Columns.Remove("WaferID" + "_After");
            dt.Columns.Remove("System" + "_After");

            return dt;
        }

        public DataTable JoinDataTableSpectrum(DataTable dt_1, DataTable dt_2, string tableName) {
            var dt = new DataTable();

            foreach(DataColumn col in dt_1.Columns) {
                dt.Columns.Add(col.ColumnName + "_Initial", typeof(string));
            }

            foreach(DataColumn col in dt_2.Columns) {
                dt.Columns.Add(col.ColumnName + "_After", typeof(string));
            }

            for(int rows = 0; rows < dt_1.Rows.Count; rows++) {
                var nRow = dt.NewRow();
                foreach(DataColumn col in dt.Columns) {
                    if(col.ColumnName.Contains("_Initial")) {
                        nRow[col.ColumnName] = dt_1.Rows[rows][col.ColumnName.Substring(0, col.ColumnName.IndexOf("_Initial"))];
                    } else {
                        nRow[col.ColumnName] = dt_2.Rows[rows][col.ColumnName.Substring(0, col.ColumnName.IndexOf("_After"))];
                    }
                }
                dt.Rows.Add(nRow);
            }
            dt.TableName = tableName;
            return dt;
        }//End

        public int Exist(string wafer, TEST_TYPE type) {
            int retVal = 0;
            try {
                using(MySqlConnection connect = new MySqlConnection(connectionString)) {
                    connect.Open();
                    string query = "check";
                    MySqlCommand cmd = new MySqlCommand(query, connect);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.Parameters.AddWithValue("@wafer", wafer);
                    cmd.Parameters.AddWithValue("@test", (int)type);
                    cmd.Parameters.AddWithValue("?isentry", MySqlDbType.Int32);
                    cmd.Parameters["?isentry"].Direction = System.Data.ParameterDirection.Output;
                    cmd.ExecuteNonQuery();
                    retVal = (int)cmd.Parameters["?isentry"].Value;
                }
                return retVal;
            } catch(MySqlException ex) {
                return -1;
            }
        }

        public void OutToExcel(DataSet dSet) {
            Globals.ThisAddIn.Application.Calculation = ExcelApp.XlCalculation.xlCalculationManual;
            ExcelApp.Application Excel= ((ExcelApp.Application)Globals.ThisAddIn.Application);
            var wb = Excel.Workbooks.Add();
            Excel.Visible = true;
            foreach(DataTable dt in dSet.Tables) {
                int ColumnsCount = dt.Columns.Count;
                ExcelApp._Worksheet ws = wb.Sheets.Add();
                if(dt.TableName == "PointData") {

                    ws.Name = "Point Data";
                    object[] Header = new object[ColumnsCount];

                    for(int i = 0; i < ColumnsCount; i++)
                        Header[i] = dt.Columns[i].ColumnName;

                    ExcelApp.Range HeaderRange = ws.get_Range((ExcelApp.Range)(ws.Cells[1, 1]), (ExcelApp.Range)(ws.Cells[1, ColumnsCount]));
                    HeaderRange.Value = Header;

                    HeaderRange.Font.Bold = true;
                    int RowsCount = dt.Rows.Count;
                    object[,] Cells = new object[RowsCount, ColumnsCount];

                    for(int j = 0; j < RowsCount; j++)
                        for(int i = 0; i < ColumnsCount; i++)
                            Cells[j, i] = dt.Rows[j][i];

                    ws.get_Range((ExcelApp.Range)(ws.Cells[2, 1]),(ExcelApp.Range)(ws.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;

                } else {
                    ws.Name = dt.TableName;

                    object[] Header = new object[ColumnsCount];

                    for(int i = 0; i < ColumnsCount; i++)
                        Header[i] = dt.Columns[i].ColumnName;

                    ExcelApp.Range HeaderRange = ws.get_Range((ExcelApp.Range)(ws.Cells[1, 1]), (ExcelApp.Range)(ws.Cells[1, ColumnsCount]));
                    HeaderRange.Value = Header;

                    HeaderRange.Font.Bold = true;
                    int RowsCount = dt.Rows.Count;

                    object[,] Cells = new object[RowsCount, ColumnsCount];

                    for(int j = 0; j < RowsCount; j++)
                        for(int i = 0; i < ColumnsCount; i++)
                            Cells[j, i] = dt.Rows[j][i];

                    ws.get_Range((ExcelApp.Range)(ws.Cells[2, 1]), (ExcelApp.Range)(ws.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;
                }
            }//End DataSet loop through
            Globals.ThisAddIn.Application.Calculation = ExcelApp.XlCalculation.xlCalculationSemiautomatic;
        }//End out to excel
    }
}
