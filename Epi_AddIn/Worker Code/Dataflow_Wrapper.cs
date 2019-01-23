using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Office = Microsoft.Office.Core;
using ExcelApp = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.ComponentModel;
using System.Collections.Specialized;

namespace Epi_AddIn.Worker_Code {
    public enum TEST_TYPE {
        AFTER = 2,
        INITIAL = 1,
        NOTSET = 3
    };
    public enum TEST_AREA {
        [Description("CenterA")] CENTERA = 1,
        [Description("CenterB")] CENTERB = 2,
        [Description("CenterC")] CENTERC = 3,
        [Description("Right")] RIGHT = 4,
        [Description("Top")] TOP = 5,
        [Description("Left")] LEFT = 6,
        [Description("Not Set")] NOTSET = 7
    };

    public static class Area_Extensions {

        public static string GetAreaString(this TEST_AREA area) {
            DescriptionAttribute[] attributes = (DescriptionAttribute[])area.GetType().GetField(area.ToString()).GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attributes.Length > 0 ? attributes[0].Description : string.Empty;
        }

        public static NameValueCollection ToList<T>() where T : struct {
            var result = new NameValueCollection();
            if(!typeof(T).IsEnum) return result;
            var enumType = typeof(T);
            var values = Enum.GetValues(enumType);
            foreach(var value in values) {
                var memInfo = enumType.GetMember(enumType.GetEnumName(value));
                var descriptionAttributes = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                var description = descriptionAttributes.Length > 0
                    ? ((DescriptionAttribute)descriptionAttributes.First()).Description
                    : value.ToString();
                result.Add(description, value.ToString());
            }
            return result;
        }

        public static TEST_AREA GetTestArea(this string val) {
            TEST_AREA[] enumVals = (TEST_AREA[])Enum.GetValues(typeof(TEST_AREA));

            for(int i = 0; i < enumVals.Length; i++) {
                if(val.Contains(enumVals[i].GetAreaString())) {
                    return enumVals[i];
                }
            }
            return TEST_AREA.NOTSET;
        }
    }

    public class Spectrum {
        public const int ARRAY_SIZE = 2048;
        private double[] _wl, _intensity;
        private int _current;
        private TEST_AREA _area;

        public double[] Wl { get => this._wl; set => this._wl = value; }
        public double[] Intensity { get => this._intensity; set => this._intensity = value; }
        public TEST_AREA Area { get => this._area; set => this._area = value; }
        public int Current { get => this._current; set => this._current = value; }

        public Spectrum() {
            this._wl = new double[ARRAY_SIZE];
            this._intensity = new double[ARRAY_SIZE];
            this._area = TEST_AREA.NOTSET;
        }

        public Spectrum(TEST_AREA area, int current, double[] wl, double[] intensity) {
            this._area = area;
            this._current = current;
            this._wl = wl;
            this._intensity = intensity;
        }

        public void Set(TEST_AREA area, double[] wl, double[] intensity) {
            this._area = area;
            Array.Copy(wl, this._wl, wl.Length);
            Array.Copy(intensity, this._intensity, intensity.Length);
        }
    }

    public class Dataflow_Wrapper {
        // The head of the dataflow network.
        public ITargetBlock<string> headBlock = null;
        private string server, uid, password, database, connectionString;
        private static readonly int SPECTRUM_SIZE = 2048;
        public static readonly int COLUMNCOUNT = 24;
        public static readonly string[] COLUMNS = { "CenterA_WL","CenterA_Spect","CenterA_WL_50mA","CenterA_Spect_50mA",
        "CenterB_WL", "CenterB_Spect", "CenterB_WL_50mA", "CenterB_Spect_50mA", "CenterC_WL", "CenterC_Spect", "CenterC_WL_50mA",
            "CenterC_Spect_50mA", "Right_WL", "Right_Spect", "Right_WL_50mA", "Right_Spect_50mA", "Top_WL", "Top_Spect", "Top_WL_50mA",
            "Top_Spect_50mA", "Left_WL", "Left_Spect", "Left_WL_50mA", "Left_Spect_50mA"};

        public Dataflow_Wrapper() {
            server = "172.20.4.20";
            uid = "aelmendorf";
            password = "Drizzle123!";
            database = "epi";
            connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";" + "SslMode=none";
        }

        private int Exist(string wafer, TEST_TYPE type) {
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

        private DataTable GetWaferSpectrum(string wafer, TEST_TYPE type) {
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
                return tbl;
            } catch(MySqlException ex) {
                return null;
            }
        }//End GetWaferSpectrum

        private IEnumerable<Spectrum> ExtractSpectrum(DataTable dt) {
            List<Spectrum> spectList = new List<Spectrum>();
            double[] wl = new double[Spectrum.ARRAY_SIZE];
            double[] inten = new double[Spectrum.ARRAY_SIZE];

            int count = 0;
            foreach(DataRow row in dt.Rows) {
                foreach(DataColumn col in dt.Columns) {
                    TEST_AREA area = col.ColumnName.GetTestArea();
                    if(!DBNull.Value.Equals(row[col])) {

                        if(col.ColumnName.Contains("WL")) {
                            wl = JsonConvert.DeserializeObject<double[]>((string)row[col]);

                        } else {
                            inten = JsonConvert.DeserializeObject<double[]>((string)row[col]);
                        }
                    } else {
                        if(col.ColumnName.Contains("WL")) {
                            wl = new double[Spectrum.ARRAY_SIZE];

                        } else {
                            inten = new double[Spectrum.ARRAY_SIZE];
                        }
                    }
                    count += 1;
                    if(count % 2 == 1) {
                        int cur = col.ColumnName.Contains("50mA") == false ? 20 : 50;
                        try {
                            spectList.Add(new Spectrum(area, cur,
                                wl,
                                inten));

                        } catch(ArgumentNullException e) {

                        }
                    }
                }
            }//End transpose/convert
            return spectList;
        }//End Extract Spectrums

        public void Run(string wafer) {
            TransformBlock<string, DataTable> loadWafers = new TransformBlock<string, DataTable>(
                w => {
                    return GetWaferSpectrum(w, TEST_TYPE.INITIAL);
                });

            TransformBlock<DataTable, IEnumerable<Spectrum>> convert = new TransformBlock<DataTable, IEnumerable<Spectrum>>(
                dt => {
                    return ExtractSpectrum(dt);
                });

            ActionBlock<IEnumerable<Spectrum>> display = new ActionBlock<IEnumerable<Spectrum>>(list => {
                List<Spectrum> tmp = list.ToList<Spectrum>();
                Globals.ThisAddIn.Application.Calculation = ExcelApp.XlCalculation.xlCalculationManual;
                ExcelApp.Application Excel = ((ExcelApp.Application)Globals.ThisAddIn.Application);
                var wb = Excel.Workbooks.Add();
                Excel.Visible = true;
                ExcelApp._Worksheet ws = wb.Sheets.Add();
                ws.Name = "TestBlock";

                for(int i= 0;i< tmp.Count;i+=2) {
                    ExcelApp.Range wl = ws.get_Range((ExcelApp.Range)(ws.Cells[1, i]), (ExcelApp.Range)(ws.Cells[tmp[i].Wl.Length,i]));
                    ExcelApp.Range intensity = ws.get_Range((ExcelApp.Range)(ws.Cells[1, i]), (ExcelApp.Range)(ws.Cells[tmp[i].Intensity.Length, i]));
                    wl.Value = tmp[i].Wl;
                    intensity.Value = tmp[i].Intensity;
                }
                Globals.ThisAddIn.Application.Calculation = ExcelApp.XlCalculation.xlCalculationSemiautomatic;
              });

            loadWafers.LinkTo(convert);
            convert.LinkTo(display);
            //loadWafers.Post(wafer);
            this.headBlock = loadWafers;
            this.headBlock.Post(wafer);
        }

    }

}
