using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelIdea
{
    public partial class ThisAddIn
    {
        public string dataSource = "file";
        public bool CCInBackground = false;
        private GeneralSettings settings;
        private RibbonControl main_ribbon;
        private CoreVBA vba_core;
        private EventWaitHandle CC_up;
        public SharedMemoryIPC smem_ipc;

        protected override object RequestComAddInAutomationService()
        {
            if (vba_core == null)
                vba_core = new CoreVBA();

            return vba_core;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            settings = new GeneralSettings();
            dataSource = settings.ExcelDataSource;
            CCInBackground = settings.CCInBackground;
            main_ribbon = Globals.Ribbons.RibbonControl;

            //check if the Compute Core has been started
            //if (!EventWaitHandle.TryOpenExisting("ijiBearExcelIdeaCCUp", out CC_up))
            //{
            //    //bring up the compute core
            //    string hash = Tools.SHA512String("!hello server h4x0r!");
            //    ProcessStartInfo proc = new ProcessStartInfo(@"C:\Users\Bob\Documents\Visual Studio 2013\Projects\ExcelIdeaConsole\ExcelIdeaConsole\bin\Debug\ExcelIdea.exe", "server " + hash);
            //    proc.UseShellExecute = false;
            //    proc.WorkingDirectory = @"C:\Users\Bob\Documents\Visual Studio 2013\Projects\ExcelIdeaConsole\ExcelIdeaConsole\bin\Debug\";
            //    //Process.Start(settings.CCPath, "ipc " + hash);
            //    Process.Start(proc);
            //}

            smem_ipc = new SharedMemoryIPC();

        }

        public void Start()
        {
            settings.Save();
            Application.ScreenUpdating = false;

            string file_path = Application.ActiveWorkbook.FullName;

            switch (dataSource)
            {
                case "file":
                    ComputeFile();
                    break;
                case "ipc":
                    break;
                default:
                    break;
            }
            
            Application.ScreenUpdating = true;
        }
        
        private void ComputeFile()
        {
            bool success = true;
            string filePath;
            string destPath;
            string exePath = @"C:\Users\Bob\Documents\Visual Studio 2013\Projects\ExcelIdeaConsole\ExcelIdeaConsole\bin\Release\ExcelIdea.exe"; //Temporary for testing

            if (Application.ActiveWorkbook.Path == "") //if this workbook hasn't been saved
            {
                //show dialog informing user that they need to save
                //show file save dialog
                //save filePath
                filePath = "";
            }
            else
                filePath = Application.ActiveWorkbook.FullName;

            //destPath = Path.GetTempPath() + @"\bob\ExcelIdeaAddIn\";
            destPath = settings.DefaultResultDir + @"\result.xlsx";

            //generate auth key for exe
                /* Auth key ideas:
                 * hash with salt & secret phrase (easily breakable by decomp)
                 * use unmanaged dll which has a hard-coded secret phrase as a trusted hasher & verifier
                 */
            string auth_key = Tools.SHA512String("!hello" + " file " + filePath + "|" + destPath + "h4x0r!");
            CCInBackground = false;
            //contacting exe or service?
            if (CCInBackground)
            {
                //contacting service
                //send message with source and destination file paths
            }
            else
            {
                //contacting exe with arguments
                ProcessStartInfo proc = new ProcessStartInfo(exePath, "file " + filePath + " " + destPath + " " + auth_key);
                proc.WindowStyle = ProcessWindowStyle.Hidden;
                proc.CreateNoWindow = true;
                proc.UseShellExecute = false;
                proc.WorkingDirectory = @"C:\Users\Bob\Documents\Visual Studio 2013\Projects\ExcelIdeaConsole\ExcelIdeaConsole\bin\Release\";
                proc.RedirectStandardOutput = true;
                using (Process ace = Process.Start(proc))
                {
                    StreamReader out_reader = ace.StandardOutput;
                    StatusForm sform = new StatusForm();
                    ProgressBar barProgress = (ProgressBar)sform.Controls.Find("barProgress", false)[0];
                    sform.Show();
                    string console_output;
                    TextBox txtStatus = (TextBox)sform.Controls.Find("txtConsoleLog", false)[0];
                    
                    while(!out_reader.EndOfStream)
                    {
                        console_output = out_reader.ReadLine();
                        if (success && console_output.Contains("ERROR"))
                            success = false;
                        txtStatus.AppendText(console_output + Environment.NewLine);
                        barProgress.PerformStep();
                    }
                    ace.WaitForExit();
                    sform.Hide();
                    sform.Dispose();
                }

                if (success && main_ribbon.chkOpenWhenDone.Checked)
                {
                    //open the result file in Excel

                }
            }
            
        }

        private void GetSheetDataIPC()
        {
            Excel.Workbook activeBook;

            Application.ScreenUpdating = false;

            try
            {
                activeBook = Application.ActiveWorkbook;
            }
            catch (Exception)
            {
                return;
            }

            Dictionary<string, int> dict_sheet_name_map = new Dictionary<string, int>(activeBook.Sheets.Count);
            List<ConcurrentDictionary<string, float>> dict_constants = new List<ConcurrentDictionary<string, float>>(activeBook.Sheets.Count);
            List<ConcurrentDictionary<string, string>> dict_string_constants = new List<ConcurrentDictionary<string, string>>(activeBook.Sheets.Count);
            List<ConcurrentDictionary<string, Tools.ArrayFormulaInfo>> dict_loader_array_formulas = new List<ConcurrentDictionary<string, Tools.ArrayFormulaInfo>>(activeBook.Sheets.Count);
            List<ConcurrentDictionary<string, string>> dict_formulas = new List<ConcurrentDictionary<string, string>>(activeBook.Sheets.Count);

            //get sheets
            for (int i = 1; i <= activeBook.Sheets.Count; i++)
            {
                Excel.Worksheet sheet = activeBook.Sheets.get_Item(i);
                dict_sheet_name_map.Add(sheet.Name, i);
                dict_constants.Add(new ConcurrentDictionary<string, float>(StringComparer.Ordinal));
                dict_formulas.Add(new ConcurrentDictionary<string, string>(StringComparer.Ordinal));
                dict_string_constants.Add(new ConcurrentDictionary<string, string>(StringComparer.Ordinal));
                dict_loader_array_formulas.Add(new ConcurrentDictionary<string, Tools.ArrayFormulaInfo>(StringComparer.Ordinal));
            }

            Stopwatch sw = new Stopwatch();
            sw.Start();
            Parallel.For(0, dict_sheet_name_map.Count, sh =>
            {
                Excel.Range usedRange;

                try
                {
                    usedRange = activeBook.Sheets.get_Item(sh + 1).UsedRange;
                }
                catch (Exception)
                {
                    return;
                }

                string start_address = usedRange.Cells[1, 1].Address.Replace("$", "");
                object[,] arr_formulas = usedRange.Formula as object[,];
                object[,] arr_values = usedRange.Value2 as object[,];
                //object[,] arr_hasarray = usedRange.
                int[] arr_size = new int[2];
                arr_size[0] = arr_formulas.GetLength(0);
                arr_size[1] = arr_formulas.GetLength(1);
                string end_address = Tools.NextCell(start_address, arr_size[0] - 1, arr_size[1] - 1);
                ConcurrentBag<string> hs_arr_exclusion = new ConcurrentBag<string>();

                //for each row r in R
                Parallel.For(1, arr_size[0] + 1, r =>
                {
                    //for (int r = 1; r <= arr_size[0]; r++)
                    //{
                    //for each column c in C
                    Parallel.For(1, arr_size[1] + 1, c =>
                    {
                        //for (int c = 1; c <= arr_size[1]; c++)
                        //{
                        //if arr_formula is empty and value is null, continue
                        //if ((string)arr_formulas[r, c] == "" && arr_values[r, c] == null) continue;
                        if ((string)arr_formulas[r, c] == "" && arr_values[r, c] == null) return;

                        string curr_address = Tools.NextCell(start_address, r - 1, c - 1);
                        //if there is no formula
                        if (!((string)arr_formulas[r, c]).Contains("="))
                        {
                            //get value type
                            string type = arr_values[r, c].GetType().ToString();

                            //add to relevant constant dictionary
                            switch (type)
                            {
                                case "System.Double":
                                    dict_constants[sh].TryAdd(curr_address, (float)(double)arr_values[r, c]);
                                    break;
                                case "System.String":
                                    dict_string_constants[sh].TryAdd(curr_address, (string)arr_values[r, c]);
                                    break;
                                default:
                                    break;
                            }
                        }
                        else //else
                        {
                            //if the cell address is in the exclusion hashset, continue
                            //if (hs_arr_exclusion.Contains(curr_address)) continue;
                            if (hs_arr_exclusion.Contains(curr_address)) return;

                            //Excel.Range cell = usedRange.get_Range(curr_address);
                            //if (cell.HasArray)
                            //{
                            //    Excel.Range array_range = cell.CurrentArray;
                            //    string first_addr = array_range.Cells[1, 1].Address.Replace("$", "");

                            //    //if the cell is an array formula and it's not the first one, continue
                            //    //if (curr_address != first_addr) continue;
                            //    if (curr_address != first_addr) return;

                            //    //else
                            //    //calculate the array formula range
                            //    Tools.ArrayFormulaInfo afi_temp;
                            //    afi_temp.Cols = array_range.Formula.GetLength(1);
                            //    afi_temp.Rows = array_range.Formula.GetLength(0);
                            //    afi_temp.StartCell = first_addr;
                            //    afi_temp.EndCell = Tools.NextCell(first_addr, afi_temp.Rows - 1, afi_temp.Cols - 1);

                            //    //add to dictionary
                            //    dict_loader_array_formulas[sh].TryAdd(first_addr, afi_temp);

                            //    //add to exclusion hashset all the array formula cells
                            //    string[] exc_cells = Tools.RangeCells(first_addr, afi_temp.EndCell);
                            //    for (int i = 0; i < exc_cells.Length; i++)
                            //        hs_arr_exclusion.Add(exc_cells[i]);
                            //}
                            dict_formulas[sh].TryAdd(curr_address, ((string)arr_formulas[r, c]).Replace("=", ""));
                        }
                    });
                });
                Marshal.ReleaseComObject(usedRange);
            });
            sw.Stop();
            Marshal.ReleaseComObject(activeBook);

            int num_sheets = dict_sheet_name_map.Count;
            sw.Start();
            PipeClient pipe_client = new PipeClient("ijiBearExcelIdeaEXCELSOURCE");
            PipeMessage msg = new PipeMessage();

            msg.dict_loader_array_formulas = new List<Dictionary<string, Tools.ArrayFormulaInfo>>(num_sheets);
            msg.dict_loader_array_formulas.AddRange(new Dictionary<string, Tools.ArrayFormulaInfo>[num_sheets]);

            msg.dict_loader_constants = new List<Dictionary<string, float>>(num_sheets);
            msg.dict_loader_constants.AddRange(new Dictionary<string, float>[num_sheets]);

            msg.dict_loader_formulas = new List<Dictionary<string, string>>(num_sheets);
            msg.dict_loader_formulas.AddRange(new Dictionary<string, string>[num_sheets]);

            msg.dict_loader_string_constants = new List<Dictionary<string, string>>(num_sheets);
            msg.dict_loader_string_constants.AddRange(new Dictionary<string, string>[num_sheets]);

            msg.dict_sheet_name_map = dict_sheet_name_map;

            for (int i = 0; i < num_sheets; i++)
            {
                msg.dict_loader_array_formulas[i] = dict_loader_array_formulas[i].ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                msg.dict_loader_constants[i] = dict_constants[i].ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                msg.dict_loader_formulas[i] = dict_formulas[i].ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                msg.dict_loader_string_constants[i] = dict_string_constants[i].ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            }

            pipe_client.Send(msg);
            sw.Stop();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
