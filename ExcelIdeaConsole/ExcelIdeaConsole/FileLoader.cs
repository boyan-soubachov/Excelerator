/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Ionic.Zip;
using System.Diagnostics;

namespace ExcelIdea
{
    class FileLoader
    {
        private string file_path = "";

        public FileLoader()
        { }

        public FileLoader(string file_path)
        {
            this.file_path = file_path;
        }

        private string LoadSheet(int sheet_number, string temp_path, out Dictionary<string, string> dict_formulas, out Dictionary<string, float> dict_constants, out Dictionary<string, Tools.ArrayFormulaInfo> dict_array_formulas, out Dictionary<string, string> dict_string_constants, Dictionary<string, string> dict_shared_strings)
        {
            temp_path += @"\xl\worksheets\sheet" + sheet_number.ToString() + ".xml";
            if (!File.Exists(temp_path))
            {
                //Error here for wrong sheet
                dict_formulas = null;
                dict_constants = null;
                dict_array_formulas = null;
                dict_string_constants = null;
                return "ERROR: The specified sheet number does not exist in the file";
            }

            //initialise vars
            dict_formulas = new Dictionary<string, string>(1024, StringComparer.Ordinal);
            dict_constants = new Dictionary<string, float>(1024, StringComparer.Ordinal);
            dict_string_constants = new Dictionary<string, string>(1024, StringComparer.Ordinal);
            dict_array_formulas = new Dictionary<string, Tools.ArrayFormulaInfo>(128, StringComparer.Ordinal);
            HashSet<string> hset_array_exclude = new HashSet<string>(StringComparer.Ordinal);
            
            //load sheet file into output dictionaries
            using (XmlReader reader = XmlReader.Create(temp_path))
            {
                string cell_id;
                bool is_formula = false;
                bool is_shared_string;
                string reader_value;
                string reader_name;
                string[] cell_list;                

                while (reader.Read())
                {
                    //load formulas and constants from sheet
                    if (reader.Name == "c" && reader.IsStartElement())
                    {
                        if (hset_array_exclude.Contains(reader["r"])) //check if we're in an array cell that needs to be excluded
                            continue;
                        cell_id = reader["r"];
                        is_shared_string = reader["t"] == "s";

                        reader.Read();
                        reader_name = reader.Name;
                        is_formula = reader_name == "f";
                        if (!is_formula && reader_name != "v")
                        {
                            //error here, some unexpected cell format.
                        }
                        //handle shared formulas
                        if (is_formula && reader.GetAttribute("t") != null && reader.GetAttribute("ref") != null)
                        {
                            if (reader["t"] == "shared")
                            {
                                string[] ref_array = reader["ref"].Split(':');
                                string[] range = Tools.RangeCells(ref_array[0], ref_array[1]);

                                reader.Read();
                                reader_value = reader.Value;
                                List<string> addresses = Tools.GetCellsFromFormula(reader_value);

                                //find direction of increment
                                bool increment_horizontal = false; //increment direction, false = vertical, true = horizontal
                                if (addresses.Count > 0) //if the formula has at least one cell reference
                                {
                                    increment_horizontal = Tools.HorizontalIncrement(ref_array[0], ref_array[1]);

                                    //handle static addresses (contains $)
                                    for (int i = 0; i < addresses.Count; i++)
                                    {
                                        if (addresses[i].Contains("$"))
                                            if ((increment_horizontal && addresses[i].IndexOf('$') == 0) || (!increment_horizontal && addresses[i].IndexOf('$') > 0))
                                                addresses.RemoveAt(i);
                                    }
                                }

                                HashSet<Tuple<string, string>> address_bag = new HashSet<Tuple<string, string>>();

                                //for each formula, increment the cells to match the new formula. This section is faster with no Parallel.For!
                                for (int i = 0; i < range.Length; i++)
                                {
                                    if (addresses.Count > 0)
                                    {
                                        string[] new_addresses = addresses.ToArray();
                                        //for each address in the formula
                                        for (int j = 0; j < addresses.Count; j++) //faster single-threaded rather than Parallel.for
                                        {
                                            if (!increment_horizontal)
                                                new_addresses[j] = Tools.NextCell(new_addresses[j], i, 0);
                                            else
                                                new_addresses[j] = Tools.NextCell(new_addresses[j], 0, i);
                                        }
                                        StringBuilder output_formula = new StringBuilder(reader_value);
                                        for (int j = 0; j < addresses.Count; j++)
                                            output_formula.Replace(addresses[j], new_addresses[j]);

                                        address_bag.Add(new Tuple<string, string>(range[i], output_formula.ToString()));
                                    }
                                    else //if there's no cell addresses in the formula, just repeat the same
                                        address_bag.Add(new Tuple<string, string>(range[i], reader_value));
                                }

                                foreach (Tuple<string, string> item in address_bag)
                                    dict_formulas.Add(item.Item1, item.Item2);
                                continue;
                            }
                            else if (reader["t"] == "array")
                            {
                                Tools.ArrayFormulaInfo afi_temp = new Tools.ArrayFormulaInfo();
                                string[] range = reader["ref"].Split(':');
                                afi_temp.StartCell = range[0];
                                if (range.Length > 1)
                                {
                                    Tuple<int, int> dimensions = Tools.RangeSize(range[0], range[1]);
                                    afi_temp.EndCell = range[1];
                                    afi_temp.Rows = dimensions.Item1;
                                    afi_temp.Cols = dimensions.Item2;
                                    cell_list = Tools.RangeCells(range[0], range[1]);
                                    for (int i = 0; i < cell_list.Length; i++)
                                        hset_array_exclude.Add(cell_list[i]);
                                }
                                else
                                {
                                    afi_temp.EndCell = range[0];
                                    afi_temp.Rows = 1;
                                    afi_temp.Cols = 1;
                                }

                                dict_array_formulas.Add(cell_id, afi_temp);
                                reader.Read();
                                reader_value = reader.Value.Replace("$", "");
                                dict_formulas.Add(cell_id, reader_value);
                                continue;
                            }
                        }
                        else if (is_shared_string)
                        {
                            reader.Read();
                            dict_string_constants.Add(cell_id, dict_shared_strings[reader.Value]);
                            continue;
                        }

                        reader.Read();
                        reader_value = reader.Value.Replace("$", "");

                        if (reader_value == "" || reader_value[0] == '#')
                            continue;

                        if (is_formula)
                            dict_formulas.Add(cell_id, reader_value);
                        else
                            dict_constants.Add(cell_id, float.Parse(reader_value));
                    }
                }
            }

            return "";
        }

        //TODO: add function here which may be called in a separate thread to parse a formula.
        //TODO: replace all the dict_formulas.add here with the spawning of a new thread to add parsed formulas to Compute_Core.dict_formulas

        /// <summary>
        /// Loads the specified sheet of an Excel file into dictionaries indexed by cells
        /// </summary>
        /// <param name="sheet_number">The number of the excel sheet being loaded</param>
        /// <param name="dict_formulas">The formula dictionary output</param>
        /// <param name="dict_constants">The constants dictionary output</param>
        public string LoadFile(out List<Dictionary<string,string>> dict_formulas, out List<Dictionary<string,float>> dict_constants, out List<Dictionary<string,Tools.ArrayFormulaInfo>> dict_array_formulas, out List<Dictionary<string,string>> dict_string_constants, out Dictionary<string,int> dict_sheet_name_map)
        {
            Dictionary<string, string> dict_shared_strings = new Dictionary<string, string>(StringComparer.Ordinal);

            dict_string_constants = null;

            if (!File.Exists(file_path))
            {
                //Error here for non-existent file
                dict_constants = null;
                dict_formulas = null;
                dict_array_formulas = null;
                dict_string_constants = null;
                dict_sheet_name_map = null;
                return "ERROR: Could not find the specified file";
            }

            string temp_path = Path.GetTempPath() + @"\bob\temp_sheet";

            if (Directory.Exists(temp_path))
                Directory.Delete(temp_path, true);
            
            //unzip file into temp directory
            try
            {
                using(ZipFile zip = ZipFile.Read(file_path))
                {
                    foreach (ZipEntry item in zip)
                        if (item.FileName.Contains("sheet") || item.FileName.Contains("sharedStrings") || item.FileName.Contains("workbook"))
                            item.Extract(temp_path, ExtractExistingFileAction.OverwriteSilently);                        
                }
            }
            catch
            {
                //probably not a zip file
                dict_constants = null;
                dict_formulas = null;
                dict_array_formulas = null;
                dict_string_constants = null;
                dict_sheet_name_map = null;
                return "ERROR: File does not seem to be of the XLSX (Office 2007+) type";
            }

            //populate the sheet map
            using (XmlReader reader = XmlReader.Create(temp_path + @"\xl\workbook.xml"))
            {
                dict_sheet_name_map = new Dictionary<string, int>(8);

                while (reader.ReadToFollowing("sheet"))
                    dict_sheet_name_map.Add(reader["name"].ToUpper(), int.Parse(reader["sheetId"]) - 1);
            }

            //check if any shared strings
            if (File.Exists(temp_path + @"\xl\sharedStrings.xml"))
                using (XmlReader reader = XmlReader.Create(temp_path+@"\xl\sharedStrings.xml"))
                {
                    reader.ReadToNextSibling("sst");
                    dict_shared_strings = new Dictionary<string, string>(int.Parse(reader["uniqueCount"]));

                    int current_id = 0;

                    while (reader.Read())
                    {
                        if (reader.IsStartElement() && reader.Name == "si")
                        {
                            reader.Read();
                            reader.Read();
                            dict_shared_strings.Add(current_id++.ToString(), reader.Value);
                        }
                    }
                }

            int num_sheets = dict_sheet_name_map.Count;

            List<Dictionary<string, string>> par_temp_dict_formulas = new List<Dictionary<string, string>>(num_sheets);
            par_temp_dict_formulas.AddRange(new Dictionary<string, string>[num_sheets]);

            List<Dictionary<string, float>> par_temp_dict_constants = new List<Dictionary<string, float>>(num_sheets);
            par_temp_dict_constants.AddRange(new Dictionary<string, float>[num_sheets]);

            List<Dictionary<string, string>> par_temp_dict_string_constants = new List<Dictionary<string, string>>(dict_sheet_name_map.Count);
            par_temp_dict_string_constants.AddRange(new Dictionary<string, string>[num_sheets]);

            List<Dictionary<string, Tools.ArrayFormulaInfo>> par_temp_dict_array_formulas = new List<Dictionary<string, Tools.ArrayFormulaInfo>>(dict_sheet_name_map.Count);
            par_temp_dict_array_formulas.AddRange(new Dictionary<string, Tools.ArrayFormulaInfo>[num_sheets]);

            Parallel.For(0, num_sheets, i => { 
                Dictionary<string, string> temp_dict_formulas;
                Dictionary<string, float> temp_dict_constants;
                Dictionary<string, string> temp_dict_string_constants;
                Dictionary<string, Tools.ArrayFormulaInfo> temp_dict_array_formulas;
                LoadSheet(i + 1, temp_path, out temp_dict_formulas, out temp_dict_constants, out temp_dict_array_formulas, out temp_dict_string_constants, dict_shared_strings);

                par_temp_dict_formulas[i] = temp_dict_formulas;
                par_temp_dict_constants[i] = temp_dict_constants;
                par_temp_dict_string_constants[i] = temp_dict_string_constants;
                par_temp_dict_array_formulas[i] = temp_dict_array_formulas;
            });
            dict_formulas = par_temp_dict_formulas;
            dict_constants = par_temp_dict_constants;
            dict_string_constants = par_temp_dict_string_constants;
            dict_array_formulas = par_temp_dict_array_formulas;
            return "";
        }
    }
}
