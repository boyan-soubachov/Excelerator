/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelIdea
{
    class ComputeCore
    {
        private string str_error_message = "";
        public static List<ConcurrentDictionary<string, List<Token>>> dict_formulas;
        private List<ConcurrentDictionary<string, float>> dict_results;
        private List<ConcurrentDictionary<string, string>> dict_string_results;
        private List<ConcurrentDictionary<string, string>> dict_derivation_error;
        private Dictionary<string, int> dict_sheet_name_map;
        private List<Dictionary<string, float>> dict_loader_constants;
        private List<Dictionary<string, string>> dict_loader_formulas;
        private List<Dictionary<string, Tools.ArrayFormulaInfo>> dict_loader_array_formulas;
        private List<Dictionary<string, string>> dict_loader_string_constants;
        private static object semaphore = new object();
        private static CUDAProcessor CUDA_Core;
        private Thread cuda_thread;
        private bool load_error = false;
        private bool computed = false;

        /// <summary>
        /// Starts the compute core with no source specified, CUDA thread starts in the background
        /// </summary>
        public ComputeCore(bool init_cuda = true)
        {
            Console.WriteLine(":: ExcelIdea Advanced Computation Engine starting ::");
            Console.WriteLine("[64-bit OS: " + Environment.Is64BitOperatingSystem + ", 64-bit process: " + Environment.Is64BitProcess + ", Processors: " + Environment.ProcessorCount +
                ", High-res timer: " + Stopwatch.IsHighResolution + "]");

            if (init_cuda)
                //initialise CUDA in the background if not started
                if (CUDA_Core == null || !CUDA_Core.core_up)
                {
                    Console.WriteLine("=> Starting CUDA cores");
                    CUDA_Core = new CUDAProcessor();
                    cuda_thread = new Thread(new ThreadStart(CUDA_Core.Init));
                    cuda_thread.Start();
                }
            load_error = true;
        }

        public CUDAProcessor GetCUDAProcessor()
        {
            if (cuda_thread != null)
                cuda_thread.Join();
            CUDA_Core.SetContext();

            return CUDA_Core;
        }

        public void SetCUDAProcessor(CUDAProcessor cuda_in)
        {
            CUDA_Core = cuda_in;
            CUDA_Core.SetContext();
        }

        /// <summary>
        /// Initialise the Advanced Computation Engine with source data from an Excel file.
        /// </summary>
        /// <param name="sourcePath">The path of the source data Excel file</param>
        public ComputeCore(string sourcePath)
        {
            Stopwatch sw = new Stopwatch();

            Console.WriteLine(":: ExcelIdea Advanced Computation Engine starting ::");
            Console.WriteLine("[64-bit OS: " + Environment.Is64BitOperatingSystem + ", 64-bit process: " + Environment.Is64BitProcess + ", Processors: " + Environment.ProcessorCount +
                ", High-res timer: " + Stopwatch.IsHighResolution + "]");

            //initialise CUDA in the background if not started
            if (CUDA_Core == null || !CUDA_Core.core_up)
            {
                Console.WriteLine("=> Starting CUDA cores");
                CUDA_Core = new CUDAProcessor();
                cuda_thread = new Thread(new ThreadStart(CUDA_Core.Init));
                cuda_thread.Start();
            }

            Console.WriteLine("=> Reading Excel file");
            sw.Start();

            FileLoader loader = new FileLoader(sourcePath);
            str_error_message = loader.LoadFile(out dict_loader_formulas, out dict_loader_constants, out dict_loader_array_formulas, out dict_loader_string_constants, out dict_sheet_name_map);
            if (dict_loader_constants == null && dict_loader_formulas == null)
            {
                Console.WriteLine("!! Error: Could not open the specified file");
                if (str_error_message != "")
                    Console.WriteLine("\t Reason: " + str_error_message);
                Console.ReadLine();
                load_error = true;
                return;
            }
            sw.Stop();
            Console.WriteLine("\t --Loader time: " + sw.Elapsed.TotalMilliseconds + " ms");
            load_error = false;
        }

        /// <summary>
        /// Initialise the Compute Core with source data coming from a named pipe message
        /// </summary>
        /// <param name="sourceData">The named pipe message containing workbook data</param>
        public ComputeCore(PipeMessage sourceData)
        {
            if (sourceData.dict_sheet_name_map.Count == 0)
            {
                load_error = true;
                return;
            }
            this.dict_loader_array_formulas = sourceData.dict_loader_array_formulas;
            this.dict_loader_constants = sourceData.dict_loader_constants;
            this.dict_loader_formulas = sourceData.dict_loader_formulas;
            this.dict_loader_string_constants = sourceData.dict_loader_string_constants;
            this.dict_sheet_name_map = sourceData.dict_sheet_name_map;
        }

        public void LoadFromFile(string sourcePath)
        {
            Console.WriteLine("=> Reading Excel file");
            Stopwatch sw = new Stopwatch();
            sw.Start();

            FileLoader loader = new FileLoader(sourcePath);
            str_error_message = loader.LoadFile(out dict_loader_formulas, out dict_loader_constants, out dict_loader_array_formulas, out dict_loader_string_constants, out dict_sheet_name_map);
            if (dict_loader_constants == null && dict_loader_formulas == null)
            {
                Console.WriteLine("!! Error: Could not open the specified file");
                if (str_error_message != "")
                    Console.WriteLine("\t Reason: " + str_error_message);
                Console.ReadLine();
                load_error = true;
                return;
            }
            sw.Stop();
            Console.WriteLine("\t --Loader time: " + sw.Elapsed.TotalMilliseconds + " ms");
            load_error = false;
        }

        public bool SaveToFile(string filePath)
        {
            if (!computed)
            {
                Console.WriteLine("=! Error, trying to save an uncomputed file");
                return false;
            }
            Stopwatch sw = new Stopwatch();

            #region Save spreadsheet
            Console.WriteLine("=> Saving to file");
            sw.Reset();
            sw.Start();

            FastFileSaver sheet_saver = new FastFileSaver(dict_sheet_name_map, dict_results, dict_loader_constants, dict_loader_string_constants, dict_string_results);
            sheet_saver.SaveFile(filePath);

            sw.Stop();
            Console.WriteLine("\t --Saver time: " + sw.Elapsed.TotalMilliseconds + " ms");
            #endregion

            return true;
        }

        public bool StartCompute()
        {
            if (load_error) return false;
            #region Global repository initialisation
            RangeRepository.RangeRepositoryStore = new ConcurrentDictionary<string, RangeStruct>(StringComparer.Ordinal);
            #endregion

            Stopwatch sw = new Stopwatch();           

            #region Formula parser
            Console.WriteLine("=> Parsing file formulas");
            sw.Reset();
            sw.Start();
            dict_formulas = new List<ConcurrentDictionary<string, List<Token>>>(dict_sheet_name_map.Count);
            for (int i = 0; i < dict_sheet_name_map.Count; i++)
                dict_formulas.Add(new ConcurrentDictionary<string, List<Token>>(Environment.ProcessorCount, dict_loader_formulas.Count * 3, StringComparer.Ordinal)); //over-allocate by 3 times to reduce hash-collision probability
            FormulaParser parser = new FormulaParser();
            Parallel.ForEach(dict_sheet_name_map, sheet_item => //faster when parallelised
            {
                List<string> formula_list = dict_loader_formulas[sheet_item.Value].Keys.ToList<string>();

                Parallel.For(0, dict_loader_formulas[sheet_item.Value].Count, i => //This is faster with Parallel.for
                {
                    dict_formulas[sheet_item.Value].TryAdd(formula_list[i], parser.Parse_Tokens(dict_loader_formulas[sheet_item.Value][formula_list[i]]));
                });
            });

            Console.WriteLine("\t --Parser time: " + sw.Elapsed.TotalMilliseconds + " ms");
            #endregion

            #region Scheduler
            Console.WriteLine("=> Building calculation structure");
            sw.Reset();
            sw.Start();
            Scheduler formula_scheduler = new Scheduler();
            List<List<CellRef>> list_ProcessingQueue = new List<List<CellRef>>(1024); //CellRef stores sheet id and cell name
            List<Dictionary<string, int>> dict_CalculationLevels = new List<Dictionary<string, int>>(dict_sheet_name_map.Count);
            dict_results = new List<ConcurrentDictionary<string, float>>(dict_sheet_name_map.Count);
            dict_string_results = new List<ConcurrentDictionary<string, string>>(dict_sheet_name_map.Count);
            bool scheduler_addition;
            for (int i = 0; i < dict_sheet_name_map.Count; i++)
            {
                dict_results.Add(new ConcurrentDictionary<string, float>(Environment.ProcessorCount, dict_loader_formulas.Count * 3, StringComparer.Ordinal));
                dict_string_results.Add(new ConcurrentDictionary<string, string>(Environment.ProcessorCount, dict_loader_formulas.Count * 3, StringComparer.Ordinal));
                dict_CalculationLevels.Add(new Dictionary<string, int>(64, StringComparer.Ordinal));
            }

            do
            {
                scheduler_addition = formula_scheduler.Assign_Calculation_Levels(dict_formulas, ref dict_CalculationLevels, ref list_ProcessingQueue, dict_loader_constants, dict_sheet_name_map);
            } while (scheduler_addition);

            //if (!scheduler_addition && dict_CalculationLevels.Count < dict_formulas.Count)
            //{
            //    //error here for circularity

            //}
            sw.Stop();
            Console.WriteLine("\t --Scheduler time: " + sw.Elapsed.TotalMilliseconds + " ms");
            #endregion

            #region Calculation
            if (cuda_thread != null)
                cuda_thread.Join();
            CUDA_Core.SetContext();
            Console.WriteLine("=> Calculating sheets");
            sw.Reset();
            sw.Start();

            //load all constants into result set
            dict_derivation_error = new List<ConcurrentDictionary<string, string>>();
            for (int i = 0; i < dict_sheet_name_map.Count; i++)
                dict_derivation_error.Add(new ConcurrentDictionary<string, string>());

            //calculate formulas
            for (int current_level = 1; current_level < list_ProcessingQueue.Count; current_level++) //for each level
            {
                //build list of signatures for current level
                ConcurrentDictionary<string, List<CellRef>> dict_levelSignatures = new ConcurrentDictionary<string, List<CellRef>>();
                for (int i = 0; i < list_ProcessingQueue[current_level].Count; i++) //faster without parallelisation and multithreading
                {
                    CellRef cell = list_ProcessingQueue[current_level][i];
                    string signature = formula_scheduler.FormulaSignature(dict_formulas[cell.Sheet_Number][cell.Cell_Ref]);

                    if (dict_levelSignatures.ContainsKey(signature)) //signature already in dictionary, add it to the list
                        dict_levelSignatures[signature].Add(cell);
                    else //no such signature in dictionary
                    {
                        List<CellRef> temp_list = new List<CellRef>();
                        temp_list.Add(cell);
                        dict_levelSignatures.TryAdd(signature, temp_list); //store in dictionary
                    }
                }

                //for each signature. NOTE: These can be parallelised since all calculations on the same level are independent of each other.
                Parallel.ForEach(dict_levelSignatures, signature =>
                {
                    float float_temp;
                    string str_temp;
                    List<Token> current_signature = dict_formulas[signature.Value[0].Sheet_Number][signature.Value[0].Cell_Ref];
                    int signature_parallels = signature.Value.Count;
                    List<List<ResultStruct>> parameters = new List<List<ResultStruct>>();
                    List<ResultStruct> results = new List<ResultStruct>();
                    bool[] error_markers = new bool[signature_parallels];
                    int[] sheet_scopes = null;

                    //for each token
                    for (int i = 0; i < current_signature.Count; i++)
                    {
                        Token current_token = current_signature[i];

                        switch (current_token.Type)
                        {
                            case Token_Type.SheetReferenceStart:
                                sheet_scopes = new int[signature_parallels];
                                Parallel.For(0, signature_parallels, j =>
                                {  //parallel is faster here.
                                    sheet_scopes[j] = dict_sheet_name_map[dict_formulas[signature.Value[j].Sheet_Number][signature.Value[j].Cell_Ref][i].GetStringValue()];
                                });
                                break;
                            case Token_Type.SheetReferenceEnd:
                                sheet_scopes = null;
                                break;
                            case Token_Type.Constant: //if constant or cell reference, add to map and array for gpu
                                //add new dimension to the parameters list
                                parameters.Add(new List<ResultStruct>(signature_parallels));

                                for (int j = 0; j < signature_parallels; j++) //attempted parallelising, much slower
                                {
                                    if (error_markers[j]) continue; //if already an error, don't bother
                                    List<Token> parallel_formula = dict_formulas[signature.Value[j].Sheet_Number][signature.Value[j].Cell_Ref];
                                    //add as row to the new list of parameters
                                    ResultStruct temp_desc = new ResultStruct();
                                    temp_desc.Value = parallel_formula[i].GetNumericValue();
                                    temp_desc.Type = ResultEnum.Unit | ResultEnum.Float;
                                    temp_desc.Cols = 1;
                                    temp_desc.Rows = 1;
                                    parameters[parameters.Count - 1].Add(temp_desc);
                                }
                                break;

                            case Token_Type.Date:
                                break;

                            case Token_Type.String:
                                parameters.Add(new List<ResultStruct>(signature_parallels));

                                for (int j = 0; j < signature_parallels; j++) //attempted parallelising, much slower
                                {
                                    if (error_markers[j]) continue; //if already an error, don't bother

                                    List<Token> parallel_formula = dict_formulas[signature.Value[j].Sheet_Number][signature.Value[j].Cell_Ref];
                                    //add as row to the new list of parameters
                                    ResultStruct temp_desc = new ResultStruct();
                                    temp_desc.Value = parallel_formula[i].GetStringValue();
                                    temp_desc.Type = ResultEnum.Unit | ResultEnum.String;
                                    temp_desc.Cols = 1;
                                    temp_desc.Rows = 1;
                                    parameters[parameters.Count - 1].Add(temp_desc);
                                }
                                break;

                            case Token_Type.Cell: //if range, add all cells included in range to map and array for gpu
                                parameters.Add(new List<ResultStruct>(signature_parallels));
                                for (int j = 0; j < signature_parallels; j++)
                                {
                                    List<Token> parallel_formula = dict_formulas[signature.Value[j].Sheet_Number][signature.Value[j].Cell_Ref];
                                    ResultStruct temp_desc = new ResultStruct();
                                    temp_desc.Type = ResultEnum.Empty;
                                    temp_desc.Cols = 1;
                                    temp_desc.Rows = 1;
                                    int sheet = sheet_scopes == null ? signature.Value[j].Sheet_Number : sheet_scopes[j];

                                    if (dict_loader_constants[sheet].TryGetValue(parallel_formula[i].GetStringValue(), out float_temp))
                                    {
                                        temp_desc.Type = ResultEnum.Unit | ResultEnum.Float;
                                        temp_desc.Value = float_temp;
                                    }
                                    else if (dict_loader_string_constants[sheet].TryGetValue(parallel_formula[i].GetStringValue(), out str_temp))
                                    {
                                        temp_desc.Type = ResultEnum.Unit | ResultEnum.String;
                                        temp_desc.Value = str_temp;
                                    }
                                    else if (dict_results[sheet].TryGetValue(parallel_formula[i].GetStringValue(), out float_temp))
                                    {
                                        temp_desc.Type = ResultEnum.Unit | ResultEnum.Float;
                                        temp_desc.Value = float_temp;
                                    }
                                    else if (dict_string_results[sheet].TryGetValue(parallel_formula[i].GetStringValue(), out str_temp))
                                    {
                                        //NOTE: Excel behaviour tries to convert to a number when referencing a cell stored as string
                                        float flt_temp;
                                        if (float.TryParse(str_temp, out flt_temp))
                                        {
                                            temp_desc.Type = ResultEnum.Unit | ResultEnum.Float;
                                            temp_desc.Value = flt_temp;
                                        }
                                        else
                                        {
                                            temp_desc.Type = ResultEnum.Unit | ResultEnum.String;
                                            temp_desc.Value = str_temp;
                                        }
                                    }
                                    parameters[parameters.Count - 1].Add(temp_desc);
                                }
                                break;

                            case Token_Type.Range:
                                parameters.Add(new List<ResultStruct>(signature_parallels));
                                parameters[parameters.Count - 1].AddRange(new ResultStruct[signature_parallels]);
                                for (int j = 0; j < signature_parallels; j++)
                                {
                                    //if (dict_derivation_error.ContainsKey(signature.Value[j]))
                                    //{
                                    //    error_markers[j] = true;
                                    //    continue;
                                    //}

                                    List<Token> parallel_formula = dict_formulas[signature.Value[j].Sheet_Number][signature.Value[j].Cell_Ref];
                                    int sheet = sheet_scopes == null ? signature.Value[j].Sheet_Number : sheet_scopes[j];

                                    ResultStruct temp_desc = new ResultStruct();
                                    Tuple<int, int> r_size = Tools.RangeSize(parallel_formula[i - 2].GetStringValue(), parallel_formula[i - 1].GetStringValue());
                                    temp_desc.Rows = r_size.Item1;
                                    temp_desc.Cols = r_size.Item2;
                                    if (dict_string_results[sheet].ContainsKey(parallel_formula[i - 2].GetStringValue()) || dict_loader_string_constants[sheet].ContainsKey(parallel_formula[i - 2].GetStringValue()))
                                        temp_desc.Type |= ResultEnum.String;
                                    else
                                        temp_desc.Type |= ResultEnum.Float;

                                    if (RangeRepository.RangeRepositoryStore.ContainsKey(sheet.ToString() + '!' + parallel_formula[i - 2].GetStringValue() + ':' + parallel_formula[i - 1].GetStringValue()))
                                    {
                                        temp_desc.Type |= ResultEnum.Mapped_Range;
                                        temp_desc.Value = sheet.ToString() + '!' + parallel_formula[i - 2].GetStringValue() + ':' + parallel_formula[i - 1].GetStringValue();
                                    }
                                    else
                                    {
                                        //add whole range exclusive of start (added by the preceding cell/constant tokens)
                                        string[] range = Tools.RangeCells(parallel_formula[i - 2].GetStringValue(), parallel_formula[i - 1].GetStringValue());
                                        dynamic range_values;
                                        bool is_string = false;
                                        if ((temp_desc.Type & ResultEnum.String) == ResultEnum.String) //temporary hack, in the future string constants will also be held in loader_constants
                                        {
                                            range_values = new string[range.Length];
                                            is_string = true;
                                        }
                                        else
                                            range_values = new float[range.Length];

                                        Parallel.For(0, range.Length, k =>
                                        {
                                            if (is_string)
                                            {
                                                string par_string_temp;
                                                if (dict_loader_string_constants[sheet].TryGetValue(range[k], out par_string_temp) || dict_string_results[sheet].TryGetValue(range[k], out par_string_temp))
                                                    range_values[k] = par_string_temp;
                                                else
                                                    range_values[k] = "";
                                            }
                                            else
                                            {
                                                float par_float_temp;
                                                if (dict_loader_constants[sheet].TryGetValue(range[k], out par_float_temp) || dict_results[sheet].TryGetValue(range[k], out par_float_temp))
                                                    range_values[k] = par_float_temp;
                                                else
                                                    range_values[k] = 0;
                                            }
                                        });

                                        if (range.Length > Tools.RangeMapThreshold && !RangeRepository.RangeRepositoryStore.ContainsKey(sheet.ToString() + '!' + parallel_formula[i - 2].GetStringValue() + ':' + parallel_formula[i - 1].GetStringValue()))
                                        {
                                            //add to range repository
                                            RangeStruct rangestore_temp = new RangeStruct();
                                            rangestore_temp.Values = range_values;
                                            //sort and create a map for the range
                                            int[] sorted_map = new int[temp_desc.Rows];
                                            dynamic sorting_range;
                                            if (is_string)
                                                sorting_range = new string[temp_desc.Rows];
                                            else
                                                sorting_range = new float[temp_desc.Rows];

                                            Array.Copy(range_values, sorting_range, temp_desc.Rows);
                                            for (int k = 0; k < sorted_map.Length; k++)
                                                sorted_map[k] = k;
                                            Array.Sort(sorting_range, sorted_map, StringComparer.Ordinal);
                                            rangestore_temp.Sorted_Map = sorted_map;
                                            RangeRepository.RangeRepositoryStore.TryAdd(sheet.ToString() + '!' + parallel_formula[i - 2].GetStringValue() + ':' + parallel_formula[i - 1].GetStringValue(), rangestore_temp);
                                            temp_desc.Value = sheet.ToString() + '!' + parallel_formula[i - 2].GetStringValue() + ':' + parallel_formula[i - 1].GetStringValue();
                                            temp_desc.Type |= ResultEnum.Mapped_Range;
                                        }
                                        else
                                        {
                                            temp_desc.Value = range_values;
                                            temp_desc.Type |= ResultEnum.Array;
                                        }
                                    }
                                    parameters[parameters.Count - 1][j] = temp_desc;
                                }
                                parameters.RemoveRange(parameters.Count - 3, 2);
                                break;
                            case Token_Type.GreaterThan:
                            case Token_Type.LessThan: //return "TRUE" or "FALSE"
                            case Token_Type.Exponent:
                            case Token_Type.Divide:
                            case Token_Type.Multiply:
                            case Token_Type.Subtract:
                            case Token_Type.Add: //if function or op, execute on gpu given maps and arrays
                                results = CUDA_Core.Compute(current_token.Operator_String, parameters, signature_parallels, 2);
                                parameters.RemoveRange(parameters.Count - 2, 2);
                                parameters.Add(results);
                                break;
                            case Token_Type.Function: //if function or op, execute on gpu given maps and arrays
                                results = CUDA_Core.Compute(current_token.Operator_String, parameters, signature_parallels, current_token.Function_Arguments);
                                //depending on function, remove the last n parameters
                                parameters.RemoveRange(parameters.Count - current_token.Function_Arguments, current_token.Function_Arguments);
                                parameters.Add(results);
                                break;
                            default:
                                break;
                        }
                    }

                    if (results != null)
                    {
                        if (results.Count == 0 && parameters.Count == 1) //special case, the formula didn't involve any calculations, e.g. =12 or ="ABC"
                            results = parameters[0];
                        Parallel.For(0, signature_parallels, sp =>
                        {
                            int dest_res_rows = 1, dest_res_cols = 1;
                            string[] cells;
                            ResultStruct res = (ResultStruct)results[sp];
                            Tools.ArrayFormulaInfo afi_temp;

                            if (dict_loader_array_formulas[signature.Value[sp].Sheet_Number].TryGetValue(signature.Value[sp].Cell_Ref, out afi_temp)) //if it's an array formula
                            {
                                cells = Tools.RangeCells(afi_temp.StartCell, afi_temp.EndCell);
                                dest_res_rows = afi_temp.Rows;
                                dest_res_cols = afi_temp.Cols;
                            }
                            else
                            {
                                cells = new string[1];
                                cells[0] = signature.Value[sp].Cell_Ref;
                            }

                            if (cells.Length > 1) //result may be 1 element but answer area may be more
                            {
                                int source_res_total = res.Rows * res.Cols;
                                if (source_res_total == 1) //if result is single element
                                    //populate all result cells with that same value
                                    Parallel.For(0, dest_res_cols, i =>
                                    {
                                        for (int j = 0; j < dest_res_rows; j++)
                                            if ((res.Type & ResultEnum.Float) == ResultEnum.Float)
                                                dict_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (float)res.Value);
                                            else
                                                dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (string)res.Value);
                                    });
                                else if (res.Cols == 1) //if single columned
                                {
                                    //apply to all columns the same row
                                    Parallel.For(0, dest_res_cols, i =>
                                    {
                                        for (int j = 0; j < dest_res_rows; j++)
                                            if (j >= res.Rows)
                                                dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], "#N/A");
                                            else
                                                if ((res.Type & ResultEnum.Float) == ResultEnum.Float)
                                                    dict_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (res.Value as float[])[(i * res.Rows + j) % (source_res_total)]);
                                                else
                                                    dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (res.Value as string[])[(i * res.Rows + j) % (source_res_total)]);
                                    });
                                }
                                else if (res.Rows == 1) //else if single rowed
                                {
                                    //apply to all rows the same column
                                    Parallel.For(0, dest_res_cols, i =>
                                    {
                                        for (int j = 0; j < dest_res_rows; j++)
                                            if (i >= res.Cols)
                                                dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], "#N/A");
                                            else
                                                if ((res.Type & ResultEnum.Float) == ResultEnum.Float)
                                                    dict_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (res.Value as float[])[(i * res.Rows + j) % (source_res_total)]);
                                                else
                                                    dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (res.Value as string[])[(i * res.Rows + j) % (source_res_total)]);
                                    });
                                }
                                else //else (multiple rowed & columned)
                                {
                                    //put the results in the destination
                                    Parallel.For(0, dest_res_cols, i =>
                                    {
                                        for (int j = 0; j < dest_res_rows; j++)
                                            //if the destination cell is outside the result bounds
                                            if (j >= res.Rows || i >= res.Cols)
                                                //set value to #N/A
                                                dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], "#N/A");
                                            else
                                                if ((res.Type & ResultEnum.Float) == ResultEnum.Float)
                                                    dict_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (res.Value as float[])[i * res.Rows + j]);
                                                else
                                                    dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[i * dest_res_rows + j], (res.Value as string[])[i * res.Rows + j]);
                                    });
                                }
                            }
                            else
                            {
                                if ((res.Type & ResultEnum.Float) == ResultEnum.Float)
                                    dict_results[signature.Value[sp].Sheet_Number].TryAdd(cells[0], (float)res.Value); //cells[0] because there's only 1 element (not an array)
                                else if ((res.Type & ResultEnum.String) == ResultEnum.String)
                                    dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[0], (string)res.Value);
                                else if ((res.Type & ResultEnum.Empty) == ResultEnum.Empty)
                                    dict_results[signature.Value[sp].Sheet_Number].TryAdd(cells[0], (float)0);
                                else
                                    dict_string_results[signature.Value[sp].Sheet_Number].TryAdd(cells[0], "#ERROR!");

                            }
                        });
                    }
                    else
                    {
                        //error here, null results
                        throw new SystemException("Result value missing when calculating cell [" + signature.Value[0] + "]");
                    }
                });
            }

            sw.Stop();
            Console.WriteLine("\t --Sheet calculation time: " + sw.Elapsed.TotalMilliseconds + " ms");
            #endregion

            #region Error report
            Parallel.For(0, dict_derivation_error.Count, i =>
            {
                if (dict_derivation_error[i].Count > 0)
                {
                    Console.WriteLine(":: Error report, sheet " + i + " ::");
                    foreach (KeyValuePair<string, string> item in dict_derivation_error[i])
                    {
                        Console.WriteLine("Cell: " + item.Key);
                        Console.WriteLine("\t -> " + item.Value);
                    }
                }
            });
            #endregion

            CUDA_Core.Dispose();
            Console.WriteLine(":: ExcelIdea ACE done ::");
            computed = true;
            return true;
        }
    }
}
