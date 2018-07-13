/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace ExcelIdea
{
    class Scheduler
    {
        /// <summary>
        /// Assigns formula calculation levels based on dependencies and sorts them for sequential calculation processing
        /// </summary>
        /// <param name="dict_formulas">The dictionary of formulas indexed by cell</param>
        /// <param name="dict_CalculationLevels">Returns a dictionary of calculation levels indexed by cell</param>
        /// <param name="list_ProcessingQueues">Returns a list of cells indexed by calculation levels</param>
        public bool Assign_Calculation_Levels(List<ConcurrentDictionary<string, List<Token>>> dict_formulas, ref List<Dictionary<string, int>> dict_CalculationLevels, ref List<List<CellRef>> list_ProcessingQueues, List<Dictionary<string, float>> dict_loader_constants, Dictionary<string,int> dict_sheet_name_map)
        {
            bool satisfied = true;
            bool changed = false;
            int diff_sheet = -1;
            int i_holder = 0;
            int level, temp_int;
            string start_cell, end_cell;
            Dictionary<string, int> checked_ranges = new Dictionary<string, int>();

            for (int i = 0; i < dict_formulas.Count; i++)
            {
                foreach (KeyValuePair<string, List<Token>> cell in dict_formulas[i])
                {
                    if (dict_CalculationLevels[i].ContainsKey(cell.Key)) //if already in list, get out
                        continue;
                    satisfied = true;
                    level = 1;

                    //for each item in token queue
                    for (int j = 0; j < cell.Value.Count; j++)
                    {
                        Token item = cell.Value.ElementAt<Token>(j);
                        switch (item.Type)
                        {
                            case Token_Type.Null:
                                level = -1;
                                break;
                            case Token_Type.Range:
                                start_cell = cell.Value.ElementAt<Token>(j - 2).GetStringValue();
                                end_cell = cell.Value.ElementAt<Token>(j - 1).GetStringValue();
                                if (checked_ranges.TryGetValue(start_cell + ':' + end_cell, out temp_int))
                                {
                                    level = Math.Max(temp_int, level);
                                    break;
                                }
                                string[] range_cells = Tools.RangeCells(start_cell, end_cell);

                                foreach (string ref_cell in range_cells)
                                {
                                    if (!dict_formulas[i].ContainsKey(ref_cell) && dict_CalculationLevels[i].ContainsKey(ref_cell)) //if the sub key still needs to be processed
                                    {
                                        satisfied = false;
                                        break;
                                    }
                                    else if (dict_CalculationLevels[i].TryGetValue(ref_cell, out temp_int)) //if the subkey has been processed
                                        level = Math.Max(level, temp_int + 1);
                                    else //if the sub key points to nothing or constant
                                        level = Math.Max(level, 1);
                                }
                                if (satisfied)
                                    checked_ranges.Add(start_cell + ':' + end_cell, level);
                                break;
                            case Token_Type.Cell:
                                if (dict_formulas[i].ContainsKey(item.GetStringValue()) && !dict_CalculationLevels[i].ContainsKey(item.GetStringValue())) //if in the formula list and not the levels list, wait for it to be added
                                {
                                    satisfied = false;
                                    break;
                                }
                                else if (dict_CalculationLevels[i].TryGetValue(item.GetStringValue(), out temp_int)) //if in levels list
                                    level = Math.Max(level, temp_int + 1);
                                else //in neither list, reference to empty cell
                                    level = Math.Max(1, level);
                                break;
                            case Token_Type.SheetReferenceStart:
                                diff_sheet = dict_sheet_name_map[item.GetStringValue()];
                                i_holder = i;
                                i = diff_sheet;
                                break;
                            case Token_Type.SheetReferenceEnd:
                                i = i_holder;
                                diff_sheet = -1;
                                break;
                            default:
                                break;
                        }

                        if (!satisfied)
                        {
                            if (diff_sheet > -1)
                            {
                                i = i_holder;
                                diff_sheet = -1;
                            }
                            break;
                        }
                    }

                    if (satisfied)
                    {
                        changed = true;
                        dict_CalculationLevels[i].Add(cell.Key, level);
                        while (list_ProcessingQueues.Count < level + 1)
                            list_ProcessingQueues.Add(new List<CellRef>());
                        list_ProcessingQueues[level].Add(new CellRef(i, cell.Key));
                    }
                }
            }
            return changed;
        }

        public string FormulaSignature(List<Token> formula_in)
        {
            string[] token_array = new string[formula_in.Count];
            StringBuilder sb_result = new StringBuilder();

            for (int i = 0; i < formula_in.Count; i++)
            {
                if (formula_in[i].Operator_String == ":")
                {
                    //token_array[i - 1] = formula_in[i - 1].Operator_String; //parallelises non-specific ranges
                    //token_array[i - 2] = formula_in[i - 2].Operator_String;
                    token_array[i - 1] = formula_in[i - 1].GetStringValue(); //parallelises specific ranges
                    token_array[i - 2] = formula_in[i - 2].GetStringValue();
                }
                token_array[i] = formula_in[i].Operator_String;
            }

            for (int i = 0; i < formula_in.Count; i++)
            {
                if (token_array[i] == "") continue;
                sb_result.Append(token_array[i]);
                sb_result.Append('|');
            }

            return sb_result.ToString();
        }

        
    }
}
