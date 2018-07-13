/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;

namespace ExcelIdea
{
    class FastFileSaver
    {
        private bool insert_formula = false;
        private ConcurrentDictionary<string, string> dict_Formulas; //local formula storage
        private List<List<List<string>>> lst_DepositCells; //[sheet number][row][column]
        private Dictionary<string, int> dict_sheet_name_map;
        private List<ConcurrentDictionary<string, float>> dict_FloatResults;
        private List<Dictionary<string, float>> dict_FloatConstants;
        private List<Dictionary<string, string>> dict_StringConstants;
        private List<ConcurrentDictionary<string, string>> dict_StringResults;
        private string[] sheet_names;

        public FastFileSaver(Dictionary<string, int> dict_sheet_name_map, List<ConcurrentDictionary<string, float>> dict_FloatResults, List<Dictionary<string, float>> dict_FloatConstants, List<Dictionary<string, string>> dict_StringConstants, List<ConcurrentDictionary<string, string>> dict_StringResults)
        {
            this.dict_sheet_name_map = dict_sheet_name_map;
            this.dict_FloatConstants = dict_FloatConstants;
            this.dict_FloatResults = dict_FloatResults;
            this.dict_StringConstants = dict_StringConstants;
            this.dict_StringResults = dict_StringResults;

            sheet_names = new string[dict_sheet_name_map.Count];
            foreach (var item in dict_sheet_name_map)
                sheet_names[item.Value] = item.Key;
            lst_DepositCells = new List<List<List<string>>>(dict_sheet_name_map.Count);
            lst_DepositCells.AddRange(new List<List<string>>[dict_sheet_name_map.Count]);
            Parallel.For(0, dict_sheet_name_map.Count, i =>
            {
                lst_DepositCells[i] = new List<List<string>>();

                //sort dictionary data by row here
                foreach (var item in dict_FloatResults[i])
                {
                    Tuple<string, int> coords = Tools.SplitCoordinates(item.Key);
                    while (lst_DepositCells[i].Count < coords.Item2)
                        lst_DepositCells[i].Add(new List<string>());
                    lst_DepositCells[i][coords.Item2 - 1].Add(item.Key);
                }
                foreach (var item in dict_FloatConstants[i])
                {
                    Tuple<string, int> coords = Tools.SplitCoordinates(item.Key);
                    while (lst_DepositCells[i].Count < coords.Item2)
                        lst_DepositCells[i].Add(new List<string>());
                    lst_DepositCells[i][coords.Item2 - 1].Add(item.Key);
                }
                foreach (var item in dict_StringConstants[i])
                {
                    Tuple<string, int> coords = Tools.SplitCoordinates(item.Key);
                    while (lst_DepositCells[i].Count < coords.Item2)
                        lst_DepositCells[i].Add(new List<string>());
                    lst_DepositCells[i][coords.Item2 - 1].Add(item.Key);
                }
                foreach (var item in dict_StringResults[i])
                {
                    Tuple<string, int> coords = Tools.SplitCoordinates(item.Key);
                    while (lst_DepositCells[i].Count < coords.Item2)
                        lst_DepositCells[i].Add(new List<string>());
                    lst_DepositCells[i][coords.Item2 - 1].Add(item.Key);
                }
            });
        }

        public void SaveFile(string path)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wbp = spreadsheet.AddWorkbookPart();
                wbp.Workbook = new Workbook();
                Sheets sheets = wbp.Workbook.AppendChild<Sheets>(new Sheets());

                Parallel.For(0, dict_sheet_name_map.Count, sheet_num =>
                {
                    WorksheetPart wsp;
                    float tmp_fltvalue;
                    string tmp_strvalue = "";
                    Cell tmp_cell;
                    lock ("addition")
                    {
                        wsp = wbp.AddNewPart<WorksheetPart>();
                        Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(wsp), SheetId = (UInt32)(sheet_num + 1), Name = sheet_names[sheet_num] };
                        sheets.Append(sheet);
                    }

                    Parallel.For(0, lst_DepositCells[sheet_num].Count, i =>
                    {
                        lst_DepositCells[sheet_num][i].Sort((x, y) => x.Length == y.Length ? string.Compare(x, y) : x.Length > y.Length ? 1 : -1);
                    });

                    OpenXmlWriter writer = OpenXmlWriter.Create(wsp);
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    for (int row = 0; row < lst_DepositCells[sheet_num].Count; row++)
                    {
                        writer.WriteStartElement(new Row());
                        for (int col = 0; col < lst_DepositCells[sheet_num][row].Count; col++)
                        {
                            string tmp_address = lst_DepositCells[sheet_num][row][col];
                            bool is_string = false;

                            if (dict_FloatResults[sheet_num].TryGetValue(tmp_address, out tmp_fltvalue))
                                is_string = false;
                            else if (dict_StringResults[sheet_num].TryGetValue(tmp_address, out tmp_strvalue))
                                is_string = true;
                            else if (dict_FloatConstants[sheet_num].TryGetValue(tmp_address, out tmp_fltvalue))
                                is_string = false;
                            else if (dict_StringConstants[sheet_num].TryGetValue(tmp_address, out tmp_strvalue))
                                is_string = true;
                            
                            if (is_string)
                                if (insert_formula)
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.String, CellFormula = new CellFormula(dict_Formulas[lst_DepositCells[sheet_num][row][col]]) };
                                else
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.String };
                            else
                                if (insert_formula)
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.Number, CellFormula = new CellFormula(dict_Formulas[lst_DepositCells[sheet_num][row][col]]) };
                                else
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.Number };

                            if (is_string)
                                tmp_cell.CellValue = new CellValue(tmp_strvalue);
                            else
                                tmp_cell.CellValue = new CellValue(tmp_fltvalue.ToString());

                            writer.WriteElement(tmp_cell);
                        }
                        writer.WriteEndElement(); //row end
                    }
                    writer.WriteEndElement(); //sheetdata
                    writer.WriteEndElement(); //worksheet
                    writer.Close();
                });
            }
        }

        ~FastFileSaver()
        {
            if (insert_formula)
                dict_Formulas.Clear();
            lst_DepositCells.Clear();
        }
    }
}
