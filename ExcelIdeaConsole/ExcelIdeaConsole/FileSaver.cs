/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet; //NOTE: This OpenXML toolkit is VERY slow. We need to be clever about this.
using System.Diagnostics;

namespace ExcelIdea
{
    class FileSaver
    {
        private bool insert_formula = false;
        private ConcurrentDictionary<string, string> dict_Formulas; //local formula storage
        private List<List<List<string>>> lst_DepositCells; //[sheet number][column][row]
        private Dictionary<string, int> dict_sheet_name_map;
        private List<ConcurrentDictionary<string, float>> dict_FloatResults;
        private List<Dictionary<string,float>> dict_FloatConstants;
        private List<Dictionary<string,string>> dict_StringConstants;
        private List<ConcurrentDictionary<string, string>> dict_StringResults;
        private string[] sheet_names;

        public FileSaver(Dictionary<string,int> dict_sheet_name_map, List<ConcurrentDictionary<string,float>> dict_FloatResults, List<Dictionary<string,float>> dict_FloatConstants, List<Dictionary<string,string>> dict_StringConstants, List<ConcurrentDictionary<string,string>> dict_StringResults)
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
            object lock_sheet_addition = new object();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook, true))
            {
                WorkbookPart workbookPart1 = package.AddWorkbookPart();
                GenerateWorkbookPart1Content(workbookPart1);
                Parallel.For(0, dict_sheet_name_map.Count, sheet =>
                {
                    WorksheetPart worksheetPart1;
                    lock (lock_sheet_addition)
                    {
                        worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("sh_" + sheet_names[sheet]);
                    }

                    //fill data
                    Worksheet worksheet1 = new Worksheet();
                    worksheet1.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                    SheetData sheetData1 = new SheetData();

                    Parallel.For(0, lst_DepositCells[sheet].Count, i =>
                    {
                        lst_DepositCells[sheet][i].Sort((x, y) => x.Length == y.Length ? string.Compare(x, y) : x.Length > y.Length ? 1 : -1);
                    });

                    string tmp_address;
                    string tmp_strvalue = "";
                    float tmp_fltvalue;
                    bool is_string = false;
                    Cell tmp_cell;
                    //for each row
                    for (int row = 0; row < lst_DepositCells[sheet].Count; row++)
                    {
                        Row tmp_row = new Row() { RowIndex = (UInt32)(row + 1)  };
                        for (int col = 0; col < lst_DepositCells[sheet][row].Count; col++)
                        {
                            //insert cell at desired column
                            tmp_address = lst_DepositCells[sheet][row][col];
                            if (dict_FloatResults[sheet].TryGetValue(tmp_address, out tmp_fltvalue))
                                is_string = false;
                            else if (dict_StringResults[sheet].TryGetValue(tmp_address, out tmp_strvalue))
                                is_string = true;
                            else if (dict_FloatConstants[sheet].TryGetValue(tmp_address, out tmp_fltvalue))
                                is_string = false;
                            else if (dict_StringConstants[sheet].TryGetValue(tmp_address, out tmp_strvalue))
                                is_string = true;

                            if (is_string)
                                if (insert_formula)
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.String, CellFormula = new CellFormula(dict_Formulas[lst_DepositCells[sheet][row][col]]) };
                                else
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.String };
                            else
                                if (insert_formula)
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.Number, CellFormula = new CellFormula(dict_Formulas[lst_DepositCells[sheet][row][col]]) };
                                else
                                    tmp_cell = new Cell() { CellReference = tmp_address, DataType = CellValues.Number };

                            if (is_string)
                                tmp_cell.CellValue = new CellValue(tmp_strvalue);
                            else
                                tmp_cell.CellValue = new CellValue(tmp_fltvalue.ToString());
                            tmp_row.Append(tmp_cell);
                        }
                        sheetData1.Append(tmp_row);
                    }
                    worksheet1.Append(sheetData1);
                    worksheetPart1.Worksheet = worksheet1;
                });
            }
        }

        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            Sheets sheets1 = new Sheets();

            for (int i = 0; i < sheet_names.Length; i++)
            {
                Sheet sheet1 = new Sheet() { Name = sheet_names[i], SheetId = Convert.ToUInt32(i + 1), Id = "sh_" + sheet_names[i] };
                sheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                sheets1.Append(sheet1);
            }
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }

        public void AttachFormulas(ConcurrentDictionary<string,string> dict_Formulas)
        {
            if (insert_formula)
                dict_Formulas.Clear();

            this.dict_Formulas = dict_Formulas;
            insert_formula = true;
        }

        ~FileSaver()
        {
            if (insert_formula)
                dict_Formulas.Clear();
            lst_DepositCells.Clear();
        }
    }
}
