using System;
using System.Collections.Generic;
using System.Linq;
using TrackerEntities;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Globalization;
using System.IO;
using DbLayer;
using Gmutils;

namespace XlsAutoReportWindowsService.XLSX
{
    class XLSX_GetReports
    {
        public static void MakeReport(List<PrepressJob> Jobs, string fileName, DateTime dt)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package, Jobs, fileName, dt);
            }
        }

        public static void MakeWeeklyReport(List<PrepressJob> Jobs, string fileName)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                CreateWeeklyParts(package, Jobs, fileName);
            }
        }

        public static void MakeReportByQuery(TCPExcelSend tes) {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(tes.fileName, SpreadsheetDocumentType.Workbook))
            {
                CreateQueryReportParts(package, tes);
            }
        }

        private static void CreateParts(SpreadsheetDocument document, List<PrepressJob> Jobs, string fileName, DateTime date2)
        {

            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>();
            XLSX_Helpers.GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            SharedStringTablePart sharedStringTablePart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                sharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();//лист с макетами
            worksheetPart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheetPart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheetPart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //добавим стилей
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            XLSX_Helpers.GenerateWorkbookStylesPartContent(workbookStylesPart);
            workbookStylesPart.Stylesheet.Save();

            Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Лист 1" };//имя можно задать любое, единственное ограничение - количество символов не должно превышать 40, хотя в документации указан предел в 255 символов, мистика одним словом
            sheets.Append(sheet);
            SheetDimension sheetDimension = new SheetDimension() { Reference = "A1:Z100000" };

            SheetViews sheetViews = new SheetViews();

            SheetView sheetView = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Pane pane = new Pane() { VerticalSplit = 1D, TopLeftCell = "A2", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };
            Selection selection = new Selection() { Pane = PaneValues.BottomLeft };

            sheetView.Append(pane);
            sheetView.Append(selection);


            // sheetView1.Append(selection1);

            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
            Columns columns = new Columns();
            columns.Append(new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 30D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 25D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)12U, Max = (UInt32Value)12U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)13U, Max = (UInt32Value)13U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)14U, Max = (UInt32Value)14U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)15U, Max = (UInt32Value)15U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)16U, Max = (UInt32Value)16U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)17U, Max = (UInt32Value)17U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)18U, Max = (UInt32Value)18U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)19U, Max = (UInt32Value)19U, Width = 20D, BestFit = true, CustomWidth = true });
            SheetData sheetData = new SheetData();

            string[] headerRow = new string[] { "Номер ТЗ", "Индивидуальный заказ", "Номер редакции", "Вкладка", "Приемщик", "Дата создания", "Время создания", "Дата принятия", "Время принятия", "Конец проверки", "Время проверки", "Время из препресс", "Папка", "Лицо/Оборот", "Скрепка", "Правки лицо", "Правки оборот", "Менеджер", "Статус" };
            string[] headerRow2 = new string[] { "Номер ТЗ", "Приемщик", "Дата создания", "Время создания", "Дата начала проверки", "Время начала проверки", "Дата окончания проверки", "Время окончания проверки" };
            uint rowInd = 1U;
            int cellNum = 1;

            Row rowHead = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
            foreach (string s in headerRow)
            {
                cellNum = XLSX_Helpers.InsertCellToRow(rowHead, cellNum, s, sharedStringTablePart, 3U);
            }
            sheetData.AppendChild(rowHead);
            rowInd++;
            cellNum = 1;
            CultureInfo cult = CultureInfo.CurrentCulture;
            using (Repository rep = new Repository())
            {
                foreach (var job in Jobs)
                {
                    Row row = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
                    int countUnits = 1;
                    if (job.Model.Units.Count() > 1)
                    {
                        foreach (var unit in job.Model.Units.AsEnumerable().Reverse())
                        {
                            row = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
                            if (rowInd % 2 == 0)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 1U);     //номер ТЗ
                                string indTz = "";
                                if (job.OrderType == 2)
                                    indTz = "Да";
                                else
                                    indTz = "Нет";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 1U); //инд заказ
                                if (null != job.num)                                                                                      //редакция
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 1U);
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);

                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countUnits.ToString(), sharedStringTablePart, 1U);//вкладка

                                if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 1U);
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);

                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);//дата создания, она же дата создания тз на проверку
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время создания
                                if (job.DoneDate.HasValue && job.StartDate.HasValue)
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);   // дата принятия
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время принятия в работу
                                    TimeSpan ts = job.DoneDate.Value - job.StartDate.Value;
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                }
                                if (job.HumanWorkEndDate.HasValue && job.StartDate.HasValue)
                                {
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);   //
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.HumanWorkEndDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время окончания проверки
                                    TimeSpan ts = job.HumanWorkEndDate.Value - job.StartDate.Value;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);  //время на проверку
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                }
                                if (job.DoneDate.HasValue)
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время из препресс
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                }
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 1U);   //папка


                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.FaceFile_Id).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.BackFile_Id).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 1U);     //лицо/оборот
                                if (unit.Skrepka == true)
                                {
                                    var modelFileSkr = rep.Select<ModelFile>().First(w => w.Id == unit.FaceFile_Id);
                                    int? countPages = modelFileSkr.PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 1U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 1U);     //Скрепка
                                if (faceFiles.Count() > 0)
                                {
                                    int faceRebukes = faceFiles[0].Rebukes;//правка лицо
                                    if (faceRebukes != 0)
                                    {
                                        string faceEdits = GetFileEdits(faceRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, faceEdits, sharedStringTablePart, 1U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 1U);
                                    }
                                }
                                else
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет лица", sharedStringTablePart, 1U);
                                }
                                if (backFiles.Count() > 0)
                                {
                                    int backRebukes = backFiles[0].Rebukes;//правка оборот
                                    if (backRebukes != 0)
                                    {
                                        string backEdits = GetFileEdits(backRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, backEdits, sharedStringTablePart, 1U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 1U);
                                    }

                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет оборота", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Manager.Name, sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.LastStateRef.Description, sharedStringTablePart, 1U);
                            }
                            else
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 2U);     //номер ТЗ
                                string indTz = "";
                                if (job.OrderType == 2)
                                    indTz = "Да";
                                else
                                    indTz = "Нет";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 2U); //инд заказ
                                if (null != job.num)
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 2U);  //редакция
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countUnits.ToString(), sharedStringTablePart, 2U);//вкладка
                                if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 2U);
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);//дата создания, она же дата создания тз на проверку
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время создания
                                if (job.DoneDate.HasValue && job.StartDate.HasValue)
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);   // дата принятия
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время принятия в работу
                                    TimeSpan ts = job.DoneDate.Value - job.StartDate.Value;
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                }
                                if (job.HumanWorkEndDate.HasValue && job.StartDate.HasValue)
                                {
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);   //
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.HumanWorkEndDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время окончания проверки
                                    TimeSpan ts = job.HumanWorkEndDate.Value - job.StartDate.Value;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 2U);  //время на проверку
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                }
                                if (job.DoneDate.HasValue)
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время из препресс
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                }
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 2U);   //папка
                                                                                                                                     // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 2U);
                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.FaceFile_Id).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.BackFile_Id).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 2U);     //лицо/оборот
                                if (unit.Skrepka == true)
                                {
                                    var modelFileSkr = rep.Select<ModelFile>().First(w => w.Id == unit.FaceFile_Id);
                                    int? countPages = modelFileSkr.PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 2U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 2U);     //Скрепка
                                if (faceFiles.Count() > 0)
                                {
                                    int faceRebukes = faceFiles[0].Rebukes;//правка лицо
                                    if (faceRebukes != 0)
                                    {
                                        string faceEdits = GetFileEdits(faceRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, faceEdits, sharedStringTablePart, 2U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 2U);
                                    }
                                }
                                else
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет лица", sharedStringTablePart, 2U);
                                }
                                if (backFiles.Count() > 0)
                                {
                                    int backRebukes = backFiles[0].Rebukes;//правка оборот
                                    if (backRebukes != 0)
                                    {
                                        string backEdits = GetFileEdits(backRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, backEdits, sharedStringTablePart, 2U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 2U);
                                    }

                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет оборота", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Manager.Name, sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.LastStateRef.Description, sharedStringTablePart, 2U);
                            }

                            sheetData.Append(row);
                            rowInd++;
                            cellNum = 1;
                            countUnits++;
                        }
                        if (countUnits == job.Model.Units.Count())
                            continue;
                    }
                    else
                    {
                        if (rowInd % 2 == 0)
                        {
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 1U);     //номер ТЗ
                            string indTz = "";
                            if (job.OrderType == 2)
                                indTz = "Да";
                            else
                                indTz = "Нет";
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 1U); //инд заказ
                            if (null != job.num)
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 1U);  //редакция
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "1", sharedStringTablePart, 1U);//вкладка
                            if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 1U);
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);//дата создания, она же дата создания тз на проверку
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время создания
                            if (job.DoneDate.HasValue && job.StartDate.HasValue)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);   // дата принятия
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время принятия в работу
                                TimeSpan ts = job.DoneDate.Value - job.StartDate.Value;
                                //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            }
                            if (job.HumanWorkEndDate.HasValue && job.StartDate.HasValue)
                            {
                                //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);   //
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.HumanWorkEndDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время окончания проверки
                                TimeSpan ts = job.HumanWorkEndDate.Value - job.StartDate.Value;
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);  //время на проверку
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            }
                            if (job.DoneDate.HasValue)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 1U);//время из препресс
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 1U);   //папка
                            if (job.Model.Units.Count > 0)
                            {                                                                                                     // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 1U);
                                int? faceId = job.Model.Units[0].FaceFile_Id;
                                int? backId = job.Model.Units[0].BackFile_Id;
                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == faceId).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == backId).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 1U);     //лицо/оборот
                                if (job.Model.Units[0].Skrepka == true)
                                {
                                    int? countPages = faceFiles[0].PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 1U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 1U);     //Скрепка
                                if (faceFiles.Count() > 0)
                                {
                                    int faceRebukes = faceFiles[0].Rebukes;//правка лицо
                                    if (faceRebukes != 0)
                                    {
                                        string faceEdits = GetFileEdits(faceRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, faceEdits, sharedStringTablePart, 1U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 1U);
                                    }
                                }
                                else
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет лица", sharedStringTablePart, 1U);
                                }
                                if (backFiles.Count() > 0)
                                {
                                    int backRebukes = backFiles[0].Rebukes;//правка оборот
                                    if (backRebukes != 0)
                                    {
                                        string backEdits = GetFileEdits(backRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, backEdits, sharedStringTablePart, 1U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 1U);
                                    }

                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет оборота", sharedStringTablePart, 1U);
                            }
                            else
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Manager.Name, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.LastStateRef.Description, sharedStringTablePart, 1U);
                        }
                        else
                        {
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 2U);     //номер ТЗ
                            string indTz = "";
                            if (job.OrderType == 2)
                                indTz = "Да";
                            else
                                indTz = "Нет";
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 2U); //инд заказ
                            if (null != job.num)
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 2U);  //редакция
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "1", sharedStringTablePart, 2U);//вкладка
                            if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 2U);
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);//дата создания, она же дата создания тз на проверку
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.CreatedDate.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время создания
                            if (job.DoneDate.HasValue && job.StartDate.HasValue)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);   // дата принятия
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время принятия в работу
                                TimeSpan ts = job.DoneDate.Value - job.StartDate.Value;
                                //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            }
                            if (job.HumanWorkEndDate.HasValue && job.StartDate.HasValue)
                            {
                                //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);   //
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.HumanWorkEndDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время окончания проверки
                                TimeSpan ts = job.HumanWorkEndDate.Value - job.StartDate.Value;
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 2U);  //время на проверку
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            }
                            if (job.DoneDate.HasValue)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.DoneDate.Value.ToString("HH:mm:ss", cult), sharedStringTablePart, 2U);//время из препресс
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 2U);   //папка
                                                                                                                                 // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 2U);
                            if (job.Model.Units.Count > 0)
                            {                                                                                                     // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 1U);
                                int? faceId = job.Model.Units[0].FaceFile_Id;
                                int? backId = job.Model.Units[0].BackFile_Id;
                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == faceId).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == backId).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 2U);     //лицо/оборот
                                if (job.Model.Units[0].Skrepka == true)
                                {
                                    int? countPages = faceFiles[0].PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 2U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 2U);     //Скрепка
                                if (faceFiles.Count() > 0)
                                {
                                    int faceRebukes = faceFiles[0].Rebukes;//правка лицо
                                    if (faceRebukes != 0)
                                    {
                                        string faceEdits = GetFileEdits(faceRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, faceEdits, sharedStringTablePart, 2U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 2U);
                                    }
                                }
                                else
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет лица", sharedStringTablePart, 2U);
                                }
                                if (backFiles.Count() > 0)
                                {
                                    int backRebukes = backFiles[0].Rebukes;//правка оборот
                                    if (backRebukes != 0)
                                    {
                                        string backEdits = GetFileEdits(backRebukes, job.JobType);
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, backEdits, sharedStringTablePart, 2U);
                                    }
                                    else
                                    {
                                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "правок не было", sharedStringTablePart, 2U);
                                    }

                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "нет оборота", sharedStringTablePart, 2U);
                            }
                            else
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Manager.Name, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.LastStateRef.Description, sharedStringTablePart, 2U);
                        }
                        sheetData.Append(row);
                        rowInd++;
                        cellNum = 1;
                    }

                }
            }
            uint saleEnd = rowInd;
            cellNum = 1;
            AutoFilter autoFilter = new AutoFilter() { Reference = "A1:S100000" };
            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup = new PageSetup() { PaperSize = (UInt32Value)9U, FirstPageNumber = (UInt32Value)0U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U };
            worksheetPart.Worksheet.Append(sheetDimension);
            worksheetPart.Worksheet.Append(sheetViews);
            worksheetPart.Worksheet.Append(sheetFormatProperties);
            worksheetPart.Worksheet.Append(columns);
            worksheetPart.Worksheet.Append(sheetData);
            worksheetPart.Worksheet.Append(autoFilter);
            worksheetPart.Worksheet.Append(pageMargins);
            worksheetPart.Worksheet.Append(pageSetup);


            workbookPart.Workbook.Save();

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>();
            XLSX_Helpers.GenerateThemePartContent(themePart);

            document.Close();
           
        }

        private static void CreateWeeklyParts(SpreadsheetDocument document, List<PrepressJob> Jobs, string fileName)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>();
            XLSX_Helpers.GenerateExtendedWeeklyFilePropertiesPartContent(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            Sheets sheets = new Sheets();
            Sheet sheetZp = new Sheet() { Name = "Расчет", SheetId = (UInt32Value)1U, Id = "rId1" };
            Sheet sheetMakets = new Sheet() { Name = "Выгрузка", SheetId = (UInt32Value)2U, Id = "rId2" };

            sheets.Append(sheetZp);
            sheets.Append(sheetMakets);

            workbookPart.Workbook.Append(sheets);

            SharedStringTablePart sharedStringTablePart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                sharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            WorksheetPart worksheetPartZPSheet = workbookPart.AddNewPart<WorksheetPart>("rId1");//лист с Зарплатой
            GenerateZpSheetContent(workbookPart, worksheetPartZPSheet, sharedStringTablePart, Jobs);

            WorksheetPart worksheetPartMaketsSheet = workbookPart.AddNewPart<WorksheetPart>("rId2");//лист с макетами
            GenerateMaketsSheetContent(workbookPart, worksheetPartMaketsSheet, sharedStringTablePart, Jobs);

            //добавим стилей
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            XLSX_Helpers.GenerateWorkbookWeeklyReportStylesPartContent(workbookStylesPart);
            // workbookStylesPart.Stylesheet.Save();

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>();
            XLSX_Helpers.GenerateThemePartWeeklyReportContent(themePart);
            //  workbookPart.Workbook.Save();
            document.Close();

        }

        private static void GenerateZpSheetContent(WorkbookPart workbookPart, WorksheetPart worksheetPart, SharedStringTablePart sharedStringTablePart, List<PrepressJob> Jobs)
        {
            worksheetPart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheetPart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheetPart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");



            //Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            // Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Расчет" };//имя можно задать любое, единственное ограничение - количество символов не должно превышать 40, хотя в документации указан предел в 255 символов, мистика одним словом
            // sheets.Append(sheet);
            SheetDimension sheetDimension = new SheetDimension() { Reference = "A1:Z100000" };

            SheetViews sheetViews = new SheetViews();

            SheetView sheetView = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Pane pane = new Pane() { VerticalSplit = 1D, TopLeftCell = "A2", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };
            Selection selection = new Selection() { Pane = PaneValues.BottomLeft };

            sheetView.Append(pane);
            sheetView.Append(selection);


            // sheetView1.Append(selection1);

            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
            Columns columns = new Columns();
            columns.Append(new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 40D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 30D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 25D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 20D, BestFit = true, CustomWidth = true });
            SheetData sheetData = new SheetData();
            var dates = Jobs.Where(w=>w.DoneDate.HasValue==true).Select(s => s.DoneDate.Value.ToString("dd.MM.yyyy")).Distinct().ToList();
            List<string> headerRow = new List<string>() { "ФИО" };
            foreach (var date in dates)
            {
                headerRow.Add(date);
            }
            headerRow.Add("Всего");
            headerRow.Add("З/П");
            uint rowInd = 1U;
            int cellNum = 1;

            Row rowHead = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
            foreach (string s in headerRow)
            {
                cellNum = XLSX_Helpers.InsertCellToRow(rowHead, cellNum, s, sharedStringTablePart, 8U);
            }
            sheetData.AppendChild(rowHead);
            rowInd++;
            cellNum = 1;
            var countJobs = Jobs.Where(w => w.Executor != null).GroupBy(g => g.Executor).ToDictionary(k => k.Key);
            int counter = 1, skrepkaCounter = 0;
            foreach (var cJob in countJobs)
            {
                Row row = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
                if (counter % 2 != 0)
                {
                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, cJob.Key.Name, sharedStringTablePart, 10U);
                    for (int i = 0; i <= headerRow.Count() - 1; i++)
                    {
                        string date = headerRow[i];
                        DateTime tmp = DateTime.Now;
                        if (DateTime.TryParse(date, out tmp) != false)
                        {
                            var curDateJobs = cJob.Value.Where(w => w.HumanWorkEndDate.Value.ToString("dd.MM.yyyy") == date).ToList();
                            int counterMakets = 0;
                            foreach (var maket in curDateJobs)
                            {
                                if (maket.Model.Units.Count() > 1)
                                    counterMakets += maket.Model.Units.Count();
                                else
                                    counterMakets++;
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, counterMakets.ToString(), sharedStringTablePart, 11U);
                        }
                    }
                    int weekJobsCount = 0;
                    foreach (var maket in cJob.Value)
                    {
                        if (maket.Model.Units.Count() > 1)
                            weekJobsCount += maket.Model.Units.Count();
                        else
                            weekJobsCount++;
                    }
                    int zp = weekJobsCount * 14;
                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, weekJobsCount.ToString(), sharedStringTablePart, 10U);
                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, zp.ToString(), sharedStringTablePart, 10U);
                }
                else
                {
                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, cJob.Key.Name, sharedStringTablePart, 9U);
                    for (int i = 0; i <= headerRow.Count() - 1; i++)
                    {
                        string date = headerRow[i];
                        DateTime tmp = DateTime.Now;
                        if (DateTime.TryParse(date, out tmp) != false)
                        {
                            var curDateJobs = cJob.Value.Where(w => w.HumanWorkEndDate.Value.ToString("dd.MM.yyyy") == date).ToList();
                            int counterMakets = 0;
                            foreach (var maket in curDateJobs)
                            {
                                if (maket.Model.Units.Count() > 1)
                                    counterMakets += maket.Model.Units.Count();
                                else
                                    counterMakets++;
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, counterMakets.ToString(), sharedStringTablePart, 9U);
                        }
                    }
                    int weekJobsCount = 0;
                    foreach (var maket in cJob.Value)
                    {
                        if (maket.Model.Units.Count() > 1)
                            weekJobsCount += maket.Model.Units.Count();
                        else
                            weekJobsCount++;
                    }
                    int zp = weekJobsCount * 14;
                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, weekJobsCount.ToString(), sharedStringTablePart, 9U);
                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, zp.ToString(), sharedStringTablePart, 9U);
                }
                sheetData.AppendChild(row);
                rowInd++;
                cellNum = 1;
                counter++;
            }
            Row emptyRow = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            sheetData.AppendChild(emptyRow);
            rowInd++;
            cellNum = 1;
            Row rowAll = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            int JobsCount = 0;
            foreach (var maket in Jobs)
            {
                if (maket.Model.Units.Count() > 1)
                    JobsCount += maket.Model.Units.Count();
                else
                    JobsCount++;
            }
            cellNum = XLSX_Helpers.InsertCellToRow(rowAll, cellNum, "Всего макетов", sharedStringTablePart, 6U);
            cellNum = XLSX_Helpers.InsertCellToRow(rowAll, cellNum, JobsCount.ToString(), sharedStringTablePart, 7U);
            sheetData.AppendChild(rowAll);
            rowInd++;
            cellNum = 1;
            foreach (var job in Jobs)
            {
                if (job.Model.Units.Any(a => a.Skrepka == true))
                    skrepkaCounter++;
            }
            rowAll = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            cellNum = XLSX_Helpers.InsertCellToRow(rowAll, cellNum, "Скрепка", sharedStringTablePart, 9U);
            cellNum = XLSX_Helpers.InsertCellToRow(rowAll, cellNum, skrepkaCounter.ToString(), sharedStringTablePart, 8U);
            sheetData.AppendChild(rowAll);
            rowInd++;
            cellNum = 1;
            emptyRow = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            sheetData.AppendChild(emptyRow);
            rowInd++;
            cellNum = 1;
            Row maketCost = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            cellNum = XLSX_Helpers.InsertCellToRow(maketCost, cellNum, "Стоимость макета", sharedStringTablePart, 9U);
            cellNum = XLSX_Helpers.InsertCellToRow(maketCost, cellNum, "14", sharedStringTablePart, 8U);
            sheetData.AppendChild(maketCost);
            rowInd++;
            cellNum = 1;
            emptyRow = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            sheetData.AppendChild(emptyRow);
            rowInd++;
            cellNum = 1;
            int zpAllMackets = JobsCount * 14;
            Row zpAll = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            cellNum = XLSX_Helpers.InsertCellToRow(zpAll, cellNum, "З/П за макеты", sharedStringTablePart, 9U);
            cellNum = XLSX_Helpers.InsertCellToRow(zpAll, cellNum, zpAllMackets.ToString(), sharedStringTablePart, 8U);
            sheetData.AppendChild(zpAll);
            rowInd++;
            cellNum = 1;
            emptyRow = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            sheetData.AppendChild(emptyRow);
            rowInd++;
            cellNum = 1;
            Row rowWorkDays = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
            cellNum = XLSX_Helpers.InsertCellToRow(rowWorkDays, cellNum, "Рабочих дней", sharedStringTablePart, 9U);
            cellNum = XLSX_Helpers.InsertCellToRow(rowWorkDays, cellNum, "5", sharedStringTablePart, 8U);
            sheetData.AppendChild(rowWorkDays);
            rowInd++;
            cellNum = 1;
            // AutoFilter autoFilter = new AutoFilter() { Reference = "A1:S100000" };
            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup = new PageSetup() { PaperSize = (UInt32Value)9U, FirstPageNumber = (UInt32Value)0U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U };
            worksheetPart.Worksheet.Append(sheetDimension);
            worksheetPart.Worksheet.Append(sheetViews);
            worksheetPart.Worksheet.Append(sheetFormatProperties);
            worksheetPart.Worksheet.Append(columns);
            worksheetPart.Worksheet.Append(sheetData);
            //worksheetPart.Worksheet.Append(autoFilter);
            worksheetPart.Worksheet.Append(pageMargins);
            worksheetPart.Worksheet.Append(pageSetup);

        }

        private static void GenerateMaketsSheetContent(WorkbookPart workbookPart, WorksheetPart worksheetPart, SharedStringTablePart sharedStringTablePart, List<PrepressJob> Jobs)
        {
            worksheetPart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheetPart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheetPart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            //Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 2U, Name = "Выгрузка" };//имя можно задать любое, единственное ограничение - количество символов не должно превышать 40, хотя в документации указан предел в 255 символов, мистика одним словом
            // sheets.Append(sheet);
            SheetDimension sheetDimension = new SheetDimension() { Reference = "A1:Z100000" };

            SheetViews sheetViews = new SheetViews();

            SheetView sheetView = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Pane pane = new Pane() { VerticalSplit = 1D, TopLeftCell = "A2", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };
            Selection selection = new Selection() { Pane = PaneValues.BottomLeft };

            sheetView.Append(pane);
            sheetView.Append(selection);


            // sheetView1.Append(selection1);

            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
            Columns columns = new Columns();
            columns.Append(new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 30D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 25D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 20D, BestFit = true, CustomWidth = true });
            columns.Append(new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 20D, BestFit = true, CustomWidth = true });

            SheetData sheetData = new SheetData();

            string[] headerRow = new string[] { "Номер ТЗ", "Индивидуальный заказ", "Номер редакции", "Вкладка", "Приемщик", "Дата принятия", "Папка", "Лицо/Оборот", "Скрепка" };
            string[] headerRow2 = new string[] { "Номер ТЗ", "Приемщик", "Дата создания", "Время создания", "Дата начала проверки", "Время начала проверки", "Дата окончания проверки", "Время окончания проверки" };
            uint rowInd = 1U;
            int cellNum = 1;

            Row rowHead = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
            foreach (string s in headerRow)
            {
                cellNum = XLSX_Helpers.InsertCellToRow(rowHead, cellNum, s, sharedStringTablePart, 3U);
            }
            sheetData.AppendChild(rowHead);
            rowInd++;
            cellNum = 1;
            CultureInfo cult = CultureInfo.CurrentCulture;
            using (Repository rep = new Repository())
            {
                foreach (var job in Jobs)
                {
                    Row row = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
                    int countUnits = 1; string indTz = "";
                    if (job.Model.Units.Count() > 1)
                    {
                        foreach (var unit in job.Model.Units.AsEnumerable().Reverse())
                        {
                            row = new Row() { RowIndex = rowInd, DyDescent = 0.25D };
                            if (rowInd % 2 == 0)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 1U);     //номер ТЗ
                                //string indTz = "";
                                if (job.OrderType == 2)
                                    indTz = "Да";
                                else
                                    indTz = "Нет";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 1U); //инд заказ
                                if (null != job.num)                                                                                      //редакция
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 1U);
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);

                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countUnits.ToString(), sharedStringTablePart, 1U);//вкладка

                                if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 1U);
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);

                                if (job.StartDate.HasValue)
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);   // дата принятия
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                }
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 1U);   //папка


                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.FaceFile_Id).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.BackFile_Id).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 1U);     //лицо/оборот
                                if (unit.Skrepka == true)
                                {
                                    var modelFileSkr = rep.Select<ModelFile>().First(w => w.Id == unit.FaceFile_Id);
                                    int? countPages = modelFileSkr.PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 1U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 1U);     //Скрепка
                            }
                            else
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 2U);     //номер ТЗ
                                //string indTz = "";
                                if (job.OrderType == 2)
                                    indTz = "Да";
                                else
                                    indTz = "Нет";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 2U); //инд заказ
                                if (null != job.num)
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 2U);  //редакция
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countUnits.ToString(), sharedStringTablePart, 2U);//вкладка
                                if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 2U);
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                if (job.StartDate.HasValue)
                                {
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);   // дата принятия
                                                                                                                                                                         //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                                }
                                else
                                {
                                    // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                    //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                }

                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 2U);   //папка
                                                                                                                                     // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 2U);
                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.FaceFile_Id).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == unit.BackFile_Id).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 2U);     //лицо/оборот
                                if (unit.Skrepka == true)
                                {
                                    var modelFileSkr = rep.Select<ModelFile>().First(w => w.Id == unit.FaceFile_Id);
                                    int? countPages = modelFileSkr.PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 2U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 2U);     //Скрепка
                            }

                            sheetData.Append(row);
                            rowInd++;
                            cellNum = 1;
                            countUnits++;
                        }
                        if (countUnits == job.Model.Units.Count())
                            continue;
                    }
                    else
                    {
                        if (rowInd % 2 == 0)
                        {
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 1U);     //номер ТЗ

                            if (job.OrderType == 2)
                                indTz = "Да";
                            else
                                indTz = "Нет";
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 1U); //инд заказ
                            if (null != job.num)
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 1U);  //редакция
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "1", sharedStringTablePart, 1U);//вкладка
                            if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 1U);
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            if (job.StartDate.HasValue)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 1U);   // дата принятия
                                //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 1U);
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 1U);   //папка
                            if (job.Model.Units.Count > 0)
                            {                                                                                                     // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 1U);
                                int? faceId = job.Model.Units[0].FaceFile_Id;
                                int? backId = job.Model.Units[0].BackFile_Id;
                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == faceId).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == backId).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 1U);     //лицо/оборот
                                if (job.Model.Units[0].Skrepka == true)
                                {
                                    int? countPages = faceFiles[0].PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 1U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 1U);     //Скрепка
                            }
                        }
                        else
                        {
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.IdStr, sharedStringTablePart, 2U);     //номер ТЗ

                            if (job.OrderType == 2)
                                indTz = "Да";
                            else
                                indTz = "Нет";
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, indTz, sharedStringTablePart, 2U); //инд заказ
                            if (null != job.num)
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.num.ToString(), sharedStringTablePart, 2U);  //редакция
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "1", sharedStringTablePart, 2U);//вкладка
                            if (job.Executor != null && !string.IsNullOrEmpty(job.Executor.Name))                                  //приемщик
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Executor.Name, sharedStringTablePart, 2U);
                            else
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            if (job.StartDate.HasValue)
                            {
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.StartDate.Value.ToString("dd.MM.yyyy", cult), sharedStringTablePart, 2U);   // дата принятия
                                                                                                                                                                     //cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds, sharedStringTablePart, 1U);
                            }
                            else
                            {
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                                // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "", sharedStringTablePart, 2U);
                            }
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.ResultFolder, sharedStringTablePart, 2U);   //папка
                                                                                                                                 // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 2U);
                            if (job.Model.Units.Count > 0)
                            {                                                                                                     // cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, job.Model.ActualUnitsCount.ToString(), sharedStringTablePart, 1U);
                                int? faceId = job.Model.Units[0].FaceFile_Id;
                                int? backId = job.Model.Units[0].BackFile_Id;
                                var faceFiles = rep.Select<ModelFile>().Where(w => w.Id == faceId).ToList();
                                var backFiles = rep.Select<ModelFile>().Where(w => w.Id == backId).ToList();
                                string FaceBack = "";
                                if (backFiles.Count() > 0)
                                    FaceBack = "2";
                                else
                                    FaceBack = "1";
                                cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, FaceBack, sharedStringTablePart, 2U);     //лицо/оборот
                                if (job.Model.Units[0].Skrepka == true)
                                {
                                    int? countPages = faceFiles[0].PagesCount;
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, countPages.ToString(), sharedStringTablePart, 2U);     //Скрепка
                                }
                                else
                                    cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, "-", sharedStringTablePart, 2U);     //Скрепка
                            }
                        }
                        sheetData.Append(row);
                        rowInd++;
                        cellNum = 1;
                    }
                }
            }
            uint saleEnd = rowInd;
            cellNum = 1;
            AutoFilter autoFilter = new AutoFilter() { Reference = "A1:S100000" };
            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup = new PageSetup() { PaperSize = (UInt32Value)9U, FirstPageNumber = (UInt32Value)0U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U };
            worksheetPart.Worksheet.Append(sheetDimension);
            worksheetPart.Worksheet.Append(sheetViews);
            worksheetPart.Worksheet.Append(sheetFormatProperties);
            worksheetPart.Worksheet.Append(columns);
            worksheetPart.Worksheet.Append(sheetData);
            worksheetPart.Worksheet.Append(autoFilter);
            worksheetPart.Worksheet.Append(pageMargins);
            worksheetPart.Worksheet.Append(pageSetup);

        }

        private static string GetFileEdits(int fileRebukes, int jobType)
        {
            string str = "";
            switch (fileRebukes)
            {
                case 10:
                    str = "нет вылетов - , близко к резу -";
                    break;
                case 6:
                    str = "нет вылетов - , близко к резу +";
                    break;
            }
            if (!string.IsNullOrEmpty(str))
                return str;
            else
            {
                using (Repository r = new Repository())
                {
                    var rebList = r.Select<RebukesRef>().Where(w => w.JobType == jobType).ToList();
                    if (rebList.Any(a => a.ErrorMask == fileRebukes))
                    {
                        return str = rebList.Where(w => w.ErrorMask == fileRebukes).Select(s => s.Name).First() + " - ";
                    }
                    else if (rebList.Any(a => a.ErrorMask <= 32 && a.ErrorMask / 2 == fileRebukes))
                    {
                        return str = rebList.Where(w => w.ErrorMask <= 32 && w.ErrorMask / 2 == fileRebukes).Select(s => s.Name).First() + " + ";
                    }
                    else
                    {
                        foreach (var reb in rebList.AsEnumerable().Reverse())
                        {
                            bool check = false;
                            int val = 0, tempRebuke = 0;
                            string tmp = "";
                            if (reb.ErrorMask > 32)
                                val = reb.ErrorMask;
                            else
                            {
                                val = reb.ErrorMask / 2;
                            }
                            do
                            {
                                tempRebuke = fileRebukes - val;
                                if (tempRebuke >= 0)
                                {
                                    fileRebukes = fileRebukes - val;
                                    if (val <= 16)
                                        str += reb.Name + " + , ";
                                    else
                                        str += reb.Name + " - , ";
                                    if (GetFileEdits(fileRebukes, jobType, out tmp))
                                    {
                                        str += tmp;
                                        check = true;
                                        break;
                                    }
                                    else
                                    {
                                        check = false;
                                        break;
                                    }
                                }
                                else
                                {
                                    check = false;
                                    break;
                                }
                            } while (fileRebukes > 0);
                            if (check == true)
                                break;
                            else
                                continue;
                        }
                    }
                }
                return str.TrimEnd(',');
            }
        }

        private static bool GetFileEdits(int fileRebukes, int jobType, out string str)
        {
            using (Repository rep = new Repository())
            {
                bool flag = false;
                var rebList = rep.Select<RebukesRef>().Where(w => w.JobType == jobType).ToList();
                if (rebList.Any(a => a.ErrorMask == fileRebukes))
                {
                    str = rebList.Where(w => w.ErrorMask == fileRebukes).Select(s => s.Name).First() + " - ";
                    return flag = true;
                }
                else if (rebList.Any(a => a.ErrorMask <= 32 && a.ErrorMask / 2 == fileRebukes))
                {
                    str = rebList.Where(w => w.ErrorMask <= 32 && w.ErrorMask / 2 == fileRebukes).Select(s => s.Name).First() + " + ";
                    return flag = true;
                }
                else
                    str = "";
                return flag = false;
            }
        }

        private static void CreateQueryReportParts(SpreadsheetDocument document, TCPExcelSend tes) {
            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>();
            XLSX_Helpers.GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            SharedStringTablePart sharedStringTablePart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                sharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();//лист с макетами
            worksheetPart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheetPart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheetPart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //добавим стилей
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            XLSX_Helpers.GenerateWorkbookStylesPartContent(workbookStylesPart);
            workbookStylesPart.Stylesheet.Save();

            Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Лист 1" };//имя можно задать любое, единственное ограничение - количество символов не должно превышать 40, хотя в документации указан предел в 255 символов, мистика одним словом
            sheets.Append(sheet);
            SheetDimension sheetDimension = new SheetDimension() { Reference = "A1:Z100000" };

            SheetViews sheetViews = new SheetViews();

            SheetView sheetView = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Pane pane = new Pane() { VerticalSplit = 1D, TopLeftCell = "A2", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };
            Selection selection = new Selection() { Pane = PaneValues.BottomLeft };

            sheetView.Append(pane);
            sheetView.Append(selection);


            // sheetView1.Append(selection1);

            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
            Columns columns = new Columns();
            uint columnCounter = 1;
            foreach (var elem in tes.data[0].header) {
                columns.Append(new Column() { Min = (UInt32Value)columnCounter, Max = (UInt32Value)columnCounter, Width = 20D, BestFit = true, CustomWidth = true });
                columnCounter++;
            }
            SheetData sheetData = new SheetData();
            uint rowInd = 1U;
            int cellNum = 1;
            Row rowHead = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
            foreach (string s in tes.data[0].header)
            {
                cellNum = XLSX_Helpers.InsertCellToRow(rowHead, cellNum, s, sharedStringTablePart, 3U);
            }
            sheetData.AppendChild(rowHead);
            rowInd++;
            cellNum = 1;
            tes.data.RemoveAt(0);
            foreach (var elem in tes.data) {
                Row row = new Row () { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                int countUnits = 1;
                if (elem.Orders.Count > 1)
                {
                    foreach (var order in elem.Orders)
                    {
                        row = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                        if (rowInd % 2 == 0)
                        {
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.accNum, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, order, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Data, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Customer, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.CustOrganization, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Payment, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Debt, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Summ, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Comment, sharedStringTablePart, 1U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Manager, sharedStringTablePart, 1U);
                        }
                        else
                        {
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.accNum, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, order, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Data, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Customer, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.CustOrganization, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Payment, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Debt, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Summ, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Comment, sharedStringTablePart, 2U);
                            cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Manager, sharedStringTablePart, 2U);
                        }
                        sheetData.Append(row);
                        rowInd++;
                        cellNum = 1;
                    }
                    
                    countUnits++;
                    if (countUnits == elem.Orders.Count())
                        continue;
                }
                else
                {
                    if (rowInd % 2 == 0)
                    {
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.accNum, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Orders[0], sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Data, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Customer, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.CustOrganization, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Payment, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Debt, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Summ, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Comment, sharedStringTablePart, 1U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Manager, sharedStringTablePart, 1U);
                    }
                    else
                    {
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.accNum, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Orders[0], sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Data, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Customer, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.CustOrganization, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Payment, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Debt, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Summ, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Comment, sharedStringTablePart, 2U);
                        cellNum = XLSX_Helpers.InsertCellToRow(row, cellNum, elem.Manager, sharedStringTablePart, 2U);
                    }
                    sheetData.Append(row);
                    rowInd++;
                    cellNum = 1;
                }
            }
            AutoFilter autoFilter = new AutoFilter() { Reference = "A1:S100000" };
            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup = new PageSetup() { PaperSize = (UInt32Value)9U, FirstPageNumber = (UInt32Value)0U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U };
            worksheetPart.Worksheet.Append(sheetDimension);
            worksheetPart.Worksheet.Append(sheetViews);
            worksheetPart.Worksheet.Append(sheetFormatProperties);
            worksheetPart.Worksheet.Append(columns);
            worksheetPart.Worksheet.Append(sheetData);
            worksheetPart.Worksheet.Append(autoFilter);
            worksheetPart.Worksheet.Append(pageMargins);
            worksheetPart.Worksheet.Append(pageSetup);


            workbookPart.Workbook.Save();

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>();
            XLSX_Helpers.GenerateThemePartContent(themePart);

            document.Close();
        }

    }
}
