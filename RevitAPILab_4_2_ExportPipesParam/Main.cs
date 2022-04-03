using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;

namespace RevitAPILab_4_2_ExportPipesParam
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        //Экспортирует параметры труб в excell
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            #region Диалог сохранения файла
            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "pipeInfo.xlsx",
                DefaultExt = ".xlsx"
            };

            string selectedFilePath = string.Empty;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            #endregion

            List<Pipe> allPipes = new FilteredElementCollector(doc)
                 .OfClass(typeof(Pipe))
                 .Cast<Pipe>()
                 .ToList();

            List < PipeParam > pipeParams = new List<PipeParam>();

            string allText = string.Empty;
            foreach (var pipe in allPipes)
            {

                PipeParam pipeParam = new PipeParam
                {
                    PipeType = pipe.Name,
                   // PipeType = pipe.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).ToString(),
                    PipeLength = UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), DisplayUnitType.DUT_METERS).ToString(),
                    PipeInnerDiam = UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble(), DisplayUnitType.DUT_MILLIMETERS).ToString(),
                    PipeOuterDiam = UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(), DisplayUnitType.DUT_MILLIMETERS).ToString()
                };
                pipeParams.Add(pipeParam);

            }

                #region
                using (var fs = new FileStream(selectedFilePath, FileMode.Create, FileAccess.Write))
                {

                    IWorkbook workbook = new XSSFWorkbook();

                    ISheet sheet1 = workbook.CreateSheet("Sheet1");

                    //sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 3));
                    IRow row = sheet1.CreateRow(0);
                    row.Height = 80 * 10;
                    row.CreateCell(0).SetCellValue("Имя типа");
                    row.CreateCell(1).SetCellValue("Длина трубы,м");
                    row.CreateCell(2).SetCellValue("Внутр. диам, мм");
                    row.CreateCell(3).SetCellValue("Внеш. диам, мм");
                    sheet1.AutoSizeColumn(0);

                    var rowIndex = 1;

                    foreach (var pipe in pipeParams)
                    {
                   
                        row = sheet1.CreateRow(rowIndex);
                        row.Height = 10 * 80;
                        row.CreateCell(0).SetCellValue(pipe.PipeType);
                        row.CreateCell(1).SetCellValue(pipe.PipeLength);
                        row.CreateCell(2).SetCellValue(pipe.PipeInnerDiam);
                        row.CreateCell(3).SetCellValue(pipe.PipeOuterDiam);
                        sheet1.AutoSizeColumn(0);
                        rowIndex++;
                    }
                    

                    //var sheet2 = workbook.CreateSheet("Sheet2");
                    //var style1 = workbook.CreateCellStyle();
                    //style1.FillForegroundColor = HSSFColor.Blue.Index2;
                    //style1.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;

                    //var style2 = workbook.CreateCellStyle();
                    //style2.FillForegroundColor = HSSFColor.Yellow.Index2;
                    //style2.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;

                    //var cell2 = sheet2.CreateRow(0).CreateCell(0);
                    //cell2.CellStyle = style1;
                    //cell2.SetCellValue(0);

                    //cell2 = sheet2.CreateRow(1).CreateCell(0);
                    //cell2.CellStyle = style2;
                    //cell2.SetCellValue(1);

                    workbook.Write(fs);
                fs.Close();
                }
                #endregion

            



            TaskDialog.Show("Экспорт", $"Файл сохранен по указанному пути:\n{selectedFilePath}");
            return Result.Succeeded;
        }

    }
}
