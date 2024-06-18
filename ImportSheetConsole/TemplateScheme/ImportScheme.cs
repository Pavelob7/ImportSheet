using System.Collections.Generic;
using System.Linq;
using ImportSheetConsole.Global;
using ObjStudioClasses;

namespace ImportSheetConsole.TemplateScheme
{
    /// <summary>
    /// Схема импорта опросного листа (шаблон)
    /// </summary>
    internal class ScriptImplementation : Script
    {
        /// <inheritdoc />
        /// <summary>
        /// </summary>
        /// <returns></returns>
        protected override ImportSheetCheckResultData Check()
        {
            #region Check

            var result = new ImportSheetCheckResultData();
            
            Helper.Sheets = new List<Sheet>();

            // проверить каждый лист
            using (AsyncTasksHelper.PublishedTaskContext.CreateRelativePercentBorders(0, 100))
            {
                var percentPerSheet = 100 / (double)WorkbookNonExcel.Worksheets.Count;
                foreach (var excelSheet in WorkbookNonExcel.Worksheets)
                {
                    Sheet sheet = null;
                    switch (excelSheet.Name)
                    {
                        case "Лист1":
                            sheet = new Sheet1(excelSheet);
                            break;
                    }
                    if (sheet != null)
                    {
                        Helper.Sheets.Add(sheet);
                        sheet.Check(result);
                    }
                    AsyncTasksHelper.PublishedTaskContext.NotifyPercent(percentPerSheet);
                }
            }

            // проверить после полной загрузки
            foreach (var sheet in Helper.Sheets)
            {
                sheet.CheckAfterAllLoading(result);
            }
            
            AddLogDebug(
                "TotalCheckedLinesCount = {0}\n" +
                "TotalErrorsCount = {1}\n" +
                "TotalWarningsCount = {2}\n",
                result.TotalCheckedLinesCount, result.TotalErrorsCount, result.TotalWarningsCount);

            return result;

            #endregion Check
        }

        /// <inheritdoc />
        /// <summary>
        /// </summary>
        /// <returns></returns>
        protected override ImportSheetProcessedResultData Import()
        {
            #region Import
            
            var result = new ImportSheetProcessedResultData();
            
            // фиктивное обращение, чтобы инициализировать книгу
            WorkbookNonExcel.Protected = false;

            using (AsyncTasksHelper.PublishedTaskContext.CreateRelativePercentBorders(0, 100))
            {
                Helper.PercentPerRow = 100 / (double)Helper.Sheets.Sum(x => x.RowValues.Count);
                ImportHelper.Import(result);
            }
            
            AddLogDebug(
                "TotalCheckedLinesCount = {0}\n" +
                "TotalErrorsCount = {1}\n" +
                "TotalWarningsCount = {2}\n" + 
                "ImportedEntitiesCount = {3}\n",
                result.TotalCheckedLinesCount, result.TotalErrorsCount, result.TotalWarningsCount, result.ImportedEntitiesCount);

            return result;

            #endregion Import
        }
    }
}