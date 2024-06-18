using System.Collections.Generic;
using System.Linq;
using ImportSheetConsole.Global;
using ObjStudioClasses;

namespace ImportSheetConsole.DefaultScheme
{
    /// <summary>
    /// Схема импорта опросного листа по умолчанию
    /// </summary>
    internal class ScriptImplemetation : Script
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
                        case "ТУ":
                            sheet = new MeterPoints(excelSheet);
                            break;
                        case "УСПД":
                            sheet = new Rtus(excelSheet);
                            break;
                        case "ТТ":
                            sheet = new TransCurrents(excelSheet);
                            break;
                        case "ТН":
                            sheet = new TransVoltages(excelSheet);
                            break;
                        case "ФЛ":
                            sheet = new NaturalPersons(excelSheet);
                            break;
                        case "ЮЛ":
                            sheet = new LegalEntities(excelSheet);
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