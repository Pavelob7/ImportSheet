using System.Collections.Generic;
using GemBox.Spreadsheet;
using ObjStudioClasses;

namespace ImportSheetConsole.TemplateScheme
{
    #region CheckHelper

    /// <summary>
    /// Утилиты валидации (внутренние)
    /// </summary>
    internal static class CheckHelper
    {
        // проверка листа
        public static Dictionary<int, object> CheckSheet1Row(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var values = new Dictionary<int, object>();

            //// Договор: номер
            //values[MeterPoints.COL_CONTRACT] =
            //    sheet.Cells[row.Index, MeterPoints.COL_CONTRACT].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            
            return values;
        }
    }

    #endregion CheckHelper
}