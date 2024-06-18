using System;
using System.Collections.Generic;
using GemBox.Spreadsheet;
using ImportSheetConsole.Global;
using ObjStudioClasses;

namespace ImportSheetConsole.TemplateScheme
{
    #region Helper

    // класс описания листа
    internal abstract class Sheet
    {
        /// <summary>
        /// Колонка номера
        /// </summary>
        public const int COL_NUM = 1 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Информация о значении строки
        /// </summary>
        public class RowValueInfo
        {
            /// <summary>
            /// Лист
            /// </summary>
            public readonly Sheet Sheet;

            /// <summary>
            /// Индекс строки
            /// </summary>
            public readonly int RowIndex;

            /// <summary>
            /// Значения
            /// </summary>
            public readonly Dictionary<int, object> Values;

            /// <summary>
            /// Конструктор
            /// </summary>
            public RowValueInfo(Sheet sheet, int rowIndex, Dictionary<int, object> values)
            {
                Sheet = sheet;
                RowIndex = rowIndex;
                Values = values;
            }
        }

        /// <summary>
        /// Значения ячеек по строкам
        /// </summary>
        public readonly List<RowValueInfo> RowValues = new List<RowValueInfo>();

        /// <summary>
        /// Индекс листа
        /// </summary>
        private readonly int _sheetIndex;

        /// <summary>
        /// Лист
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get { return ScriptImpl.Instance.WorkbookNonExcel.Worksheets[_sheetIndex]; }

        }
        
        // добавить значение строки
        protected void AddRowValue(int rowIndex, Dictionary<int, object> rowValue)
        {
            var rowValueInfo = new RowValueInfo(this, rowIndex, rowValue);
            RowValues.Add(rowValueInfo);
        }
        
        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        protected Sheet(ExcelWorksheet sheet)
        {
            _sheetIndex = sheet.Index;
        }

        /// <summary>
        /// Проверить
        /// </summary>
        /// <param name="result"></param>
        public abstract void Check(ImportSheetCheckResultData result);

        /// <summary>
        /// Проверить после загрузки всех листов
        /// </summary>
        /// <param name="result"></param>
        public abstract void CheckAfterAllLoading(ImportSheetCheckResultData result);
    }

    // лист 1
    internal class Sheet1 : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        // ...
        
        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public Sheet1(ExcelWorksheet sheet)
            : base(sheet)
        {
        }
        
        // Проверить
        public override void Check(ImportSheetCheckResultData result)
        {
            // очистить лист от результатов предыдущей проверки
            Worksheet.ClearCheckResults(START_ROW, COL_NUM);

            foreach (var row in Worksheet.Rows)
            {
                if (row.Index < START_ROW)
                    continue;
                if (row.Cells[COL_NUM].Value == null)
                    break;
                try
                {
                    var rowValue = CheckHelper.CheckSheet1Row(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue);
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
                }
                finally
                {
                    result.TotalCheckedLinesCount++;
                }
            }
        }

        // Проверить после загрузки всех листов
        public override void CheckAfterAllLoading(ImportSheetCheckResultData result)
        {
        }
    }
    
    /// <summary>
    /// Общие утилиты (внутренние)
    /// </summary>
    internal static class Helper
    {
        /// <summary>
        /// Смещение индексов для разных версий поставщика Excel
        /// </summary>
        public const byte OFFSET_IF_NOT_EXCEL = 1;

        /// <summary>
        /// Листы
        /// </summary>
        [ThreadStatic]
        public static List<Sheet> Sheets;

        /// <summary>
        /// Процент на одну строку
        /// </summary>
        [ThreadStatic]
        public static double PercentPerRow;

        // Функции получения значения
        // ...
        

        /// <summary>
        /// Уведомить об изменении процента по строкам
        /// </summary>
        public static void NotifyPercentRow()
        {
            AsyncTasksHelper.PublishedTaskContext.NotifyPercent(PercentPerRow);
        }
    }

    #endregion Helper
}