using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using CSConstants;
using GemBox.Spreadsheet;
using ObjStudioClasses;
using RDCustomEntity = ObjStudioClasses.RDCustomEntity;

namespace ImportSheetConsole.Global
{
    #region CheckHelpers
    /// <summary>
    /// Утилиты валидации
    /// </summary>
    internal static class CheckHelpers
    {
        /// <summary>
        /// Индекст столбца для подсветки наличия ошибок и предупреждений
        /// </summary>
        public static int ColumnIndexOfRowHighlight = 1;
        
        /// <summary>
        /// Описание предупреждения при пустом значении ячейки
        /// </summary>
        public const string EMPTY_CELLVALUE_WARN = "Ячейка содержит пустое значение";

        /// <summary>
        /// Описание ошибки при пустом значении ячейки
        /// </summary>
        public const string EMPTY_CELLVALUE_ERROR = "Ячейка не может содержать пустое значение";

        /// <summary>
        /// Описание ошибки при неудачном разборе объекта НСИ из значения ячейки
        /// </summary>
        public const string RDPARSE_CELLVALUE_ERROR = "Не удалось разобрать значение типа \"{0}\"";

        /// <summary>
        /// Описание ошибки при неудачном разборе объекта НСИ из значения ячейки
        /// </summary>
        public const string NOT_APPLIED_CELLVALUE_ERROR = "\"{0}\" не может быть добавлен в элемент \"{1}\"";

        /// <summary>
        /// Функция получения значения из строки
        /// </summary>
        public delegate T GetValueFunc<out T>(string value);

        /// <summary>
        /// Функция проверки значения
        /// </summary>
        public delegate bool CheckFunc<in T>(T value, string stringValue, out string warnText, out string errorText);
        
        /// <summary>
        /// Получить значение ячейки строкой
        /// </summary>
        public static string GetString(this ExcelCell source)
        {
            if (source == null || source.Value == null)
                return null;
            var result = source.Value.ToString().Trim();
            return string.IsNullOrEmpty(result) ? null : result;
        }
        
        /// <summary>
        /// Функция проверки по умолчанию
        /// </summary>
        public static bool DefaultCheck<T>(T value, string stringValue, out string warnText, out string errorText)
        {
            warnText = EMPTY_CELLVALUE_WARN;
            errorText = EMPTY_CELLVALUE_ERROR;

            // строка
            if (typeof(T) == typeof(string)) 
                return !string.IsNullOrEmpty(value as string);

            // НСИ
            if (value == null && !string.IsNullOrEmpty(stringValue))
            {
                var elementType = typeof(T);
                var t = elementType;
                while (t != null)
                {
                    elementType = t;
                    t = t.GetElementType();
                }

                if (elementType.IsSubclassOf(typeof(RDCustomEntity)))
                {
                    var attr = elementType.GetCustomAttribute<RDMetadataSourceAttribute>();
                    if (attr != null)
                    {
                        var rd = RDCustomEntity.Find(attr.SourceRefName);
                        errorText = string.Format(RDPARSE_CELLVALUE_ERROR, rd.Caption);
                        warnText = errorText;
                    }
                }
            }

            return value != null;
        }

        /// <summary>
        /// Функция проверки адреса
        /// </summary>
        public static bool CheckAddressFiasDefault(FiasSuggestionData value, string stringValue, out string warnText, out string errorText)
        {
            if (!DefaultCheck(value, stringValue, out warnText, out errorText))
                return false;

            warnText = "Адрес не проверен";
            errorText = "Ошибка проверки адреса";

            try
            {
                var nonParsedAddress = value.Value;
                if (!string.IsNullOrEmpty(value.Region))
                    nonParsedAddress = nonParsedAddress.Replace(value.Region, null);
                if (!string.IsNullOrEmpty(value.CityDistrict))
                    nonParsedAddress = nonParsedAddress.Replace(value.CityDistrict, null);
                if (!string.IsNullOrEmpty(value.Area))
                    nonParsedAddress = nonParsedAddress.Replace(value.Area, null);
                if (!string.IsNullOrEmpty(value.City))
                    nonParsedAddress = nonParsedAddress.Replace(value.City, null);
                if (!string.IsNullOrEmpty(value.Settlement))
                    nonParsedAddress = nonParsedAddress.Replace(value.Settlement, null);
                if (!string.IsNullOrEmpty(value.Street))
                    nonParsedAddress = nonParsedAddress.Replace(value.Street, null);
                if (!string.IsNullOrEmpty(value.House))
                    nonParsedAddress = nonParsedAddress.Replace(value.House, null);
                if (!string.IsNullOrEmpty(value.Block))
                    nonParsedAddress = nonParsedAddress.Replace(value.Block, null);

                nonParsedAddress = nonParsedAddress.Trim(',', ' ');

                if (!string.IsNullOrEmpty(nonParsedAddress))
                    throw new Exception(string.Format("Не удалось разобрать часть адреса - \"{0}\"", nonParsedAddress));

                return true;
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                errorText += ":\n" + ex.Message;
            }

            return false;
        }

        /// <summary>
        /// Функция получения адреса через сервис ФИАС
        /// </summary>
        public static bool CheckAddressFias(FiasSuggestionData value, string stringValue, out string warnText, out string errorText)
        {
            if (!DefaultCheck(value, stringValue, out warnText, out errorText))
                return false;

            warnText = "Адрес не подтвержден в сервисе ФИАС";
            errorText = "Ошибка подтверждения адреса через сервис ФИАС";

            try
            {
                var suggestions = AddressFiasClassInfo.Get().GetFiasAddress(stringValue).ToArray();

                if (suggestions.Length == 0)
                    throw new Exception("Для указанного адреса сервис ФИАС не возвратил ниодного соответствия");

                var address = suggestions
                    .FirstOrDefault(
                        x => string.Equals(x.Value, value.Value, StringComparison.InvariantCultureIgnoreCase)
                    );

                if (address == null && suggestions.Length > 1)
                    throw new Exception(string.Format(
                        "Для указанного адреса сервис ФИАС возвратил более одного соответствия ({0}):\n{1}",
                        suggestions.Length, string.Join("\n", suggestions.Select(x => x.Value))));

                if (address == null)
                    throw new Exception(string.Format(
                        "Для указанного адреса сервис ФИАС возвратил соответствия, в которых не содержиться искомого значения:\n" +
                        "\"{0}\"\nв то время как ожидалось: \"{1}\"",
                        string.Join("\n", suggestions.Select(x => x.Value)), value.Value));

                // проверить уровень соответствия

                if (address.FiasLevel != null && value.FiasLevel != null &&
                    address.FiasLevel != FiasLevelItem.Instances.Other &&
                    address.FiasLevel.OrderPosition < value.FiasLevel.OrderPosition)
                    throw new Exception(string.Format(
                        "Для указанного адреса сервис ФИАС возвратил уровень детализации ниже ожидаемого " +
                        "(адрес распознан до \"{0}\" в то время как ожидалось до \"{1}\"):\n{2}",
                        address.FiasLevel, value.FiasLevel, string.Join("\n", suggestions.Select(x => x.Value))));

                value.Value = address.Value;
                value.FiasId = address.FiasId;

                return true;
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                errorText += ":\n" + ex.Message;
            }

            return false;
        }

        /// <summary>
        /// Может ли быть добавлена сущность в родителя
        /// </summary>
        /// <param name="source"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static bool IsAppliedToClassifierAttributes(this CustomClassifierNodeClassInfo source, CustomClassifierNodeClassInfo parent)
        {
            if (source == null || parent == null)
                return false;

            foreach (var attribute in parent.GetClassifierAttributesFast())
            {
                if (!attribute.Options.RequiredTypeRefName.HasValue)
                    continue;
                var cl = BaseClassClassInfo.Find(attribute.Options.RequiredTypeRefName.Value);
                if (cl != null && source.InheritsFrom(cl))
                    return true;
            }

            return false;
        }
        
        /// <summary>
        /// Проверить ячейку
        /// </summary>
        public static TValue GetWithCheck<TCell, TValue>(
            // объект ячейки (ExcelCell или Range)
            this TCell source,
            // общий результат метода проверки
            ImportSheetCheckResultData result,
            // функция получения значения ячейки
            GetValueFunc<TValue> getValueFunc,
            // функция проверки значения ячейки
            CheckFunc<TValue> checkFunc,
            // флаг обязательности наличия значения в ячейки		
            bool required = false,
            // повышать требование обязательности наличия значения, если в ячейке присутствует какое либо значение
            bool forceRequiredIfCellValueExist = true)
        {
            string stringValue;

            var nonExcelCell = source as ExcelCell;

            if (nonExcelCell != null)
                stringValue = nonExcelCell.GetString();
            else
                throw new ArgumentException("Не удалось определить тип постовщика ячеек (Excel или NonExcel)");

            var value = getValueFunc(stringValue);
            string warnText;
            string errorText;
            var checkResult = (checkFunc ?? DefaultCheck)(value, stringValue, out warnText, out errorText);

            // не ошибка
            if (checkResult)
                return value;

            // повышаем требование наличия значения, если ячейка не пустая
            if (!required && forceRequiredIfCellValueExist && !string.IsNullOrEmpty(stringValue))
                required = true;

            // вывод сообщения в ячейку (ошибки или предупреждения)
            var kind = !required ? DiagnosticMessageKindData.Warning : DiagnosticMessageKindData.Error;
            var text = !required ? warnText : errorText;

            // сообщение может быть подавлено, если функция проверки вернула пустой текст
            if (string.IsNullOrEmpty(text))
                return value;

            // выводит только сообщения об ошибках
            if (required)
            {
                if (nonExcelCell != null)
                    nonExcelCell.SetResult(kind, result, text);
            }

            return value;
        }

        /// <summary>
        /// Цвет подсветки ячейки с информацией
        /// </summary>
        public static Lazy<SpreadsheetColor> InfoColor = new Lazy<SpreadsheetColor>(() => SpreadsheetColor.FromName(ColorName.LightGreen));

        /// <summary>
        /// Цвет подсветки ячейки с предупреждением
        /// </summary>
        public static Lazy<SpreadsheetColor> WarnColor = new Lazy<SpreadsheetColor>(() => SpreadsheetColor.FromName(ColorName.LightBlue));

        /// <summary>
        /// Цвет подсветки ячейки с ошибкой
        /// </summary>
        public static Lazy<SpreadsheetColor> ErrorColor = new Lazy<SpreadsheetColor>(() => SpreadsheetColor.FromName(ColorName.Yellow));

        /// <summary>
        /// Цвет подсветки ячейки с ошибкой
        /// </summary>
        public static Lazy<SpreadsheetColor> RedColor = new Lazy<SpreadsheetColor>(() => SpreadsheetColor.FromName(ColorName.Red));

        /// <summary>
        /// Функция добавления сообщения в журнал
        /// </summary>
        [ThreadStatic]
        public static Action<DiagnosticMessageKindData, string, object[]> AddLogMessage;
            
        /// <summary>
        /// Ячейки, для которох уже был установлен результат
        /// </summary>
        [ThreadStatic]
        private static HashSet<ExcelCell> _setResultProcessedCells;

        /// <summary>
        /// Зарегистрировать результат в ячейке
        /// </summary>
        public static void SetResult(this ExcelCell source, DiagnosticMessageKindData kind, ImportSheetCheckResultData result, string text)
        {
            if (_setResultProcessedCells == null)
                _setResultProcessedCells = new HashSet<ExcelCell>();

            // не изменяем состояние, если в ячейку уже было установлен результат ранее
            if (kind != DiagnosticMessageKindData.Information && !_setResultProcessedCells.Add(source))
                return;
            
            // задать цвет строки
            var setColors = new Action<Lazy<SpreadsheetColor>>(color =>
            {
                var information = color.Value == InfoColor.Value;
                if (source.Style.FillPattern.GradientColor1 != ErrorColor.Value &&
                    (!information || source.Style.FillPattern.GradientColor1.IsEmpty))
                    source.Style.FillPattern.GradientColor1 = color.Value;
                var pattern = source.Worksheet.Cells[source.Row.Index, ColumnIndexOfRowHighlight - 1].Style.FillPattern;
                if (pattern.GradientColor1 != ErrorColor.Value &&
                    (!information || pattern.GradientColor1.IsEmpty))
                    pattern.GradientColor1 = color.Value;
                if (source.Worksheet.TabColor != ErrorColor.Value &&
                    (!information || source.Worksheet.TabColor.IsEmpty))
                    source.Worksheet.TabColor = color.Value;
            });

            if (kind == DiagnosticMessageKindData.Information)
            {
                setColors(InfoColor);
            }
            else if (kind == DiagnosticMessageKindData.Warning)
            {
                if (result != null) result.TotalWarningsCount++;
                setColors(WarnColor);
            }
            else
            {
                if (result != null) result.TotalErrorsCount++;
                setColors(ErrorColor);
            }

            // комментарий
            if (!string.IsNullOrEmpty(text))
            {
                var comment = source.Comment;
                comment.Text = text;
                comment.TopLeftCell = new AnchorCell(source.Column, source.Row, true);
                comment.BottomRightCell = new AnchorCell(source.Worksheet.Columns[source.Column.Index + 3], source.Worksheet.Rows[source.Row.Index + 3], false);
                comment.IsVisible = false;

                // добавить в журнал
                if (AddLogMessage != null)
                {
                    AddLogMessage(kind, "Ячейка {0}: {1}", new object[] { source, text });
                }
            }
        }

        /// <summary>
        /// Зарегистрировать информацию в ячейке
        /// </summary>
        public static void SetInfo(this ExcelCell source, ImportSheetCheckResultData result, string text)
        {
            source.SetResult(DiagnosticMessageKindData.Information, result, text);
        }

        /// <summary>
        /// Зарегистрировать предупреждение в ячейке
        /// </summary>
        public static void SetWarn(this ExcelCell source, ImportSheetCheckResultData result, string text)
        {
            source.SetResult(DiagnosticMessageKindData.Warning, result, text);
        }

        /// <summary>
        /// Зарегистрировать ошибку в ячейке
        /// </summary>
        public static void SetError(this ExcelCell source, ImportSheetCheckResultData result, string text)
        {
            source.SetResult(DiagnosticMessageKindData.Error, result, text);
        }

        /// <summary>
        /// Очистить лист от результатов предыдущей проверки
        /// </summary>
        public static void ClearCheckResults(this ExcelWorksheet source, int startRow, int startCol)
        {
            var columnCount = source.CalculateMaxUsedColumns();
            var colorEmpty = SpreadsheetColor.FromName(ColorName.Empty);
            
            foreach (var row in source.Rows)
            {
                if (row.Index < startRow)
                    continue;
                for (var col = startCol; col < columnCount; col++)
                {
                    var cell = source.Cells[row.Index, col];
                    if (!cell.Comment.Exists)
                        continue;

                    cell.Clear(ClearOptions.Comment);
                    cell.Style.FillPattern = null;
                    cell.Worksheet.Cells[row.Index, ColumnIndexOfRowHighlight - 1].Style.FillPattern = null;
                    source.TabColor = colorEmpty;
                }
            }
        }
    }
    #endregion CheckHelpers
}