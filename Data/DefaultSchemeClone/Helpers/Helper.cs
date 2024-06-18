

// класс описания листа
internal abstract class Sheet
{
    /// <summary>
    /// № п/п
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
        get { return Script.Instance.WorkbookNonExcel.Worksheets[_sheetIndex]; }

    }

    /// <summary>
    /// Значения по серийному номеру
    /// </summary>
    public readonly Dictionary<object, List<RowValueInfo>> RowValuesBySerialNumber = new Dictionary<object, List<RowValueInfo>>();

    /// <summary>
    /// Значения по серийному номеру + типу
    /// </summary>
    public readonly Dictionary<Tuple<object, object>, List<RowValueInfo>> RowValuesBySerialNumberAndType = new Dictionary<Tuple<object, object>, List<RowValueInfo>>();

    // добавить значение строки
    protected void AddRowValue(int rowIndex, Dictionary<int, object> rowValue, int colSerialNumber)
    {
        var rowValueInfo = new RowValueInfo(this, rowIndex, rowValue);
        RowValues.Add(rowValueInfo);
        var valueSerialNumber = rowValue[colSerialNumber];
        if (valueSerialNumber != null)
        {
            List<RowValueInfo> list;
            if (!RowValuesBySerialNumber.TryGetValue(valueSerialNumber, out list))
                RowValuesBySerialNumber.Add(valueSerialNumber, list = new List<RowValueInfo>());
            list.Add(rowValueInfo);
        }
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

// лист ТУ
internal class SIMCards : Sheet
{
    // индексация в Gembox с 0
    public const int START_ROW = 2 - 1;

    /// <summary>
    /// Номер телефона
    /// </summary>
    public const int COL_PHONE_NUMBER = 1 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// ICC ID
    /// </summary>
    public const int COL_ICC_ID = 2 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// Точка доступа (APN)
    /// </summary>
    public const int COL_APN = 3 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// IP адрес УСПД
    /// </summary>
    public const int COL_IP = 4 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// Абонент (оператор связи)
    /// </summary>
    public const int COL_CELLULAR_OPERATOR = 5 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// Установлено в
    /// </summary>
    public const int COL_INSTALLED_AT = 6 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// Комментарии
    /// </summary>
    public const int COL_COMMENTS = 7 - Helper.OFFSET_IF_NOT_EXCEL;

    /// <summary>
    /// Область видимости
    /// </summary>
    public const int COL_ISOLATION_LEVEL = 8 - Helper.OFFSET_IF_NOT_EXCEL;


    /// <summary>
    /// Максимальный номер столбца
    /// </summary>
    public const int COL_MAX_NUMBER = COL_ISOLATION_LEVEL;

    /// <summary>
    /// Конструктор
    /// </summary>
    /// <param name="sheet"></param>
    public SIMCards(ExcelWorksheet sheet)
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
                var rowValue = CheckHelper.CheckSIMCardsRow(result, Worksheet, row);
                AddRowValue(row.Index, rowValue, COL_PHONE_NUMBER);
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Script.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
                Worksheet.Cells[row.Index, COL_NUM].SetError(result, ex.Message);
                //break;
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
        // проверить уникальность номеров телефонов
        foreach (var rowValues in RowValuesBySerialNumberAndType
            .Where(x => x.Value.Count > 1))
        {
            foreach (var rowValue in rowValues.Value)
            {
                Worksheet.Cells[rowValue.RowIndex, COL_PHONE_NUMBER].SetError(result,
                    "Найдено БОЛЕЕ одного номера телефона");
            }
        }
    }
    
    /// <summary>
    /// Генератор ключей по полям
    /// </summary>
    private static class KeyGenerator
    {
        public abstract class CustomCubicle
        {
            protected readonly Dictionary<int, object> Columns;
            protected readonly HashSet<int> ExcludeColumns;
            
            protected CustomCubicle(IEnumerable<int> columns, IEnumerable<int> excludeColumns)
            {
                Columns = columns.ToDictionary(x => x, x => (object)null);
                ExcludeColumns = new HashSet<int>(excludeColumns ?? new int[0]);
            }
        }
    }
}

/// <summary>
/// Общие утилиты (внутренние)
/// </summary>
internal static class Helper
{
    /// <summary>
    /// Среда выполнения Excel или NonExcel
    /// </summary>
    public const bool IS_EXCEL = false;

    /// <summary>
    /// Смещение индексов для разных версий поставщика Excel
    /// </summary>
    public const byte OFFSET_IF_NOT_EXCEL = IS_EXCEL ? 0 : 1;

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

    /// <summary>
    /// Попытаться трактовать содержимое ячейки как Guid
    /// </summary>
    /// <param name="cellValue"></param>
    /// <returns></returns>
    public static Guid? TryGetGuidIdentifier(object cellValue)
    {
        var cellValueAsStr = cellValue as string;
        if (!string.IsNullOrEmpty(cellValueAsStr))
        {
            Guid tmpGuid;
            if (Guid.TryParse(cellValueAsStr, out tmpGuid))
                return tmpGuid;
        }
        return null;
    }

    /// <summary>
    /// Получить ссылку на сущность
    /// </summary>
    public static List<Sheet.RowValueInfo> GetRowValues<T>(Sheet sheet, EntityLink link, CheckHelpers.GetValueFunc<T> getValueFunc)
    {
        if (link == null)
            return null;

        List<Sheet.RowValueInfo> list;

        // только серийный номер
        if (string.IsNullOrEmpty(link.EntityType))
        {
            sheet.RowValuesBySerialNumber.TryGetValue(link.SerialNumber, out list);
            return list;
        }

        // серийный номер + тип
        var type = getValueFunc(link.EntityType);
        sheet.RowValuesBySerialNumberAndType.TryGetValue(Tuple.Create((object)link.SerialNumber, (object)type), out list);
        return list;
    }

    /// <summary>
    /// Получить ссылку на сущность
    /// </summary>
    public static Sheet.RowValueInfo GetRowValue<T>(Sheet sheet, EntityLink link, CheckHelpers.GetValueFunc<T> getValueFunc)
    {
        return (GetRowValues(sheet, link, getValueFunc) ?? new List<Sheet.RowValueInfo>()).FirstOrDefault();
    }

    /// <summary>
    /// Уведомить об изменении процента по строкам
    /// </summary>
    public static void NotifyPercentRow()
    {
        AsyncTasksHelper.PublishedTaskContext.NotifyPercent(PercentPerRow);
    }
}

