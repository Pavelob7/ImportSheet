
/// <summary>
/// Утилиты валидации (внутренние)
/// </summary>
internal static class CheckHelper
{
    // проверка листа
    public static Dictionary<int, object> CheckSIMCardsRow(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
    {
        var values = new Dictionary<int, object>();
        values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;
        
        for (var i = Sheet.COL_NUM + 1; i <= SIMCards.COL_MAX_NUMBER; i++)
            values[i] = null;


        // Номер телефона
        values[SIMCards.COL_PHONE_NUMBER] =
            sheet.Cells[row.Index, SIMCards.COL_PHONE_NUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
        // ICC ID
        values[SIMCards.COL_ICC_ID] =
            sheet.Cells[row.Index, SIMCards.COL_ICC_ID].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
        // APN
        values[SIMCards.COL_APN] =
            sheet.Cells[row.Index, SIMCards.COL_APN].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
        // IP
        values[SIMCards.COL_IP] =
            sheet.Cells[row.Index, SIMCards.COL_IP].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
        // Сотовый оператор
        values[SIMCards.COL_CELLULAR_OPERATOR] =
            sheet.Cells[row.Index, SIMCards.COL_CELLULAR_OPERATOR].GetWithCheck(result, Helpers.GetCellularOperator, CheckHelpers.DefaultCheck);
        // Установлено в
        values[SIMCards.COL_INSTALLED_AT] =
            sheet.Cells[row.Index, SIMCards.COL_INSTALLED_AT].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
        // Комментарии
        values[SIMCards.COL_COMMENTS] =
            sheet.Cells[row.Index, SIMCards.COL_COMMENTS].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
        // Область видимости
        values[SIMCards.COL_ISOLATION_LEVEL] =
            sheet.Cells[row.Index, SIMCards.COL_ISOLATION_LEVEL].GetWithCheck(result, Helpers.GetIsolationLevel, CheckHelpers.DefaultCheck);

        // Проверка на обязательное заполнение номера телефона
        if (values[SIMCards.COL_PHONE_NUMBER] == null)
        {
            sheet.Cells[row.Index, SIMCards.COL_PHONE_NUMBER].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
        }
        if (values[SIMCards.COL_PHONE_NUMBER] != null)
        {
            string phoneNubmer = values[SIMCards.COL_PHONE_NUMBER] as string;
            bool isExist = SimCard.Where(x => x.AttributePhoneNumber == phoneNubmer).Any();
            if (isExist)
            {
                sheet.Cells[row.Index, SIMCards.COL_PHONE_NUMBER].SetError(result, $"Номер {phoneNubmer} уже числится в справочнике");
                Script.Instance.AddLogInfo($"Ошибка! Номер {phoneNubmer} уже числится в справочнике");
            }
        }

        // Проверка оператора
        if (values[SIMCards.COL_CELLULAR_OPERATOR] == null)
        {
            sheet.Cells[row.Index, SIMCards.COL_CELLULAR_OPERATOR].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
            Script.Instance.AddLogInfo($"Ошибка! оператор связи не может быть пустым");
        }
        if (values[SIMCards.COL_CELLULAR_OPERATOR] != null)
        {
            var cellularOperator = values[SIMCards.COL_CELLULAR_OPERATOR] as CellularOperator;
            bool isExists = CellularOperator.Where(x => x == cellularOperator).Any();
            if (!isExists)
            {
                sheet.Cells[row.Index, SIMCards.COL_CELLULAR_OPERATOR].SetError(result, "Оператор связи не найден");
                Script.Instance.AddLogInfo($"Ошибка! оператор связи не найден");
            }
        }

        // Проверка области видимости
        if (values[SIMCards.COL_ISOLATION_LEVEL] != null)
        {
            var isolationLevel = values[SIMCards.COL_ISOLATION_LEVEL] as IsolationLevel;
            bool isExists = IsolationLevel.Where(x => x == isolationLevel).Any();
            if (!isExists)
            {
                sheet.Cells[row.Index, SIMCards.COL_ISOLATION_LEVEL].SetError(result, "Область видимости не найдена");
                Script.Instance.AddLogInfo($"Ошибка! область видимости не найдена");
            }
        }

        return values;
    }

    /// <summary>
    /// Проверить ссылку на сущность и возвратить значение
    /// </summary>
    public static Sheet.RowValueInfo CheckEntityLinkAndGetRowValue<T>(this ExcelCell source, ImportSheetCheckResultData result,
        Sheet sheetWhereEntities, object rowValue, CheckHelpers.GetValueFunc<T> getValueFunc,
        string entityTypeCaption, string serialNumberCaption = "серийный номер", string typeCaption = "тип")
    {
        var link = rowValue as EntityLink;
        if (link == null)
            return null;
        var rowValues = Helper.GetRowValues(sheetWhereEntities, link, getValueFunc);
        if (source.CheckEntityLinks(result, new[] { Tuple.Create(rowValues, link, entityTypeCaption, serialNumberCaption, typeCaption) }))
            return rowValues.FirstOrDefault();
        return null;
    }

    /// <summary>
    /// Проверить ссылку на сущность
    /// </summary>
    public static bool CheckEntityLink<T>(this ExcelCell source, ImportSheetCheckResultData result, 
        Sheet sheetWhereEntities, object rowValue, CheckHelpers.GetValueFunc<T> getValueFunc,
        string entityTypeCaption, string serialNumberCaption = "серийный номер", string typeCaption = "тип")
    {
        var link = rowValue as EntityLink;
        if (link == null) 
            return false;
        var rowValues = Helper.GetRowValues(sheetWhereEntities, link, getValueFunc);
        return source.CheckEntityLinks(result, new[] { Tuple.Create(rowValues, link, entityTypeCaption, serialNumberCaption, typeCaption) });
    }

    /// <summary>
    /// Проверить ссылку на сущность
    /// </summary>
    public static bool CheckEntityLinks(this ExcelCell source, ImportSheetCheckResultData result,
        Tuple<List<Sheet.RowValueInfo>, EntityLink, string, string, string>[] rowEntityValues)
    {
        var count = 0;
        var text = new StringBuilder();
        foreach (var rowEntityValue in rowEntityValues)
        {
            var rowValues = rowEntityValue.Item1;
            var link = rowEntityValue.Item2;
            var entityTypeCaption = rowEntityValue.Item3;
            var serialNumberCaption = rowEntityValue.Item4;
            var typeCaption = rowEntityValue.Item5;
            if (rowValues == null)
                text.AppendLine(
                    string.IsNullOrEmpty(link.EntityType)
                        ? string.Format("В документе не удалось найти {0}. Объект поиска: {1}", entityTypeCaption, serialNumberCaption)
                        : string.Format("В документе не удалось найти {0}. Объект поиска: {1} и {2}",
                            entityTypeCaption, serialNumberCaption, typeCaption));
            else if (rowValues.Count > 1)
                text.AppendLine(
                    string.IsNullOrEmpty(link.EntityType)
                        ? string.Format("В документе найдено БОЛЕЕ одного {0}. Объект поиска: {1}", entityTypeCaption, serialNumberCaption)
                        : string.Format("В документе найдено БОЛЕЕ одного {0}. Объект поиска: {1} и {2}",
                            entityTypeCaption, serialNumberCaption, typeCaption));
            else
            {
                count++;
                text.AppendLine(
                    string.IsNullOrEmpty(link.EntityType)
                        ? string.Format("В документе найдено {0}. Объект поиска: {1}", entityTypeCaption, serialNumberCaption)
                        : string.Format("В документе найдено {0}. Объект поиска: {1} и {2}", entityTypeCaption, serialNumberCaption, typeCaption));
            }
        }

        if (count == 1) return true;

        source.SetError(result, text.ToString());
        return false;
    }
}
