

/// <summary>
/// Утилиты валидации (внутренние)
/// </summary>
internal static class ImportHelper
{
   [ThreadStatic]
    private static SIMCards _simCards;
    [ThreadStatic]
    private static HashSet<IsolationLevel> _visibleIsolationLevels;
    [ThreadStatic]
    private static RDAttributeValuesWithCreation<CommonUserDirectoryItem> _existingCards;

    /// <summary>
    /// Импортировать
    /// </summary>
    public static void Import(ImportSheetProcessedResultData result)
    {
        _existingCards = DirectoryOfCommonUserItems.OnlyInstance.AttributeCommonUserItems;
        _visibleIsolationLevels = IsolationLevel.GetInstances().ToHashSet();

        try
        {
            ImportInternal(result);
        }
        finally
        {
            _existingCards = null;
        }
    }

    /// <summary>
    /// Импортировать, внутренний вызов, обрамленный кешированием
    /// </summary>
    private static void ImportInternal(ImportSheetProcessedResultData result)
    {
        _simCards = (SIMCards)Helper.Sheets.First(x => x is SIMCards);

        foreach (var rowValue in _simCards.RowValues)
        {
            ImportHelpers.Init();
            ImportPhoneNumberRow(result, rowValue);
        }
    }

    /// <summary>
    /// Импортировать строку
    /// </summary>
    private static void ImportPhoneNumberRow(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue)
    {
        var isolationLevels = rowValue.Values[SIMCards.COL_ISOLATION_LEVEL] as IsolationLevel[];
        if (isolationLevels != null)
        {
            var allowedIsolationLevels = isolationLevels.Where(_visibleIsolationLevels.Contains).ToArray();
            if (allowedIsolationLevels.Length != isolationLevels.Length)
            {
                _simCards.Worksheet.Cells[rowValue.RowIndex, SIMCards.COL_ISOLATION_LEVEL].SetWarn(result,
                    "Не все области видимости могут быть применены для текущего пользователя");
                isolationLevels = allowedIsolationLevels.Any() ? allowedIsolationLevels : null;
                rowValue.Values[SIMCards.COL_ISOLATION_LEVEL] = isolationLevels;
            }
        }

        var phoneNumber = rowValue.Values[SIMCards.COL_PHONE_NUMBER] as string;
        var iccId = rowValue.Values[SIMCards.COL_ICC_ID] as string;
        var apn = rowValue.Values[SIMCards.COL_APN] as string;
        var ipAddress = rowValue.Values[SIMCards.COL_IP] as string;
        var cellularOperator = rowValue.Values[SIMCards.COL_CELLULAR_OPERATOR] as CellularOperator;
        var installedAt = rowValue.Values[SIMCards.COL_INSTALLED_AT] as string;
        var comments = rowValue.Values[SIMCards.COL_COMMENTS] as string;
        var isolationLevel = rowValue.Values[SIMCards.COL_ISOLATION_LEVEL] as IsolationLevel;

        try
        {
            if (phoneNumber == null && cellularOperator == null)
            {
                _simCards.Worksheet.Cells[rowValue.RowIndex, SIMCards.COL_ISOLATION_LEVEL].SetError(result,
                        $"Номер телефона и Оператор не могут быть null");
            }
            else
            {
                _existingCards = DirectoryOfCommonUserItems.OnlyInstance.AttributeCommonUserItems;
                var newSimCard = _existingCards.AppendNew<SimCard>().Value;

                newSimCard.AttributePhoneNumber = phoneNumber;

                if (!string.IsNullOrEmpty(iccId))
                    newSimCard.AttributeSimId = iccId;

                if (!string.IsNullOrEmpty(apn))
                    newSimCard.AttributeAccessPointApn = apn;

                if (!string.IsNullOrEmpty(ipAddress))
                    newSimCard.AttributeIpAddress = ipAddress;

                if (cellularOperator != null)
                    newSimCard.AttributeCellularOperator = cellularOperator;

                if (!string.IsNullOrEmpty(comments))
                    newSimCard.AttributeComment = comments;

                if (isolationLevel != null)
                    newSimCard.AttributeIsolationLevels = new[] { isolationLevel };
            }
        }
        catch (System.Threading.ThreadInterruptedException)
        {
            throw;
        }
        catch (Exception ex)
        {
            var columnCount = rowValue.Sheet.Worksheet.CalculateMaxUsedColumns();
            for (var i = Sheet.COL_NUM; i < columnCount; i++)
            {
                rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, i].SetError(result, null);
            }

            // удалить созданные сущности
            foreach (var entity in ImportHelpers.CreatedEntities)
            {
                try
                {
                    entity.Remove();
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch
                {
                    //
                }
            }

            Script.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
        }
        finally
        {
            result.TotalCheckedLinesCount++;
            Helper.NotifyPercentRow();
        }
    }

    /// <summary>
    /// Получить Оператора по наименованию
    /// </summary>
    /// <param name="caption">Наименование</param>
    /// <returns></returns>
    private static CellularOperator GetCellularOperator(string caption)
    {
        if (string.IsNullOrEmpty(caption))
            return null;

        return CellularOperator.FirstOrDefault(co => co.AttributeCaption == caption);
    }

    /// <summary>
    /// Получение Области видимости по наименованию
    /// </summary>
    /// <param name="caption">Наименование</param>
    /// <returns></returns>
    private static IsolationLevel GetIsolationLevel(string caption)
    {
        if (string.IsNullOrEmpty(caption))
            return null;

        return IsolationLevel.FirstOrDefault(x => x.Caption == caption);
    }
}

