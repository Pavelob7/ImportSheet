
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
                        case "Сим-карты":
                            sheet = new SIMCards(excelSheet);
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
