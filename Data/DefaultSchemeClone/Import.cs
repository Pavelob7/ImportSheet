
            
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
