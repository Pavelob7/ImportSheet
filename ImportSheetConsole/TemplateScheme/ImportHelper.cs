using System;
using System.Collections.Generic;
using System.Linq;
using ImportSheetConsole.Global;
using ObjStudioClasses;

namespace ImportSheetConsole.TemplateScheme
{
    #region ImportHelper

    /// <summary>
    /// Утилиты импорта (внутренние)
    /// </summary>
    internal static class ImportHelper
    {
        /// <summary>
        /// Наименование классификатор произвольных элементов
        /// </summary>
        public const string ANY_CLASSIFIER_CAPTION = "Потребители";

        /// <summary>
        /// Лист 1
        /// </summary>
        [ThreadStatic]
        private static Sheet1 _sheet1;

        /// <summary>
        /// Классификатор произвольных элементов
        /// </summary>
        [ThreadStatic]
        private static ClassifierOfAnyItems _anyClassifier;
        
        /// <summary>
        /// Импортировать
        /// </summary>
        public static void Import(ImportSheetProcessedResultData result)
        {
            _sheet1 = (Sheet1)Helper.Sheets.First(x => x is Sheet1);
            
            // классификаторы
            _anyClassifier = ImportHelpers.FindOrCreate.ClassifierNodes.OfType<ClassifierOfAnyItems>(ANY_CLASSIFIER_CAPTION, null);

            // импортировать строку
            foreach (var rowValue in _sheet1.RowValues)
            {
                ImportSheet1Row(result, rowValue);
            }
        }
        
        /// <summary>
        /// Импортировать строку
        /// </summary>
        private static void ImportSheet1Row(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue)
        {
            var entities = rowValue.Values.ToDictionary(x => x.Key, x => (RDInstance)null);
            ImportHelpers.CreatedEntities = new List<RDInstance>();

            var anyParentClassifierItem = (RDInstance)_anyClassifier;

            try
            {
                //// Договор
                //entities[MeterPoints.COL_CONTRACT] = ImportHelpers.FindOrCreateEntity(result,
                //    new[] { _sheet1.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_CONTRACT] },
                //    rowValue.Values[MeterPoints.COL_CONTRACT] as string,
                //    anyParentClassifierItem,
                //    ImportHelpers.FindOrCreate.ClassifierNodes.OfType<ConsumerContract>);
                //anyParentClassifierItem = entities[MeterPoints.COL_CONTRACT] ?? anyParentClassifierItem;
            }
            catch (Exception ex)
            {
                // удалить созданные сущности
                foreach (var entity in ImportHelpers.CreatedEntities)
                {
                    try
                    {
                        entity.Remove();
                    }
                    catch
                    {
                        //
                    }
                }

                ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
            }
            finally
            {
                result.TotalCheckedLinesCount++;
                Helper.NotifyPercentRow();
            }
        }
    }

    #endregion ImportHelper
}