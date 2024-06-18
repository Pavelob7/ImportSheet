using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GemBox.Spreadsheet;
using ImportSheetConsole.Global;
using ObjStudioClasses;

namespace ImportSheetConsole.DefaultScheme
{
    #region Helper

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
            get { return ScriptImpl.Instance.WorkbookNonExcel.Worksheets[_sheetIndex]; }

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
        protected void AddRowValue(int rowIndex, Dictionary<int, object> rowValue, int colSerialNumber, int colType)
        {
            var rowValueInfo = new RowValueInfo(this, rowIndex, rowValue);
            RowValues.Add(rowValueInfo);
            var valueSerialNumber = rowValue[colSerialNumber];
            var valueType = rowValue[colType];
            if (valueSerialNumber != null)
            {
                List<RowValueInfo> list;
                if (!RowValuesBySerialNumber.TryGetValue(valueSerialNumber, out list))
                    RowValuesBySerialNumber.Add(valueSerialNumber, list = new List<RowValueInfo>());
                list.Add(rowValueInfo);
                if (valueType != null)
                {
                    var key = Tuple.Create(valueSerialNumber, valueType);
                    if (!RowValuesBySerialNumberAndType.TryGetValue(key, out list))
                        RowValuesBySerialNumberAndType.Add(key, list = new List<RowValueInfo>());
                    list.Add(rowValueInfo);
                }
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
    internal class MeterPoints : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 4 - 1;

        /// <summary>
        /// ПЭС
        /// </summary>
        public const int COL_PES = 2 - Helper.OFFSET_IF_NOT_EXCEL;
        
        /// <summary>
        /// РЭС
        /// </summary>
        public const int COL_RES = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        // ПС	

        /// <summary>
        /// ПС
        /// </summary>
        public const int COL_SUBSTATION_SUBSTATION = 4 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Уровень напряжения ПС
        /// </summary>
        public const int COL_SUBSTATION_VOLTAGE = 5 - Helper.OFFSET_IF_NOT_EXCEL;

        // ПС - Высокая сторона

        /// <summary>
        /// ПС - Высокая сторона - Уровень напряжения РУ
        /// </summary>
        public const int COL_SUBSTATION_HI_VOLTAGE = 6 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПС - Высокая сторона - Секция шин
        /// </summary>
        public const int COL_SUBSTATION_HI_BUSBAR = 7 - Helper.OFFSET_IF_NOT_EXCEL;
        
        /// <summary>
        /// ПС - Высокая сторона - Ячейка
        /// </summary>
        public const int COL_SUBSTATION_HI_CUBICLE_NAME = 8 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПС - Высокая сторона - Тип ячейки
        /// </summary>
        public const int COL_SUBSTATION_HI_CUBICLE_TYPE = 9 - Helper.OFFSET_IF_NOT_EXCEL;

        // ПС - Низкая сторона

        /// <summary>
        /// ПС - Низкая сторона - Уровень напряжения РУ
        /// </summary>
        public const int COL_SUBSTATION_LO_VOLTAGE = 10 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПС - Низкая сторона - Секция шин
        /// </summary>
        public const int COL_SUBSTATION_LO_BUSBAR = 11 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПС - Низкая сторона - Ячейка
        /// </summary>
        public const int COL_SUBSTATION_LO_CUBICLE_NAME = 12 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПС - Низкая сторона - Тип ячейки
        /// </summary>
        public const int COL_SUBSTATION_LO_CUBICLE_TYPE = 13 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПС - Низкая сторона - Линия/фидер
        /// </summary>
        public const int COL_SUBSTATION_LO_POWERLINE = 14 - Helper.OFFSET_IF_NOT_EXCEL;

        // РП

        /// <summary>
        /// РП
        /// </summary>
        public const int COL_SWITCHGEAR_SUBSTATION_SUBSTATION = 15 - Helper.OFFSET_IF_NOT_EXCEL;
        
        /// <summary>
        /// РП - Секция шин
        /// </summary>
        public const int COL_SWITCHGEAR_SUBSTATION_BUSBAR = 16 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// РП - Ячейка, входящая от ПС
        /// </summary>
        public const int COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN = 17 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// РП - Ячейка, отходящая в ТП
        /// </summary>
        public const int COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT = 18 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// РП - Линия/фидер
        /// </summary>
        public const int COL_SWITCHGEAR_SUBSTATION_POWERLINE = 19 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТП
        /// </summary>
        public const int COL_LO_SUBSTATION_SUBSTATION = 20 - Helper.OFFSET_IF_NOT_EXCEL;

        // ТП - Высокая сторона

        /// <summary>
        /// ТП - Высокая сторона - Секция шин
        /// </summary>
        public const int COL_LO_SUBSTATION_HI_BUSBAR = 21 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТП - Высокая сторона - Ячейка
        /// </summary>
        public const int COL_LO_SUBSTATION_HI_CUBICLE_NAME = 22 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТП - Высокая сторона - Тип ячейки
        /// </summary>
        public const int COL_LO_SUBSTATION_HI_CUBICLE_TYPE = 23 - Helper.OFFSET_IF_NOT_EXCEL;

        // ТП - Низкая сторона

        /// <summary>
        /// ТП - Низкая сторона - Секция шин
        /// </summary>
        public const int COL_LO_SUBSTATION_LO_BUSBAR = 24 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТП - Низкая сторона - Ячейка
        /// </summary>
        public const int COL_LO_SUBSTATION_LO_CUBICLE_NAME = 25 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТП - Низкая сторона - Тип ячейки
        /// </summary>
        public const int COL_LO_SUBSTATION_LO_CUBICLE_TYPE = 26 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТП - Низкая сторона - Линия/фидер
        /// </summary>
        public const int COL_LO_SUBSTATION_LO_POWERLINE = 27 - Helper.OFFSET_IF_NOT_EXCEL;

        // Адрес

        /// <summary>
        /// Адрес - Адрес ФИАС
        /// </summary>
        public const int COL_ADDRESS_FIAS = 28 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Адрес - Квартира
        /// </summary>
        public const int COL_ADDRESS_FLAT = 29 - Helper.OFFSET_IF_NOT_EXCEL;

        // ПУ

        /// <summary>
        /// ПУ - Тип ПУ
        /// </summary>
        public const int COL_METER_TYPE = 30 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПУ - Серийный номер
        /// </summary>
        public const int COL_METER_SERIALNUMBER = 31 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПУ - Дата выпуска
        /// </summary>
        public const int COL_METER_CREATEDATE = 32 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПУ - Дата установки
        /// </summary>
        public const int COL_METER_INSTALLDATE = 33 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПУ - Дата последней поверки
        /// </summary>
        public const int COL_METER_CHECKDATE = 34 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПУ - Класс точности
        /// </summary>
        public const int COL_METER_ACCURACYCLASS = 35 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ПУ - Часовой пояс
        /// </summary>
        public const int COL_METER_TIMEZONE = 36 - Helper.OFFSET_IF_NOT_EXCEL;

        // Трансформаторы

        /// <summary>
        /// Трансформаторы - Ктт
        /// </summary>
        public const int COL_TRANS_KTT = 37 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Трансформаторы - Ктн
        /// </summary>
        public const int COL_TRANS_KTN = 38 - Helper.OFFSET_IF_NOT_EXCEL;
        
        /// <summary>
        /// ТТ, фаза 1
        /// </summary>
        public const int COL_TRANS_CURRENT_PHASE1 = 39 - Helper.OFFSET_IF_NOT_EXCEL;
        
        /// <summary>
        /// ТТ, фаза 2
        /// </summary>
        public const int COL_TRANS_CURRENT_PHASE2 = 40 - Helper.OFFSET_IF_NOT_EXCEL;
        
        /// <summary>
        /// ТТ, фаза 3
        /// </summary>
        public const int COL_TRANS_CURRENT_PHASE3 = 41 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - ТН, фаза 1
        /// </summary>
        public const int COL_TRANS_VOLTAGE_PHASE1 = 42 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - ТН, фаза 2
        /// </summary>
        public const int COL_TRANS_VOLTAGE_PHASE2 = 43 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - ТН, фаза 3
        /// </summary>
        public const int COL_TRANS_VOLTAGE_PHASE3 = 44 - Helper.OFFSET_IF_NOT_EXCEL;

        // Связь с ПУ

        /// <summary>
        /// Связь с ПУ - Связной номер
        /// </summary>
        public const int COL_METER_NETWORKNUMBER = 45 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с ПУ - Пользователь
        /// </summary>
        public const int COL_METER_LOGIN = 46 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с ПУ - Пароль
        /// </summary>
        public const int COL_METER_PASSWORD = 47 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с ПУ - Маршрут опроса
        /// </summary>
        public const int COL_METER_ROUTE = 48 - Helper.OFFSET_IF_NOT_EXCEL;

        ///// <summary>
        ///// Связь с ПУ - Специальные параметры
        ///// </summary>
        //public const int COL_METER_PARAMS = 49 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Параметры, которые имеются на листе
        /// </summary>
        public static readonly Guid[] KnownMeterParams =
        {
            RDMetadataGuids.Equipment.InstanceAttributeSerialNumber,
            RDMetadataGuids.Equipment.InstanceAttributeReleaseDate,
            RDMetadataGuids.ElectricityMeter.InstanceAttributeInstallDate,
            RDMetadataGuids.IMetrologicalCheckingInstance.AttributeLastCalibrationDate,
            RDMetadataGuids.ElectricityMeter.InstanceAttributeAccuracyClass,
            RDMetadataGuids.IEntityWithTimeZone.AttributeTimeZone,
            RDMetadataGuids.IEquipmentWithNetworkId.AttributeNetworkId,
            RDMetadataGuids.IEquipmentWithUserAndPassword.AttributeUser,
            RDMetadataGuids.IEquipmentWithPassword.AttributePassword,
        };

        /// <summary>
        /// УСПД
        /// </summary>
        public const int COL_RTU = 49 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Абонент
        /// </summary>
        public const int COL_CONSUMER = 50 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Уровень изоляции
        /// </summary>
        public const int COL_ISOLATION_LEVEL = 51 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Тариф
        /// </summary>
        public const int COL_TARIFF = 52 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Максимальный номер столбца
        /// </summary>
        public const int COL_MAX_NUMBER = COL_TARIFF;
        
        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public MeterPoints(ExcelWorksheet sheet)
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
                    var rowValue = CheckHelper.CheckMeterPointsRow(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue, COL_METER_SERIALNUMBER, COL_METER_TYPE);
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
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
            // проверить ссылки на сущности
            var rtus = Helper.Sheets.First(x => x is Rtus);
            var transCurrents = Helper.Sheets.First(x => x is TransCurrents);
            var transVoltages = Helper.Sheets.First(x => x is TransVoltages);
            var naturalPersons = Helper.Sheets.First(x => x is NaturalPersons);
            var legalEntities = Helper.Sheets.First(x => x is LegalEntities);

            // проверить уникальность ПУ
            foreach (var rowValues in RowValuesBySerialNumberAndType
                .Where(x => x.Value.Count > 1))
            {
                foreach (var rowValue in rowValues.Value)
                {
                    Worksheet.Cells[rowValue.RowIndex, COL_METER_SERIALNUMBER].SetError(result,
                        "Найдено БОЛЕЕ одного ПУ с указанным серийным номером и типом");
                }
            }

            // проверить корректность заполнение по разным критериям
            CheckPowerLineConnection(result);
            CheckMeterConnection(result);

            //

            var usedTransCurrents = new Dictionary<RowValueInfo, string>();
            var usedTransVoltages = new Dictionary<RowValueInfo, string>();

            var rowIndex = START_ROW;
            foreach (var rowValue in RowValues)
            {
                // ТТ

                var checkAndAddUsedTransCurrent = new Action<ExcelCell, RowValueInfo>((source, transRowValue) =>
                {
                    if (transRowValue == null)
                        return;
                    string cellAddress;
                    if (usedTransCurrents.TryGetValue(transRowValue, out cellAddress))
                        source.SetError(result, string.Format("ТТ уже указан для другой ТУ. Aдрес ячейки ТТ: {0}", cellAddress));
                    else
                        usedTransCurrents[transRowValue] = source.ToString();
                });

                // ТТ, фаза 1
                var tt1 = Worksheet.Cells[rowIndex, COL_TRANS_CURRENT_PHASE1].CheckEntityLinkAndGetRowValue(result,
                    transCurrents, rowValue.Values[COL_TRANS_CURRENT_PHASE1], Helpers.GetString, "ТТ");
                checkAndAddUsedTransCurrent(Worksheet.Cells[rowIndex, COL_TRANS_CURRENT_PHASE1], tt1);
                // ТТ, фаза 2
                var tt2 = Worksheet.Cells[rowIndex, COL_TRANS_CURRENT_PHASE2].CheckEntityLinkAndGetRowValue(result,
                    transCurrents, rowValue.Values[COL_TRANS_CURRENT_PHASE2], Helpers.GetString, "ТТ");
                checkAndAddUsedTransCurrent(Worksheet.Cells[rowIndex, COL_TRANS_CURRENT_PHASE2], tt2);
                // ТТ, фаза 3
                var tt3 = Worksheet.Cells[rowIndex, COL_TRANS_CURRENT_PHASE3].CheckEntityLinkAndGetRowValue(result,
                    transCurrents, rowValue.Values[COL_TRANS_CURRENT_PHASE3], Helpers.GetString, "ТТ");
                checkAndAddUsedTransCurrent(Worksheet.Cells[rowIndex, COL_TRANS_CURRENT_PHASE3], tt3);

                // ТН
                
                // ТН, фаза 1
                var tn1 = Worksheet.Cells[rowIndex, COL_TRANS_VOLTAGE_PHASE1].CheckEntityLinkAndGetRowValue(result,
                    transVoltages, rowValue.Values[COL_TRANS_VOLTAGE_PHASE1], Helpers.GetString, "ТН");
                // ТН, фаза 2
                var tn2 = Worksheet.Cells[rowIndex, COL_TRANS_VOLTAGE_PHASE2].CheckEntityLinkAndGetRowValue(result,
                    transVoltages, rowValue.Values[COL_TRANS_VOLTAGE_PHASE2], Helpers.GetString, "ТН");
                // ТН, фаза 3
                var tn3 = Worksheet.Cells[rowIndex, COL_TRANS_VOLTAGE_PHASE3].CheckEntityLinkAndGetRowValue(result,
                    transVoltages, rowValue.Values[COL_TRANS_VOLTAGE_PHASE3], Helpers.GetString, "ТН");
                
                // УCПД

                Worksheet.Cells[rowIndex, COL_RTU].CheckEntityLink(result, rtus, rowValue.Values[COL_RTU], Helpers.GetRtuClass, "УСПД");

                // Абонент

                var link = rowValue.Values[COL_CONSUMER] as EntityLink;
                if (link != null)
                {
                    Worksheet.Cells[rowIndex, COL_CONSUMER].CheckEntityLinks(result,
                        new[]
                        {
                            Tuple.Create(Helper.GetRowValues(naturalPersons, link, Helpers.GetString), link, "Физическое лицо", "номер лицевого счета", "фамилия"),
                            Tuple.Create(Helper.GetRowValues(legalEntities, link, Helpers.GetString), link, "Юридическое лицо", "номер лицевого счета", "наименование"),
                        });
                }

                rowIndex++;
            }

            // проверить и уточнить маршруты
            foreach (var rowValue in RowValues)
            {
                var routes = new List<ImportRoute>(rowValue.Values[COL_METER_ROUTE] as ImportRoute[] ?? new ImportRoute[0]);
                // добавить маршрут через УСПД к ПУ
                ChannelizingEquipmentClassInfo rtuClassInfo = null;
                var link = rowValue.Values[COL_RTU] as EntityLink;
                if (link != null)
                {
                    var rtuRowValue = Helper.GetRowValue(rtus, link, Helpers.GetRtuClass);
                    rtuClassInfo = rtuRowValue == null ? null : rtuRowValue.Values[Rtus.COL_TYPE] as ChannelizingEquipmentClassInfo;
                    // актуализировать список маршрутов ПУ
                    if (rtuClassInfo != null)
                    {
                        routes.Add(new ImportRoute(NonDirectRouteClassInfo.Get())
                        {
                            Params =
                            {
                                { RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, link }
                            },
                            OtherParams = link.OtherParams
                        });
                        rowValue.Values[COL_METER_ROUTE] = routes.ToArray();
                    }
                }
                // пройти по всем маршрутам
                var errors = new Dictionary<int, StringBuilder>
                {
                    // УСПД
                    { COL_RTU, new StringBuilder() },
                    // другое каналообразующее оборудование
                    { COL_METER_ROUTE, new StringBuilder() }
                };
                foreach (var route in routes)
                {
                    EquipmentClassInfo routeEquipmentClassInfo = null;
                    // возможно есть маршрут через промежуточное оборудование
                    object routeEntityLink;
                    if (route.Params.TryGetValue(RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, out routeEntityLink) && 
                        routeEntityLink is EntityLink)
                    {
                        // поиск на листе УСПД
                        var rtuEntityLinkRowValue = Helper.GetRowValues(rtus, (EntityLink)routeEntityLink, Helpers.GetRtuClass);
                        if (rtuEntityLinkRowValue != null && rtuEntityLinkRowValue.Any())
                        {
                            if (Worksheet.Cells[rowValue.RowIndex, COL_METER_ROUTE].CheckEntityLinks(result,
                                new[] { Tuple.Create(rtuEntityLinkRowValue, (EntityLink)routeEntityLink, "Каналообразующее оборудование", "Серийный номер", "Тип") }))
                                routeEquipmentClassInfo = rtuEntityLinkRowValue.First().Values[Rtus.COL_TYPE] as ChannelizingEquipmentClassInfo;
                        }
                        
                        // поиск на листе ТУ
                        if (routeEquipmentClassInfo == null)
                        {
                            var meterEntityLinkRowValue = Helper.GetRowValues(this, (EntityLink)routeEntityLink, Helpers.GetMeterClass);
                            if (meterEntityLinkRowValue != null && meterEntityLinkRowValue.Any())
                            {
                                if (Worksheet.Cells[rowValue.RowIndex, COL_METER_ROUTE].CheckEntityLinks(result,
                                    new[] { Tuple.Create(meterEntityLinkRowValue, (EntityLink)routeEntityLink, "Каналообразущее устройство", "Серийный номер", "Тип") }))
                                    routeEquipmentClassInfo = (EquipmentClassInfo)(meterEntityLinkRowValue.First().Values[COL_METER_TYPE] as IChannelizingEquipmentClass);
                            }
                        }

                        // ошибка
                        if (routeEquipmentClassInfo == null)
                        {
                            Worksheet.Cells[rowValue.RowIndex, COL_METER_ROUTE].CheckEntityLinks(result,
                                new[] { Tuple.Create(rtuEntityLinkRowValue, (EntityLink)routeEntityLink, "Каналообразующее оборудование", "Серийный номер", "Тип") });
                        }
                    }

                    // уточнить тип маршрута
                    string errorText;
                    if (!route.ParseFor(rowValue.Values[COL_METER_TYPE] as ElectricityMeterClassInfo, routeEquipmentClassInfo, out errorText))
                    {
                        var errorBuilder = errors[routeEquipmentClassInfo == rtuClassInfo ? COL_RTU : COL_METER_ROUTE];
                        errorBuilder.AppendLine(errorText);
                    }
                }
                foreach (var errorBuilder in errors)
                {
                    if (errorBuilder.Value.Length > 0)
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, errorBuilder.Key].SetError(result, errorBuilder.Value.ToString());
                    }
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

            /// <summary>
            /// ПС - Высокая сторона
            /// </summary>
            public class SubstationHi : CustomCubicle
            {
                public SubstationHi(params int[] excludeColumns)
                    : base(new[]
                    {
                        COL_PES, 
                        COL_RES,
                        COL_SUBSTATION_SUBSTATION,
                        COL_SUBSTATION_VOLTAGE,
                        COL_SUBSTATION_HI_VOLTAGE,
                        COL_SUBSTATION_HI_BUSBAR,
                        COL_SUBSTATION_HI_CUBICLE_NAME,
                        COL_SUBSTATION_HI_CUBICLE_TYPE
                    }, excludeColumns)
                {
                }

                public Tuple<Tuple<object, object>, Tuple<object, object, object, object, object, object>> GetTuple(RowValueInfo rowValue)
                {
                    foreach (var key in Columns.Keys.ToArray())
                        Columns[key] = ExcludeColumns.Contains(key) ? null : rowValue.Values[key];
                    
                    return Tuple.Create(
                        Tuple.Create(
                            Columns[COL_PES], 
                            Columns[COL_RES]),
                        Tuple.Create(
                            Columns[COL_SUBSTATION_SUBSTATION],
                            Columns[COL_SUBSTATION_VOLTAGE],
                            Columns[COL_SUBSTATION_HI_VOLTAGE],
                            Columns[COL_SUBSTATION_HI_BUSBAR],
                            Columns[COL_SUBSTATION_HI_CUBICLE_NAME],
                            Columns[COL_SUBSTATION_HI_CUBICLE_TYPE])
                    );
                }

                public bool GetWhere(Tuple<Tuple<object, object>, Tuple<object, object, object, object, object, object>> source)
                {
                    Columns[COL_PES] = source.Item1.Item1;
                    Columns[COL_RES] = source.Item1.Item2;
                    Columns[COL_SUBSTATION_SUBSTATION] = source.Item2.Item1;
                    Columns[COL_SUBSTATION_VOLTAGE] = source.Item2.Item2;
                    Columns[COL_SUBSTATION_HI_VOLTAGE] = source.Item2.Item3;
                    Columns[COL_SUBSTATION_HI_BUSBAR] = source.Item2.Item4;
                    Columns[COL_SUBSTATION_HI_CUBICLE_NAME] = source.Item2.Item5;
                    Columns[COL_SUBSTATION_HI_CUBICLE_TYPE] = source.Item2.Item6;

                    return Columns
                        .Where(x => !ExcludeColumns.Contains(x.Key))
                        .All(x => x.Value != null);
                }
            }

            /// <summary>
            /// ПС - Низкая сторона
            /// </summary>
            public class SubstationLo : CustomCubicle
            {
                public SubstationLo(params int[] excludeColumns)
                    : base(new[]
                    {
                        COL_PES, 
                        COL_RES,
                        COL_SUBSTATION_SUBSTATION,
                        COL_SUBSTATION_VOLTAGE,
                        COL_SUBSTATION_LO_VOLTAGE,
                        COL_SUBSTATION_LO_BUSBAR,
                        COL_SUBSTATION_LO_CUBICLE_NAME,
                        COL_SUBSTATION_LO_CUBICLE_TYPE,
                        COL_SUBSTATION_LO_POWERLINE
                    }, excludeColumns)
                {
                }

                public Tuple<Tuple<object, object>, Tuple<object, object, object, object, object, object, object>> GetTuple(RowValueInfo rowValue)
                {
                    foreach (var key in Columns.Keys.ToArray())
                        Columns[key] = ExcludeColumns.Contains(key) ? null : rowValue.Values[key];
                    
                    return Tuple.Create(
                        Tuple.Create(
                            Columns[COL_PES], 
                            Columns[COL_RES]),
                        Tuple.Create(
                            Columns[COL_SUBSTATION_SUBSTATION], 
                            Columns[COL_SUBSTATION_VOLTAGE], 
                            Columns[COL_SUBSTATION_LO_VOLTAGE],
                            Columns[COL_SUBSTATION_LO_BUSBAR],
                            Columns[COL_SUBSTATION_LO_CUBICLE_NAME],
                            Columns[COL_SUBSTATION_LO_CUBICLE_TYPE],
                            Columns[COL_SUBSTATION_LO_POWERLINE])
                    );
                }

                public bool GetWhere(Tuple<Tuple<object, object>, Tuple<object, object, object, object, object, object, object>> source)
                {
                    Columns[COL_PES] = source.Item1.Item1;
                    Columns[COL_RES] = source.Item1.Item2;
                    Columns[COL_SUBSTATION_SUBSTATION] = source.Item2.Item1;
                    Columns[COL_SUBSTATION_VOLTAGE] = source.Item2.Item2;
                    Columns[COL_SUBSTATION_LO_VOLTAGE] = source.Item2.Item3;
                    Columns[COL_SUBSTATION_LO_BUSBAR] = source.Item2.Item4;
                    Columns[COL_SUBSTATION_LO_CUBICLE_NAME] = source.Item2.Item5;
                    Columns[COL_SUBSTATION_LO_CUBICLE_TYPE] = source.Item2.Item6;
                    Columns[COL_SUBSTATION_LO_POWERLINE] = source.Item2.Item7;

                    return Columns
                        .Where(x => !ExcludeColumns.Contains(x.Key))
                        .All(x => x.Value != null);
                }
            }

            /// <summary>
            /// РП - Ячейка, входящая от ПС
            /// </summary>
            public class SwitchgearSubstationIn : CustomCubicle
            {
                public SwitchgearSubstationIn(params int[] excludeColumns)
                    : base(new[]
                    {
                        COL_PES, 
                        COL_RES,
                        COL_SUBSTATION_VOLTAGE,
                        COL_SWITCHGEAR_SUBSTATION_SUBSTATION,
                        COL_SWITCHGEAR_SUBSTATION_BUSBAR,
                        COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN
                    }, excludeColumns)
                {
                }

                public Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object>> GetTuple(RowValueInfo rowValue)
                {
                    foreach (var key in Columns.Keys.ToArray())
                        Columns[key] = ExcludeColumns.Contains(key) ? null : rowValue.Values[key];
                    
                    return Tuple.Create(
                        Tuple.Create(
                            Columns[COL_PES], 
                            Columns[COL_RES]),
                        Tuple.Create(
                            Columns[COL_SUBSTATION_VOLTAGE]),
                        Tuple.Create(
                            Columns[COL_SWITCHGEAR_SUBSTATION_SUBSTATION], 
                            Columns[COL_SWITCHGEAR_SUBSTATION_BUSBAR], 
                            Columns[COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN])
                    );
                }

                public bool GetWhere(Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object>> source)
                {
                    Columns[COL_PES] = source.Item1.Item1;
                    Columns[COL_RES] = source.Item1.Item2;
                    Columns[COL_SUBSTATION_VOLTAGE] = source.Item2.Item1;
                    Columns[COL_SWITCHGEAR_SUBSTATION_SUBSTATION] = source.Item3.Item1;
                    Columns[COL_SWITCHGEAR_SUBSTATION_BUSBAR] = source.Item3.Item2;
                    Columns[COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] = source.Item3.Item3;

                    return Columns
                        .Where(x => !ExcludeColumns.Contains(x.Key))
                        .All(x => x.Value != null);
                }
            }

            /// <summary>
            /// РП - Ячейка, отходящая в ТП
            /// </summary>
            public class SwitchgearSubstationOut : CustomCubicle
            {
                public SwitchgearSubstationOut(params int[] excludeColumns)
                    : base(new[]
                    {
                        COL_PES, 
                        COL_RES,
                        COL_SUBSTATION_VOLTAGE,
                        COL_SWITCHGEAR_SUBSTATION_SUBSTATION,
                        COL_SWITCHGEAR_SUBSTATION_BUSBAR,
                        COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT,
                        COL_SWITCHGEAR_SUBSTATION_POWERLINE
                    }, excludeColumns)
                {
                }

                public Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object, object>> GetTuple(RowValueInfo rowValue)
                {
                    foreach (var key in Columns.Keys.ToArray())
                        Columns[key] = ExcludeColumns.Contains(key) ? null : rowValue.Values[key];
                    
                    return Tuple.Create(
                        Tuple.Create(
                            Columns[COL_PES], 
                            Columns[COL_RES]),
                        Tuple.Create(
                            Columns[COL_SUBSTATION_VOLTAGE]),
                        Tuple.Create(
                            Columns[COL_SWITCHGEAR_SUBSTATION_SUBSTATION], 
                            Columns[COL_SWITCHGEAR_SUBSTATION_BUSBAR], 
                            Columns[COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT], 
                            Columns[COL_SWITCHGEAR_SUBSTATION_POWERLINE])
                    );
                }

                public bool GetWhere(Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object, object>> source)
                {
                    Columns[COL_PES] = source.Item1.Item1;
                    Columns[COL_RES] = source.Item1.Item2;
                    Columns[COL_SUBSTATION_VOLTAGE] = source.Item2.Item1;
                    Columns[COL_SWITCHGEAR_SUBSTATION_SUBSTATION] = source.Item3.Item1;
                    Columns[COL_SWITCHGEAR_SUBSTATION_BUSBAR] = source.Item3.Item2;
                    Columns[COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] = source.Item3.Item3;
                    Columns[COL_SWITCHGEAR_SUBSTATION_POWERLINE] = source.Item3.Item4;

                    return Columns
                        .Where(x => !ExcludeColumns.Contains(x.Key))
                        .All(x => x.Value != null);
                }
            }

            /// <summary>
            /// ТП - Высокая сторона
            /// </summary>
            public class LoSubstationHi : CustomCubicle
            {
                public LoSubstationHi(params int[] excludeColumns)
                    : base(new[]
                    {
                        COL_PES,
                        COL_RES,
                        COL_SUBSTATION_VOLTAGE,
                        COL_LO_SUBSTATION_SUBSTATION,
                        COL_LO_SUBSTATION_HI_BUSBAR,
                        COL_LO_SUBSTATION_HI_CUBICLE_NAME,
                        COL_LO_SUBSTATION_HI_CUBICLE_TYPE
                    }, excludeColumns)
                {
                }

                public Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object, object>> GetTuple(RowValueInfo rowValue)
                {
                    foreach (var key in Columns.Keys.ToArray())
                        Columns[key] = ExcludeColumns.Contains(key) ? null : rowValue.Values[key];
                    
                    return Tuple.Create(
                        Tuple.Create(
                            Columns[COL_PES], 
                            Columns[COL_RES]),
                        Tuple.Create(
                            Columns[COL_SUBSTATION_VOLTAGE]),
                        Tuple.Create(
                            Columns[COL_LO_SUBSTATION_SUBSTATION], 
                            Columns[COL_LO_SUBSTATION_HI_BUSBAR], 
                            Columns[COL_LO_SUBSTATION_HI_CUBICLE_NAME], 
                            Columns[COL_LO_SUBSTATION_HI_CUBICLE_TYPE])
                    );
                }

                public bool GetWhere(Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object, object>> source)
                {
                    Columns[COL_PES] = source.Item1.Item1;
                    Columns[COL_RES] = source.Item1.Item2;
                    Columns[COL_SUBSTATION_VOLTAGE] = source.Item2.Item1;
                    Columns[COL_LO_SUBSTATION_SUBSTATION] = source.Item3.Item1;
                    Columns[COL_LO_SUBSTATION_HI_BUSBAR] = source.Item3.Item2;
                    Columns[COL_LO_SUBSTATION_HI_CUBICLE_NAME] = source.Item3.Item3;
                    Columns[COL_LO_SUBSTATION_HI_CUBICLE_TYPE] = source.Item3.Item4;

                    return Columns
                        .Where(x => !ExcludeColumns.Contains(x.Key))
                        .All(x => x.Value != null);
                }
            }

            /// <summary>
            /// ТП - Низкая сторона
            /// </summary>
            public class LoSubstationLo : CustomCubicle
            {
                public LoSubstationLo(params int[] excludeColumns)
                    : base(new[]
                    {
                        COL_PES,
                        COL_RES,
                        COL_SUBSTATION_VOLTAGE,
                        COL_LO_SUBSTATION_SUBSTATION,
                        COL_LO_SUBSTATION_LO_BUSBAR,
                        COL_LO_SUBSTATION_LO_CUBICLE_NAME,
                        COL_LO_SUBSTATION_LO_CUBICLE_TYPE,
                        COL_LO_SUBSTATION_LO_POWERLINE
                    }, excludeColumns)
                {
                }

                public Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object, object, object>> GetTuple(RowValueInfo rowValue)
                {
                    foreach (var key in Columns.Keys.ToArray())
                        Columns[key] = ExcludeColumns.Contains(key) ? null : rowValue.Values[key];
                    
                    return Tuple.Create(
                        Tuple.Create(
                            Columns[COL_PES], 
                            Columns[COL_RES]),
                        Tuple.Create(
                            Columns[COL_SUBSTATION_VOLTAGE]),
                        Tuple.Create(
                            Columns[COL_LO_SUBSTATION_SUBSTATION], 
                            Columns[COL_LO_SUBSTATION_LO_BUSBAR], 
                            Columns[COL_LO_SUBSTATION_LO_CUBICLE_NAME], 
                            Columns[COL_LO_SUBSTATION_LO_CUBICLE_TYPE],
                            Columns[COL_LO_SUBSTATION_LO_POWERLINE])
                    );
                }

                public bool GetWhere(Tuple<Tuple<object, object>, Tuple<object>, Tuple<object, object, object, object, object>> source)
                {
                    Columns[COL_PES] = source.Item1.Item1;
                    Columns[COL_RES] = source.Item1.Item2;
                    Columns[COL_SUBSTATION_VOLTAGE] = source.Item2.Item1;
                    Columns[COL_LO_SUBSTATION_SUBSTATION] = source.Item3.Item1;
                    Columns[COL_LO_SUBSTATION_LO_BUSBAR] = source.Item3.Item2;
                    Columns[COL_LO_SUBSTATION_LO_CUBICLE_NAME] = source.Item3.Item3;
                    Columns[COL_LO_SUBSTATION_LO_CUBICLE_TYPE] = source.Item3.Item4;
                    Columns[COL_LO_SUBSTATION_LO_POWERLINE] = source.Item3.Item5;
                    
                    return Columns
                        .Where(x => !ExcludeColumns.Contains(x.Key))
                        .All(x => x.Value != null);
                }
            }
        }
        
        /// <summary>
        /// Проверить условия подключения линии/фидера
        /// </summary>
        /// <param name="result"></param>
        private void CheckPowerLineConnection(ImportSheetCheckResultData result)
        {
            // Разные СШ для оной ячейки одной и той же РП/ТП

            var substationLoKeyGen = new KeyGenerator.SubstationLo(COL_SUBSTATION_LO_BUSBAR);
            var rowValuesSubstationByPowerLine = RowValues
                .GroupBy(x => substationLoKeyGen.GetTuple(x))
                .Where(x => substationLoKeyGen.GetWhere(x.Key))
                .ToArray();

            // для ПС - РП
            // Секция шин	Ячейка	Тип ячейки	            Линия/фидер	        РП	        Секция шин	Ячейка, входящая от ПС
            // СШ-1	        яч. 3	Ячейка присоединения	ВЛ РП-1		        РП-1	    СШ-1	    яч.ПС Энергетик 35 кВ
            // СШ-2	        яч. 3	Ячейка присоединения	ВЛ РП-1		        РП-1	    СШ-1	    яч.ПС Энергетик 35 кВ
            // !! должно быть СШ-1 или СШ-2, т.к в одну ячейку РП не ляжет 2 линии

            foreach (var rowValuesByPowerLine in rowValuesSubstationByPowerLine)
            {
                foreach (var rowValueByCubicle in rowValuesByPowerLine
                    .GroupBy(x => Tuple.Create(
                        x.Values[COL_SWITCHGEAR_SUBSTATION_SUBSTATION],
                        x.Values[COL_SWITCHGEAR_SUBSTATION_BUSBAR],
                        x.Values[COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN]))
                    .Where(x =>
                        x.Key.Item1 != null &&
                        x.Key.Item2 != null &&
                        x.Key.Item3 != null))
                {
                    if (rowValueByCubicle.Select(x => x.Values[COL_SUBSTATION_LO_BUSBAR]).Distinct().Count() > 1)
                    {
                        foreach (var cell in rowValueByCubicle
                            .SelectMany(x => new[]
                            {
                                Worksheet.Cells[x.RowIndex, COL_SUBSTATION_LO_BUSBAR],
                                Worksheet.Cells[x.RowIndex, COL_SUBSTATION_LO_POWERLINE],
                                Worksheet.Cells[x.RowIndex, COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN]
                            }))
                        {
                            cell.SetError(result, string.Format(
                                "Ячейка \"{0}\". Подключена к линии/фидеру \"{1}\", расположенной на разных СШ.\n" +
                                "В одну ячейку РП не может входить более одной отходящей линии ПС",
                                rowValueByCubicle.Key, rowValuesByPowerLine.Key.Item2.Item6));
                        }
                    }
                }
            }

            // для ПС - ТП
            // Секция шин	Ячейка	Тип ячейки	            Линия/фидер	        ТП	        Секция шин	Ячейка	                Тип ячейки
            // СШ-1	        яч. 2	Ячейка присоединения	ВЛ Энергетик-2		ТП-2321	    СШ-1	    яч. ПС Энергетик 35 кВ	Ячейка присоединения
            // СШ-2	        яч. 2	Ячейка присоединения	ВЛ Энергетик-2		ТП-2321	    СШ-1	    яч. ПС Энергетик 35 кВ	Ячейка присоединения
            // !! должно быть СШ-1 или СШ-2, т.к в одну ячейку ТП не ляжет 2 линии

            foreach (var rowValuesByPowerLine in rowValuesSubstationByPowerLine)
            {
                foreach (var rowValueByCubicle in rowValuesByPowerLine
                    // отфильтровать РП
                    .Where(x => new[]
                        {
                            x.Values[COL_SWITCHGEAR_SUBSTATION_SUBSTATION],
                            x.Values[COL_SWITCHGEAR_SUBSTATION_BUSBAR],
                            x.Values[COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT],
                            x.Values[COL_SWITCHGEAR_SUBSTATION_POWERLINE]
                        }
                        .All(y => y == null))
                    .GroupBy(x => Tuple.Create(
                        x.Values[COL_LO_SUBSTATION_SUBSTATION],
                        x.Values[COL_LO_SUBSTATION_HI_BUSBAR],
                        x.Values[COL_LO_SUBSTATION_HI_CUBICLE_TYPE],
                        x.Values[COL_LO_SUBSTATION_HI_CUBICLE_NAME]))
                    .Where(x =>
                        x.Key.Item1 != null &&
                        x.Key.Item2 != null &&
                        x.Key.Item3 != null &&
                        x.Key.Item4 != null))
                {
                    if (rowValueByCubicle.Select(x => x.Values[COL_SUBSTATION_LO_BUSBAR]).Distinct().Count() > 1)
                    {
                        foreach (var cell in rowValueByCubicle
                            .SelectMany(x => new[]
                            {
                                Worksheet.Cells[x.RowIndex, COL_SUBSTATION_LO_BUSBAR],
                                Worksheet.Cells[x.RowIndex, COL_SUBSTATION_LO_POWERLINE],
                                Worksheet.Cells[x.RowIndex, COL_LO_SUBSTATION_HI_CUBICLE_NAME],
                            }))
                        {
                            cell.SetError(result, string.Format(
                                "Ячейка \"{0}\". Подключена к линии/фидеру \"{1}\", расположенной на разных СШ.\n" +
                                "В одну ячейку ТП не может входить более одной отходящей линии ПС",
                                rowValueByCubicle.Key, rowValuesByPowerLine.Key.Item2.Item7));
                        }
                    }
                }
            }

            // для РП - ТП
            // Секция шин	Ячейка, входящая от ПС	Ячейка, отходящая в ТП	Линия/фидер     ТП	        Секция шин	Ячейка	    Тип ячейки
            // СШ-1	        яч.ПС Энергетик 35 кВ	яч.ТП-2325	            ВЛ ТП-2325	    ТП-2325	    СШ-1	    яч. РП      Ячейка присоединения
            // СШ-2	        яч.ПС Энергетик 35 кВ	яч.ТП-2325	            ВЛ ТП-2325	    ТП-2325	    СШ-1	    яч. РП	    Ячейка присоединения
            // !! должно быть СШ-1 или СШ-2, т.к в одну ячейку ТП не ляжет 2 линии
            
            var switchgearSubstationKeyGen = new KeyGenerator.SwitchgearSubstationOut(COL_SWITCHGEAR_SUBSTATION_BUSBAR);
            foreach (var rowValuesByPowerLine in RowValues
                .GroupBy(x => switchgearSubstationKeyGen.GetTuple(x))
                .Where(x => switchgearSubstationKeyGen.GetWhere(x.Key)))
            {
                foreach (var rowValueByCubicle in rowValuesByPowerLine
                    .GroupBy(x => Tuple.Create(
                        x.Values[COL_LO_SUBSTATION_SUBSTATION],
                        x.Values[COL_LO_SUBSTATION_HI_BUSBAR],
                        x.Values[COL_LO_SUBSTATION_HI_CUBICLE_TYPE],
                        x.Values[COL_LO_SUBSTATION_HI_CUBICLE_NAME]))
                    .Where(x =>
                        x.Key.Item1 != null &&
                        x.Key.Item2 != null &&
                        x.Key.Item3 != null &&
                        x.Key.Item4 != null))
                {
                    if (rowValueByCubicle.Select(x => x.Values[COL_SWITCHGEAR_SUBSTATION_BUSBAR]).Distinct().Count() > 1)
                    {
                        foreach (var cell in rowValueByCubicle
                            .SelectMany(x => new[]
                            {
                                Worksheet.Cells[x.RowIndex, COL_SWITCHGEAR_SUBSTATION_BUSBAR],
                                Worksheet.Cells[x.RowIndex, COL_SWITCHGEAR_SUBSTATION_POWERLINE],
                                Worksheet.Cells[x.RowIndex, COL_LO_SUBSTATION_HI_CUBICLE_NAME],
                            }))
                        {
                            cell.SetError(result, string.Format(
                                "Ячейка \"{0}\". Подключена к линии/фидеру \"{1}\", расположенной на разных СШ.\n" +
                                "В одну ячейку ТП не может входить более одной отходящей линии РП",
                                rowValueByCubicle.Key, rowValuesByPowerLine.Key.Item3.Item4));
                        }
                    }
                }
            }

            // Указание разных линий из одной ячейки

            // для ПС
            // Уровень напряжения РУ	Секция шин  Ячейка	Тип ячейки	            Линия/фидер
            // 10 кВ	                СШ-1	    яч. 2	Ячейка присоединения	ВЛ Энергетик-2
            // 10 кВ	                СШ-1	    яч. 2	Ячейка присоединения	ВЛ Энергетик-3(!)

            substationLoKeyGen = new KeyGenerator.SubstationLo(COL_SUBSTATION_LO_BUSBAR);
            foreach (var rowValuesByCubicle in RowValues
                .GroupBy(x => substationLoKeyGen.GetTuple(x))
                .Where(x => substationLoKeyGen.GetWhere(x.Key)))
            {
                var rowValuesWithPowerLines = rowValuesByCubicle
                    .Where(x => x.Values[COL_SUBSTATION_LO_POWERLINE] != null)
                    .ToArray();
                if (rowValuesWithPowerLines.Select(x => x.Values[COL_SUBSTATION_LO_POWERLINE]).Distinct().Count() > 1)
                {
                    foreach (var cell in rowValuesWithPowerLines
                        .SelectMany(x => new[]
                        {
                            Worksheet.Cells[x.RowIndex, COL_SUBSTATION_LO_CUBICLE_NAME],
                            Worksheet.Cells[x.RowIndex, COL_SUBSTATION_LO_POWERLINE],
                        }))
                    {
                        cell.SetError(result, string.Format(
                            "Ячейка \"{0}\" не может содержать более одной отходящей линии",
                            rowValuesByCubicle.Key.Item2));
                    }
                }
            }

            // для РП
            // РП	    Секция шин	Ячейка, входящая от ПС	Ячейка, отходящая в ТП	Линия/фидер
            // РП-1	    СШ-1	    яч.ПС Энергетик 35 кВ	яч.ТП-2325	            ВЛ ТП-2325
            // РП-1	    СШ-1	    яч.ПС Энергетик 35 кВ	яч.ТП-2325	            ВЛ ТП-2321 (!)

            switchgearSubstationKeyGen = new KeyGenerator.SwitchgearSubstationOut(COL_SWITCHGEAR_SUBSTATION_POWERLINE);
            foreach (var rowValuesByCubicle in RowValues
                .GroupBy(x => switchgearSubstationKeyGen.GetTuple(x))
                .Where(x => switchgearSubstationKeyGen.GetWhere(x.Key)))
            {
                var rowValuesWithPowerLines = rowValuesByCubicle
                    .Where(x => x.Values[COL_SWITCHGEAR_SUBSTATION_POWERLINE] != null)
                    .ToArray();
                if (rowValuesWithPowerLines.Select(x => x.Values[COL_SWITCHGEAR_SUBSTATION_POWERLINE]).Distinct().Count() > 1)
                {
                    foreach (var cell in rowValuesWithPowerLines
                        .SelectMany(x => new[]
                        {
                            Worksheet.Cells[x.RowIndex, COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT],
                            Worksheet.Cells[x.RowIndex, COL_SWITCHGEAR_SUBSTATION_POWERLINE],
                        }))
                    {
                        cell.SetError(result, string.Format(
                            "Ячейка \"{0}\" не может содержать более одной отходящей линии",
                            rowValuesByCubicle.Key.Item3));
                    }
                }
            }

            // для ТП
            // Секция шин	Ячейка	Тип ячейки	            Линия/фидер
            // СШ-1	        яч. 2	Ячейка присоединения	линия 2321-2
            // СШ-1	        яч. 2	Ячейка присоединения	линия 2321-3(!)

            var loSubstationLoKeyGen = new KeyGenerator.LoSubstationLo(COL_LO_SUBSTATION_LO_POWERLINE);
            foreach (var rowValuesByCubicle in RowValues
                .GroupBy(x => loSubstationLoKeyGen.GetTuple(x))
                .Where(x => loSubstationLoKeyGen.GetWhere(x.Key)))
            {
                var rowValuesWithPowerLines = rowValuesByCubicle
                    .Where(x => x.Values[COL_LO_SUBSTATION_LO_POWERLINE] != null)
                    .ToArray();
                if (rowValuesWithPowerLines.Select(x => x.Values[COL_LO_SUBSTATION_LO_POWERLINE]).Distinct().Count() > 1)
                {
                    foreach (var cell in rowValuesWithPowerLines
                        .SelectMany(x => new[]
                        {
                            Worksheet.Cells[x.RowIndex, COL_LO_SUBSTATION_LO_CUBICLE_NAME],
                            Worksheet.Cells[x.RowIndex, COL_LO_SUBSTATION_LO_POWERLINE],
                        }))
                    {
                        cell.SetError(result, string.Format(
                            "Ячейка \"{0}\" не может содержать более одной отходящей линии",
                            rowValuesByCubicle.Key.Item3));
                    }
                }
            }
        }

        /// <summary>
        /// Проверить условия подключения ПУ
        /// </summary>
        /// <param name="result"></param>
        private void CheckMeterConnection(ImportSheetCheckResultData result)
        {
            // Только одно местоустановки должно быть для ПУ
            
            var excludeColumns = new List<int>();

            var rowValuesWithMeters = RowValues
                .Select(x => Tuple.Create(x, Tuple.Create(x.Values[COL_METER_TYPE], x.Values[COL_METER_SERIALNUMBER])))
                .Where(x => x.Item2.Item1 != null || x.Item2.Item2 != null)
                .Select(x => x.Item1)
                .ToArray();

            //// география
            //foreach (var rowValuesWithMeterPlacement in rowValuesWithMeters
            //    .GroupBy(x => x.Values[COL_ADDRESS_FIAS] as FiasSuggestionData, new FiasSuggectionDataComparer())
            //    .Where(x => x.Key != null && x.Count() > 1))
            //{
            //        foreach (var cell in rowValuesWithMeterPlacement
            //            .SelectMany(x => new[]
            //            {
            //                Worksheet.Cells[x.RowIndex, COL_ADDRESS_FIAS],
            //                Worksheet.Cells[x.RowIndex, COL_METER_SERIALNUMBER]
            //            }))
            //        {
            //            cell.SetError(result, string.Format(
            //                "Для места установки \"{0}\" указано более одного ПУ",
            //                Worksheet.Cells[cell.Row.Index, COL_ADDRESS_FIAS].Value));
            //        }
            //}
            excludeColumns.Add(COL_ADDRESS_FIAS);

            // ТП - низкая сторона - ячейка
            var loSubstationLoKeyGen = new KeyGenerator.LoSubstationLo(COL_LO_SUBSTATION_LO_POWERLINE);
            ProcessCubicleMeterConnection(result, rowValuesWithMeters, excludeColumns, COL_LO_SUBSTATION_LO_CUBICLE_NAME,
                loSubstationLoKeyGen.GetTuple, loSubstationLoKeyGen.GetWhere);
            
            // ТП - высокая сторона - ячейка
            var loSubstationHiKeyGen = new KeyGenerator.LoSubstationHi();
            ProcessCubicleMeterConnection(result, rowValuesWithMeters, excludeColumns, COL_LO_SUBSTATION_HI_CUBICLE_NAME,
                loSubstationHiKeyGen.GetTuple, loSubstationHiKeyGen.GetWhere);

            // РП - Ячейка, отходящая в ТП
            var switchgearSubstationOutKeyGen = new KeyGenerator.SwitchgearSubstationOut(COL_SWITCHGEAR_SUBSTATION_POWERLINE);
            ProcessCubicleMeterConnection(result, rowValuesWithMeters, excludeColumns, COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT,
                switchgearSubstationOutKeyGen.GetTuple, switchgearSubstationOutKeyGen.GetWhere);
            
            // РП - Ячейка, входящая от ПС
            var switchgearSubstationInKeyGen = new KeyGenerator.SwitchgearSubstationIn(COL_SWITCHGEAR_SUBSTATION_POWERLINE);
            ProcessCubicleMeterConnection(result, rowValuesWithMeters, excludeColumns, COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN,
                switchgearSubstationInKeyGen.GetTuple, switchgearSubstationInKeyGen.GetWhere);

            // ПС - низкая сторона - ячейка
            var substationLoKeyGen = new KeyGenerator.SubstationLo(COL_SUBSTATION_LO_POWERLINE);
            ProcessCubicleMeterConnection(result, rowValuesWithMeters, excludeColumns, COL_SUBSTATION_LO_CUBICLE_NAME,
                substationLoKeyGen.GetTuple, substationLoKeyGen.GetWhere);
            
            //    // ПС - высокая сторона - ячейка
            //    x.Values[COL_SUBSTATION_HI_CUBICLE_NAME]
            var substationHiKeyGen = new KeyGenerator.SubstationHi();
            ProcessCubicleMeterConnection(result, rowValuesWithMeters, excludeColumns, COL_SUBSTATION_HI_CUBICLE_NAME,
                substationHiKeyGen.GetTuple, substationHiKeyGen.GetWhere);
        }

        /// <summary>
        /// Обработать подключение ПУ к присоединению ТУ
        /// </summary>
        /// <typeparam name="TTuple"></typeparam>
        private void ProcessCubicleMeterConnection<TTuple>(ImportSheetCheckResultData result, IEnumerable<RowValueInfo> rowValuesWithMeters, 
            ICollection<int> excludeColumns, int colCubicleName,
            Func<RowValueInfo, TTuple> getTuple, Func<TTuple, bool> getWhere)
        {
            foreach (var rowValuesWithMeterPlacement in rowValuesWithMeters
                .Where(x => excludeColumns.All(y => x.Values[y] == null))
                .GroupBy(getTuple)
                .Where(x => getWhere(x.Key))
                .Where(x => x.Count() > 1))
            {
                foreach (var cell in rowValuesWithMeterPlacement
                    .SelectMany(x => new[]
                    {
                        Worksheet.Cells[x.RowIndex, colCubicleName],
                        Worksheet.Cells[x.RowIndex, COL_METER_SERIALNUMBER]
                    }))
                {
                    cell.SetError(result, string.Format(
                        "Для места установки \"{0}\" указано более одного ПУ",
                        rowValuesWithMeterPlacement.Key));
                }
            }

            excludeColumns.Add(colCubicleName);
        }
    }

    // лист УСПД
    internal class Rtus : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 3 - 1;

        /// <summary>
        /// Место установки
        /// </summary>
        public const int COL_PLACEMENT = 2 - Helper.OFFSET_IF_NOT_EXCEL;

        // УСПД

        /// <summary>
        /// УСПД - Тип УСПД
        /// </summary>
        public const int COL_TYPE = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// УСПД - Серийный номер
        /// </summary>
        public const int COL_SERIALNUMBER = 4 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// УСПД - Дата выпуска
        /// </summary>
        public const int COL_CREATEDATE = 5 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// УСПД - Дата установки
        /// </summary>
        public const int COL_INSTALLDATE = 6 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// УСПД - Дата последней поверки
        /// </summary>
        public const int COL_CHECKDATE = 7 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// УСПД - Часовой пояс
        /// </summary>
        public const int COL_TIMEZONE = 8 - Helper.OFFSET_IF_NOT_EXCEL;

        // Связь с УСПД

        /// <summary>
        /// Связь с УСПД - Связной номер
        /// </summary>
        public const int COL_NETWORKNUMBER = 9 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с УСПД - Пользователь
        /// </summary>
        public const int COL_LOGIN = 10 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с УСПД - Пароль
        /// </summary>
        public const int COL_PASSWORD = 11 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с УСПД - Маршрут опроса
        /// </summary>
        public const int COL_ROUTE = 12 - Helper.OFFSET_IF_NOT_EXCEL;

        ///// <summary>
        ///// Связь с УСПД - Специальные параметры
        ///// </summary>
        //public const int COL_PARAMS = 13 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Параметры, которые имеются на листе
        /// </summary>
        public static readonly Guid[] KnownMeterParams =
        {
            RDMetadataGuids.Equipment.InstanceAttributeSerialNumber,
            RDMetadataGuids.Equipment.InstanceAttributeReleaseDate,
            //RDMetadataGuids.ElectricityMeter.InstanceAttributeInstallDate,
            RDMetadataGuids.IEntityWithTimeZone.AttributeTimeZone,
            RDMetadataGuids.IEquipmentWithNetworkId.AttributeNetworkId,
            RDMetadataGuids.IEquipmentWithUserAndPassword.AttributeUser,
            RDMetadataGuids.IEquipmentWithPassword.AttributePassword,
        };

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public Rtus(ExcelWorksheet sheet)
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
                    var rowValue = CheckHelper.CheckRtus(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue, COL_SERIALNUMBER, COL_TYPE);
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
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
            // проверить уникальность
            foreach (var rowValues in RowValuesBySerialNumberAndType
                .Where(x => x.Value.Count > 1))
            {
                foreach (var rowValue in rowValues.Value)
                {
                    Worksheet.Cells[rowValue.RowIndex, COL_SERIALNUMBER].SetError(result,
                        "Найдено БОЛЕЕ одного УСПД с указанным серийным номером и типом");
                }
            }
            
            // проверить и уточнить маршруты
            foreach (var rowValue in RowValues)
            {
                var errors = new StringBuilder();
                foreach (var route in rowValue.Values[COL_ROUTE] as ImportRoute[] ?? new ImportRoute[0])
                {
                    ChannelizingEquipmentClassInfo routeEntityClassInfo = null;
                    object routeEntityLink;
                    if (route.Params.TryGetValue(RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, out routeEntityLink) && 
                        routeEntityLink is EntityLink)
                    {
                        var routeEntityLinkRowValue = Worksheet.Cells[rowValue.RowIndex, COL_ROUTE].CheckEntityLinkAndGetRowValue(
                            result, this, routeEntityLink, Helpers.GetRtuClass, "Каналообразующее оборудование");
                        routeEntityClassInfo = routeEntityLinkRowValue == null ? null : routeEntityLinkRowValue.Values[COL_TYPE] as ChannelizingEquipmentClassInfo;
                    }
                    string errorText;
                    if (!route.ParseFor(rowValue.Values[COL_TYPE] as ChannelizingEquipmentClassInfo, routeEntityClassInfo, out errorText))
                    {
                        errors.AppendLine(errorText);
                    }
                }
                if (errors.Length > 0)
                {
                    rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, COL_ROUTE].SetError(result, errors.ToString());
                }
            }
        }
    }

    // лист ТТ
    internal class TransCurrents : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 3 - 1;

        // ТТ

        /// <summary>
        /// ТТ - Тип ТТ
        /// </summary>
        public const int COL_TYPE = 2 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТТ - Класс точности
        /// </summary>
        public const int COL_ACCURACYCLASS = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТТ - Серийный номер
        /// </summary>
        public const int COL_SERIALNUMBER = 4 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТТ - Дата выпуска
        /// </summary>
        public const int COL_CREATEDATE = 5 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТТ - Дата установки
        /// </summary>
        public const int COL_INSTALLDATE = 6 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТТ - Дата последней поверки
        /// </summary>
        public const int COL_CHECKDATE = 7 - Helper.OFFSET_IF_NOT_EXCEL;

        // Параметры ТТ

        /// <summary>
        /// Параметры ТТ - I ном. перв, А
        /// </summary>
        public const int COL_PRIMARY_NOM = 8 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Параметры ТТ - I ном. втор, А
        /// </summary>
        public const int COL_SECONDARY_NOM = 9 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public TransCurrents(ExcelWorksheet sheet)
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
                    var rowValue = CheckHelper.CheckTransCurrents(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue, COL_SERIALNUMBER, COL_TYPE);
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
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
            // проверить уникальность
            foreach (var rowValues in RowValuesBySerialNumberAndType
                .Where(x => x.Value.Count > 1))
            {
                foreach (var rowValue in rowValues.Value)
                {
                    Worksheet.Cells[rowValue.RowIndex, COL_SERIALNUMBER].SetError(result,
                        "Найдено БОЛЕЕ одного ТТ с указанным серийным номером и типом");
                }
            }
        }
    }

    // лист ТН
    internal class TransVoltages : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 3 - 1;

        // ТН

        /// <summary>
        /// ТН - Тип ТН
        /// </summary>
        public const int COL_TYPE = 2 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - Класс точности
        /// </summary>
        public const int COL_ACCURACYCLASS = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - Серийный номер
        /// </summary>
        public const int COL_SERIALNUMBER = 4 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - Дата выпуска
        /// </summary>
        public const int COL_CREATEDATE = 5 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - Дата установки
        /// </summary>
        public const int COL_INSTALLDATE = 6 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ТН - Дата последней поверки
        /// </summary>
        public const int COL_CHECKDATE = 7 - Helper.OFFSET_IF_NOT_EXCEL;

        // Параметры ТН

        /// <summary>
        /// Параметры ТН - U ном. перв, В
        /// </summary>
        public const int COL_PRIMARY_NOM = 8 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Параметры ТН - U ном. втор, В
        /// </summary>
        public const int COL_SECONDARY_NOM = 9 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public TransVoltages(ExcelWorksheet sheet)
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
                    var rowValue = CheckHelper.CheckTransVoltages(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue, COL_SERIALNUMBER, COL_TYPE);
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
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
            // проверить уникальность
            foreach (var rowValues in RowValuesBySerialNumberAndType
                .Where(x => x.Value.Count > 1))
            {
                foreach (var rowValue in rowValues.Value)
                {
                    Worksheet.Cells[rowValue.RowIndex, COL_SERIALNUMBER].SetError(result,
                        "Найдено БОЛЕЕ одного ТН с указанным серийным номером и типом");
                }
            }
        }
    }

    // лист ФЛ
    internal class NaturalPersons : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 3 - 1;

        // ФИО

        /// <summary>
        /// ФИО - Фамилия
        /// </summary>
        public const int COL_LASTNAME = 2 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ФИО - Имя
        /// </summary>
        public const int COL_FIRSTNAME = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// ФИО - Отчетство
        /// </summary>
        public const int COL_MIDDLENAME = 4 - Helper.OFFSET_IF_NOT_EXCEL;

        // Связь с абонентом

        /// <summary>
        /// Связь с абонентом - Эл.почта
        /// </summary>
        public const int COL_EMAIL = 5 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с абонентом - Телефон
        /// </summary>
        public const int COL_PHONE = 6 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Номер лицевого счета
        /// </summary>
        public const int COL_CURRENTACCOUNT = 7 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public NaturalPersons(ExcelWorksheet sheet)
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
                    var rowValue = CheckHelper.CheckNaturalPersons(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue, COL_CURRENTACCOUNT, COL_LASTNAME);
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
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
            // проверить уникальность
            foreach (var rowValues in RowValuesBySerialNumberAndType
                .Where(x => x.Value.Count > 1))
            {
                foreach (var rowValue in rowValues.Value)
                {
                    Worksheet.Cells[rowValue.RowIndex, COL_CURRENTACCOUNT].SetError(result,
                        "Найдено БОЛЕЕ одного Физического лица с указанным лицевым счетом и фамилией");
                }
            }
        }
    }

    // лист ЮЛ
    internal class LegalEntities : Sheet
    {
        // индексация в Gembox с 0
        public const int START_ROW = 3 - 1;

        /// <summary>
        /// Наименование
        /// </summary>
        public const int COL_CAPTION = 2 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Юридический адрес
        /// </summary>
        public const int COL_ADDRESS = 3 - Helper.OFFSET_IF_NOT_EXCEL;

        // Связь с абонентом

        /// <summary>
        /// Связь с абонентом - Эл.почта
        /// </summary>
        public const int COL_EMAIL = 4 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Связь с абонентом - Телефон
        /// </summary>
        public const int COL_PHONE = 5 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Номер лицевого счета
        /// </summary>
        public const int COL_CURRENTACCOUNT = 6 - Helper.OFFSET_IF_NOT_EXCEL;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sheet"></param>
        public LegalEntities(ExcelWorksheet sheet)
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
                    var rowValue = CheckHelper.CheckLegalEntities(result, Worksheet, row);
                    AddRowValue(row.Index, rowValue, COL_CURRENTACCOUNT, COL_CAPTION);
                }
                catch (System.Threading.ThreadInterruptedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", Worksheet.Name, row.Index, ex);
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
            // проверить уникальность
            foreach (var rowValues in RowValuesBySerialNumberAndType
                .Where(x => x.Value.Count > 1))
            {
                foreach (var rowValue in rowValues.Value)
                {
                    Worksheet.Cells[rowValue.RowIndex, COL_CURRENTACCOUNT].SetError(result,
                        "Найдено БОЛЕЕ одного Юридического лица с указанным лицевым счетом и наименованием");
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
        /// Наименование вводной ТУ МКД
        /// </summary>
        public const string TENEMENT_HOUSE_POWERLINE_METERPOINT_NAME = "Ввод";

        /// <summary>
        /// Наименование помещения (вмето квартиры)
        /// </summary>
        public const string ROOM_PRIFIX_NAME = "Помещение";

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
    
    #endregion Helper
}