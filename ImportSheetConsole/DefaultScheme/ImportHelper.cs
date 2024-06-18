using System;
using System.Collections.Generic;
using System.Linq;
using CommonTools;
using ImportSheetConsole.Global;
using ObjStudioClasses;
using TimeZone = ObjStudioClasses.TimeZone;

namespace ImportSheetConsole.DefaultScheme
{
    #region ImportHelper

    /// <summary>
    /// Утилиты валидации (внутренние)
    /// </summary>
    internal static class ImportHelper
    {
        /// <summary>
        /// Наименование классифкатора энергообъектов
        /// </summary>
        public const string ENERGY_CLASSIFIER_CAPTION = "Иерархия сети";

        /// <summary>
        /// Наименование классифкатора географических регионов
        /// </summary>
        public const string GEO_CLASSIFIER_CAPTION = "Географические объекты";

        /// <summary>
        /// Лист ПУ
        /// </summary>
        [ThreadStatic]
        private static MeterPoints _meterPoints;

        /// <summary>
        /// Лист УСПД
        /// </summary>
        [ThreadStatic]
        private static Rtus _rtus;

        /// <summary>
        /// Лист ТТ
        /// </summary>
        [ThreadStatic]
        private static TransCurrents _transCurrents;

        /// <summary>
        /// Лист ТН
        /// </summary>
        [ThreadStatic]
        private static TransVoltages _transVoltages;

        /// <summary>
        /// Лист ФЛ
        /// </summary>
        [ThreadStatic]
        private static NaturalPersons _naturalPersons;

        /// <summary>
        /// Лист ЮЛ
        /// </summary>
        [ThreadStatic]
        private static LegalEntities _legalEntities;

        /// <summary>
        /// Классификатор энергообъектов
        /// </summary>
        [ThreadStatic]
        private static Lazy<ClassifierOfMeterPointsByEnergyEntities> _energyClassifier;
        
        /// <summary>
        /// Классификатор географических регионов
        /// </summary>
        [ThreadStatic]
        private static Lazy<ClassifierOfMeterPointsByGeoLocation> _geoClassifier;

        /// <summary>
        /// Импортированные места установки
        /// </summary>
        [ThreadStatic]
        private static Dictionary<string, ClassifierItem> _placements;

        /// <summary>
        /// Импортированные ПУ
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, MeterEquipment> _importedMeters;

        /// <summary>
        /// Импортированные УСПД
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, ChannelizingEquipment> _importedRtus;

        /// <summary>
        /// Импортированные ТТ
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, CurrentTransformer> _importedTransCurrents;

        /// <summary>
        /// Импортированные ТН
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, VoltageTransformer> _importedTransVoltages;

        /// <summary>
        /// Импортированные ФЛ
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, NaturalPerson> _importedNaturalPersons;

        /// <summary>
        /// Импортированные ЮЛ
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, LegalEntity> _importedLegalEntities;

        /// <summary>
        /// Сопоставленные сущностям области видимости
        /// </summary>
        [ThreadStatic]
        private static Dictionary<RDInstance, IsolationLevel[]> _importedEntityIsolationLevels;

        /// <summary>
        /// Список видимых пользователю областей видимости
        /// </summary>
        [ThreadStatic]
        private static HashSet<IsolationLevel> _visibleIsolationLevels;
        
        /// <summary>
        /// Скешированный список существующих приборов учета, 
        /// вынесен с целью ускорения пробежки по коллекции
        /// </summary>
        [ThreadStatic]
        private static List<MeterEquipment> _existingMeters;
	
        /// <summary>
        /// Скешированный список существующих каналообразующих устройств, 
        /// вынесен с целью ускорения пробежки по коллекции
        /// </summary>
        [ThreadStatic]
        private static List<ChannelizingEquipment> _existingRtus;
	
        /// <summary>
        /// Скешированный список существующих физ. лиц, 
        /// вынесен с целью ускорения пробежки по коллекции
        /// </summary>
        [ThreadStatic]
        private static List<NaturalPerson> _existingNaturals;
	
        /// <summary>
        /// Скешированный список существующих юр. лиц, 
        /// вынесен с целью ускорения пробежки по коллекции
        /// </summary>
        [ThreadStatic]
        private static List<LegalEntity> _existingLegals;

        /// <summary>
        /// Импортировать
        /// </summary>
        public static void Import(ImportSheetProcessedResultData result)
        {
            _existingMeters = MeterEquipment.GetInstances().ToList();
            _existingRtus = ChannelizingEquipment.GetInstances().ToList();
            _existingNaturals = NaturalPerson.GetInstances().ToList();
            _existingLegals = LegalEntity.GetInstances().ToList();
            try
            {
                ImportInternal(result);
            }
            finally
            {
                _existingMeters = null;
                _existingRtus = null;
                _existingNaturals = null;
                _existingLegals = null;
            }
        }

        /// <summary>
        /// Импортировать, внутренний вызов, обрамленный кешированием
        /// </summary>
        private static void ImportInternal(ImportSheetProcessedResultData result)
        {
            _meterPoints = (MeterPoints)Helper.Sheets.First(x => x is MeterPoints);
            _rtus = (Rtus)Helper.Sheets.First(x => x is Rtus);
            _transCurrents = (TransCurrents)Helper.Sheets.First(x => x is TransCurrents);
            _transVoltages = (TransVoltages)Helper.Sheets.First(x => x is TransVoltages);
            _naturalPersons = (NaturalPersons)Helper.Sheets.First(x => x is NaturalPersons);
            _legalEntities = (LegalEntities)Helper.Sheets.First(x => x is LegalEntities);

            // классификаторы
            _energyClassifier = new Lazy<ClassifierOfMeterPointsByEnergyEntities>(() =>
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<ClassifierOfMeterPointsByEnergyEntities>(ENERGY_CLASSIFIER_CAPTION, null));
            _geoClassifier = new Lazy<ClassifierOfMeterPointsByGeoLocation>(() =>
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<ClassifierOfMeterPointsByGeoLocation>(GEO_CLASSIFIER_CAPTION, null));

            _placements = new Dictionary<string, ClassifierItem>();

            _importedMeters = new Dictionary<int, MeterEquipment>();
            _importedRtus = new Dictionary<int, ChannelizingEquipment>();
            _importedTransCurrents = new Dictionary<int, CurrentTransformer>();
            _importedTransVoltages = new Dictionary<int, VoltageTransformer>();
            _importedNaturalPersons = new Dictionary<int, NaturalPerson>();
            _importedLegalEntities = new Dictionary<int, LegalEntity>();
            _visibleIsolationLevels = new HashSet<IsolationLevel>(IsolationLevel.GetInstances());

            _importedEntityIsolationLevels = new Dictionary<RDInstance, IsolationLevel[]>();

            // импортировать ТУ
            // в процессе импорта также создаются УСПД, ТТ, ТН, абоненты, на которые ссылаюется ТУ
            // !! если в процессе импорта ТУ возникает ошибка, то связанные с ней УСПД, абоненты и тп, не импортируются
            // !! и попадают в списки _imported*, т.е помечаются как импортированные
            foreach (var rowValue in _meterPoints.RowValues)
            { 
                ImportMeterPointRow(result, rowValue);
            }
            //
            // импортировать оставшиеся УСПД (см. замечание выше)
            foreach (var rowValue in _rtus.RowValues.Where(x => !_importedRtus.ContainsKey(x.RowIndex)))
            {
                var value = ImportRtuRow(result, rowValue, false, null);
                _importedRtus[rowValue.RowIndex] = value;
            }

            // установить местоустановки для УСПД
            foreach (var rowValue in _rtus.RowValues)
            {
                ChannelizingEquipment rtu;
                if (!_importedRtus.TryGetValue(rowValue.RowIndex, out rtu) || rtu == null)
                    continue;
                // строкам, где описаны существующие УСПД, установку местоположения не производим
                Guid rtuId;
                if (Guid.TryParse(rowValue.Values[Sheet.COL_NUM] as string, out rtuId))
                    continue;
                var placementValues = rowValue.Values[Rtus.COL_PLACEMENT] as string[];
                if (placementValues == null)
                    continue;
                // в процессе может возникнуть исключение, если будет осуществлен доступ к УСПД, 
                // которое было удалено параллельно с импортом
                try
                {
                    ClassifierItem classInfo;
                    var placements = placementValues
                        .Select(x => _placements.TryGetValue(x, out classInfo) ? classInfo : null)
                        .Where(x => x != null)
                        .ToArray();
                    ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_PLACEMENT] },
                        placements,
                        (value, parent, args) =>
                        {
                            var values = new HashSet<ClassifierItem>(rtu.AttributeManualPlacement.GetValues());
                            foreach (var placement in placements)
                            {
                                if (!values.Contains(placement))
                                    rtu.AttributeManualPlacement.Add(placement);
                            }
                            return rtu;
                        });
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

            // установка маршрутов через промежуточное оборудование

            // для ПУ
            foreach (var rowValue in _meterPoints.RowValues)
            { 
                MeterEquipment meter;
                if (!_importedMeters.TryGetValue(rowValue.RowIndex, out meter) || meter == null)
                    continue;
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_ROUTE] },
                    rowValue.Values[MeterPoints.COL_METER_ROUTE] as ImportRoute[],
                    (value, parent, args) =>
                    {
                        foreach (var route in value)
                        {
                            object routeEntityLink;
                            if (!route.Params.TryGetValue(RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, out routeEntityLink) || 
                                !(routeEntityLink is EntityLink))
                                continue;
                            
                            Equipment viaRtu = null;

                            // лист УСПД
                            var rtuRowValue = Helper.GetRowValue(_rtus, (EntityLink)routeEntityLink, Helpers.GetRtuClass);
                            if (rtuRowValue != null)
                            {
                                ChannelizingEquipment rtu;
                                if (_importedRtus.TryGetValue(rtuRowValue.RowIndex, out rtu) && rtu != null)
                                    viaRtu = rtu;
                            }

                            // лист ТУ
                            var meterRowValue = Helper.GetRowValue(_meterPoints, (EntityLink)routeEntityLink, Helpers.GetMeterClass);
                            if (meterRowValue != null)
                            {
                                MeterEquipment meterEquipment;
                                if (_importedMeters.TryGetValue(meterRowValue.RowIndex, out meterEquipment) && meterEquipment != null)
                                    viaRtu = meterEquipment;
                            }

                            if (viaRtu == null)
                                continue;
                            
                            // установить экземпляр УСПД, вместо описания ссылки
                            route.Params[RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment] = viaRtu;
                            // не добавляется, если маршруты равны
                            var route1 = route;
                            var rdRoute = meter.AttributeRoutes.GetValues().FirstOrDefault(x => route1.Equals(x));
                            if (rdRoute == null)
                                rdRoute = ImportHelpers.FindOrCreate.Entity(meter.AttributeRoutes, route.ClassInfo);
                            // связывание специфичных параметров
                            route.Assign(rdRoute);

                            // добавить область видимости промежуточному оборудованию
                            IsolationLevel[] isolationLevels;
                            if (_importedEntityIsolationLevels.TryGetValue(meter, out isolationLevels))
                            {
                                _importedEntityIsolationLevels[viaRtu] = isolationLevels;
                                viaRtu.AppendIsolationLevels(isolationLevels);
                            }
                        }

                        return meter;
                    });
            }

            // для УСПД
            foreach (var rowValue in _rtus.RowValues)
            { 
                ChannelizingEquipment rtu;
                if (!_importedRtus.TryGetValue(rowValue.RowIndex, out rtu) || rtu == null)
                    continue;
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_ROUTE] },
                    rowValue.Values[Rtus.COL_ROUTE] as ImportRoute[],
                    (value, parent, args) =>
                    {
                        foreach (var route in value)
                        {
                            object routeEntityLink;
                            if (!route.Params.TryGetValue(RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, out routeEntityLink) || 
                                !(routeEntityLink is EntityLink))
                                continue;
                            var rtuRowValue = Helper.GetRowValue(_rtus, (EntityLink)routeEntityLink, Helpers.GetRtuClass);
                            if (rtuRowValue == null)
                                continue;
                            ChannelizingEquipment viaRtu;
                            if (!_importedRtus.TryGetValue(rtuRowValue.RowIndex, out viaRtu) || viaRtu == null)
                                continue;
                            // установить экземпляр УСПД, вместо описания ссылки
                            route.Params[RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment] = viaRtu;
                            // не добавляется, если маршруты равны
                            var route1 = route;
                            var rdRoute = rtu.AttributeRoutes.GetValues().FirstOrDefault(x => route1.Equals(x));
                            if (rdRoute == null)
                                rdRoute = ImportHelpers.FindOrCreate.Entity(rtu.AttributeRoutes, route.ClassInfo);
                            // связывание специфичных параметров
                            route.Assign(rdRoute);
                            
                            // добавить область видимости промежуточному оборудованию
                            IsolationLevel[] isolationLevels;
                            if (_importedEntityIsolationLevels.TryGetValue(rtu, out isolationLevels))
                            {
                                _importedEntityIsolationLevels[viaRtu] = isolationLevels;
                                viaRtu.AppendIsolationLevels(isolationLevels);
                            }
                        }

                        return rtu;
                    });
            }
            
            // импортировать оставшиеся ТТ
            foreach (var rowValue in _transCurrents.RowValues.Where(x => !_importedTransCurrents.ContainsKey(x.RowIndex)))
            {
                ImportTransCurrentRow(result, rowValue);
            }

            // импортировать оставшиеся ТН
            foreach (var rowValue in _transVoltages.RowValues.Where(x => !_importedTransVoltages.ContainsKey(x.RowIndex)))
            {
                ImportTransVoltageRow(result, rowValue);
            }

            // импортировать оставшихся ФЛ
            foreach (var rowValue in _naturalPersons.RowValues.Where(x => !_importedNaturalPersons.ContainsKey(x.RowIndex)))
            {
                ImportConsumerRow(result, rowValue);
            }

            // импортировать оставшихся ЮЛ
            foreach (var rowValue in _legalEntities.RowValues.Where(x => !_importedLegalEntities.ContainsKey(x.RowIndex)))
            {
                ImportConsumerRow(result, rowValue);
            }
        }

        /// <summary>
        /// Импортировать строку
        /// </summary>
        private static void ImportMeterPointRow(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue)
        {
            // Корректируем значение областей видимости
            // Если пользователь видит лишь часть областей, добавляем предупреждение
            var isolationLevels = rowValue.Values[MeterPoints.COL_ISOLATION_LEVEL] as IsolationLevel[];
            if (isolationLevels != null)
            {
                var allowedIsolationLevels = isolationLevels.Where(_visibleIsolationLevels.Contains).ToArray();
                if (allowedIsolationLevels.Length != isolationLevels.Length)
                {
                    _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ISOLATION_LEVEL].SetWarn(result, 
                        "Не все области видимости могут быть применены для текущего пользователя");
                    isolationLevels = allowedIsolationLevels.Any() ? allowedIsolationLevels : null;
                    rowValue.Values[MeterPoints.COL_ISOLATION_LEVEL] = isolationLevels;
                }
            }
            
            var possibleId = Helper.TryGetGuidIdentifier(rowValue.Values[Sheet.COL_NUM]);
            var existingEntity = possibleId.HasValue ? MeterPoint.Find(possibleId.Value) : null;

            // Очищаем связи, если ТУ уже существует:
            if (existingEntity != null)
            {
                // 1. ТУ с абонентом 
                var existingEntityConsumer = existingEntity.AttributeConsumer.FirstOrDefault();
                if (existingEntityConsumer != null)
                {
                    var recordToRemove = existingEntityConsumer.AttributeMeterPoints.GetValuesInfo().FirstOrDefault(x => x.Value == existingEntity);
                    if (recordToRemove != null)
                        recordToRemove.Remove();
                }
                // 2. маршруты ПУ
                var existingEntityMeter = existingEntity.AttributeElectricityMeter;
                if (existingEntityMeter != null)
                    existingEntityMeter.AttributeRoutes.Clear();
                // 3. Тариф
                existingEntity.SetCurrentTariff(null);
            }

            var entities = rowValue.Values.ToDictionary(x => x.Key, x => (RDInstance)null);
            ImportHelpers.CreatedEntities = new List<RDInstance>();

            var energyParentClassifierItem = (RDInstance)_energyClassifier.Value;
            var geoParentClassifierItem = (RDInstance)_geoClassifier.Value;

            // строка УСПД
            Sheet.RowValueInfo rtuRowValue = null;
            // строка ТТ, Фаза 1
            Sheet.RowValueInfo tt1RowValue = null;
            // строка ТТ, Фаза 2
            Sheet.RowValueInfo tt2RowValue = null;
            // строка ТТ, Фаза 3
            Sheet.RowValueInfo tt3RowValue = null;
            // строка ТН, Фаза 1
            Sheet.RowValueInfo tn1RowValue = null;
            // строка ТН, Фаза 2
            Sheet.RowValueInfo tn2RowValue = null;
            // строка ТН, Фаза 3
            Sheet.RowValueInfo tn3RowValue = null;
            // строка ФЛ
            Sheet.RowValueInfo naturalPersonRowValue = null;
            // строка ЮЛ
            Sheet.RowValueInfo legalEntityRowValue = null;
            
            try
            {
                if (existingEntity == null)
                {
                    // ПЭС
                    entities[MeterPoints.COL_PES] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_PES] },
                        rowValue.Values[MeterPoints.COL_PES] as string,
                        energyParentClassifierItem,
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<ElectricalNetworksSubsidiary>);
                    energyParentClassifierItem = entities[MeterPoints.COL_PES] ?? energyParentClassifierItem;
                    _placements[rowValue.Values[MeterPoints.COL_PES] as string ?? string.Empty] = (ClassifierItem)entities[MeterPoints.COL_PES];

                    // РЭС
                    entities[MeterPoints.COL_RES] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_RES] },
                        rowValue.Values[MeterPoints.COL_RES] as string,
                        energyParentClassifierItem,
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<ElectricalNetworksDistrict>);
                    energyParentClassifierItem = entities[MeterPoints.COL_RES] ?? energyParentClassifierItem;
                    _placements[rowValue.Values[MeterPoints.COL_RES] as string ?? string.Empty] = (ClassifierItem)entities[MeterPoints.COL_RES];

                    // ПС
                    var substationParentClassifier = ImportMeterPointRowSubstation(result, rowValue, entities, energyParentClassifierItem);

                    // РП
                    ImportMeterPointRowSwitchgearSubstation(result, rowValue, entities, substationParentClassifier);

                    // ТП
                    ImportMeterPointRowLoSubstation(result, rowValue, entities, energyParentClassifierItem);

                    // Адрес ФИАС
                    ImportMeterPointRowAddress(result, rowValue, entities, geoParentClassifierItem);
                }

                // УСПД
                var rtuLink = rowValue.Values[MeterPoints.COL_RTU] as EntityLink;
                rtuRowValue = rtuLink != null
                    ? Helper.GetRowValue(_rtus, rtuLink, Helpers.GetRtuClass)
                    : null;
                entities[MeterPoints.COL_RTU] = ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_RTU] },
                    rtuRowValue,
                    (value, parent, args) =>
                    {
                        // исключение, если УСПД уже был импортирован ранее с ошибкой
                        ChannelizingEquipment resultRtu;
                        // импорт с пробросом исключения
                        return _importedRtus.TryGetValue(value.RowIndex, out resultRtu)
                            ? resultRtu
                            : ImportRtuRow(result, value, true, isolationLevels);
                    });

                // ПУ
                var meter = ImportMeterPointRowMeter(result, rowValue, existingEntity);
                entities[MeterPoints.COL_METER_TYPE] = meter;
                if (isolationLevels != null)
                {
                    _importedEntityIsolationLevels[meter] = isolationLevels;
                    meter.AppendIsolationLevels(isolationLevels);
                }

                // Абонент
                
                // строка ФЛ
                var naturalPersonLink = rowValue.Values[MeterPoints.COL_CONSUMER] as EntityLink;
                naturalPersonRowValue = naturalPersonLink != null
                    ? Helper.GetRowValue(_naturalPersons, naturalPersonLink, Helpers.GetString)
                    : null;
                // строка ЮЛ
                EntityLink legalEntityLink;
                legalEntityRowValue = naturalPersonRowValue == null && 
                    (legalEntityLink = rowValue.Values[MeterPoints.COL_CONSUMER] as EntityLink) != null 
                        ? Helper.GetRowValue(_legalEntities, legalEntityLink, Helpers.GetString)
                        : null;

                entities[MeterPoints.COL_CONSUMER] = ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_CONSUMER] },
                    naturalPersonRowValue ?? legalEntityRowValue,
                    (value, parent, args) =>
                    {
                        // импорт с пробросом исключения
                        var resultConsumer = ImportConsumerRow(result, value, true);
                        return resultConsumer;
                    });

                var meterPoint = existingEntity;
                // ТУ, если есть ПУ или абонент
                if (meterPoint == null)
                {
                    if (entities[MeterPoints.COL_METER_TYPE] != null || entities[MeterPoints.COL_CONSUMER] != null)
                    {
                        var meterPointParent =
                            // география
                            entities[MeterPoints.COL_ADDRESS_FIAS] ??
                            // ТП - низкая сторона - ячейка
                            entities[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME] ??
                            // ТП - высокая сторона - ячейка
                            entities[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME] ??
                            // РП - Ячейка, отходящая в ТП
                            entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] ??
                            // РП - Ячейка, входящая от ПС
                            entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] ??
                            // ПС - низкая сторона - ячейка
                            entities[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME] ??
                            // ПС - высокая сторона - ячейка
                            entities[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME];

                        // наименование ТУ
                        var meterPointName = (entities[MeterPoints.COL_METER_TYPE] ?? entities[MeterPoints.COL_CONSUMER]).ToString();
                        var meterPointCell = _meterPoints.Worksheet.Cells[rowValue.RowIndex, Sheet.COL_NUM];

                        // для многоквартирного дома, добаляем номер квартиры
                        if (meterPointParent == entities[MeterPoints.COL_ADDRESS_FIAS] &&
                            meterPointParent is TenementHouse &&
                            rowValue.Values[MeterPoints.COL_ADDRESS_FLAT] is string)
                        {
                            var building = meterPointParent as TenementHouse;
                            var flatName = rowValue.Values[MeterPoints.COL_ADDRESS_FLAT] as string;
                            
                            // вводная ТУ
                            if (flatName?.IndexOf(Helper.TENEMENT_HOUSE_POWERLINE_METERPOINT_NAME, StringComparison.InvariantCultureIgnoreCase) != -1 &&
                                entities[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] is PowerLine)
                            {
                                var powerLine = entities[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] as PowerLine;
                                var powerLineConnection = building.AttributePowerLineConnections.GetValues()
                                    .FirstOrDefault(x => x.AttributePowerLine == powerLine);
                                meterPointParent = powerLineConnection;
                            }

                            meterPointName = string.Format("{0}, {1}", flatName, meterPointName);
                            meterPointCell = _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FLAT];
                        }
                        // для строения в случае наличия присоединений, нобходимо выполнить поиск по линии в нем
                        else if (meterPointParent == entities[MeterPoints.COL_ADDRESS_FIAS] &&
                            meterPointParent is CustomBuilding &&
                            entities[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] is PowerLine)
                        {
                            var building = meterPointParent as CustomBuilding;
                            var powerLine = entities[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] as PowerLine;
                            var powerLineConnection = building.AttributePowerLineConnections.GetValues()
                                .FirstOrDefault(x => x.AttributePowerLine == powerLine);
                            if (powerLineConnection != null)
                                meterPointParent = powerLineConnection;
                        }
                            
                        meterPoint = ImportHelpers.FindOrCreateEntity(result,
                            new[] { meterPointCell },
                            meterPointName,
                            meterPointParent,
                            ImportHelpers.FindOrCreate.ClassifierNodes.MeterPoint,
                            true,
                            // ПУ
                            entities[MeterPoints.COL_METER_TYPE]);
                    }
                }

                if (meterPoint != null)
                {
                    // Принудительно выставляем область видимости ТУ, если этой области у ТУ не было
                    if (isolationLevels != null && isolationLevels.Any())
                    {
                        ImportHelpers.FindOrCreateEntity(result,
                            new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ISOLATION_LEVEL] },
                            isolationLevels,
                            (value, parent, args) =>
                            {
                                _importedEntityIsolationLevels[meterPoint] = value;
                                var meterPointExistingLevels = meterPoint.AttributeIsolationLevels.ToArray();
                                var meterPointIsolationLevelsToImport = value.Where(x => !meterPointExistingLevels.Contains(x)).ToArray();
                                if (meterPointIsolationLevelsToImport.Any())
                                {
                                    meterPoint.AttributeIsolationLevelsInternal.SetValues(value);
                                    // Если модифицировали области видимости ТУ, то выставляем ее же и в линии абонентского присоединения, чтобы
                                    // указанная ТУ была доступна из классификатора сети по этой линии
                                    PowerLine[] consumerConnectionLines;
                                    RDClassesAndInstances.SecurityManager.PushSecurityContext();
                                    try
                                    {
                                        consumerConnectionLines = meterPoint.RelationsIPowerLineConnectionWithConsumerMeterPointsAttributeConsumerMeterPoints
                                            .Select(x => x.AttributePowerLine)
                                            .Where(x => x != null)
                                            .ToArray();
                                        foreach (var connectionLine in consumerConnectionLines)
                                        {
                                            _importedEntityIsolationLevels[connectionLine] = value;
                                            var connectionLineExistingLevels = connectionLine.AttributeIsolationLevels.ToArray();
                                            meterPointIsolationLevelsToImport = value.Where(x => !connectionLineExistingLevels.Contains(x)).ToArray();
                                            if (meterPointIsolationLevelsToImport.Any())
                                                connectionLine.AttributeIsolationLevelsInternal.SetValues(connectionLineExistingLevels.Concat(meterPointIsolationLevelsToImport));
                                        }
                                    }
                                    finally
                                    {
                                        RDClassesAndInstances.SecurityManager.PopSecurityContext();
                                    }
                                }
                                return meterPoint;
                            });
                    }

                    // абонент
                    var consumer = entities[MeterPoints.COL_CONSUMER] as Consumer;
                    if (consumer != null)
                    {
                        if (!consumer.AttributeMeterPoints.GetValues().Contains(meterPoint))
                            consumer.AttributeMeterPoints.Add(meterPoint);
                        if (isolationLevels != null)
                        {
                            _importedEntityIsolationLevels[consumer] = isolationLevels;
                            consumer.AppendIsolationLevels(isolationLevels);
                        }
                    }

                    // Дата установки
                    ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_INSTALLDATE] },
                        rowValue.Values[MeterPoints.COL_METER_INSTALLDATE] as DateTime?,
                        (value, parent, args) =>
                        {
                            ImportHelpers.FindOrCreate.SetMeterToMeterPoint(
                                meterPoint,
                                entities[MeterPoints.COL_METER_TYPE] as ElectricityMeter,
                                value);
                            return meterPoint;
                        });

                    // Тариф
                    entities[MeterPoints.COL_TARIFF] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TARIFF] },
                        rowValue.Values[MeterPoints.COL_TARIFF] as Tariff,
                        (value, parent, args) =>
                        {
                            if (existingEntity == null && meterPoint.AttributeTariff != null && meterPoint.AttributeTariff != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                meterPoint.SetCurrentTariff(value);
                            return value;
                        });
                }

                // Трансформаторы

                // Ктт, Ктн
                ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_KTT],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_KTN]
                    },
                    Tuple.Create(rowValue.Values[MeterPoints.COL_TRANS_KTT] as double?, rowValue.Values[MeterPoints.COL_TRANS_KTN] as double?),
                    (value, parent, args) =>
                    {
                        ImportHelpers.FindOrCreate.SetReductiveTransformerToMeterPoint(meterPoint, value.Item1, value.Item2);
                        return meterPoint;
                    });

                // ТТ
                // строка ТТ, Фаза 1
                var tt1Link = rowValue.Values[MeterPoints.COL_TRANS_CURRENT_PHASE1] as EntityLink;
                tt1RowValue = tt1Link != null
                    ? Helper.GetRowValue(_transCurrents, tt1Link, Helpers.GetString)
                    : null;
                // строка ТТ, Фаза 2
                var tt2Link = rowValue.Values[MeterPoints.COL_TRANS_CURRENT_PHASE2] as EntityLink;
                tt2RowValue = tt2Link != null
                    ? Helper.GetRowValue(_transCurrents, tt2Link, Helpers.GetString)
                    : null;
                // строка ТТ, Фаза 3
                var tt3Link = rowValue.Values[MeterPoints.COL_TRANS_CURRENT_PHASE3] as EntityLink;
                tt3RowValue = tt3Link != null
                    ? Helper.GetRowValue(_transCurrents, tt3Link, Helpers.GetString)
                    : null;
                // чистим историю замен ТТ свежее дат установки
                if (existingEntity != null)
                {
                    var dt = new []
                    {
                        tt1RowValue == null ? null : tt1RowValue.Values[TransCurrents.COL_INSTALLDATE] as DateTime?,
                        tt2RowValue == null ? null : tt2RowValue.Values[TransCurrents.COL_INSTALLDATE] as DateTime?,
                        tt3RowValue == null ? null : tt3RowValue.Values[TransCurrents.COL_INSTALLDATE] as DateTime?
                    }
                    .Max();
                    if (dt.HasValue)
                    {
                        var linkSettings = existingEntity.AttributeMeterPointToMeterLinkSettings;
                        if (linkSettings != null)
                        {
                            var historyToRemove = linkSettings.AttributeMeasureCurrentTransformersInfo.GetValuesInfo()
                                .Where(x => x.Value.AttributeInstallDt.GetValueOrDefault() > dt.Value)
                                .ToArray();
                            foreach (var attributeValue in historyToRemove)
                                attributeValue.Remove();
                        }
                    }
                }
                ImportMeterPointRowCurrentTransformers(result, rowValue, meterPoint, tt1RowValue, tt2RowValue, tt3RowValue, entities, isolationLevels);

                // ТН
                // строка ТН, Фаза 1
                var tn1Link = rowValue.Values[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] as EntityLink;
                tn1RowValue = tn1Link != null
                    ? Helper.GetRowValue(_transVoltages, tn1Link, Helpers.GetString)
                    : null;
                // строка ТН, Фаза 2
                var tn2Link = rowValue.Values[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] as EntityLink;
                tn2RowValue = tn2Link != null
                    ? Helper.GetRowValue(_transVoltages, tn2Link, Helpers.GetString)
                    : null;
                // строка ТН, Фаза 3
                var tn3Link = rowValue.Values[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] as EntityLink;
                tn3RowValue = tn3Link != null
                    ? Helper.GetRowValue(_transVoltages, tn3Link, Helpers.GetString)
                    : null;
                ImportMeterPointRowVoltageTransformers(result, rowValue, meterPoint, tn1RowValue, tn2RowValue, tn3RowValue, entities, isolationLevels);
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
                    if (entities[MeterPoints.COL_METER_TYPE] == entity)
                        entities[MeterPoints.COL_METER_TYPE] = null;
                    if (entities[MeterPoints.COL_RTU] == entity)
                        entities[MeterPoints.COL_RTU] = null;
                    if (entities[MeterPoints.COL_TRANS_CURRENT_PHASE1] == entity)
                        entities[MeterPoints.COL_TRANS_CURRENT_PHASE1] = null;
                    if (entities[MeterPoints.COL_TRANS_CURRENT_PHASE2] == entity)
                        entities[MeterPoints.COL_TRANS_CURRENT_PHASE2] = null;
                    if (entities[MeterPoints.COL_TRANS_CURRENT_PHASE3] == entity)
                        entities[MeterPoints.COL_TRANS_CURRENT_PHASE3] = null;
                    if (entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] == entity)
                        entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] = null;
                    if (entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] == entity)
                        entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] = null;
                    if (entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] == entity)
                        entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] = null;
                    if (entities[MeterPoints.COL_CONSUMER] == entity)
                        entities[MeterPoints.COL_CONSUMER] = null;
                    
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

                ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
            }
            finally
            {
                _importedMeters[rowValue.RowIndex] = (MeterEquipment)entities[MeterPoints.COL_METER_TYPE];
                if (rtuRowValue != null && entities[MeterPoints.COL_RTU] != null)
                    _importedRtus[rtuRowValue.RowIndex] = (ChannelizingEquipment)entities[MeterPoints.COL_RTU];
                if (tt1RowValue != null && entities[MeterPoints.COL_TRANS_CURRENT_PHASE1] != null)
                    _importedTransCurrents[tt1RowValue.RowIndex] = (CurrentTransformer)entities[MeterPoints.COL_TRANS_CURRENT_PHASE1];
                if (tt2RowValue != null && entities[MeterPoints.COL_TRANS_CURRENT_PHASE2] != null)
                    _importedTransCurrents[tt2RowValue.RowIndex] = (CurrentTransformer)entities[MeterPoints.COL_TRANS_CURRENT_PHASE2];
                if (tt3RowValue != null && entities[MeterPoints.COL_TRANS_CURRENT_PHASE3] != null)
                    _importedTransCurrents[tt3RowValue.RowIndex] = (CurrentTransformer)entities[MeterPoints.COL_TRANS_CURRENT_PHASE3];
                if (tn1RowValue != null && entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] != null)
                    _importedTransVoltages[tn1RowValue.RowIndex] = (VoltageTransformer)entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1];
                if (tn2RowValue != null && entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] != null)
                    _importedTransVoltages[tn2RowValue.RowIndex] = (VoltageTransformer)entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2];
                if (tn3RowValue != null && entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] != null)
                    _importedTransVoltages[tn3RowValue.RowIndex] = (VoltageTransformer)entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3];
                if (naturalPersonRowValue != null && entities[MeterPoints.COL_CONSUMER] as NaturalPerson != null)
                    _importedNaturalPersons[naturalPersonRowValue.RowIndex] = entities[MeterPoints.COL_CONSUMER] as NaturalPerson;
                if (legalEntityRowValue != null && entities[MeterPoints.COL_CONSUMER] as LegalEntity != null)
                    _importedLegalEntities[legalEntityRowValue.RowIndex] = entities[MeterPoints.COL_CONSUMER] as LegalEntity;

                result.TotalCheckedLinesCount++;
                Helper.NotifyPercentRow();
            }
        }

        // импортировать ПС
        private static RDInstance ImportMeterPointRowSubstation(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue,
            Dictionary<int, RDInstance> entities, RDInstance lastParentClassifierItem)
        {
            // ПС

            // Уровень напряжения ПС
            entities[MeterPoints.COL_SUBSTATION_VOLTAGE] = null;
            var voltage = rowValue.Values[MeterPoints.COL_SUBSTATION_VOLTAGE] as VoltageEnumItem;
            if (voltage != null)
            {
                // создать группу подстанций по уровню напряжения
                entities[MeterPoints.COL_SUBSTATION_VOLTAGE] = ImportHelpers.FindOrCreateEntity(result,
                    new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_VOLTAGE] },
                    string.Format("ПС {0}", voltage.Caption),
                    lastParentClassifierItem,
                    ImportHelpers.FindOrCreate.ClassifierNodes.OfType<SubstationsGroupByVoltage>);
            }
            lastParentClassifierItem = entities[MeterPoints.COL_SUBSTATION_VOLTAGE] ?? lastParentClassifierItem;

            // ПС
            entities[MeterPoints.COL_SUBSTATION_SUBSTATION] = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_SUBSTATION] },
                rowValue.Values[MeterPoints.COL_SUBSTATION_SUBSTATION] as string,
                lastParentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<HighVoltageSubstation>);
            _placements[rowValue.Values[MeterPoints.COL_SUBSTATION_SUBSTATION] as string ?? string.Empty] = (ClassifierItem)entities[MeterPoints.COL_SUBSTATION_SUBSTATION];

            if (entities[MeterPoints.COL_SUBSTATION_SUBSTATION] != null)
            {
                // ПС - Высокая сторона

                // Уровень напряжения РУ
                entities[MeterPoints.COL_SUBSTATION_HI_VOLTAGE] = ImportHelpers.FindOrCreateEntity(result,
                    new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_HI_VOLTAGE] },
                    rowValue.Values[MeterPoints.COL_SUBSTATION_HI_VOLTAGE] as VoltageEnumItem,
                    entities[MeterPoints.COL_SUBSTATION_SUBSTATION],
                    ImportHelpers.FindOrCreate.ClassifierNodes.Switchgear);

                if (rowValue.Values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME] != null)
                {
                    // Секция шин
                    entities[MeterPoints.COL_SUBSTATION_HI_BUSBAR] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_HI_BUSBAR] },
                        rowValue.Values[MeterPoints.COL_SUBSTATION_HI_BUSBAR] as string ?? "СШ",
                        entities[MeterPoints.COL_SUBSTATION_HI_VOLTAGE],
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<BusbarSection>);

                    //Ячейка
                    //Тип ячейки
                    entities[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME] = ImportHelpers.FindOrCreateEntity(result,
                        new[]
                        {
                            _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME],
                            _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_HI_CUBICLE_TYPE]
                        },
                        rowValue.Values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME] as string,
                        entities[MeterPoints.COL_SUBSTATION_HI_BUSBAR],
                        ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle,
                        true,
                        // тип ячейки
                        rowValue.Values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_TYPE]);
                }

                // ПС - Низкая сторона

                // Уровень напряжения РУ
                entities[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] = ImportHelpers.FindOrCreateEntity(result,
                    new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_LO_VOLTAGE] },
                    rowValue.Values[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] as VoltageEnumItem,
                    entities[MeterPoints.COL_SUBSTATION_SUBSTATION],
                    ImportHelpers.FindOrCreate.ClassifierNodes.Switchgear);

                if (rowValue.Values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME] != null)
                {
                    // Секция шин
                    entities[MeterPoints.COL_SUBSTATION_LO_BUSBAR] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_LO_BUSBAR] },
                        rowValue.Values[MeterPoints.COL_SUBSTATION_LO_BUSBAR] as string ?? "СШ",
                        entities[MeterPoints.COL_SUBSTATION_LO_VOLTAGE],
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<BusbarSection>);

                    //Ячейка
                    //Тип ячейки
                    entities[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME] = ImportHelpers.FindOrCreateEntity(result,
                        new[]
                        {
                            _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME],
                            _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_LO_CUBICLE_TYPE]
                        },
                        rowValue.Values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME] as string,
                        entities[MeterPoints.COL_SUBSTATION_LO_BUSBAR],
                        ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle,
                        true,
                        // тип ячейки
                        rowValue.Values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_TYPE]);

                    // Линия/фидер
                    entities[MeterPoints.COL_SUBSTATION_LO_POWERLINE] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SUBSTATION_LO_POWERLINE] },
                        rowValue.Values[MeterPoints.COL_SUBSTATION_LO_POWERLINE] as string,
                        entities[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME],
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<PowerLine>);
                }
            }

            return lastParentClassifierItem;
        }

        // импортировать РП
        private static void ImportMeterPointRowSwitchgearSubstation(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue,
            Dictionary<int, RDInstance> entities, RDInstance parentClassifierItem)
        {
            // РП

            var substationLoVoltage = entities[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] as Switchgear;
            if (substationLoVoltage != null)
            {
                entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] = ImportHelpers.FindOrCreateEntity(result,
                    new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] },
                    rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] as string,
                    parentClassifierItem,
                    ImportHelpers.FindOrCreate.ClassifierNodes.SwitchGearSubstation,
                    true,
                    // уровень напряжения
                    substationLoVoltage.AttributeVoltage);
                _placements[rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] as string ?? string.Empty] = (ClassifierItem)entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION];

                if (entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] != null &&
                    (rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] != null ||
                        rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] != null))
                {
                    // Секция шин
                    entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR] },
                        rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR] as string ?? "СШ",
                        entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION],
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<BusbarSection>);

                    // Ячейка, входящая от ПС
                    entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] },
                        rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] as string,
                        entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR],
                        ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle,
                        true,
                        // линия от ПС
                        entities[MeterPoints.COL_SUBSTATION_LO_POWERLINE]);

                    // Ячейка, отходящая в ТП
                    entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] },
                        rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] as string,
                        entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR],
                        ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle);

                    // Линия/фидер
                    entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] },
                        rowValue.Values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] as string,
                        entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT],
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<PowerLine>);
                }
            }
        }

        // импортировать ТП
        private static void ImportMeterPointRowLoSubstation(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue,
            Dictionary<int, RDInstance> entities, RDInstance parentClassifierItem)
        {
            // ТП

            // Уровень напряжения ТП, если есть уровень напряжения ПС
            RDInstance loSubstationVoltageGroup = null;
            if (rowValue.Values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] != null &&
                rowValue.Values[MeterPoints.COL_SUBSTATION_VOLTAGE] != null)
            {
                // создать группу подстанций по уровню напряжения ТП
                loSubstationVoltageGroup = ImportHelpers.FindOrCreateEntity(result,
                    new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_SUBSTATION] },
                    string.Format("ТП {0}", VoltageEnumItem.Instances.Near0dot4kV.Caption),
                    parentClassifierItem,
                    ImportHelpers.FindOrCreate.ClassifierNodes.OfType<SubstationsGroupByVoltage>);
            }

            parentClassifierItem = loSubstationVoltageGroup ?? parentClassifierItem;

            entities[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_SUBSTATION] },
                rowValue.Values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] as string,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<LowVoltageSubstation>);
            _placements[rowValue.Values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] as string ?? string.Empty] = (ClassifierItem)entities[MeterPoints.COL_LO_SUBSTATION_SUBSTATION];

            if (entities[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] != null)
            {
                // ТП - Высокая сторона, при наличии ПС - низкая сторона

                if (entities[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] != null)
                {
                    // Уровень напряжения РУ (по уровню напряжения ПС - Низкая сторона)
                    var loSubstationHiVoltage = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_SUBSTATION] },
                        rowValue.Values[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] as VoltageEnumItem,
                        entities[MeterPoints.COL_LO_SUBSTATION_SUBSTATION],
                        ImportHelpers.FindOrCreate.ClassifierNodes.Switchgear);

                    if (rowValue.Values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME] != null)
                    {
                        // Секция шин
                        entities[MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR] = ImportHelpers.FindOrCreateEntity(result,
                            new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR] },
                            rowValue.Values[MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR] as string ?? "СШ",
                            loSubstationHiVoltage,
                            ImportHelpers.FindOrCreate.ClassifierNodes.OfType<BusbarSection>);

                        //Ячейка
                        //Тип ячейки
                        entities[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME] = ImportHelpers.FindOrCreateEntity(result,
                            new[]
                            {
                                    _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME],
                                    _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_TYPE]
                            },
                            rowValue.Values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME] as string,
                            entities[MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR],
                            ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle,
                            true,
                            // тип ячейки
                            rowValue.Values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_TYPE],
                            // линия от ПС/РП
                            entities[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] ?? entities[MeterPoints.COL_SUBSTATION_LO_POWERLINE]);
                    }
                }

                // ТП - Низкая сторона

                // Уровень напряжения РУ
                var loSubstationLoVoltage = ImportHelpers.FindOrCreateEntity(result,
                    new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_SUBSTATION] },
                    VoltageEnumItem.Instances.Near0dot4kV,
                    entities[MeterPoints.COL_LO_SUBSTATION_SUBSTATION],
                    ImportHelpers.FindOrCreate.ClassifierNodes.Switchgear);

                if (rowValue.Values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME] != null)
                {
                    // Секция шин
                    entities[MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR] },
                        rowValue.Values[MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR] as string ?? "СШ",
                        loSubstationLoVoltage,
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<BusbarSection>);

                    //Ячейка
                    //Тип ячейки
                    entities[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME] = ImportHelpers.FindOrCreateEntity(result,
                        new[]
                        {
                                _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME],
                                _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_TYPE]
                        },
                        rowValue.Values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME] as string,
                        entities[MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR],
                        ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle,
                        true,
                        // тип ячейки
                        rowValue.Values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_TYPE]);

                    // Линия/фидер
                    entities[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] = ImportHelpers.FindOrCreateEntity(result,
                        new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] },
                        rowValue.Values[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] as string,
                        entities[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME],
                        ImportHelpers.FindOrCreate.ClassifierNodes.OfType<PowerLine>);
                }
            }
        }

        // импортировать адрес
        private static void ImportMeterPointRowAddress(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue,
            Dictionary<int, RDInstance> entities, RDInstance parentClassifierItem)
        {
            var addressFias = rowValue.Values[MeterPoints.COL_ADDRESS_FIAS] as FiasSuggestionData;
            if (addressFias == null) return;

            // Субъект РФ
            var subjectRf = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS] },
                addressFias.Region,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<SubjectRF>);
            parentClassifierItem = subjectRf ?? parentClassifierItem;

            // Район административного деления
            var district = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS] },
                addressFias.Area,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<District>);
            parentClassifierItem = district ?? parentClassifierItem;

            // Населенный пункт + район населенного пункта
            
            var centerOfPopulationValue = addressFias.City;
            var centerOfPopulationZoneValue = addressFias.CityDistrict;
            var centerOfPopulationZoneValue2 = addressFias.Settlement;
            
            // если отсутствует город, то используем населенный пункт
            // иначе все что ниже города - районы населеднного пункта
            if (string.IsNullOrEmpty(centerOfPopulationValue))
            {
                centerOfPopulationValue = addressFias.Settlement;
                centerOfPopulationZoneValue = null;
                centerOfPopulationZoneValue2 = null;
            }
            var centerOfPopulation = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS] },
                centerOfPopulationValue,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<CenterOfPopulation>);
            parentClassifierItem = centerOfPopulation ?? parentClassifierItem;

            // район населенного пункта 1
            var centerOfPopulationZone = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS] },
                centerOfPopulationZoneValue,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<CenterOfPopulationZone>);
            parentClassifierItem = centerOfPopulationZone ?? parentClassifierItem;

            // район населенного пункта 2
            var centerOfPopulationZone2 = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS] },
                centerOfPopulationZoneValue2,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<CenterOfPopulationZone>);
            parentClassifierItem = centerOfPopulationZone2 ?? parentClassifierItem;

            // улица
            var street = ImportHelpers.FindOrCreateEntity(result,
                new[] { _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS] },
                addressFias.Street,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.OfType<Street>);
            parentClassifierItem = street ?? parentClassifierItem;

            // дом
            var building = ImportHelpers.FindOrCreateEntity(result,
                new[] 
                { 
                    _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS],
                    _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FLAT]
                },
                addressFias.House,
                parentClassifierItem,
                ImportHelpers.FindOrCreate.ClassifierNodes.CustomBuilding,
                true,
                // тип сооружения (частный дом - если не указана квартира)
                rowValue.Values[MeterPoints.COL_ADDRESS_FLAT] == null ? (BaseClassClassInfo)DwellingHouse.GetClassInfo() : TenementHouse.GetClassInfo(),
                // линия/фидер
                entities[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE]);
            parentClassifierItem = building ?? parentClassifierItem;

            // помещение в доме
            if (rowValue.Values[MeterPoints.COL_ADDRESS_FLAT] is string flatName &&
                flatName.IndexOf(Helper.ROOM_PRIFIX_NAME, StringComparison.InvariantCultureIgnoreCase) != -1)
            {
                var room = ImportHelpers.FindOrCreateEntity(result,
                    new[] 
                    { 
                        _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FIAS],
                        _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_ADDRESS_FLAT]
                    },
                    flatName,
                    parentClassifierItem,
                    ImportHelpers.FindOrCreate.ClassifierNodes.Room,
                    true);
                parentClassifierItem = room ?? parentClassifierItem;
            }

            entities[MeterPoints.COL_ADDRESS_FIAS] = parentClassifierItem;
        }

        // импортировать ПУ
        private static MeterEquipment ImportMeterPointRowMeter(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue, MeterPoint existingMeterPoint)
        {
            var existingMeter = existingMeterPoint == null ? null :  existingMeterPoint.AttributeElectricityMeter;
            var meter = existingMeter ?? ImportHelpers.FindOrCreateEntity(result,
                new[]
                {
                    rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_TYPE],
                    rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_SERIALNUMBER],
                    rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_ACCURACYCLASS]
                },
                rowValue.Values[MeterPoints.COL_METER_TYPE] as MeterEquipmentClassInfo,
                (value, parent, args) =>
                {
                    var meterResult = _existingMeters
                        .Where(x => x.Class == value)
                        .AsEnumerable()
                        .FirstOrDefault(x => Equals(x.AttributeSerialNumber, rowValue.Values[MeterPoints.COL_METER_SERIALNUMBER]));
                    if (meterResult == null)
                    {
                        meterResult = (MeterEquipment)ImportHelpers.FindOrCreate.Entity(value);
                        meterResult.AttributeSerialNumber = rowValue.Values[MeterPoints.COL_METER_SERIALNUMBER] as string;
                        _existingMeters.Add(meterResult);
                    }

                    return meterResult;
                });

            // Дата выпуска
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_CREATEDATE] },
                rowValue.Values[MeterPoints.COL_METER_CREATEDATE] as DateTime?,
                (value, parent, args) =>
                {
                    if (existingMeter == null && meter.AttributeReleaseDate.HasValue && meter.AttributeReleaseDate != value)
                        ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        meter.AttributeReleaseDate = value;
                    return meter;
                });
            
            // Дата последней поверки
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_CHECKDATE] },
                rowValue.Values[MeterPoints.COL_METER_CHECKDATE] as DateTime?,
                (value, parent, args) =>
                {
                    var checkingInstance = meter as IMetrologicalCheckingInstance;
                    if (checkingInstance == null)
                        return meter;
                    if (existingMeter == null && checkingInstance.AttributeLastCalibrationDate.HasValue && checkingInstance.AttributeLastCalibrationDate != value)
                        ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        checkingInstance.AttributeLastCalibrationDate = value;
                    return meter;
                });

            // Часовой пояс
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_TIMEZONE] },
                rowValue.Values[MeterPoints.COL_METER_TIMEZONE] as TimeZone,
                (value, parent, args) =>
                {
                    var timezoneEntity = meter as IEntityWithTimeZone;
                    if (timezoneEntity == null)
                        return meter;
                    if (existingMeter == null && timezoneEntity.AttributeTimeZone != null && timezoneEntity.AttributeTimeZone != value)
                        ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        timezoneEntity.AttributeTimeZone = value;
                    return meter;
                });

            // Связной номер
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_NETWORKNUMBER] },
                rowValue.Values[MeterPoints.COL_METER_NETWORKNUMBER] as string,
                (value, parent, args) =>
                {
                    // костыльная установка атрибута "Номер модема" для ПУ Фобос
                    var fobosMeter = meter as FobosMeter;
                    if (fobosMeter != null)
                    {
                        if (existingMeter == null && !string.IsNullOrEmpty(fobosMeter.AttributeModemId) && fobosMeter.AttributeModemId != value)
                            ImportHelpers.AddWarn("Атрибут \"Номер модема\" не был установлен, т.к. уже содержит другое значение");
                        else
                            fobosMeter.AttributeModemId = value;
                    }

                    // физический адрес для оборудования DLMS
                    var dlmsWithPhysicalAddress = meter as IEquipmentDlmsWithPhysicalAddress;
                    int physicalAddress;
                    if (dlmsWithPhysicalAddress != null && int.TryParse(value, out physicalAddress))
                        dlmsWithPhysicalAddress.AttributePhysicalAddress = physicalAddress;

                    // логический адрес для оборудования DLMS
                    var dlmsWithPhysicalAndLogicalAddress = meter as IEquipmentDlmsWithPhysicalAndLogicalAddresses;
                    if (dlmsWithPhysicalAndLogicalAddress != null)
                        dlmsWithPhysicalAndLogicalAddress.AttributeLogicalAddress = 1;

                    // связной номер
                    var newEquipment = meter as IEquipmentWithNetworkId;
                    if (newEquipment == null)
                        return meter;

                    if (existingMeter == null && !string.IsNullOrEmpty(newEquipment.AttributeNetworkId) && newEquipment.AttributeNetworkId != value)
                        ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        newEquipment.AttributeNetworkId = value;
                    
                    return meter;
                });

            // Пользователь
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_LOGIN] },
                rowValue.Values[MeterPoints.COL_METER_LOGIN] as string,
                (value, parent, args) =>
                {
                    // уровень доступа для оборудования DLMS
                    var dlmsWithAccessLevel = meter as IEquipmentDlmsWithPasswordAndAccessLevel;
                    AccessLevelToEquipmentItem dlmcAccessLevel;
                    if (dlmsWithAccessLevel != null && Helpers.AccessLevelToEquipmentItems.Value.TryGetValue(value ?? string.Empty, out dlmcAccessLevel))
                        dlmsWithAccessLevel.AttributeEquipmentDlmsAccessLevel = dlmcAccessLevel;

                    // уровень доступа для счетчика C12xx
                    var c12xxWithAccessLevel = meter as IEquipmentC12xxWithUserAndPassword;
                    C12xxAccessLevelItem c12xxAccessLevel;
                    if (c12xxWithAccessLevel != null && Helpers.AccessLevelToC12xx.Value.TryGetValue(value ?? string.Empty, out c12xxAccessLevel))
                        c12xxWithAccessLevel.AttributeAccessLevel = c12xxAccessLevel;

                    // уровень доступа для счетчика CascadeSoft
                    var cascadeSoftWithAccessLevel = meter as ICascadeSoftAccessSettings;
                    CascadeSoftAccessLevelsItem cascadeSoftAccessLevel;
                    if (cascadeSoftWithAccessLevel != null && Helpers.AccessLevelToCascadeSoft.Value.TryGetValue(value ?? string.Empty, out cascadeSoftAccessLevel))
                        cascadeSoftWithAccessLevel.AttributeAccessLevel = cascadeSoftAccessLevel;

                    // уровень доступа для счетчика Mercury23x
                    var mercury23xWithAccessLevel = meter as IMercury23xAccessSettings;
                    Mercury23xAccessLevelsItem mercury23xAccessLevel;
                    if (mercury23xWithAccessLevel != null && Helpers.AccessLevelToMercury23x.Value.TryGetValue(value ?? string.Empty, out mercury23xAccessLevel))
                        mercury23xWithAccessLevel.AttributeAccessLevel = mercury23xAccessLevel;

                    // уровень доступа для счетчика MIR
                    var mirWithAccessLevel = meter as IEquipmentWithPasswordAndMIRAccessLevel;
                    MIRAccessLevelsItem mirAccessLevel;
                    if (mirWithAccessLevel != null && Helpers.AccessLevelToMIR.Value.TryGetValue(value ?? string.Empty, out mirAccessLevel))
                        mirWithAccessLevel.AttributeAccessLevel = mirAccessLevel;

                    // уровень доступа для счетчика Милур
                    var milandWithAccessLevel = meter as Milandr;
                    MilurX07AccessLevelItem milandrAccessLevel;
                    if (milandWithAccessLevel != null && Helpers.AccessLevelToMilandr.Value.TryGetValue(value ?? string.Empty, out milandrAccessLevel))
                        milandWithAccessLevel.AttributeAccessLevel = milandrAccessLevel;

                    var userAndPasswordEquipment = meter as IEquipmentWithUserAndPassword;
                    if (userAndPasswordEquipment == null)
                        return meter;

                    if (existingMeter == null && !string.IsNullOrEmpty(userAndPasswordEquipment.AttributeUser) && userAndPasswordEquipment.AttributeUser != value)
                        ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        userAndPasswordEquipment.AttributeUser = value;
                    return meter;
                });

            // Пароль
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_PASSWORD] },
                rowValue.Values[MeterPoints.COL_METER_PASSWORD] as string,
                (value, parent, args) =>
                {
                    var passwordEquipment = meter as IEquipmentWithPassword;
                    if (passwordEquipment == null)
                        return meter;
                    if (existingMeter == null && !string.IsNullOrEmpty(passwordEquipment.AttributePassword) && passwordEquipment.AttributePassword != value)
                        ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        passwordEquipment.AttributePassword = value;
                    return meter;
                });

            // маршруты (только прямые, через другое каналообразующее оборудование проставляется после импорта всего)
            ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_ROUTE] },
                rowValue.Values[MeterPoints.COL_METER_ROUTE] as ImportRoute[],
                (value, parent, args) =>
                {
                    // прямые маршруты
                    foreach (var route in value)
                    {
                        if (route.ClassInfo.InheritsFrom(NonDirectRouteClassInfo.Get()))
                            continue;
                        // не добавляется, если маршруты равны
                        // приоритеты выставляются при добавлении триггером, если не задано явно в ОЛ 
                        var route1 = route;
                        var rdRoute = meter.AttributeRoutes.GetValues().FirstOrDefault(x => route1.Equals(x));
                        if (rdRoute == null)
                            rdRoute = ImportHelpers.FindOrCreate.Entity(meter.AttributeRoutes, route.ClassInfo);
                        // связывание специфичных параметров
                        route.Assign(rdRoute);
                    }

                    return meter;
                });

            //// специальные параметры
            //ImportHelpers.FindOrCreateEntity(result,
            //    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_METER_PARAMS] },
            //    rowValue.Values[MeterPoints.COL_METER_PARAMS] as Dictionary<Guid, object>,
            //    (value, parent, args) =>
            //    {
            //        var meterClassInfo = rowValue.Values[MeterPoints.COL_METER_TYPE] as MeterEquipmentClassInfo;
            //        if (meterClassInfo == null)
            //            return meter;
                    
            //        new Helpers.ClassAttributesParser(meterClassInfo).Assign(meter, value);
                    
            //        return meter;
            //    });

            return meter;
        }

        // импортировать ТТ ТУ
        private static void ImportMeterPointRowCurrentTransformers(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue,
            MeterPoint meterPoint, Sheet.RowValueInfo tt1RowValue, Sheet.RowValueInfo tt2RowValue, Sheet.RowValueInfo tt3RowValue,
            Dictionary<int, RDInstance> entities, IsolationLevel[] isolationLevels)
        {
            // ТТ, Фаза 1
            entities[MeterPoints.COL_TRANS_CURRENT_PHASE1] = ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_CURRENT_PHASE1] },
                tt1RowValue,
                (value, parent, args) =>
                {
                    // импорт с пробросом исключения
                    var resultTrans = ImportTransCurrentRow(result, value, true);
                    if (isolationLevels != null)
                    {
                        _importedEntityIsolationLevels[resultTrans] = isolationLevels;
                        resultTrans.AppendIsolationLevels(isolationLevels);
                    }

                    return resultTrans;
                });

            // ТТ, Фаза 2
            entities[MeterPoints.COL_TRANS_CURRENT_PHASE2] = ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_CURRENT_PHASE2] },
                tt2RowValue,
                (value, parent, args) =>
                {
                    // импорт с пробросом исключения
                    var resultTrans = ImportTransCurrentRow(result, value, true);
                    if (isolationLevels != null)
                    {
                        _importedEntityIsolationLevels[resultTrans] = isolationLevels;
                        resultTrans.AppendIsolationLevels(isolationLevels);
                    }

                    return resultTrans;
                });

            // ТТ, Фаза 3
            entities[MeterPoints.COL_TRANS_CURRENT_PHASE3] = ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_CURRENT_PHASE3] },
                tt3RowValue,
                (value, parent, args) =>
                {
                    // импорт с пробросом исключения
                    var resultTrans = ImportTransCurrentRow(result, value, true);
                    if (isolationLevels != null)
                    {
                        _importedEntityIsolationLevels[resultTrans] = isolationLevels;
                        resultTrans.AppendIsolationLevels(isolationLevels);
                    }

                    return resultTrans;
                });

            var transCurrents = new[]
            {
                (CurrentTransformer)entities[MeterPoints.COL_TRANS_CURRENT_PHASE1],
                (CurrentTransformer)entities[MeterPoints.COL_TRANS_CURRENT_PHASE2],
                (CurrentTransformer)entities[MeterPoints.COL_TRANS_CURRENT_PHASE3],
            };

            // ТТ - Дата установки
            ImportHelpers.FindOrCreateEntity(result,
                new[]
                {
                    entities[MeterPoints.COL_TRANS_CURRENT_PHASE1] != null ? rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_CURRENT_PHASE1] : null,
                    entities[MeterPoints.COL_TRANS_CURRENT_PHASE2] != null ? rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_CURRENT_PHASE2] : null,
                    entities[MeterPoints.COL_TRANS_CURRENT_PHASE3] != null ? rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_CURRENT_PHASE3] : null,
                    tt1RowValue == null ? null : tt1RowValue.Sheet.Worksheet.Cells[tt1RowValue.RowIndex, TransCurrents.COL_INSTALLDATE],
                    tt2RowValue == null ? null : tt2RowValue.Sheet.Worksheet.Cells[tt2RowValue.RowIndex, TransCurrents.COL_INSTALLDATE],
                    tt3RowValue == null ? null : tt3RowValue.Sheet.Worksheet.Cells[tt3RowValue.RowIndex, TransCurrents.COL_INSTALLDATE],
                }.Where(x => x != null).ToArray(),
                new[]
                {
                    tt1RowValue == null ? null : tt1RowValue.Values[TransCurrents.COL_INSTALLDATE] as DateTime?,
                    tt2RowValue == null ? null : tt2RowValue.Values[TransCurrents.COL_INSTALLDATE] as DateTime?,
                    tt3RowValue == null ? null : tt3RowValue.Values[TransCurrents.COL_INSTALLDATE] as DateTime?
                }.Max(),
                (value, parent, args) =>
                {
                    ImportHelpers.FindOrCreate.SetCurrentTransformerToMeterPoint(meterPoint, transCurrents, value);
                    return meterPoint;
                });
        }

        // импортировать ТН ТУ
        private static void ImportMeterPointRowVoltageTransformers(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue,
            MeterPoint meterPoint, Sheet.RowValueInfo tn1RowValue, Sheet.RowValueInfo tn2RowValue, Sheet.RowValueInfo tn3RowValue,
            Dictionary<int, RDInstance> entities, IsolationLevel[] isolationLevels)
        {
            // СШ ТУ
            var meterPointBusbar = meterPoint?.RelationsCubicleAttributeMeterPoint
                .SelectMany(x => x.RelationsBusbarSectionAttributeCubicles)
                .FirstOrDefault();

            // ТН, Фаза 1
            entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] = ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_VOLTAGE_PHASE1] },
                tn1RowValue,
                (value, parent, args) =>
                {
                    VoltageTransformer resultTrans = null;
                    if (meterPointBusbar == null)
                        ImportHelpers.AddWarn(string.Format(
                            "ТН не был сопоставлен ТУ {0}, " +
                            "т.к. не удалось найти Секцию шин, к которой относится ТУ",
                            meterPoint));
                    else
                    {
                        // импорт с пробросом исключения
                        resultTrans = ImportTransVoltageRow(result, value, true);
                        if (isolationLevels != null)
                        {
                            _importedEntityIsolationLevels[resultTrans] = isolationLevels;
                            resultTrans.AppendIsolationLevels(isolationLevels);
                        }
                    }

                    return resultTrans;
                });

            // ТН, Фаза 2
            entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] = ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_VOLTAGE_PHASE2] },
                tn2RowValue,
                (value, parent, args) =>
                {
                    VoltageTransformer resultTrans = null;
                    if (meterPointBusbar == null)
                        ImportHelpers.AddWarn(string.Format(
                            "ТН не был сопоставлен ТУ {0}, " +
                            "т.к. не удалось найти Секцию шин, к которой относится ТУ",
                            meterPoint));
                    else
                    {
                        // импорт с пробросом исключения
                        resultTrans = ImportTransVoltageRow(result, value, true);
                        if (isolationLevels != null)
                        {
                            _importedEntityIsolationLevels[resultTrans] = isolationLevels;
                            resultTrans.AppendIsolationLevels(isolationLevels);
                        }
                    }

                    return resultTrans;
                });

            // ТН, Фаза 3
            entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] = ImportHelpers.FindOrCreateEntity(result,
                new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_VOLTAGE_PHASE3] },
                tn3RowValue,
                (value, parent, args) =>
                {
                    VoltageTransformer resultTrans = null;
                    if (meterPointBusbar == null)
                        ImportHelpers.AddWarn(string.Format(
                            "ТН не был сопоставлен ТУ {0}, " +
                            "т.к. не удалось найти Секцию шин, к которой относится ТУ",
                            meterPoint));
                    else
                    {
                        // импорт с пробросом исключения
                        resultTrans = ImportTransVoltageRow(result, value, true);
                        if (isolationLevels != null)
                        {
                            _importedEntityIsolationLevels[resultTrans] = isolationLevels;
                            resultTrans.AppendIsolationLevels(isolationLevels);
                        }
                    }

                    return resultTrans;
                });

            var transVoltages = new[]
            {
                    (VoltageTransformer)entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1],
                    (VoltageTransformer)entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2],
                    (VoltageTransformer)entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3],
                };

            // ячейка ТН
            if (meterPointBusbar != null && transVoltages.Any(x => x != null))
            {
                ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] != null ? _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_VOLTAGE_PHASE1] : null,
                        entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] != null ? _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_VOLTAGE_PHASE2] : null,
                        entities[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] != null ? _meterPoints.Worksheet.Cells[rowValue.RowIndex, MeterPoints.COL_TRANS_VOLTAGE_PHASE3] : null,
                        tn1RowValue == null ? null : tn1RowValue.Sheet.Worksheet.Cells[tn1RowValue.RowIndex, TransVoltages.COL_INSTALLDATE],
                        tn2RowValue == null ? null : tn2RowValue.Sheet.Worksheet.Cells[tn2RowValue.RowIndex, TransVoltages.COL_INSTALLDATE],
                        tn3RowValue == null ? null : tn3RowValue.Sheet.Worksheet.Cells[tn3RowValue.RowIndex, TransVoltages.COL_INSTALLDATE],
                    }.Where(x => x != null).ToArray(),
                    "яч. ТН",
                    meterPointBusbar,
                    ImportHelpers.FindOrCreate.ClassifierNodes.Cubicle,
                    true,
                    // тип ячейки
                    CubicleMeasureVoltageTransformerClassInfo.Get(),
                    // трансформаторы
                    transVoltages,
                    // ТН - Дата установки
                    new KeyValuePair<string, DateTime?>("VoltageTransformerInstallDate",
                        new[]
                        {
                            tn1RowValue == null ? null : tn1RowValue.Values[TransVoltages.COL_INSTALLDATE] as DateTime?,
                            tn2RowValue == null ? null : tn2RowValue.Values[TransVoltages.COL_INSTALLDATE] as DateTime?,
                            tn3RowValue == null ? null : tn3RowValue.Values[TransVoltages.COL_INSTALLDATE] as DateTime?
                        }.Max())
                );
            }
        }

        // импортировать УСПД
        private static ChannelizingEquipment ImportRtuRow(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue, bool throwIfException, 
            IsolationLevel[] isolationLevels)
        {
            if (rowValue == null || rowValue.Sheet.GetType() != typeof(Rtus))
                return null;
            var possibleId = Helper.TryGetGuidIdentifier(rowValue.Values[Sheet.COL_NUM]);
            var existingEntity = possibleId.HasValue ? ChannelizingEquipment.Find(possibleId.Value) : null;
            // Если переимпортируется УСПД, чистим ему маршруты (перезатягиваем заново)
            if (existingEntity != null)
                existingEntity.AttributeRoutes.Clear();
            try
            {
                var rtu = ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_TYPE],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_SERIALNUMBER]
                    },
                    rowValue.Values[Rtus.COL_TYPE] as ChannelizingEquipmentClassInfo,
                    (value, parent, args) =>
                    {
                        var rtuResult = existingEntity ?? _existingRtus
                            .Where(x => x.Class == value)
                            .AsEnumerable()
                            .FirstOrDefault(x => Equals(x.AttributeSerialNumber, rowValue.Values[Rtus.COL_SERIALNUMBER]));
                        if (rtuResult == null)
                        {
                            rtuResult = (ChannelizingEquipment)ImportHelpers.FindOrCreate.Entity(value);
                            rtuResult.AttributeSerialNumber = rowValue.Values[Rtus.COL_SERIALNUMBER] as string;
                            _existingRtus.Add(rtuResult);
                        }
                        else if (existingEntity != null)
                            rtuResult.AttributeSerialNumber = rowValue.Values[Rtus.COL_SERIALNUMBER] as string;
                        return rtuResult;
                    });
                // Если заданы уровни изоляции, то добавляем их к тому, что уже есть у УСПД 
                // (именно добавляем, а не подменяем т.к. один УCПД может присутствовать в нескольких опросниках 
                // в разных контекстах изоляции)
                if (isolationLevels != null && isolationLevels.Any())
                {
                    _importedEntityIsolationLevels[rtu] = isolationLevels;
                    rtu.AppendIsolationLevels(isolationLevels);
                }

                // Дата выпуска
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_CREATEDATE] },
                    rowValue.Values[Rtus.COL_CREATEDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && rtu.AttributeReleaseDate.HasValue && rtu.AttributeReleaseDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            rtu.AttributeReleaseDate = value;
                        return rtu;
                    });

                // Дата установки
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_INSTALLDATE] },
                    rowValue.Values[Rtus.COL_INSTALLDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && rtu.AttributeInstallDate.HasValue && rtu.AttributeInstallDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            rtu.AttributeInstallDate = value;
                        return rtu;
                    });

                // Дата последней поверки
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_CHECKDATE] },
                    rowValue.Values[Rtus.COL_CHECKDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        var checkingInstance = rtu as IMetrologicalCheckingInstance;
                        if (checkingInstance == null)
                            return rtu;
                        if (existingEntity == null &&  checkingInstance.AttributeLastCalibrationDate.HasValue && checkingInstance.AttributeLastCalibrationDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            checkingInstance.AttributeLastCalibrationDate = value;
                        return rtu;
                    });

                // Часовой пояс
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_TIMEZONE] },
                    rowValue.Values[Rtus.COL_TIMEZONE] as TimeZone,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && rtu.AttributeTimeZone != null && rtu.AttributeTimeZone != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            rtu.AttributeTimeZone = value;
                        return rtu;
                    });

                // Связной номер
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_NETWORKNUMBER] },
                    rowValue.Values[Rtus.COL_NETWORKNUMBER] as string,
                    (value, parent, args) =>
                    {
                        var newEquipment = rtu as IEquipmentWithNetworkId;
                        if (newEquipment == null)
                            return rtu;
                        if (existingEntity == null && !string.IsNullOrEmpty(newEquipment.AttributeNetworkId) && newEquipment.AttributeNetworkId != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            newEquipment.AttributeNetworkId = value;
                        return rtu;
                    });
                var bs = rtu as BaseStationNB300;
                // Пользователь
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_LOGIN] },
                    rowValue.Values[Rtus.COL_LOGIN] as string,
                    (value, parent, args) =>
                    {
                        // Аметов 13.01.21 так как BaseStationNB300 не поддерживает интерфейсы IEquipmentWithUserAndPassword IEquipmentWithPassword
                        if (bs != null)
                            if (existingEntity == null && !string.IsNullOrEmpty(bs.AttributeLogin) && bs.AttributeLogin != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                bs.AttributeLogin = value;

                        var userAndPasswordEquipment = rtu as IEquipmentWithUserAndPassword;
                        if (userAndPasswordEquipment == null)
                            return rtu;
                        if (existingEntity == null && !string.IsNullOrEmpty(userAndPasswordEquipment.AttributeUser) && userAndPasswordEquipment.AttributeUser != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            userAndPasswordEquipment.AttributeUser = value;
                        return rtu;
                    });
                
                // Пароль
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_PASSWORD] },
                    rowValue.Values[Rtus.COL_PASSWORD] as string,
                    (value, parent, args) =>
                    {

                        // Аметов 13.01.21 Так как BaseStationNB300 не поддерживает интерфейсы IEquipmentWithUserAndPassword IEquipmentWithPassword
                        if (bs != null)
                            if (existingEntity == null && !string.IsNullOrEmpty(bs.AttributePassword) && bs.AttributePassword != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                bs.AttributePassword = value;

                        var passwordEquipment = rtu as IEquipmentWithPassword;
                        if (passwordEquipment == null)
                            return rtu;
                        if (existingEntity == null && !string.IsNullOrEmpty(passwordEquipment.AttributePassword) && passwordEquipment.AttributePassword != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            passwordEquipment.AttributePassword = value;
                        return rtu;
                    });

                // Аметов 13.01.21 Если введен логин и пароль, ставим галочку использовать авторизацию
                if (bs != null)
                    if (!string.IsNullOrEmpty(bs.AttributePassword) && !string.IsNullOrEmpty(bs.AttributeLogin))
                        bs.AttributeUseAuth = true;

                // маршруты (только прямые, через другое каналообразующее оборудование проставляется после импорта всего)
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_ROUTE] },
                    rowValue.Values[Rtus.COL_ROUTE] as ImportRoute[],
                    (value, parent, args) =>
                    {
                        foreach (var route in value)
                        {
                            // не добавляется, если маршруты равны или маршрут через другое УСПД
                            if (route.ClassInfo.InheritsFrom(NonDirectRouteClassInfo.Get()))
                                continue;
                            // приоритеты выставляются при добавлении триггером, если не задано явно в ОЛ 
                            var route1 = route;
                            var rdRoute = rtu.AttributeRoutes.GetValues().FirstOrDefault(x => route1.Equals(x));
                            if (rdRoute == null)
                                rdRoute = ImportHelpers.FindOrCreate.Entity(rtu.AttributeRoutes, route.ClassInfo);
                            // связывание специфичных параметров
                            route.Assign(rdRoute);
                        }
                        return rtu;
                    });

                //// специальные параметры
                //ImportHelpers.FindOrCreateEntity(result,
                //    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, Rtus.COL_PARAMS] },
                //    rowValue.Values[Rtus.COL_PARAMS] as Dictionary<Guid, object>,
                //    (value, parent, args) =>
                //    {
                //        var rtuClassInfo = rowValue.Values[Rtus.COL_TYPE] as ChannelizingEquipmentClassInfo;
                //        if (rtuClassInfo == null)
                //            return rtu;
                    
                //        new Helpers.ClassAttributesParser(rtuClassInfo).Assign(rtu, value);
                    
                //        return rtu;
                //    });
                
                return rtu;
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
                if (throwIfException) throw;
            }
            finally
            {
                result.TotalCheckedLinesCount++;
                Helper.NotifyPercentRow();
            }
            return null;
        }

        // импортировать абонента
        private static Consumer ImportConsumerRow(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue, bool throwIfException = false)
        {
            if (rowValue == null ||
                !new[] { typeof(NaturalPersons), typeof(LegalEntities) }.Contains(rowValue.Sheet.GetType()))
                return null;
            try
            {
                // Физ лица

                if (rowValue.Sheet == _naturalPersons)
                {
                    var possibleId = Helper.TryGetGuidIdentifier(rowValue.Values[Sheet.COL_NUM]);
                    var existingEntity = possibleId.HasValue ? NaturalPerson.Find(possibleId.Value) : null;
                    var naturalPerson = ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, NaturalPersons.COL_CURRENTACCOUNT],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, NaturalPersons.COL_LASTNAME],
                    },
                    NaturalPersonClassInfo.Get(),
                    (value, parent, args) =>
                    {
                        var wantedType = value.TypeOfInstance();
                        var consumerResult = existingEntity ?? _existingNaturals
                            .Where(x => x.GetType() == wantedType)
                            .AsEnumerable()
                            .FirstOrDefault(x =>
                                Equals(x.AttributeCurrentAccount, rowValue.Values[NaturalPersons.COL_CURRENTACCOUNT]) &&
                                Equals(x.AttributeLastName, rowValue.Values[NaturalPersons.COL_LASTNAME] ?? x.AttributeLastName));
                        if (consumerResult == null)
                        {
                            consumerResult = (NaturalPerson)ImportHelpers.FindOrCreate.Entity(value);
                            consumerResult.AttributeCurrentAccount = rowValue.Values[NaturalPersons.COL_CURRENTACCOUNT] as string;
                            consumerResult.AttributeLastName = rowValue.Values[NaturalPersons.COL_LASTNAME] as string;
                            _existingNaturals.Add(consumerResult);
                        }
                        else if (existingEntity != null)
                        {
                            consumerResult.AttributeCurrentAccount = rowValue.Values[NaturalPersons.COL_CURRENTACCOUNT] as string;
                            consumerResult.AttributeLastName = rowValue.Values[NaturalPersons.COL_LASTNAME] as string;
                        }

                        return consumerResult;
                    });

                    // Имя
                    ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, NaturalPersons.COL_FIRSTNAME] },
                        rowValue.Values[NaturalPersons.COL_FIRSTNAME] as string,
                        (value, parent, args) =>
                        {
                            if (existingEntity == null && !string.IsNullOrEmpty(naturalPerson.AttributeFirstName) && naturalPerson.AttributeFirstName != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                naturalPerson.AttributeFirstName = value;
                            return naturalPerson;
                        });

                    // Отчество
                    ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, NaturalPersons.COL_MIDDLENAME] },
                        rowValue.Values[NaturalPersons.COL_MIDDLENAME] as string,
                        (value, parent, args) =>
                        {
                            if (existingEntity == null && !string.IsNullOrEmpty(naturalPerson.AttributeMiddleName) && naturalPerson.AttributeMiddleName != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                naturalPerson.AttributeMiddleName = value;
                            return naturalPerson;
                        });

                    // Email
                    ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, NaturalPersons.COL_EMAIL] },
                        rowValue.Values[NaturalPersons.COL_EMAIL] as string,
                        (value, parent, args) =>
                        {
                            if (existingEntity == null && !string.IsNullOrEmpty(naturalPerson.AttributeEmail) && naturalPerson.AttributeEmail != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                naturalPerson.AttributeEmail = value;
                            return naturalPerson;
                        });

                    // Телефон
                    ImportHelpers.FindOrCreateEntity(result,
                        new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, NaturalPersons.COL_PHONE] },
                        rowValue.Values[NaturalPersons.COL_PHONE] as string,
                        (value, parent, args) =>
                        {
                            if (existingEntity == null && !string.IsNullOrEmpty(naturalPerson.AttributePhone) && naturalPerson.AttributePhone != value)
                                ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                            else
                                naturalPerson.AttributePhone = value;
                            return naturalPerson;
                        });
                    
                    return naturalPerson;
                }

                // Юр лица
                var possibleLeId = Helper.TryGetGuidIdentifier(rowValue.Values[Sheet.COL_NUM]);
                var existingLeEntity = possibleLeId.HasValue ? LegalEntity.Find(possibleLeId.Value) : null;

                var legalEntity = ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, LegalEntities.COL_CURRENTACCOUNT],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, LegalEntities.COL_CAPTION],
                    },
                    LegalEntityClassInfo.Get(),
                    (value, parent, args) =>
                    {
                        var wantedType = value.TypeOfInstance();
                        var consumerResult = existingLeEntity ?? _existingLegals
                            .Where(x => x.GetType() == wantedType)
                            .AsEnumerable()
                            .FirstOrDefault(x =>
                                Equals(x.AttributeCurrentAccount, rowValue.Values[LegalEntities.COL_CURRENTACCOUNT]) &&
                                Equals(x.AttributeLegalEntityCaption, rowValue.Values[LegalEntities.COL_CAPTION] ?? x.AttributeLegalEntityCaption));
                        if (consumerResult == null)
                        {
                            consumerResult = (LegalEntity)ImportHelpers.FindOrCreate.Entity(value);
                            consumerResult.AttributeCurrentAccount = rowValue.Values[LegalEntities.COL_CURRENTACCOUNT] as string;
                            consumerResult.AttributeLegalEntityCaption = rowValue.Values[LegalEntities.COL_CAPTION] as string;
                            _existingLegals.Add(consumerResult);
                        }
                        else if (existingLeEntity != null)
                        {
                            consumerResult.AttributeCurrentAccount = rowValue.Values[LegalEntities.COL_CURRENTACCOUNT] as string;
                            consumerResult.AttributeLegalEntityCaption = rowValue.Values[LegalEntities.COL_CAPTION] as string;
                        }

                        return consumerResult;
                    });

                // Адрес
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, LegalEntities.COL_ADDRESS] },
                    rowValue.Values[LegalEntities.COL_ADDRESS] as string,
                    (value, parent, args) =>
                    {
                        if (existingLeEntity == null && !string.IsNullOrEmpty(legalEntity.AttributeAddress) && legalEntity.AttributeAddress != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            legalEntity.AttributeAddress = value;
                        return legalEntity;
                    });

                // Email
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, LegalEntities.COL_EMAIL] },
                    rowValue.Values[LegalEntities.COL_EMAIL] as string,
                    (value, parent, args) =>
                    {
                        if (existingLeEntity == null && !string.IsNullOrEmpty(legalEntity.AttributeEmail) && legalEntity.AttributeEmail != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            legalEntity.AttributeEmail = value;
                        return legalEntity;
                    });

                // Телефон
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, LegalEntities.COL_PHONE] },
                    rowValue.Values[LegalEntities.COL_PHONE] as string,
                    (value, parent, args) =>
                    {
                        if (existingLeEntity == null && !string.IsNullOrEmpty(legalEntity.AttributePhone) && legalEntity.AttributePhone != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            legalEntity.AttributePhone = value;
                        return legalEntity;
                    });

                return legalEntity;
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
                if (throwIfException) throw;
            }
            finally
            {
                result.TotalCheckedLinesCount++;
                Helper.NotifyPercentRow();
            }
            return null;
        }

        // импортировать ТТ
        private static CurrentTransformer ImportTransCurrentRow(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue, bool throwIfException = false)
        {
            if (rowValue == null || rowValue.Sheet.GetType() != typeof(TransCurrents))
                return null;
            var possibleId = Helper.TryGetGuidIdentifier(rowValue.Values[Sheet.COL_NUM]);
            var existingEntity = possibleId.HasValue ? CurrentTransformer.Find(possibleId.Value) : null;

            try
            {
                var transCurrent = ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_TYPE],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_SERIALNUMBER],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_ACCURACYCLASS],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_PRIMARY_NOM],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_SECONDARY_NOM],
                    },
                    rowValue.Values[TransCurrents.COL_TYPE] as string,
                    (value, parent, args) =>
                    {
                        var serialNumber = rowValue.Values[TransCurrents.COL_SERIALNUMBER] as string;
                        var accuracyClass = rowValue.Values[TransCurrents.COL_ACCURACYCLASS] as AccuracyRatingEnumItem;
                        var primaryNom = rowValue.Values[TransCurrents.COL_PRIMARY_NOM] as CurrentEnumItem;
                        var secondaryNom = rowValue.Values[TransCurrents.COL_SECONDARY_NOM] as MeasureCurrentEnumItem;

                        var classInfo = ImportHelpers.FindOrCreate.TransformerClassInfo<CurrentTransformerClassInfo>(value, 
                            x =>
                                x.AttributeAccuracyClass == accuracyClass &&
                                x.AttributeHighSideNominalCurrent == primaryNom &&
                                x.AttributeLowSideNominalCurrent == secondaryNom, 
                            x =>
                            {
                                x.AttributeAccuracyClass = accuracyClass;
                                x.AttributeHighSideNominalCurrent_this_class = primaryNom;
                                x.AttributeLowSideNominalCurrent_this_class = secondaryNom;
                            });

                        var transResult = existingEntity ?? classInfo.GetInstances()
                            .AsEnumerable()
                            .FirstOrDefault(x => Equals(x.AttributeSerialNumber, serialNumber));
                        if (transResult == null)
                        {
                            transResult = (CurrentTransformer)ImportHelpers.FindOrCreate.Entity(classInfo);
                            transResult.AttributeSerialNumber = serialNumber;
                        }
                        else if (existingEntity != null)
                            transResult.AttributeSerialNumber = serialNumber;

                        return transResult;
                    });

                // Дата выпуска
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_CREATEDATE] },
                    rowValue.Values[TransCurrents.COL_CREATEDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && transCurrent.AttributeReleaseDate.HasValue && transCurrent.AttributeReleaseDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            transCurrent.AttributeReleaseDate = value;
                        return transCurrent;
                    });

                // Дата последней поверки
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransCurrents.COL_CHECKDATE] },
                    rowValue.Values[TransCurrents.COL_CHECKDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && transCurrent.AttributeLastCalibrationDate.HasValue && transCurrent.AttributeLastCalibrationDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            transCurrent.AttributeLastCalibrationDate = value;
                        return transCurrent;
                    });

                return transCurrent;
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
                if (throwIfException) throw;
            }
            finally
            {
                result.TotalCheckedLinesCount++;
                Helper.NotifyPercentRow();
            }

            return null;
        }

        // импортировать ТН
        private static VoltageTransformer ImportTransVoltageRow(ImportSheetProcessedResultData result, Sheet.RowValueInfo rowValue, bool throwIfException = false)
        {
            if (rowValue == null || rowValue.Sheet.GetType() != typeof(TransVoltages))
                return null;
            var possibleId = Helper.TryGetGuidIdentifier(rowValue.Values[Sheet.COL_NUM]);
            var existingEntity = possibleId.HasValue ? VoltageTransformer.Find(possibleId.Value) : null;
            try
            {
                var transVoltage = ImportHelpers.FindOrCreateEntity(result,
                    new[]
                    {
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_TYPE],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_SERIALNUMBER],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_ACCURACYCLASS],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_PRIMARY_NOM],
                        rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_SECONDARY_NOM]
                    },
                    rowValue.Values[TransVoltages.COL_TYPE] as string,
                    (value, parent, args) =>
                    {
                        var serialNumber = rowValue.Values[TransVoltages.COL_SERIALNUMBER] as string;
                        var accuracyClass = rowValue.Values[TransVoltages.COL_ACCURACYCLASS] as AccuracyRatingEnumItem;
                        var primaryNom = rowValue.Values[TransVoltages.COL_PRIMARY_NOM] as VoltageEnumItem;
                        var secondaryNom = rowValue.Values[TransVoltages.COL_SECONDARY_NOM] as MeasureVoltageEnumItem;

                        var classInfo = ImportHelpers.FindOrCreate.TransformerClassInfo<VoltageTransformerClassInfo>(value, 
                            x =>
                                x.AttributeAccuracyClass == accuracyClass &&
                                x.AttributeHighSideNominalVoltage == primaryNom &&
                                x.AttributeLowSideNominalVoltage == secondaryNom,
                            x =>
                            {
                                x.AttributeAccuracyClass = accuracyClass;
                                x.AttributeHighSideNominalVoltage_this_class = primaryNom;
                                x.AttributeLowSideNominalVoltage_this_class = secondaryNom;
                            });
                        
                        var transResult = existingEntity ?? classInfo.GetInstances()
                            .AsEnumerable()
                            .FirstOrDefault(x => Equals(x.AttributeSerialNumber, serialNumber));
                        if (transResult == null)
                        {
                            transResult = (VoltageTransformer)ImportHelpers.FindOrCreate.Entity(classInfo);
                            transResult.AttributeSerialNumber = serialNumber;
                        }
                        else if (existingEntity != null)
                            transResult.AttributeSerialNumber = serialNumber;

                        return transResult;
                    });

                // Дата выпуска
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_CREATEDATE] },
                    rowValue.Values[TransVoltages.COL_CREATEDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && transVoltage.AttributeReleaseDate.HasValue && transVoltage.AttributeReleaseDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            transVoltage.AttributeReleaseDate = value;
                        return transVoltage;
                    });

                // Дата последней поверки
                ImportHelpers.FindOrCreateEntity(result,
                    new[] { rowValue.Sheet.Worksheet.Cells[rowValue.RowIndex, TransVoltages.COL_CHECKDATE] },
                    rowValue.Values[TransVoltages.COL_CHECKDATE] as DateTime?,
                    (value, parent, args) =>
                    {
                        if (existingEntity == null && transVoltage.AttributeLastCalibrationDate.HasValue && transVoltage.AttributeLastCalibrationDate != value)
                            ImportHelpers.AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                        else
                            transVoltage.AttributeLastCalibrationDate = value;
                        return transVoltage;
                    });

                return transVoltage;
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ScriptImpl.Instance.AddLogInfo("Ошибка [Лист: {0}, строка: {1}]: {2}", rowValue.Sheet.Worksheet.Name, rowValue.RowIndex + 1, ex);
                if (throwIfException) throw;
            }
            finally
            {
                result.TotalCheckedLinesCount++;
                Helper.NotifyPercentRow();
            }

            return null;
        }
    }

    #endregion ImportHelper
}