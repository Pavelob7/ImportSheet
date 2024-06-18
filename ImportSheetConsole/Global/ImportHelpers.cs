using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using CommonTools;
using CSConstants;
using GemBox.Spreadsheet;
using ObjStudioClasses;
using ArgumentException = System.ArgumentException;
using RDAttributeValues = ObjStudioClasses.RDAttributeValues;
using RDCustomEntity = ObjStudioClasses.RDCustomEntity;
using RDInstance = ObjStudioClasses.RDInstance;

namespace ImportSheetConsole.Global
{
    #region ImportHelpers

    /// <summary>
    /// Утилиты импорта
    /// </summary>
    internal static class ImportHelpers
    {
        /// <summary>
        /// Функция поиска или добавления сущности
        /// </summary>
        public delegate TEntity FindOrCreateEntityFunc<out TEntity, in TValue>(TValue value, RDInstance parent, params object[] args)
            where TEntity : RDInstance;

        /// <summary>
        /// Справочник элементов пользователя
        /// </summary>
        public static readonly Lazy<RDAttributeValuesWithCreation<CommonUserDirectoryItem>> CommonDirectory =
            new Lazy<RDAttributeValuesWithCreation<CommonUserDirectoryItem>>(() =>
                DirectoryOfCommonUserItems.OnlyInstance.AttributeCommonUserItems);

        // текущий контекст безопасности, под которым выполняется ОЛ
        [ThreadStatic]
        private static PersonalizatedSource? _currentSecurityContextSource;

        // время начала импорта, для отслеживания
        [ThreadStatic]
        private static DateTime _importStartTime;

        /// <summary>
        /// Инициализировать стартовую информацию
        /// </summary>
        public static void Init()
        {
            _currentSecurityContextSource = RDClassesAndInstances.SecurityManager.CurrentSource as PersonalizatedSource;
            _importStartTime = DateTime.Now;
        }

        // результат создания сущности (инициализируется методом FindOrCreateEntity)
        [ThreadStatic]
        private static ImportSheetProcessedResultData _createEntityResult;

        // предупреждения, формируемые методами FindOrCreateEntityFunc
        [ThreadStatic]
        private static List<HashSet<string>> _warnLists;

        /// <summary>
        /// Предупреждения, формируемые методами FindOrCreateEntityFunc
        /// </summary>
        private static HashSet<string> WarnList
        {
            get { return _warnLists.LastOrDefault(); }
        }

        /// <summary>
        /// Добавить предупреждение
        /// </summary>
        /// <param name="text"></param>
        public static void AddWarn(string text)
        {
            if (WarnList != null)
                WarnList.Add(text);
        }

        /// <summary>
        /// Созданные сущности (создается извне, заполняется методом Entity и др)
        /// </summary>
        [ThreadStatic]
        public static List<RDInstance> CreatedEntities;

        /// <summary>
        /// Утилиты создания сущностей
        /// </summary>
        public static class FindOrCreate
        {
            /// <summary>
            /// Получить описание класса по типу экземпляра класса
            /// </summary>
            private static BaseClassClassInfo BaseClassInfoOfType(Type instanceType)
            {
                var attr = instanceType.GetCustomAttribute<RDMetadataSourceAttribute>();
                if (attr == null)
                    throw new ArgumentException(string.Format("Не удалось получить атрибут RDMetadataSourceAttribute для типа НСИ сущности {0}", instanceType.Name));
                return RDCustomEntity.Find(attr.SourceRefName) as BaseClassClassInfo;
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>(BaseClassClassInfo classInfo, Func<T, bool> searchEntityFunc)
                where T : RDInstance
            {
                if (classInfo == null)
                    throw new ArgumentException("Не задано описание класса элемента (classInfo == null)");

                T result = null;
                if (searchEntityFunc != null)
                {
                    result = classInfo.GetInstances()
                        .OfType<T>()
                        .AsEnumerable()
                        .FirstOrDefault(searchEntityFunc);
                }

                if (result == null)
                {
                    result = CommonDirectory.Value.AppendNew(classInfo.RefName).Value as T;
                    // изменение списка и статистики созданных сущностей
                    if (_createEntityResult != null)
                        _createEntityResult.ImportedEntitiesCount++;
                    if (CreatedEntities != null)
                        CreatedEntities.Add(result);
                }

                return result;
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static CommonUserDirectoryItem Entity(BaseClassClassInfo classInfo)
            {
                return Entity<RDInstance>(classInfo, null) as CommonUserDirectoryItem;
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>(Func<T, bool> searchEntityFunc)
                where T : RDInstance
            {
                return Entity(BaseClassInfoOfType(typeof(T)), searchEntityFunc);
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>()
                where T : RDInstance
            {
                return Entity(BaseClassInfoOfType(typeof(T))) as T;
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>(RDAttributeValuesWithCreation<T> source, BaseClassClassInfo classInfo, Func<T, bool> searchEntityFunc)
                where T : RDInstance
            {
                if (classInfo == null)
                    throw new ArgumentException("Не задано описание класса элемента (classInfo == null)");

                T result = null;
                if (searchEntityFunc != null)
                {
                    result = source.GetValues()
                        .AsEnumerable()
                        .FirstOrDefault(searchEntityFunc);
                }

                if (result == null)
                {
                    result = source.AppendNew(classInfo.RefName).Value;
                    // изменение списка и статистики созданных сущностей
                    if (_createEntityResult != null)
                        _createEntityResult.ImportedEntitiesCount++;
                    if (CreatedEntities != null)
                        CreatedEntities.Add(result);
                }

                return result;
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>(RDAttributeValuesWithCreation<T> source, BaseClassClassInfo classInfo)
                where T : RDInstance
            {
                if (classInfo == null)
                    throw new ArgumentException("Не задано описание класса элемента (classInfo == null)");
                var result = source.AppendNew(classInfo.RefName).Value;
                // изменение списка и статистики созданных сущностей
                if (_createEntityResult != null)
                    _createEntityResult.ImportedEntitiesCount++;
                if (CreatedEntities != null)
                    CreatedEntities.Add(result);
                return result;
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>(RDAttributeValuesWithCreation<T> source, Func<T, bool> searchEntityFunc)
                where T : RDInstance
            {
                return Entity(source, BaseClassInfoOfType(typeof(T)), searchEntityFunc);
            }

            /// <summary>
            /// Cоздать сущность
            /// </summary>
            /// <returns></returns>
            public static T Entity<T>(RDAttributeValuesWithCreation<T> source)
                where T : RDInstance
            {
                return Entity(source, BaseClassInfoOfType(typeof(T)));
            }

            /// <summary>
            /// Информация о блокировке
            /// </summary>
            private class EntityWithCaptionLockInfo
            {
                public int Counter;
            }

            /// <summary>
            /// Объекты синхронизации поиска по наименованию
            /// </summary>
            private static readonly Dictionary<(BaseClassClassInfo classInfo, string caption, RDInstance parent), EntityWithCaptionLockInfo> _entityWithCaptionLocks = new Dictionary<(BaseClassClassInfo classInfo, string caption, RDInstance parent), EntityWithCaptionLockInfo>();

            /// <summary>
            /// Получить объект синхронизации для поиска элемента по наименованию
            /// </summary>
            /// <param name="classInfo"></param>
            /// <param name="caption"></param>
            /// <param name="parent"></param>
            /// <returns></returns>
            private static void LockEntityWithCaption(BaseClassClassInfo classInfo, string caption, RDInstance parent)
            {
                EntityWithCaptionLockInfo result;
                lock (_entityWithCaptionLocks)
                {
                    if (!_entityWithCaptionLocks.TryGetValue((classInfo, caption, parent), out result))
                        _entityWithCaptionLocks[(classInfo, caption, parent)] = result = new EntityWithCaptionLockInfo();
                    result.Counter++;
                }
                Monitor.Enter(result);
            }

            /// <summary>
            /// Освободить объект синхронизации для поиска элемента по наименованию
            /// </summary>
            /// <param name="classInfo"></param>
            /// <param name="caption"></param>
            /// <param name="parent"></param>
            /// <returns></returns>
            private static void UnlockEntityWithCaption(BaseClassClassInfo classInfo, string caption, RDInstance parent)
            {
                lock (_entityWithCaptionLocks)
                {
                    if (!_entityWithCaptionLocks.TryGetValue((classInfo, caption, parent), out var result))
                        return;
                    result.Counter--;
                    if (result.Counter <= 0)
                        _entityWithCaptionLocks.Remove((classInfo, caption, parent));
                    Monitor.Exit(result);
                }
            }

            /// <summary>
            /// Найти или создать сущность по наименованию
            /// </summary>
            /// <returns></returns>
            public static T EntityWithCaption<T>(BaseClassClassInfo classInfo, string caption, RDInstance parent,
                Func<T, bool> searchEntityFunc)
                where T : RDInstance
            {
                if (classInfo == null)
                    throw new ArgumentException("Не задано описание класса элемента (classInfo == null)");
                if (caption == null)
                    throw new ArgumentException("Не задано наименование элемента (caption == null)");
                LockEntityWithCaption(classInfo, caption, parent);
                try
                {
                    var parentNode = parent as CustomClassifierNode;
                    var result =
                        (
                            // поиск в родителе, если он есть
                            (parentNode == null ? null : parentNode.GetLowerItems()) ??
                            // иначе все элементы, указанного типа
                            classInfo.GetInstances().OfType<RDInstance>()
                        )
                        .Where(x => x.Class == classInfo)
                        .OfType<IEntityWithCaption>()
                        .AsEnumerable()
                        .FirstOrDefault(x => searchEntityFunc != null
                            ? searchEntityFunc((T)x)
                            : x.AttributeCaption == caption) as T;
                    if (result == null)
                    {
                        result = Entity(classInfo) as T;
                        if (result != null)
                            ((IEntityWithCaption)result).AttributeCaption = caption;
                    }

                    return result;
                }
                finally
                {
                    UnlockEntityWithCaption(classInfo, caption, parent);
                }
            }

            /// <summary>
            /// Найти или создать сущность по наименованию
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <returns></returns>
            public static T EntityWithCaption<T>(string caption, RDInstance parent, Func<RDInstance, bool> searchEntityFunc)
                where T : CommonUserDirectoryItem
            {
                return EntityWithCaption(BaseClassInfoOfType(typeof(T)), caption, parent, searchEntityFunc) as T;
            }

            /// <summary>
            /// Найти или создать тип трансформатора
            /// </summary>
            public static T TransformerClassInfo<T>(string typeCaption, Func<T, bool> searchFunc, Action<T> setFunc)
                where T : MeasuringTransformerClassInfo
            {
                var baseClassInfo = BaseClassInfoOfType(typeof(T));
                var classInfo = baseClassInfo
                    .GetAllDescendants()
                    .OfType<T>()
                    .FirstOrDefault(x => x.Caption == typeCaption && searchFunc(x));

                if (classInfo == null)
                {
                    classInfo = (T)baseClassInfo.CreateDescendant();
                    classInfo.Caption = typeCaption;
                    setFunc(classInfo);
                }
                return (T)classInfo;
            }

            /// <summary>
            /// Установить линию в строение в присоединение
            /// </summary>
            /// <returns></returns>
            private static PowerLineConnection SetBuildingPowerLineConnection(CustomBuilding source, PowerLine powerLine)
            {
                if (source == null || powerLine == null)
                    return null;

                PowerLineConnection result = null;

                foreach (var powerLineConnection in source.AttributePowerLineConnections.GetValues())
                {
                    // 1. есть присоединение без линии, но совпадающее по наименованию (вероятно линия лежала в нем ранее!?)
                    if (powerLineConnection.AttributePowerLine == null && powerLineConnection.AttributeCaption == powerLine.AttributeCaption)
                    {
                        result = powerLineConnection;
                        break;
                    }

                    if (powerLineConnection.AttributePowerLine == null)
                        continue;

                    // 2. или присоединение содержит саму линию
                    if (powerLineConnection.AttributePowerLine == powerLine)
                    {
                        result = powerLineConnection;
                        break;
                    }

                    // 3. или в присоединении есть линия без ТП, совпадающая по наименованию
                    if (!EnumPowerLineOwners(powerLineConnection.AttributePowerLine).OfType<CubiclePowerLine>().Any() &&
                        powerLineConnection.AttributePowerLine.AttributeCaption == powerLine.AttributeCaption)
                    {
                        result = powerLineConnection;
                        break;
                    }
                }

                // добавить присоединение с новой линией
                if (result == null)
                {
                    result = EntityWithCaption<PowerLineConnection>(powerLine.AttributeCaption, null, x => false);
                    source.AttributePowerLineConnections.Add(result);
                }

                // установка новой линии
                result.AttributePowerLine = powerLine;

                return result;
            }

            /// <summary>
            /// Получить рекурсивно владельцев линии
            /// </summary>
            private static IEnumerable<RDEntityWithAttributes> EnumPowerLineOwners(PowerLine source)
            {
                if (source == null)
                    yield break;

                foreach (var relation in source.GetRelations())
                {
                    var relationPowerLine = relation.Owner.Owner as PowerLine;
                    if (relationPowerLine == null)
                    {
                        yield return relation.Owner.Owner;
                        continue;
                    }

                    foreach (var owner in EnumPowerLineOwners(relationPowerLine) ?? new RDEntityWithAttributes[0])
                        yield return owner;

                }
            }

            /// <summary>
            /// Удалить линию, если она не имеет питающих ячеек
            /// </summary>
            private static void CheckAndRemoveOrphanedPowerLines(params PowerLine[] sources)
            {
                foreach (var source in sources ?? new PowerLine[0])
                {
                    if (source == null)
                        continue;
                    if (!EnumPowerLineOwners(source).OfType<CubiclePowerLine>().Any())
                        source.Remove();
                }
            }

            /// <summary>
            /// Установить линию
            /// </summary>
            public static void SetPowerLine(RDInstance source, PowerLine powerLine)
            {
                if (powerLine == null)
                    return;

                // реализует интерфейс IPowerLineConnection
                var powerLineConnection = source as IPowerLineConnection;
                // присоединение к линии/фидеру ТУ потребителей
                var powerLineConnectionWithConsumerMeterPoints = source as IPowerLineConnectionWithConsumerMeterPoints;
                // строение
                var building = source as CustomBuilding;

                // удалить линии без питающих ячеек
                if (powerLineConnection != null)
                    CheckAndRemoveOrphanedPowerLines(powerLineConnection.AttributePowerLine);

                if (powerLineConnection != null)
                {
                    if (powerLineConnection.AttributePowerLine != null &&
                        powerLineConnection.AttributePowerLine != powerLine)
                    {
                        if (building == null)
                            throw new ArgumentException(string.Format(
                                "Элементу {0} не удалось установить атрибут Линия, т.к. атрибут уже содержит другое значение. " +
                                "Элемент не был импоритрован.",
                                source));

                        // для строения, необходимо создать присоединения, в случае нескольких линий

                        // 1. присоединение с линией от интерфейса IPowerLineConnection
                        if (!building.AttributePowerLineConnections.GetValues().Select(x => x.AttributePowerLine).Contains(powerLineConnection.AttributePowerLine))
                        {
                            var connection = SetBuildingPowerLineConnection(building, powerLineConnection.AttributePowerLine);
                            // перекинуть ТУ в присоединение
                            if (powerLineConnectionWithConsumerMeterPoints != null)
                            {
                                foreach (var meterPoint in powerLineConnectionWithConsumerMeterPoints.AttributeConsumerMeterPoints.GetValues())
                                    connection.AttributeConsumerMeterPoints.Add(meterPoint);
                                powerLineConnectionWithConsumerMeterPoints.AttributeConsumerMeterPoints.Clear();
                            }
                        }

                        // удалить линию интерфейса IPowerLineConnection
                        powerLineConnection.AttributePowerLine = null;

                        // 2. присоединение с линией powerLine (будет создано ниже)
                    }

                    // для строения необходимо проверить наличие присоединений
                    if (powerLineConnection.AttributePowerLine == null && building != null)
                    {
                        if (building.AttributePowerLineConnections.GetValues().Any(x => x.AttributePowerLine != null))
                        {
                            SetBuildingPowerLineConnection(building, powerLine);
                            return;
                        }
                    }

                    powerLineConnection.AttributePowerLine = powerLine;
                    return;
                }

                // строение
                if (building != null)
                {
                    SetBuildingPowerLineConnection(building, powerLine);
                }
            }

            /// <summary>
            /// Установить показания на момент снятия/установки ПУ
            /// </summary>
            /// <param name="source"></param>
            /// <param name="integratedData">Значение в кВт</param>
            private static void SetIntegratedData(RDAttributeValuesWithCreation<MeterIntegratedDataInfo> source, (Parameter parameter, double value)[] integratedData)
            {
                source.Clear();
                if (integratedData == null)
                    return;

                foreach (var (parameter, integratedValue) in integratedData)
                {
                    var value = source.AppendNew().Value;
                    value.AttributeParameter = parameter;
                    // перевод в Вт*ч
                    value.AttributeValue = integratedValue * 1000;
                }
            }

            /// <summary>
            /// Снять текущий ПУ в ТУ
            /// </summary>
            public static void UninstallMeterFromMeterPoint(MeterPoint source, DateTime uninstallDate, EquipmentReplacementReason reason, string comment, (Parameter parameter, double value)[] integratedData)
            {
                if (source == null)
                    return;
                
                var meter = source.AttributeElectricityMeter;
                if (meter == null)
                    return;

                var installedMeterInfo = meter.RelationsInstalledMeterInfoAttributeMeter.FirstOrDefault(x => x.AttributeUninstallInfo == null);
                if (installedMeterInfo == null)
                    return;

                var unisntallInfo = installedMeterInfo.AttributeUninstallInfoNullSafe();
                unisntallInfo.AttributeExecutionDate = uninstallDate;
                unisntallInfo.AttributeReplacementReason = reason;
                unisntallInfo.AttributeComment = comment;

                SetIntegratedData(unisntallInfo.AttributeIntegratedData, integratedData);
            }

            /// <summary>
            /// Установить ПУ в ТУ
            /// </summary>
            public static void SetMeterToMeterPoint(MeterPoint source, ElectricityMeter meter, DateTime? installDate)
            {
                if (source == null || meter == null || !installDate.HasValue)
                    return;
                var meterPointMeter = source.AttributeElectricityMeter;
                if (meterPointMeter != null && meterPointMeter != meter)
                {
                    AddWarn(string.Format(
                        "ПУ {0} не был сопоставлен ТУ {1}, " +
                        "т.к. ТУ уже содержит другой ПУ {2}",
                        meter, source, meterPointMeter));
                }
                else
                {
                    var link = source.AttributeMeterPointToMeterLinkSettingsNullSafe();
                    var installedInfos = link.AttributeInstalledMetersInfo.GetValues().ToArrayOptimized();
                    if (installedInfos
                        .Any(x => x.AttributeInstallInfo != null &&
                            x.AttributeInstallInfo.AttributeExecutionDate >= installDate &&
                            x.AttributeMeter != meter))
                    {
                        AddWarn(string.Format(
                            "ПУ {0} не был сопоставлен ТУ {1}, " +
                            "т.к. ТУ содержит другой ПУ с датой установки больше или равной {2}",
                            meter, source, installDate));
                    }
                    else
                    {
                        var prevInstalledInfo = installedInfos
                            .OrderBy(x => x.AttributeInstallInfo == null ? null : x.AttributeInstallInfo.AttributeExecutionDate)
                            .LastOrDefault();
                        // установить дату снятия, предыдущему ПУ
                        if (prevInstalledInfo != null && prevInstalledInfo.AttributeMeter != meter &&
                            prevInstalledInfo.AttributeUninstallInfo != null &&
                            prevInstalledInfo.AttributeUninstallInfo.AttributeExecutionDate == null)
                        {
                            var unistallInfo = prevInstalledInfo.AttributeUninstallInfoNullSafe();
                            unistallInfo.AttributeExecutionDate = installDate;
                        }

                        // создать или обновить, существующую запись
                        var installInfo = prevInstalledInfo != null && prevInstalledInfo.AttributeMeter == meter
                            ? prevInstalledInfo.AttributeInstallInfoNullSafe()
                            : Entity(link.AttributeInstalledMetersInfo).AttributeInstallInfoNullSafe();
                        var asInstalledMeterInfo = installInfo.OwnerInstalledMeterCommonInfoAttributeInstallInfo as InstalledMeterInfo;
                        if (asInstalledMeterInfo != null)
                            asInstalledMeterInfo.AttributeMeter = meter;
                        installInfo.AttributeExecutionDate = installDate;

                        // снять ПУ в других ТУ
                        foreach (var otherMeterPointInstalledMeterInfo in meter.RelationsInstalledMeterInfoAttributeMeter
                            .Where(x => x != asInstalledMeterInfo && x.AttributeUninstallInfo == null))
                        {
                            var unInstallInfo = otherMeterPointInstalledMeterInfo.AttributeUninstallInfoNullSafe();
                            // дата снятия = дате установки в новой ТУ или дате установки в старой ТУ, если в новой дата меньше
                            var dt = asInstalledMeterInfo?.AttributeInstallInfo?.AttributeExecutionDate ?? DateTime.Now;
                            var otherInstallDate = otherMeterPointInstalledMeterInfo.AttributeInstallInfo?.AttributeExecutionDate;
                            unInstallInfo.AttributeExecutionDate = dt < otherInstallDate ? otherInstallDate : dt;
                            unInstallInfo.AttributeReplacementReason = EquipmentReplacementReason.Instances.Other;
                            if (asInstalledMeterInfo != null)
                            {
                                unInstallInfo.AttributeComment = string.Format("ПУ снят в результате установки в ТУ {0} опросным листом",
                                    asInstalledMeterInfo.OwnerMeterPointToMeterLinkSettingsAttributeInstalledMetersInfo?.OwnerMeterPointAttributeMeterPointToMeterLinkSettings);
                            }
                        }
                    }
                }
            }

            /// <summary>
            /// Установить ПУ в ТУ
            /// </summary>
            public static void SetMeterToMeterPoint(MeterPoint source, ElectricityMeter meter, DateTime? installDate, 
                bool? coefInMeter, MeterJoinSchemesItem joinScheme, bool? directionEnverted, string comment, 
                (Parameter parameter, double value)[] integratedData)
            {
                if (source == null || meter == null || !installDate.HasValue)
                    return;
                
                var link = source.AttributeMeterPointToMeterLinkSettingsNullSafe();
                var installedInfos = link.AttributeInstalledMetersInfo.GetValues().ToArrayOptimized();
                if (installedInfos
                    .Any(x => x.AttributeInstallInfo != null &&
                        x.AttributeInstallInfo.AttributeExecutionDate >= installDate &&
                        x.AttributeMeter != meter))
                {
                    AddWarn(string.Format(
                        "ПУ {0} не был сопоставлен ТУ {1}, " +
                        "т.к. ТУ содержит другой ПУ с датой установки больше или равной {2}",
                        meter, source, installDate));
                }
                else
                {
                    var prevInstalledInfo = installedInfos
                        .OrderBy(x => x.AttributeInstallInfo == null ? null : x.AttributeInstallInfo.AttributeExecutionDate)
                        .LastOrDefault();
                    // установить дату снятия, предыдущему ПУ
                    if (prevInstalledInfo != null && prevInstalledInfo.AttributeMeter != meter &&
                        prevInstalledInfo.AttributeUninstallInfo != null &&
                        prevInstalledInfo.AttributeUninstallInfo.AttributeExecutionDate == null)
                    {
                        var unistallInfo = prevInstalledInfo.AttributeUninstallInfoNullSafe();
                        unistallInfo.AttributeExecutionDate = installDate;
                    }

                    // создать или обновить, существующую запись
                    var installInfo = prevInstalledInfo != null && prevInstalledInfo.AttributeMeter == meter
                        ? prevInstalledInfo.AttributeInstallInfoNullSafe()
                        : Entity(link.AttributeInstalledMetersInfo).AttributeInstallInfoNullSafe();
                    if (installInfo.OwnerInstalledMeterCommonInfoAttributeInstallInfo is InstalledMeterInfo asInstalledMeterInfo)
                    {
                        asInstalledMeterInfo.AttributeMeter = meter;
                        asInstalledMeterInfo.AttributeMeasureTransformerScalesAreInMeter = coefInMeter;
                        asInstalledMeterInfo.AttributeMeterJoinScheme = joinScheme;
                        asInstalledMeterInfo.AttributeDirectionInverted = directionEnverted;
                    }
                    installInfo.AttributeExecutionDate = installDate;
                    installInfo.AttributeComment = comment;
                    SetIntegratedData(installInfo.AttributeIntegratedData, integratedData);
                }
            }

            /// <summary>
            /// Установить упрощенные коэф трансформации в ТУ
            /// </summary>
            public static void SetReductiveTransformerToMeterPoint(MeterPoint source, double? currentCoef, double? voltageCoef)
            {
                if (source == null)
                    return;

                // удалить упрощенные настройки ТУ в случае их отсутствия
                if (!currentCoef.HasValue && !voltageCoef.HasValue)
                {
                    var tmpLink = source.AttributeMeterPointToMeterLinkSettings;
                    if (tmpLink != null)
                        tmpLink.AttributeTransformersReductive = null;
                    return;
                }

                // установить
                var link = source.AttributeMeterPointToMeterLinkSettingsNullSafe();
                var info = link.AttributeTransformersReductive;
                if (!(link.AttributeTransformersReductive is MeterPointTransformersInfoReductive))
                {
                    link.AttributeTransformersReductive = null;
                    info = link.AttributeTransformersReductiveNullSafe(MeterPointTransformersInfoReductive.GetClassInfo().RefName);
                }

                if (currentCoef < 1) currentCoef = null;
                if (voltageCoef < 1) voltageCoef = null;
                ((MeterPointTransformersInfoReductive)info).AttributeCurrentTurnRatio = currentCoef ?? 1;
                ((MeterPointTransformersInfoReductive)info).AttributeVoltageTurnRatio = voltageCoef ?? 1;
            }

            /// <summary>
            /// Установить ТТ в ТУ
            /// </summary>
            public static void SetCurrentTransformerToMeterPoint(MeterPoint source, CurrentTransformer[] transformers, DateTime? installDate)
            {
                if (source == null || transformers == null || !transformers.Any() || transformers.All(x => x == null) || !installDate.HasValue)
                    return;
                var link = source.AttributeMeterPointToMeterLinkSettingsNullSafe();
                var installedInfos = link.AttributeMeasureCurrentTransformersInfo.GetValues().ToArrayOptimized();
                if (installedInfos.Any(x =>
                    x.AttributeInstallDt >= installDate &&
                    (!transformers.SequenceEqual(new[]
                    {
                        x.AttributeCurrentTransformerOnPhase1,
                        x.AttributeCurrentTransformerOnPhase2,
                        x.AttributeCurrentTransformerOnPhase3
                    })
                    || !(new[]
                    {
                        x.AttributeCurrentTransformerOnPhase1,
                        x.AttributeCurrentTransformerOnPhase2,
                        x.AttributeCurrentTransformerOnPhase3
                    }.All(z => z == null)))
                    ))
                {
                    AddWarn(string.Format(
                        "Трансформаторы тока не были сопоставлены ТУ {0}, " +
                        "т.к. ТУ содержит другие ТТ с датой установки больше или равной {1}",
                        source, installDate));
                }
                else
                {
                    var prevInstalledInfo = installedInfos
                        .OrderBy(x => x.AttributeInstallDt)
                        .LastOrDefault();
                    // создать или обновить, существующую запись
                    var installInfo = prevInstalledInfo != null && (transformers.SequenceEqual(new[]
                    {
                        prevInstalledInfo.AttributeCurrentTransformerOnPhase1,
                        prevInstalledInfo.AttributeCurrentTransformerOnPhase2,
                        prevInstalledInfo.AttributeCurrentTransformerOnPhase3
                    })
                    || new[]
                    {
                        prevInstalledInfo.AttributeCurrentTransformerOnPhase1,
                        prevInstalledInfo.AttributeCurrentTransformerOnPhase2,
                        prevInstalledInfo.AttributeCurrentTransformerOnPhase3
                    }.All(z => z == null)
                    )

                        ? prevInstalledInfo
                        : Entity(link.AttributeMeasureCurrentTransformersInfo);
                    installInfo.AttributeCurrentTransformerOnPhase1 = transformers.FirstOrDefault();
                    installInfo.AttributeCurrentTransformerOnPhase2 = transformers.Skip(1).FirstOrDefault();
                    installInfo.AttributeCurrentTransformerOnPhase3 = transformers.Skip(2).FirstOrDefault();
                    installInfo.AttributeInstallDt = installDate;
                }
            }

            /// <summary>
            /// Установить ТН в ячейку ТН
            /// </summary>
            public static void SetVoltageTransformerToCubicle(CubicleMeasureVoltageTransformer source, VoltageTransformer[] transformers, DateTime? installDate)
            {
                if (source == null || transformers == null || !transformers.Any() || !installDate.HasValue)
                    return;
                var installedInfos = source.AttributeMeasureVoltageTransformersInfo.GetValues().ToArrayOptimized();
                if (installedInfos.Any(x =>
                    x.AttributeInstallDt >= installDate &&
                    !transformers.SequenceEqual(new[]
                    {
                        x.AttributeVoltageTransformerOnPhase1,
                        x.AttributeVoltageTransformerOnPhase2,
                        x.AttributeVoltageTransformerOnPhase3
                    })))
                {
                    AddWarn(string.Format(
                        "Трансформаторы напряжения не были сопоставлены ячейке ТН {0}, " +
                        "т.к. ячейка ТН содержит другие ТН с датой установки больше или равной {1}",
                        source, installDate));
                }
                else
                {
                    var prevInstalledInfo = installedInfos
                        .OrderBy(x => x.AttributeInstallDt)
                        .LastOrDefault();
                    // создать или обновить, существующую запись
                    var installInfo = prevInstalledInfo != null && transformers.SequenceEqual(new[]
                    {
                        prevInstalledInfo.AttributeVoltageTransformerOnPhase1,
                        prevInstalledInfo.AttributeVoltageTransformerOnPhase2,
                        prevInstalledInfo.AttributeVoltageTransformerOnPhase3
                    })
                        ? prevInstalledInfo
                        : Entity(source.AttributeMeasureVoltageTransformersInfo);
                    installInfo.AttributeVoltageTransformerOnPhase1 = transformers.FirstOrDefault();
                    installInfo.AttributeVoltageTransformerOnPhase2 = transformers.Skip(1).FirstOrDefault();
                    installInfo.AttributeVoltageTransformerOnPhase3 = transformers.Skip(2).FirstOrDefault();
                    installInfo.AttributeInstallDt = installDate;
                }
            }

            /// <summary>
            /// Установить сетевой номер
            /// </summary>
            public static void SetNetworkNumber(MeterEquipment meter, string value)
            {
                // костыльная установка атрибута "Номер модема" для ПУ Фобос
                var fobosMeter = meter as FobosMeter;
                if (fobosMeter != null)
                {
                    if (!string.IsNullOrEmpty(fobosMeter.AttributeModemId) && fobosMeter.AttributeModemId != value)
                        AddWarn("Атрибут \"Номер модема\" не был установлен, т.к. уже содержит другое значение");
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
                    return;

                if (!string.IsNullOrEmpty(newEquipment.AttributeNetworkId) && newEquipment.AttributeNetworkId != value)
                    AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                else
                    newEquipment.AttributeNetworkId = value;
            }

            /// <summary>
            /// Устанвоить логин/пользователя для ПУ
            /// </summary>
            public static void SetMeterEquipmentUser(MeterEquipment meter, string value)
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

                var userAndPasswordEquipment = meter as IEquipmentWithUserAndPassword;
                if (userAndPasswordEquipment == null)
                    return;

                if (!string.IsNullOrEmpty(userAndPasswordEquipment.AttributeUser) && userAndPasswordEquipment.AttributeUser != value)
                    AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                else
                    userAndPasswordEquipment.AttributeUser = value;
            }

            /// <summary>
            /// Установить логин/пользователя для УСПД
            /// </summary>
            public static void SetChannelizingEquipmentUser(ChannelizingEquipment rtu, string value)
            {
                var stationNB300 = rtu as BaseStationNB300;

                // Аметов 13.01.21 так как BaseStationNB300 не поддерживает интерфейсы IEquipmentWithUserAndPassword IEquipmentWithPassword
                if (stationNB300 != null)
                {
                    if (!string.IsNullOrEmpty(stationNB300.AttributeLogin) && stationNB300.AttributeLogin != value)
                        AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        stationNB300.AttributeLogin = value;
                }

                var userAndPasswordEquipment = rtu as IEquipmentWithUserAndPassword;
                if (userAndPasswordEquipment == null)
                    return;
                
                if (!string.IsNullOrEmpty(userAndPasswordEquipment.AttributeUser) && userAndPasswordEquipment.AttributeUser != value)
                    AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                else
                    userAndPasswordEquipment.AttributeUser = value;

                // Аметов 13.01.21 Если введен логин и пароль, ставим галочку использовать авторизацию
                if (stationNB300 != null)
                {
                    if (!string.IsNullOrEmpty(stationNB300.AttributePassword) && !string.IsNullOrEmpty(stationNB300.AttributeLogin))
                        stationNB300.AttributeUseAuth = true;
                }
            }

            /// <summary>
            /// Установить пароль для УСПД
            /// </summary>
            public static void SetChannelizingEquipmentPassword(ChannelizingEquipment rtu, string value)
            {
                var stationNB300 = rtu as BaseStationNB300;

                // Аметов 13.01.21 Так как BaseStationNB300 не поддерживает интерфейсы IEquipmentWithUserAndPassword IEquipmentWithPassword
                if (stationNB300 != null)
                {
                    if (!string.IsNullOrEmpty(stationNB300.AttributePassword) && stationNB300.AttributePassword != value)
                        AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                    else
                        stationNB300.AttributePassword = value;
                }

                var passwordEquipment = rtu as IEquipmentWithPassword;
                if (passwordEquipment == null)
                    return;
                
                if (!string.IsNullOrEmpty(passwordEquipment.AttributePassword) && passwordEquipment.AttributePassword != value)
                    AddWarn("Атрибут не был установлен, т.к. уже содержит другое значение");
                else
                    passwordEquipment.AttributePassword = value;

                // Аметов 13.01.21 Если введен логин и пароль, ставим галочку использовать авторизацию
                if (stationNB300 != null)
                {
                    if (!string.IsNullOrEmpty(stationNB300.AttributePassword) && !string.IsNullOrEmpty(stationNB300.AttributeLogin))
                        stationNB300.AttributeUseAuth = true;
                }
            }

            /// <summary>
            /// Установить прямые маршруты
            /// </summary>
            public static void SetDirectRoutes(IDirectlyPollingEquipment equipment, ImportRoute[] routes)
            {
                // прямые маршруты
                foreach (var route in routes ?? new ImportRoute[0])
                {
                    if (route.ClassInfo.InheritsFrom(NonDirectRouteClassInfo.Get()))
                        continue;
                    // не добавляется, если маршруты равны
                    // приоритеты выставляются при добавлении триггером, если не задано явно в ОЛ 
                    var route1 = route;
                    var rdRoute = equipment.AttributeRoutes.GetValues().FirstOrDefault(x => route1.Equals(x));
                    if (rdRoute == null)
                        rdRoute = Entity(equipment.AttributeRoutes, route.ClassInfo);
                    // связывание специфичных параметров
                    route.Assign(rdRoute);
                }
            }

            /// <summary>
            /// Установить маршруты через промежуточное оборудование
            /// </summary>
            public static void SetNonDirectRoutes(IDirectlyPollingEquipment equipment, ImportRoute[] routes, Func<object, Equipment> getViaRtu, Action<Equipment, ImportRoute> postSetRouteAction)
            {
                foreach (var route in routes)
                {
                    object routeEntityLink;
                    if (!route.Params.TryGetValue(RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, out routeEntityLink) ||
                        !(routeEntityLink is EntityLink))
                        continue;

                    var viaRtu = getViaRtu?.Invoke(routeEntityLink);
                    if (viaRtu == null)
                        continue;

                    // установить экземпляр УСПД, вместо описания ссылки
                    route.Params[RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment] = viaRtu;
                    // не добавляется, если маршруты равны
                    var route1 = route;
                    var rdRoute = equipment.AttributeRoutes.GetValues().FirstOrDefault(x => route1.Equals(x));
                    if (rdRoute == null)
                        rdRoute = Entity(equipment.AttributeRoutes, route.ClassInfo);
                    // связывание специфичных параметров
                    route.Assign(rdRoute);

                    postSetRouteAction?.Invoke(viaRtu, route);
                }
            }

            /// <summary>
            /// Утилиты создания классифицируемых сущностей
            /// </summary>
            public static class ClassifierNodes
            {
                /// <summary>
                /// Добавить сущность по связям в классифицируемые атрибуты
                /// </summary>
                public static void AddToClassifierAttributes(CustomClassifierNode source, params object[] args)
                {
                    if (source == null)
                        return;
                    var f = false;
                    var classifierNodes = args.OfType<CustomClassifierNode>().ToArray();
                    foreach (var instance in classifierNodes)
                    {
                        foreach (var attribute in instance.Class.GetClassifierAttributesFast())
                        {
                            if (!attribute.Options.RequiredTypeRefName.HasValue)
                                continue;
                            var cl = BaseClassClassInfo.Find(attribute.Options.RequiredTypeRefName.Value);
                            RDAttributeValues values;
                            if (source.Class.InheritsFrom(cl) &&
                                (values = instance.GetAttributeValues(attribute) as RDAttributeValues) != null)
                            {
                                // массив
                                if (attribute.Options.MaxOccurs > 1)
                                {
                                    if (!values.GetValues().Select(x => x.Value).Contains(source))
                                        values.Add(source);
                                    f = true;
                                }
                                // не массив
                                else
                                {
                                    var v = values.GetValues().ToArrayOptimized();
                                    if (v.Any(x => ((RDCustomEntity)x.Value).RefName != source.RefName))
                                    {
                                        AddWarn(string.Format(
                                            "Элемент \"{0}\" не был сопоставлен атрибуту \"{1}.{2}\", " +
                                            "т.к. атрибут уже содержит другое значение",
                                            Tuple.Create(source.Class, source.AttributeCaption), instance, attribute));
                                    }
                                    else if (!v.Any())
                                    {
                                        values.Add(source);
                                        f = true;
                                    }
                                    else
                                        f = true;
                                }
                            }
                        }
                    }

                    if (!f)
                        throw new ArgumentException(string.Format(
                            "Элемент \"{0}\" не был сопоставлен ни одному из родителей (всего: {1}): \"{2}\".\n" +
                            "Отсутствуют атрибуты родителей, которые могут содержать указанный элемент",
                            Tuple.Create(source.Class, source.AttributeCaption),
                            classifierNodes.Length,
                            string.Join("; ", classifierNodes.Select(x => x.Caption))));
                }

                /// <summary>
                /// Добавить сущность с базовым классов "Обобщенный элемент классификатора (CustomClassifierNode)"
                /// </summary>
                public static CustomClassifierNode OfType<TEntity>(BaseClassClassInfo classInfo, string value, RDInstance parent,
                    Func<TEntity, bool> searchEntityFunc, params object[] args)
                    where TEntity : RDInstance
                {
                    LockEntityWithCaption(classInfo, value, parent);
                    try
                    {
                        var result = EntityWithCaption(classInfo, value, parent, searchEntityFunc) as CustomClassifierNode;
                        if (!(result is Classifier))
                            AddToClassifierAttributes(result, args.Union(parent.Yield()).ToArray());
                        return result;
                    }
                    finally
                    {
                        UnlockEntityWithCaption(classInfo, value, parent);
                    }
                }

                /// <summary>
                /// Добавить сущность с базовым классов "Обобщенный элемент классификатора (CustomClassifierNode)"
                /// </summary>
                public static CustomClassifierNode OfType(BaseClassClassInfo classInfo, string value, RDInstance parent, params object[] args)
                {
                    return OfType<CustomClassifierNode>(classInfo, value, parent, null, args);
                }

                /// <summary>
                /// Добавить сущность с базовым классов "Обобщенный элемент классификатора (CustomClassifierNode)"
                /// </summary>
                public static TEntity OfType<TEntity>(string value, RDInstance parent, Func<TEntity, bool> searchEntityFunc, params object[] args)
                    where TEntity : CustomClassifierNode
                {
                    return (TEntity)OfType(BaseClassInfoOfType(typeof(TEntity)), value, parent, searchEntityFunc, args);
                }

                /// <summary>
                /// Добавить сущность с базовым классов "Обобщенный элемент классификатора (CustomClassifierNode)"
                /// </summary>
                public static TEntity OfType<TEntity>(string value, RDInstance parent, params object[] args)
                    where TEntity : CustomClassifierNode
                {
                    return (TEntity)OfType(BaseClassInfoOfType(typeof(TEntity)), value, parent, args);
                }

                /// <summary>
                /// Добавить сущность "Распределительное устройство (Switchgear)"
                /// </summary>
                public static Switchgear Switchgear(VoltageEnumItem value, RDInstance parent, params object[] args)
                {
                    var result = OfType<Switchgear>(string.Format("РУ-{0}", value), parent, args);
                    if (result != null)
                        result.AttributeVoltage = value;
                    return result;
                }

                /// <summary>
                /// Добавить сущность "Распределительная подстанция (SwitchGearSubstation)"
                /// args:
                /// - Уровень напряжения
                /// </summary>
                public static SwitchGearSubstation SwitchGearSubstation(string value, RDInstance parent, params object[] args)
                {
                    var result = OfType<SwitchGearSubstation>(value, parent, args);
                    if (result != null)
                        result.AttributeVoltage = args.OfType<VoltageEnumItem>().FirstOrDefault();
                    return result;
                }

                /// <summary>
                /// Добавить сущность "Ячейка оборудования ПС (Cubicle)"
                /// args:
                /// - Тип ячейки. По умолчанию CubiclePowerLineClassInfo
                /// - Линия/фидер
                /// - ТН [VoltageTransfomer[]]
                /// - Дата установки ТН [KeyValuePair("VoltageTransformerInstallDate"=DateTime?)]
                /// </summary>
                public static Cubicle Cubicle(string value, RDInstance parent, params object[] args)
                {
                    var classInfo = args.OfType<CubicleClassInfo>().FirstOrDefault() ?? CubiclePowerLineClassInfo.Get();
                    var result = OfType(classInfo, value, parent, args) as Cubicle;

                    // установить линию
                    SetPowerLine(result, args.OfType<PowerLine>().FirstOrDefault());

                    // ТН
                    var transCubicle = result as CubicleMeasureVoltageTransformer;
                    if (transCubicle != null)
                    {
                        var transVoltages = args.OfType<VoltageTransformer[]>().FirstOrDefault();
                        var transVoltageInstallDate = args.OfType<KeyValuePair<string, DateTime?>>().FirstOrDefault(x => x.Key == "VoltageTransformerInstallDate").Value;
                        SetVoltageTransformerToCubicle(transCubicle, transVoltages, transVoltageInstallDate);
                    }

                    return result;
                }

                /// <summary>
                /// Добавить сущность "Сооружение (обобщенно) (CustomBuilding)"
                /// args:
                /// - Тип сооружения. По умолчанию DwellingHouseClassInfo
                /// - Линия/фидер
                /// </summary>
                public static CustomBuilding CustomBuilding(string value, RDInstance parent, params object[] args)
                {
                    var classInfo = args.OfType<CustomBuildingClassInfo>().FirstOrDefault() ?? DwellingHouseClassInfo.Get();
                    var result = OfType(classInfo, value, parent, args) as CustomBuilding;
                    // установить линию
                    SetPowerLine(result, args.OfType<PowerLine>().FirstOrDefault());
                    return result;
                }

                /// <summary>
                /// Добавить сущность "Помещение (Room)"
                /// args:
                /// </summary>
                public static Room Room(string value, RDInstance parent, params object[] args)
                {
                    var result = OfType(RoomClassInfo.Get(), value, parent, args) as Room;
                    return result;
                }

                /// <summary>
                /// Добавить сущность "ТУ (MeterPoint)"
                /// args:
                /// - ПУ
                /// - Абонент 
                /// - ТТ [CurrentTransfomer[]]
                /// - Дата установки ПУ [KeyValuePair("InstallDate"=DateTime?)]
                /// - Дата установки ТТ [KeyValuePair("CurrentTransformerInstallDate"=DateTime?)]
                /// - Коэф.трансф по току [KeyValuePair("CurrentCoef"=double)]
                /// - Коэф.трансф по напряжению [KeyValuePair("VoltageCoef"=double)]
                /// </summary>
                public static MeterPoint MeterPoint(string value, RDInstance parent, params object[] args)
                {
                    var meter = args.OfType<ElectricityMeter>().FirstOrDefault();
                    var consumer = args.OfType<Consumer>().FirstOrDefault();
                    var transCurrents = args.OfType<CurrentTransformer[]>().FirstOrDefault();
                    var installDate = args.OfType<KeyValuePair<string, DateTime?>>().FirstOrDefault(x => x.Key == "InstallDate").Value;
                    var transCurrentInstallDate = args.OfType<KeyValuePair<string, DateTime?>>().FirstOrDefault(x => x.Key == "CurrentTransformerInstallDate").Value;
                    var currentCoef = args.OfType<KeyValuePair<string, double?>>().FirstOrDefault(x => x.Key == "CurrentCoef").Value;
                    var voltageCoef = args.OfType<KeyValuePair<string, double?>>().FirstOrDefault(x => x.Key == "VoltageCoef").Value;

                    var result = OfType<MeterPoint>(value, parent, args);

                    // абонент
                    if (consumer != null)
                    {
                        if (!consumer.AttributeMeterPoints.GetValues().Contains(result))
                            consumer.AttributeMeterPoints.Add(result);
                    }

                    // ПУ
                    SetMeterToMeterPoint(result, meter, installDate ?? DateTime.Today);

                    // коэф.
                    SetReductiveTransformerToMeterPoint(result, currentCoef, voltageCoef);

                    // ТТ
                    SetCurrentTransformerToMeterPoint(result, transCurrents, transCurrentInstallDate ?? DateTime.Today);

                    return result;
                }

                /// <summary>
                /// Создать географический классификатор по адресу ФИАС
                /// Возвращает нижний созданный элемент
                /// args:
                /// - тип сооружения (CustomBuilding)
                /// - линия/фидер
                /// </summary>
                public static CustomClassifierNode AddressFias(FiasSuggestionData value, RDInstance parent, params object[] args)
                {
                    if (value == null)
                        return null;

                    CustomClassifierNode result = null;
                    var parentClassifierItem = parent;

                    // Субъект РФ
                    if (value.Region != null && parentClassifierItem != null)
                    {
                        result = OfType<SubjectRF>(value.Region, parentClassifierItem, args);
                        parentClassifierItem = result ?? parentClassifierItem;
                    }

                    // Район административного деления
                    if (value.Area != null && parentClassifierItem != null)
                    {
                        result = OfType<District>(value.Area, parentClassifierItem, args);
                        parentClassifierItem = result ?? parentClassifierItem;
                    }

                    // Населенный пункт + район населенного пункта

                    var centerOfPopulationValue = value.City;
                    var centerOfPopulationZoneValue = value.CityDistrict;
                    var centerOfPopulationZoneValue2 = value.Settlement;

                    // если отсутствует город, то используем населенный пункт
                    // иначе все что ниже города - районы населеднного пункта
                    if (string.IsNullOrEmpty(centerOfPopulationValue))
                    {
                        centerOfPopulationValue = value.Settlement;
                        centerOfPopulationZoneValue = null;
                        centerOfPopulationZoneValue2 = null;
                    }

                    if (centerOfPopulationValue != null && parentClassifierItem != null)
                    {
                        result = OfType<CenterOfPopulation>(centerOfPopulationValue, parentClassifierItem, args);
                        parentClassifierItem = result ?? parentClassifierItem;
                    }

                    // район населенного пункта 1
                    if (centerOfPopulationZoneValue != null && parentClassifierItem != null)
                    {
                        result = OfType<CenterOfPopulationZone>(centerOfPopulationZoneValue, parentClassifierItem, args);
                        parentClassifierItem = result ?? parentClassifierItem;
                    }

                    // район населенного пункта 2
                    if (centerOfPopulationZoneValue2 != null && parentClassifierItem != null)
                    {
                        result = OfType<CenterOfPopulationZone>(centerOfPopulationZoneValue2, parentClassifierItem, args);
                        parentClassifierItem = result ?? parentClassifierItem;
                    }

                    // улица
                    if (value.Street != null && parentClassifierItem != null)
                    {
                        result = OfType<Street>(value.Street, parentClassifierItem, args);
                        parentClassifierItem = result ?? parentClassifierItem;
                    }

                    // дом
                    if (value.House != null && parentClassifierItem != null)
                    {
                        result = CustomBuilding(value.House, parentClassifierItem, args);
                    }

                    return result ?? parentClassifierItem as CustomClassifierNode;
                }
            }

            /// <summary>
            /// Утилиты создания неклассифицируемых сущностей
            /// </summary>
            public static class NonClassifierNodes
            {   
            }
        }

        /// <summary>
        /// Добавить ОВ
        /// </summary>
        public static void AppendIsolationLevels(this NonClassifierItem source, IsolationLevel[] isolationLevels)
        {
            if (source == null || isolationLevels == null)
                return;

            // UPD ОВ заменяются указанными, для сущностей, которые были созданы в ходе работы ОЛ, иначе добавляются
            // !!! перед началом импорта необходим вызов ImportHelpers.Init();
            if (source.CreationSource == _currentSecurityContextSource && source.CreationStamp > _importStartTime)
            {
                source.AttributeIsolationLevels = isolationLevels;
            }
            else
            {
                var existingLevels = source.AttributeIsolationLevels.ToArray();
                var isolationLevelsToAppend = isolationLevels.Where(x => !existingLevels.Contains(x)).ToArray();
                if (isolationLevelsToAppend.Length > 0)
                    source.AttributeIsolationLevels = existingLevels.Concat(isolationLevelsToAppend);
            }
        }

        /// <summary>
        /// Получить или создать сущность, если ее еще нет
        /// </summary>
        public static TEntity FindOrCreateEntity<TCell, TValue, TEntity>(
            // общий результат метода импорта
            ImportSheetProcessedResultData result,
            // объект ячейки (ExcelCell или Range)
            TCell[] source,
            // значение
            TValue value,
            // родитель
            RDInstance parent,
            // функция добавления сущности
            FindOrCreateEntityFunc<TEntity, TValue> findOrCreateEntityFunc,
            // наличие родителя обязательно для элемента
            bool parentRequared = true,
            // аргументы
            params object[] args)
            where TEntity : RDInstance
        {
            if (value == null)
                return null;
            if (parentRequared && parent == null)
                return null;
            var kind = DiagnosticMessageKindData.Information;
            string text = null;

            _createEntityResult = result;
            (_warnLists ?? (_warnLists = new List<HashSet<string>>())).Add(new HashSet<string>());
            TEntity resultEntity;
            try
            {
                resultEntity = findOrCreateEntityFunc(value, parent, args);
                if (WarnList.Any())
                {
                    kind = DiagnosticMessageKindData.Warning;
                    text = string.Join("\n", WarnList);
                }
            }
            catch (System.Threading.ThreadInterruptedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                kind = DiagnosticMessageKindData.Error;
                text = ex.Message;
                throw;
            }
            finally
            {
                foreach (var cell in source.OfType<ExcelCell>())
                    cell.SetResult(kind, result, text);
                //foreach (var cell in source.OfType<Range>())
                //    cell.SetResult(kind, result, text);
                _warnLists.RemoveAt(_warnLists.Count - 1);
                _createEntityResult = null;
            }

            return resultEntity;
        }

        /// <summary>
        /// Получить или создать сущность, если ее еще нет
        /// </summary>
        public static TEntity FindOrCreateEntity<TCell, TValue, TEntity>(
            // общий результат метода импорта
            ImportSheetProcessedResultData result,
            // объект ячейки (ExcelCell или Range)
            TCell[] source,
            // значение
            TValue value,
            // функция добавления сущности
            FindOrCreateEntityFunc<TEntity, TValue> findOrCreateEntityFunc,
            // аргументы
            params object[] args)
            where TEntity : RDInstance
        {
            return FindOrCreateEntity(result, source, value, null, findOrCreateEntityFunc, false, args);
        }
    }

    #endregion ImportHelpers
}