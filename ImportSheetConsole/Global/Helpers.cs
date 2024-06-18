using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using CommonTools;
using ConsoleUsage.UserScripts.ImportSheets.Global;
using ObjStudioClasses;
using static ConsoleUsage.UserScripts.ImportSheets.Global.CSServiceTriggers;
using ClassAreaOfAction = ObjStudioClasses.ClassAreaOfAction;
using InheritanceLevelKind = ObjStudioClasses.InheritanceLevelKind;
using TimeZone = ObjStudioClasses.TimeZone;

namespace ImportSheetConsole.Global
{
    #region Helpers
    /// <summary>
    /// Общие утилиты
    /// </summary>
    internal static class Helpers
    {
        /// <summary>
        /// Конструктор
        /// </summary>
        static Helpers()
        {
            CSServiceTriggers.Implementation.RegisterExternalTrigger(new TriggerEquipmentClass(() =>
            {
                _meterClassInfos = null;
                _rtuClassInfos = null;
            }));
        }

        /// <summary>
        /// Разделитель нескольких значений в ячейке
        /// </summary>
        public static string[] MultiValueSeparators = { "/" };

        /// <summary>
        /// Подразделитель значений в ячейке, при наличие разделителей нескольких значений ячейки
        /// </summary>
        public static string[] MultiValueSubSeparators = { "|" };

        /// <summary>
        /// Разделитель массива значений специфичных атрибутов
        /// </summary>
        public static char ArraySeparator = '_';

        /// <summary>
        /// Заменитель разделителей в значении
        /// </summary>
        public static readonly Dictionary<string, char> SeparatorStubs = new Dictionary<string, char>
        {
            { "&#8260;", '/' },
            { "&brvbar;", '|' },
            { "&#8254;", '_' },
            { "&nbsp;", ' ' }
        };

        /// <summary>
        /// Псевдонимы атрибутов
        /// </summary>
        public static Dictionary<Guid, string> AttributesAliases = new Dictionary<Guid, string>
        {
            // Идентификатор класса
            { RDMetadataGuids.BaseClass.RefName, "Класс" },

            // Маршрут через УСПД с подустройствами. Номер прибора учета
            { RDMetadataGuids.RouteViaRtuWithSubdevices.InstanceAttributeDeviceNumber, "Номер ПУ" },

            // Маршрут через СИКОН TC65/SDM-TC65.Порт подключения (1 или 2) или идентификатор устройства (hex)
            { RDMetadataGuids.RouteViaSiconTC65.InstanceAttributeComPortNumber, "Порт" },
        };

        /// <summary>
        /// Псевдонимы ячеек
        /// </summary>
        public static Dictionary<string, string> CubicleAliases = new Dictionary<string, string>
        {
            { "Ячейка ввода силового трансформатора", "Ячейка высокой стороны силового трансформатора" },
        };
        
        /// <summary>
        /// Перечисление уровней напряжения
        /// </summary>
        private static readonly Lazy<VoltageEnumItem[]> _voltageEnumItems = new Lazy<VoltageEnumItem[]>(() =>
            VoltageEnumItem.GetInstances().ToArray());

        /// <summary>
        /// Перечисление классов точности
        /// </summary>
        private static readonly Lazy<AccuracyRatingEnumItem[]> _accuracyRatingEnumItems = new Lazy<AccuracyRatingEnumItem[]>(() =>
            AccuracyRatingEnumItem.GetInstances().ToArray());

        /// <summary>
        /// Перечисление схем включения
        /// </summary>
        private static readonly Lazy<MeterJoinSchemesItem[]> _meterJoinSchemesItems = new Lazy<MeterJoinSchemesItem[]>(() =>
            MeterJoinSchemesItem.GetInstances().ToArray());

        /// <summary>
        /// Перечисление причины замены оборудования
        /// </summary>
        private static readonly Lazy<EquipmentReplacementReason[]> _equipmentReplacementReasons = new Lazy<EquipmentReplacementReason[]>(() =>
            EquipmentReplacementReason.GetInstances().ToArray());

        /// <summary>
        /// Часовые пояса
        /// </summary>
        private static readonly Lazy<TimeZone[]> _timeZones = new Lazy<TimeZone[]>(() =>
            TimeZone.GetInstances().ToArray());
        
        /// <summary>
        /// Перечисление номинальных токов
        /// </summary>
        private static readonly Lazy<CurrentEnumItem[]> _nominalCurrentEnumItems = new Lazy<CurrentEnumItem[]>(() =>
            CurrentEnumItem.GetInstances().ToArray());

        /// <summary>
        /// Перечисление номинальных токов измерительных цепей
        /// </summary>
        private static readonly Lazy<MeasureCurrentEnumItem[]> _nominalMeasureCurrentEnumItems = new Lazy<MeasureCurrentEnumItem[]>(() =>
            MeasureCurrentEnumItem.GetInstances().ToArray());

        /// <summary>
        /// Перечисление номинальных напряжений
        /// </summary>
        private static readonly Lazy<VoltageEnumItem[]> _nominalVoltageEnumItems = new Lazy<VoltageEnumItem[]>(() =>
            VoltageEnumItem.GetInstances().ToArray());

        /// <summary>
        /// Перечисление номинальных напряжений измерительных цепей
        /// </summary>
        private static readonly Lazy<MeasureVoltageEnumItem[]> _nominalMeasureVoltageEnumItems = new Lazy<MeasureVoltageEnumItem[]>(() =>
            MeasureVoltageEnumItem.GetInstances().ToArray());
        
        /// <summary>
        /// Классы с подклассами "Ячейка оборудования ПС"
        /// </summary>
        public static readonly Lazy<Dictionary<string, CubicleClassInfo>> CubicleClassInfos = new Lazy<Dictionary<string, CubicleClassInfo>>(() =>
            CubicleClassInfo.Get()
                .GetAllDescendants()
                .ToDictionary(x => x.Caption.Trim(), x => x)
        );

        /// <summary>
        /// Классы с подклассами "Приборы учета"
        /// </summary>
        private static Lazy<Dictionary<string, MeterEquipmentClassInfo>> _meterClassInfos;

        /// <summary>
        /// Классы с подклассами "Приборы учета"
        /// </summary>
        public static Lazy<Dictionary<string, MeterEquipmentClassInfo>> MeterClassInfos => _meterClassInfos ??= new Lazy<Dictionary<string, MeterEquipmentClassInfo>>(() =>
            ElectricityMeterClassInfo.Get()
                .GetAllDescendants()
                .Where(x => !x.GetDescendants().Where(y => y.IsSystemEntity).Any())
                .Select(x => Tuple.Create(
                    (MeterEquipmentClassInfo)x,
                    x.GetAllAccessors()
                        .TakeWhile(y => y.GetType() != typeof(ElectricityMeterClassInfo))
                        .OfType<MeterEquipmentClassInfo>()
                        .LastOrDefault()))
                .Concat(ResourceMeterClassInfo.Get()
                    .GetAllDescendants()
                    .Where(x => !x.GetDescendants().Where(y => y.IsSystemEntity).Any())
                    .Select(x => Tuple.Create(
                        (MeterEquipmentClassInfo)x,
                        x.GetAllAccessors()
                            .TakeWhile(y => y.GetType() != typeof(ResourceMeterClassInfo))
                            .OfType<MeterEquipmentClassInfo>()
                            .LastOrDefault())))
                .Select(x => Tuple.Create(
                    string.Format("{0} - {1}",
                        x.Item2 == null ? x.Item1.Caption : x.Item2.Caption,
                        x.Item1.Caption),
                    x.Item1))
                .GroupBy(x => x.Item1, x => x.Item2)
                .ToDictionary(x => x.Key.Trim(), x => x.FirstOrDefault())
        );

        /// <summary>
        /// Классы с подклассами "УСПД"
        /// </summary>
        private static Lazy<Dictionary<string, ChannelizingEquipmentClassInfo>> _rtuClassInfos;

        /// <summary>
        /// Классы с подклассами "УСПД"
        /// </summary>
        public static Lazy<Dictionary<string, ChannelizingEquipmentClassInfo>> RtuClassInfos => _rtuClassInfos ??= new Lazy<Dictionary<string, ChannelizingEquipmentClassInfo>>(() =>
            ChannelizingEquipmentClassInfo.Get()
                .GetAllDescendants()
                .Where(x => !x.GetDescendants().Where(y => y.IsSystemEntity).Any())
                .Select(x => Tuple.Create(
                    x,
                    x.GetAllAccessors()
                        .TakeWhile(y => y.GetType() != typeof(ChannelizingEquipmentClassInfo))
                        .LastOrDefault()))
                .Select(x => Tuple.Create(
                    string.Format("{0} - {1}",
                        x.Item2 == null ? x.Item1.Caption : x.Item2.Caption,
                        x.Item1.Caption),
                    x.Item1))
                .GroupBy(x => x.Item1, x => x.Item2)
                .ToDictionary(x => x.Key.Trim(), x => x.FirstOrDefault())
        );

        /// <summary>
        /// Классы с подклассами "Маршрут доступа"
        /// </summary>
        public static readonly Lazy<Dictionary<string, RouteClassInfo>> RouteClassInfos = new Lazy<Dictionary<string, RouteClassInfo>>(() =>
            RouteClassInfo.Get()
                .GetAllDescendants()
                .Distinct()
                .ToDictionary(x => x.Caption.Trim())
        );

        /// <summary>
        /// Поддерживаемые классы маршрутов оборудованием
        /// </summary>
        public static readonly ThreadSafeCache<EquipmentClassInfo, HashSet<RouteClassInfo>> SupportedRouteClassInfosByEquipment = new ThreadSafeCache<EquipmentClassInfo, HashSet<RouteClassInfo>>(x =>
            new HashSet<RouteClassInfo>(((IDirectlyPollingEquipmentType)x).AttributeAvaibleRouteTypes.GetValues())
        );

        /// <summary>
        /// Классы маршрутов через УСПД по классу оборудования
        /// </summary>
        public static readonly Lazy<Dictionary<EquipmentClassInfo, NonDirectRouteClassInfo[]>> NonDirectRouteByRtuClassInfos = new Lazy<Dictionary<EquipmentClassInfo, NonDirectRouteClassInfo[]>>(() =>
        {
            var channelizingEquipmentClassInfo = ChannelizingEquipmentClassInfo.Get();
            return NonDirectRouteClassInfo.Get()
                .GetAllDescendants()
                .SelectMany(x => x.AttributeChannelizingEquipmentType.GetValues(), Tuple.Create)
                .Where(x => x.Item2 != channelizingEquipmentClassInfo)
                .SelectMany(x => x.Item2.GetThisAndAllDescendants(), (x, y) => Tuple.Create(x.Item1, y))
                .GroupBy(x => x.Item2)
                .Where(x => x.Key is IChannelizingEquipmentClass)
                .ToDictionary(x => x.Key, x => x.Select(y => y.Item1).ToArray());
        });

        /// <summary>
        /// Области видимости
        /// </summary>
        [ThreadStatic]
        private static Dictionary<string, IsolationLevel> _isolationLevels;

        /// <summary>
        /// Области видимости
        /// </summary>
        private static Dictionary<string, IsolationLevel> IsolationLevels
        {
            get
            {
                if (_isolationLevels == null)
                {
                    RDClassesAndInstances.SecurityManager.PushSecurityContext();
                    try
                    {
                        _isolationLevels = new Dictionary<string, IsolationLevel>();
                        foreach (var isolationLevel in IsolationLevel.GetInstances())
                            _isolationLevels[isolationLevel.ToString().Replace("/", "\\")] = isolationLevel;
                    }
                    finally
                    {
                        RDClassesAndInstances.SecurityManager.PopSecurityContext();
                    }
                }

                return _isolationLevels;
            }
        }

        /// <summary>
        /// Тарифы
        /// </summary>
        [ThreadStatic]
        private static Dictionary<string, Tariff> _tariffs;

        /// <summary>
        /// Тарифы
        /// </summary>
        public static Dictionary<string, Tariff> Tariffs
        {
            get
            {
                return _tariffs ?? (_tariffs = Tariff
                    .GetInstances()
                    .ToDictionary(x => x.Caption.Trim()));
            }
        }

        /// <summary>
        /// Уровень доступа к оборудованию DLMS
        /// </summary>
        public static readonly Lazy<Dictionary<string, AccessLevelToEquipmentItem>> AccessLevelToEquipmentItems = new Lazy<Dictionary<string, AccessLevelToEquipmentItem>>(() =>
            AccessLevelToEquipmentItem.GetInstances()
                .GroupBy(x => x.Caption.Trim())
                .ToDictionary(x => x.Key, x => x.FirstOrDefault())
        );

        /// <summary>
        /// Уровень доступа счетчика C12xx
        /// </summary>
        public static readonly Lazy<Dictionary<string, C12xxAccessLevelItem>> AccessLevelToC12xx = new Lazy<Dictionary<string, C12xxAccessLevelItem>>(() =>
            C12xxAccessLevelItem.GetInstances()
                .GroupBy(x => x.Caption.Trim())
                .ToDictionary(x => x.Key, x => x.FirstOrDefault())
        );

        /// <summary>
        /// Уровень доступа счетчика CascadeSoft
        /// </summary>
        public static readonly Lazy<Dictionary<string, CascadeSoftAccessLevelsItem>> AccessLevelToCascadeSoft = new Lazy<Dictionary<string, CascadeSoftAccessLevelsItem>>(() =>
            CascadeSoftAccessLevelsItem.GetInstances()
                .GroupBy(x => x.Caption.Trim())
                .ToDictionary(x => x.Key, x => x.FirstOrDefault())
        );

        /// <summary>
        /// Уровень доступа счетчика Меркурий 23x
        /// </summary>
        public static readonly Lazy<Dictionary<string, Mercury23xAccessLevelsItem>> AccessLevelToMercury23x = new Lazy<Dictionary<string, Mercury23xAccessLevelsItem>>(() =>
            Mercury23xAccessLevelsItem.GetInstances()
                .GroupBy(x => x.Caption.Trim())
                .ToDictionary(x => x.Key, x => x.FirstOrDefault())
        );

        /// <summary>
        /// Уровень доступа счетчика MIR
        /// </summary>
        public static readonly Lazy<Dictionary<string, MIRAccessLevelsItem>> AccessLevelToMIR = new Lazy<Dictionary<string, MIRAccessLevelsItem>>(() =>
            MIRAccessLevelsItem.GetInstances()
                .GroupBy(x => x.Caption.Trim())
                .ToDictionary(x => x.Key, x => x.FirstOrDefault())
        );

        /// <summary>
        /// Уровень доступа счетчика Милур (Миландр)
        /// </summary>
        public static readonly Lazy<Dictionary<string, MilurX07AccessLevelItem>> AccessLevelToMilandr = new Lazy<Dictionary<string, MilurX07AccessLevelItem>>(() =>
            MilurX07AccessLevelItem.GetInstances().AsEnumerable()
                .GroupBy(x => x.Caption.Trim())
                .ToDictionary(x => x.Key, x => x.FirstOrDefault())
        );

        /// <summary>
        /// SIM карты
        /// </summary>
        [ThreadStatic]
        public static Dictionary<string, SimCard> _simCards;

        /// <summary>
        /// SIM карты
        /// </summary>
        public static Dictionary<string, SimCard> SimCards
        {
            get
            {
                return _simCards ?? (_simCards = SimCard
                    .GetInstances()
                    .AsEnumerable()
                    .GroupBy(x => x.AttributePhoneNumber?.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x.Key))
                    .ToDictionary(x => x.Key, x => x.FirstOrDefault()));
            }
        }

        /// <summary>
        /// Оператор
        /// </summary>
        [ThreadStatic]
        private static Dictionary<string, CellularOperator> _cellularOperator;

        /// <summary>
        /// Оператор
        /// </summary>
        public static Dictionary<string, CellularOperator> CellularOperators
        {
            get
            {
                return _cellularOperator ?? (_cellularOperator = CellularOperator
                    .GetInstances()
                    .ToDictionary(x => x.Caption.Trim()));
            }
        }

        /// <summary>
        /// Найти позицию первго совпадения массива строк
        /// </summary>
        /// <param name="source"></param>
        /// <param name="separators"></param>
        /// <returns></returns>
        public static int IndexOfAny(this string source, string[] separators)
        {
            var result = -1;
            foreach (var separator in separators)
            {
                var index = source.IndexOf(separator, StringComparison.InvariantCulture);
                if (index != -1)
                {
                    result = index;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// Определяем функцию нормализации значения даты/времени
        /// в этой реализации создаем новый экземпляр с точностью до секунд
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime? TruncateMSec(this DateTime? dt)
        {
            return dt == null ? default : new DateTime(dt.Value.Year, dt.Value.Month, dt.Value.Day, dt.Value.Hour, dt.Value.Minute, dt.Value.Second);
        }

        /// <summary>
        /// Исключить заменители из строки
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static string WithoutStubs(this string source)
        {
            return SeparatorStubs.Aggregate(source, (x, stub) =>
                x.Replace(stub.Key, stub.Value.ToString()));
        }
        
        /// <summary>
        /// Функция получения строки
        /// </summary>
        public static string GetString(string value)
        {
            return value;
        }

        /// <summary>
        /// Функция получения строк
        /// </summary>
        public static string[] GetStrings(string value)
        {
            return value == null ? null : value.Split(MultiValueSeparators, StringSplitOptions.RemoveEmptyEntries);
        }

        /// <summary>
        /// Функция получения целого числа
        /// </summary>
        public static int? GetInt(string value)
        {
            if (value == null)
                return null;

            int result;
            if (int.TryParse(value, out result))
                return result;

            return null;
        }

        /// <summary>
        /// Функция получения булевого значения
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool? GetBool(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;
            bool result;
            if (bool.TryParse(value, out result))
                return result;
            return null;
        }

        /// <summary>
        /// Функция получения булевого значения
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool? GetBoolStr(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;
            return string.Equals(value, "да", StringComparison.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Функция получения вещественного числа
        /// </summary>
        public static double? GetDouble(string value)
        {
            if (value == null)
                return null;
            
            double result;
            if (double.TryParse(value
                    .Replace(".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                    .Replace(",", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator),
                out result))
                return result;

            return null;
        }

        /// <summary>
        /// Функция получения даты и времени
        /// </summary>
        public static DateTime? GetDateTime(string value)
        {
            if (value == null)
                return null;

            DateTime result;
            if (DateTime.TryParseExact(value, "dd.MM.yyyy HH:mm:ss", null, DateTimeStyles.None, out result))
                return result;
            if (DateTime.TryParseExact(value, "dd.MM.yyyy", null, DateTimeStyles.None, out result))
                return result;
            if (DateTime.TryParse(value, out result))
                return result;

            return null;
        }

        /// <summary>
        /// Функция получения уровня напряжения
        /// </summary>
        public static VoltageEnumItem GetVoltage(string source)
        {
            if (source == null)
                return null;
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _voltageEnumItems.Value
                    .FirstOrDefault(x => Math.Abs((x.AttributeValue ?? 0.0) - value * 1000) < double.Epsilon);
            // значение как текст
            return _voltageEnumItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения класса точности
        /// </summary>
        public static AccuracyRatingEnumItem GetAccuracyRating(string source)
        {
            if (source == null)
                return null;
            
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _accuracyRatingEnumItems.Value
                    .Where(x => x.AttributeAccuracyRatingType == null)
                    .FirstOrDefault(x => Math.Abs((x.AttributeValue ?? 0.0) - value) < double.Epsilon);
            // значение как текст
            return _accuracyRatingEnumItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения схемы включения
        /// </summary>
        public static MeterJoinSchemesItem GetMeterJoinSchemes(string source)
        {
            if (source == null)
                return null;

            // значение как текст
            return _meterJoinSchemesItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения причин замены
        /// </summary>
        public static EquipmentReplacementReason GetEquipmentReplacementReasons(string source)
        {
            if (source == null)
                return null;

            // значение как текст
            return _equipmentReplacementReasons.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения часового пояса
        /// </summary>
        public static TimeZone GetTimeZone(string source)
        {
            if (source == null)
                return null;
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _timeZones.Value
                    .FirstOrDefault(x => Math.Abs((x.AttributeUTCOffsetHours ?? 0.0) - value) < double.Epsilon);
            // значение как текст
            return _timeZones.Value
                .FirstOrDefault(x => x.Caption.IndexOf(source, StringComparison.InvariantCultureIgnoreCase) != -1);
        }

        /// <summary>
		/// Функция получения области видимости
		/// </summary>
		public static IsolationLevel GetIsolationLevel(string source)
		{
			if (source == null)
			    return null;

		    return IsolationLevels
		        .Where(x => x.Key.EndsWith(source.Replace("/", "\\"), StringComparison.InvariantCultureIgnoreCase))
		        .Select(x => x.Value)
		        .FirstOrDefault();
		}

        /// <summary>
        /// Функция получения области видимости
        /// </summary>
        public static IsolationLevel[] GetIsolationLevels(string source)
        {
            var zoneNames = GetStrings(source);
            if (zoneNames == null)
                return null;
            var result = new List<IsolationLevel>();
            foreach (var zoneName in zoneNames.Where(x => x != null).Select(x => x.Replace("/", "\\")))
            {
                var foundZones = IsolationLevels
                    .Where(x => x.Key.EndsWith(zoneName, StringComparison.InvariantCultureIgnoreCase))
                    .Select(x => x.Value)
                    .ToArray();
                
                if (foundZones.Length > 1)
                    return null;

                result.AddRange(foundZones);
            }

            return result.Any() ? result.ToArray() : null;
        }
        
        /// <summary>
        /// Функция получения тарифа
        /// </summary>
        public static Tariff GetTariff(string source)
        {
            if (source == null)
                return null;

            Tariff result;
            return Tariffs.TryGetValue(source, out result) ? result : null;
        }

        /// <summary>
        /// Функция получения оператора
        /// </summary>
        public static CellularOperator GetCellularOperator(string source)
        {
            if (source == null)
                return null;

            CellularOperator result;
            return CellularOperators.TryGetValue(source, out result) ? result : null;
        }

        /// <summary>
        /// Функция получения номинала тока
        /// </summary>
        public static CurrentEnumItem GetNominalCurrent(string source)
        {
            if (source == null)
                return null;
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _nominalCurrentEnumItems.Value
                    .FirstOrDefault(x => Math.Abs((x.AttributeValue ?? 0.0) - value) < double.Epsilon);
            // значение как текст
            return _nominalCurrentEnumItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения номинала напряжения
        /// </summary>
        public static VoltageEnumItem GetNominalVoltage(string source)
        {
            if (source == null)
                return null;
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _nominalVoltageEnumItems.Value
                    .FirstOrDefault(x => Math.Abs((x.AttributeValue ?? 0.0) - value) < double.Epsilon);
            // значение как текст
            return _nominalVoltageEnumItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения номинала тока измерительный цепей
        /// </summary>
        public static MeasureCurrentEnumItem GetNominalMeasureCurrent(string source)
        {
            if (source == null)
                return null;
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _nominalMeasureCurrentEnumItems.Value
                    .FirstOrDefault(x => Math.Abs((x.AttributeValue ?? 0.0) - value) < double.Epsilon);
            // значение как текст
            return _nominalMeasureCurrentEnumItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения номинала напряжения измерительный цепей
        /// </summary>
        public static MeasureVoltageEnumItem GetNominalMeasureVoltage(string source)
        {
            if (source == null)
                return null;
            // значение как число
            double value;
            if (double.TryParse(source, out value))
                return _nominalMeasureVoltageEnumItems.Value
                    .FirstOrDefault(x => Math.Abs((x.AttributeValue ?? 0.0) - value) < double.Epsilon);
            // значение как текст
            return _nominalMeasureVoltageEnumItems.Value
                .FirstOrDefault(x => x.Caption.Equals(source, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Функция получения типа ячейки
        /// </summary>
        public static CubicleClassInfo GetCubicleClass(string source)
        {
            if (source == null)
                return null;
            CubicleClassInfo result;
            string alias;
            if (!CubicleClassInfos.Value.TryGetValue(source, out result) &&
                CubicleAliases.TryGetValue(source, out alias))
                CubicleClassInfos.Value.TryGetValue(alias, out result);

            return result;
        }

        /// <summary>
        /// Функция получения типа ПУ
        /// </summary>
        public static MeterEquipmentClassInfo GetMeterClass(string source)
        {
            if (source == null)
                return null;
            MeterEquipmentClassInfo result;
            MeterClassInfos.Value.TryGetValue(source, out result);
            return result;
        }

        /// <summary>
        /// Функция получения типа УСПД
        /// </summary>
        public static ChannelizingEquipmentClassInfo GetRtuClass(string source)
        {
            if (source == null)
                return null;
            ChannelizingEquipmentClassInfo result;
            RtuClassInfos.Value.TryGetValue(source, out result);
            return result;
        }

        /// <summary>
        /// Функция получения адреса ФИАС
        /// </summary>
        public static FiasSuggestionData GetAddressFias(string source)
        {
            if (source == null)
                return null;

            // функция установки уровня ФИАС
            var setFiasLevel = new Func<FiasSuggestionData, FiasLevelItem, bool>((fiasSuggestion, fiasLevel) =>
            {
                if (fiasSuggestion.FiasLevel == null || fiasSuggestion.FiasLevel.OrderPosition < fiasLevel.OrderPosition)
                {
                    fiasSuggestion.FiasLevel = fiasLevel;
                    return true;
                }

                return false;
            });

            // Удаление лишних пробелов в адресе ФИАС
            source = Regex.Replace(source, "[ ]+", " ");
            var result = new FiasSuggestionData { Value = source };
            var address = source.Split(',');
            foreach (var part in address.Select(x => (x ?? string.Empty).Trim()))
            {
                var parseNeeded = true;
                // Регион (обл, респ, край)
                if (parseNeeded && (
                    part.EndsWith(" обл", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" край", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" респ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" Аобл", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" АО", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("респ ", StringComparison.InvariantCultureIgnoreCase)))
                {
                    if (setFiasLevel(result, FiasLevelItem.Instances.Region))
                    {
                        result.Region = part;
                        parseNeeded = false;
                    }
                }

                // Район в регионе (р-н)
                // Район города (р-н)
                if (parseNeeded && (
                    part.EndsWith(" р-н", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("р-н ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" улус", StringComparison.InvariantCultureIgnoreCase)))
                {
                    // если город имеется, то район города
                    if (!string.IsNullOrEmpty(result.City))
                    {
                        if (setFiasLevel(result, FiasLevelItem.Instances.CityDistrict))
                        {
                            result.CityDistrict = part;
                            parseNeeded = false;
                        }
                    }
                    // иначе район региона
                    else
                    {
                        if (setFiasLevel(result, FiasLevelItem.Instances.Area))
                        {
                            result.Area = part;
                            parseNeeded = false;
                        }
                    }
                }

                // Город (г)
                if (parseNeeded && (
                    part.StartsWith("г ", StringComparison.InvariantCultureIgnoreCase)))
                {
                    var f = true;
                    // если город уже имеется, то город перезаписывается, а старое значение становится Регионом
                    // пример: г Санкт-Петербург, г Колпино, пр-кт Ленина, д 8
                    if (!string.IsNullOrEmpty(result.City))
                    {
                        result.Region = result.City;
                    }
                    else
                    {
                        f = setFiasLevel(result, FiasLevelItem.Instances.City);
                    }

                    if (f)
                    {
                        result.City = part;
                        parseNeeded = false;
                    }
                }

                // Населенный пункт (нп, деревня, поселок, село, ст, тер., )
                if (parseNeeded && (
                    part.StartsWith("нп ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("аал ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("деревня ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("поселок ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("село ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("хутор ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("ст ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("п/ст ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("тер.", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("тер ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("кв-л ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("мкр ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("пгт ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("жилзона ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("ст-ца ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("гп ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("рп ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" слобода", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("аул ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" сл", StringComparison.InvariantCultureIgnoreCase)))
                {
                    // если населенный пункт уже имеется, то новое значение становится улицей
                    if (!string.IsNullOrEmpty(result.Settlement))
                    {
                        if (setFiasLevel(result, FiasLevelItem.Instances.Street))
                        {
                            result.Street = part;
                            parseNeeded = false;
                        }
                    }
                    else
                    {
                        if (setFiasLevel(result, FiasLevelItem.Instances.Settlement))
                        {
                            result.Settlement = part;
                            parseNeeded = false;
                        }
                    }
                }

                // Улица (ул)
                if (parseNeeded && (
                    part.StartsWith("ул ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("шоссе ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" шоссе", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" пер", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("пер ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" тупик", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" спуск", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" пр-кт", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("пр-кт ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("проезд ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" проезд", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("сад ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("наб ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" мкр", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" пл", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("пл ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.StartsWith("б-р ", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" б-р", StringComparison.InvariantCultureIgnoreCase) ||
                    part.EndsWith(" наб", StringComparison.InvariantCultureIgnoreCase)))
                {
                    if (setFiasLevel(result, FiasLevelItem.Instances.Street))
                    {
                        result.Street = part;
                        parseNeeded = false;
                    }
                }

                // Дом (д)
                if (parseNeeded && (
                    part.StartsWith("д ", StringComparison.InvariantCultureIgnoreCase)))
                {
                    if (setFiasLevel(result, FiasLevelItem.Instances.House))
                    {
                        result.House = part;
                        parseNeeded = false;
                    }
                }

                // Корпус/строение (стр)
                if (parseNeeded && (
                    part.StartsWith("стр ", StringComparison.InvariantCultureIgnoreCase)))
                {
                    result.Block = part;
                    //parseNeeded = false;
                }
            }

            // принять строение как дом, в случае отсутствия дома
            if (string.IsNullOrEmpty(result.House) && !string.IsNullOrEmpty(result.Block))
            {
                if (setFiasLevel(result, FiasLevelItem.Instances.House))
                {
                    result.House = result.Block;
                }
            }

            if (result.FiasLevel == null)
                result = null;

            return result;
        }

        /// <summary>
        /// Функция получения ссылки на сущность
        /// </summary>
        public static EntityLink GetEntityLink(string source)
        {
            if (source == null)
                return null;

            // читаем серийный номер и тип до первого разделителя,
            // все остальное - произвольное описание параметров
            //var idxFirstSeparator = source.IndexOfAny(MultiValueSeparators);
            //if (idxFirstSeparator == -1)
            var idxFirstSeparator = source.IndexOfAny(MultiValueSubSeparators);

            Dictionary<string, string> otherParams = null;
            if (idxFirstSeparator != -1)
            {
                otherParams = source.Substring(idxFirstSeparator + 1).Trim()
                    .Split(MultiValueSubSeparators, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x =>
                    {
                        var equalIndex = x.IndexOf("=", StringComparison.Ordinal);
                        return Tuple.Create(
                            equalIndex != -1 ? x.Substring(0, equalIndex) : string.Empty,
                            equalIndex != -1 ? x.Substring(equalIndex + 1) : string.Empty);
                    })
                    .GroupBy(x => x.Item1, x => x.Item2)
                    .ToDictionary(x => x.Key, x => x.FirstOrDefault());
                source = source.Substring(0, idxFirstSeparator).Trim();
            }

            // подмена заменителей, на реальные знаки
            source = source.WithoutStubs();

            var idxTypeStart = source.IndexOf("(", StringComparison.InvariantCultureIgnoreCase);
            var idxTypeEnd = source.LastIndexOf(")", StringComparison.InvariantCultureIgnoreCase);

            var result = new EntityLink
            {
                SerialNumber = source,
                OtherParams = otherParams ?? new Dictionary<string, string>()
            };

            if (idxTypeStart != -1 && idxTypeEnd != -1)
            {
                result.SerialNumber = source.Substring(0, idxTypeStart).Trim();
                result.EntityType = source.Substring(idxTypeStart + 1, idxTypeEnd - idxTypeStart - 1).Trim();
            }

            return string.IsNullOrEmpty(result.SerialNumber) ? null : result;
        }
        
        /// <summary>
        /// Функция получения маршрута
        /// </summary>
        public static ImportRoute[] GetRoutes(string source)
        {
            if (source == null)
                return null;

            var result = new List<ImportRoute>(source
                .Split(MultiValueSeparators, StringSplitOptions.RemoveEmptyEntries)
                .SelectMany(x => ParseRoute(x.Trim()))
            );

            return !result.Any() ? null : result.ToArray();
        }

        /// <summary>
        /// Функция получения параметров (ключ=значение)
        /// </summary>
        public static Dictionary<string, string> GetParams(string source)
        {
            if (source == null)
                return null;

            var result = new Dictionary<string, string>();

            foreach (var pair in source
                .Split(MultiValueSubSeparators, StringSplitOptions.RemoveEmptyEntries))
            {
                var equalIndex = pair.IndexOf("=", StringComparison.Ordinal);
                if (equalIndex == -1)
                    continue;
                result[pair.Substring(0, equalIndex).WithoutStubs()] = pair.Substring(equalIndex + 1).WithoutStubs();
            }

            return result;
        }

        /// <summary>
        /// Распарсить строку маршрута
        /// </summary>
        /// <returns></returns>
        private static IEnumerable<ImportRoute> ParseRoute(string source)
        {
            if (string.IsNullOrEmpty(source))
                yield break;

            var idxFirstSeparator = source.IndexOfAny(MultiValueSubSeparators);
            var otherParams = new Dictionary<string, string>();
            if (idxFirstSeparator != -1)
            {
                otherParams = source.Substring(idxFirstSeparator + 1).Trim()
                    .Split(MultiValueSubSeparators, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x =>
                    {
                        var equalIndex = x.IndexOf("=", StringComparison.Ordinal);
                        return Tuple.Create(
                            equalIndex != -1 ? x.Substring(0, equalIndex) : string.Empty,
                            equalIndex != -1 ? x.Substring(equalIndex + 1) : string.Empty);
                    })
                    .GroupBy(x => x.Item1, x => x.Item2)
                    .ToDictionary(x => x.Key, x => x.FirstOrDefault());
                source = source.Substring(0, idxFirstSeparator).Trim();
            }

            // маршрут TCP
            var index = source.IndexOf(":", StringComparison.Ordinal);
            if (index > -1)
            {
                var ipAddress = source.Substring(0, index).Trim();
                var portStr = source.Substring(index + 1).Trim();
                int port;
                if (int.TryParse(portStr, out port) && port > 0)
                {
                    ImportRoute result;
                    if (!string.IsNullOrEmpty(ipAddress))
                    {
                        result = new ImportRoute(TCPClientRouteClassInfo.Get())
                        {
                            Params =
                            {
                                { RDMetadataGuids.CustomIPClientRoute.InstanceAttributeHostOrIP, ipAddress },
                                { RDMetadataGuids.CustomIPClientRoute.InstanceAttributePort, port },
                            },
                            OtherParams = otherParams
                        };
                    }
                    else
                    {
                        result = new ImportRoute(TCPServerRouteClassInfo.Get())
                        {
                            Params =
                            {
                                {
                                    RDMetadataGuids.ITCPServerRoute.AttributeListeningTCPSocket,
                                    // !! серверный порт, должен быть создан на этапе установки значения в экземпляр (метод Assign)
                                    port
                                },
                            },
                            OtherParams = otherParams
                        };
                    }
                    
                    yield return result;
                    yield break;
                }
            }

            // маршрут телефонный номер
            const string POOL_OF_MODEMS_PATTERN = "<(.*?)>";

            var phoneStr = Regex.Replace(source, POOL_OF_MODEMS_PATTERN, "");
            if (phoneStr.Length > 5 && phoneStr.Replace(" ", "").All(x => "1234567890+pwt,".Contains(x)))
            {
                var poolOfModemsStr = Regex.Match(source, POOL_OF_MODEMS_PATTERN).Value.Trim();
                poolOfModemsStr = Regex.Replace(poolOfModemsStr, "[<>]", "").Trim();
                phoneStr = Regex.Replace(source, POOL_OF_MODEMS_PATTERN, "").Trim();
                // Добавить пул модемов, если есть
                var phoneNumbers = phoneStr.Replace(" ", "").Split(',');
                var f = false;
                foreach (var phoneNumber in phoneNumbers.Select(x => x.Trim()))
                {
                    yield return new ImportRoute(ModemRouteClassInfo.Get())
                    {
                        Params =
                        {
                            { RDMetadataGuids.CustomModemRoute.InstanceAttributePhoneNumber, phoneNumber.WithoutStubs() },
                            // !! модемный пул, должен быть создан на этапе установки значения в экземпляр (метод Assign)
                            { RDMetadataGuids.CustomModemRoute.InstanceAttributeModemsPool, poolOfModemsStr.WithoutStubs() },
                        },
                        OtherParams = otherParams
                    };
                    f = true;
                }
                if (f) yield break;
            }

            // специальный маршрут
            RouteClassInfo routeClassInfoDummy;
            if (RouteClassInfos.Value.TryGetValue(source.WithoutStubs(), out routeClassInfoDummy))
            {
                yield return new ImportRoute(routeClassInfoDummy)
                {
                    OtherParams = otherParams
                };
                yield break;
            }

            // маршрут через промежуточное оборудование
            var link = GetEntityLink(source);
            if (link != null)
            {
                yield return new ImportRoute(NonDirectRouteClassInfo.Get())
                {
                    Params =
                    {
                        { RDMetadataGuids.NonDirectRoute.InstanceAttributeChannelizingEquipment, link },
                    },
                    OtherParams = otherParams
                };
            }
        }

        /// <summary>
        /// Обработчик атрибутов класса
        /// </summary>
        public class ClassAttributesParser
        {
            /// <summary>
            /// Атрибуты, значения которых могут быть получены из строковых описателей при установке в экземпляр
            /// </summary>
            private static readonly HashSet<Guid> _entityAttributeRefNamesWithUnderlyingValues = new HashSet<Guid>
            {
                // серверный сокет
                RDMetadataGuids.ITCPServerRoute.AttributeListeningTCPSocket
            };

            /// <summary>
            /// Кэш атрибутов класса
            /// </summary>
            private static readonly ThreadSafeCache<BaseClassClassInfo, RDAttribute[]> _classAttributesCache = new ThreadSafeCache<BaseClassClassInfo, RDAttribute[]>(x =>
                x.GetOwnAttributes(
                        ClassAreaOfAction.ForInstance,
                        InheritanceLevelKind.CurrentLevelAndAncestors,
                        true)
                    .ToArray()
            );

            /// <summary>
            /// Кэш значений перечисления
            /// </summary>
            private static readonly ThreadSafeCache<EnumerationItemClassInfo, EnumerationItem[]> _enumItemsCache = new ThreadSafeCache<EnumerationItemClassInfo, EnumerationItem[]>(x =>
                x.GetInstances().ToArray()
            );

            // класс
            private readonly BaseClassClassInfo _classInfo;

            /// <summary>
            /// Конструктор
            /// </summary>
            /// <param name="classInfo"></param>
            public ClassAttributesParser(BaseClassClassInfo classInfo)
            {
                _classInfo = classInfo;
            }

            /// <summary>
            /// Установить результирующее значения при парсинге атрибутов
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="attribute"></param>
            /// <param name="values"></param>
            /// <param name="resultValue"></param>
            /// <returns></returns>
            private static bool SetParseResultValue<T>(RDAttribute attribute, T[] values, out object resultValue)
            {
                resultValue = null;

                // массив
                if (attribute.Options.MaxOccurs > 1)
                {
                    resultValue = values;
                    return values != null;
                }

                resultValue = values.FirstOrDefault();
                return attribute.Options.MinOccurs == 1 ? resultValue != null : values.Any();
            }

            /// <summary>
            /// Разбор специфичных атрибутов
            /// </summary>
            /// <param name="otherParams"></param>
            /// <param name="resultParams"></param>
            /// <param name="errorText"></param>
            /// <returns></returns>
            public bool Parse(IEnumerable<KeyValuePair<string, string>> otherParams,
                Dictionary<Guid, object> resultParams, out string errorText)
            {
                errorText = null;
                if (otherParams == null || resultParams == null)
                    return true;

                var otherParamsDic = otherParams
                    .GroupBy(x => x.Key, x => x.Value.WithoutStubs())
                    .ToDictionary(x => x.Key, x => x.FirstOrDefault());

                if (otherParamsDic.Any())
                {
                    foreach (var attribute in _classAttributesCache.Get(_classInfo))
                    {
                        string value;
                        string alias;
                        
                        var attributeCaption = string.Empty;
                        var f = otherParamsDic.TryGetValue(attribute.Caption, out value);
                        if (f)
                            attributeCaption = attribute.Caption;
                        else if (AttributesAliases.TryGetValue(attribute.RefName, out alias) && otherParamsDic.TryGetValue(alias, out value))
                        {
                            attributeCaption = alias;
                            f = true;
                        }

                        if (f)
                        {
                            var result = false;
                            object resultValue = null;
                            switch (attribute.AttributeType)
                            {
                                case AttributeValueType.Integer:
                                    int intValue;
                                    var intValues = value.Split(ArraySeparator)
                                        .Select(x => int.TryParse(x, out intValue) ? intValue : (int?)null)
                                        .ToArray();
                                    result = SetParseResultValue(attribute, intValues, out resultValue);
                                    break;
                                case AttributeValueType.Double:
                                    double doubleValue;
                                    var doubleValues = value.Split(ArraySeparator)
                                        .Select(x => double.TryParse(x, out doubleValue) ? doubleValue : (double?)null)
                                        .ToArray();
                                    result = SetParseResultValue(attribute, doubleValues, out resultValue);
                                    break;
                                case AttributeValueType.String:
                                    var stringValues = value.Split(ArraySeparator);
                                    result = SetParseResultValue(attribute, stringValues, out resultValue);
                                    break;
                                case AttributeValueType.Boolean:
                                    bool boolValue;
                                    var boolValues = value.Split(ArraySeparator)
                                        .Select(x =>
                                            bool.TryParse(x, out boolValue)
                                                ? boolValue
                                                : int.TryParse(x, out intValue)
                                                    ? intValue > 0
                                                    : (bool?)null)
                                        .ToArray();
                                    result = SetParseResultValue(attribute, boolValues, out resultValue);
                                    break;
                                case AttributeValueType.DateTime:
                                    DateTime dateTimeValue;
                                    var dateTimeValues = value.Split(ArraySeparator)
                                        .Select(x => DateTime.TryParseExact(x, "dd.MM.yyyy HH:mm:ss",
                                            DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out dateTimeValue)
                                            ? dateTimeValue
                                            : (DateTime?)null)
                                        .ToArray();
                                    result = SetParseResultValue(attribute, dateTimeValues, out resultValue);
                                    break;
                                case AttributeValueType.Entity:
                                   var classInfo = RDCustomEntity.Find(attribute.Options.RequiredTypeRefName.GetValueOrDefault()) as BaseClassClassInfo;
                                    var enumerationClassInfo = classInfo as EnumerationItemClassInfo;
                                    if (enumerationClassInfo != null)
                                    {
                                        var enumItems = _enumItemsCache.Get(enumerationClassInfo);
                                        var enumValues = value.Split(ArraySeparator)
                                            .Select(x => enumItems.FirstOrDefault(y => string.Equals(y.Caption, x, StringComparison.InvariantCultureIgnoreCase)))
                                            .ToArray();
                                        result = SetParseResultValue(attribute, enumValues, out resultValue);
                                    }
                                    else
                                    {
                                        Guid guidValue;
                                        var entityValues = value.Split(ArraySeparator)
                                            .Select(x => Guid.TryParse(x, out guidValue) ? guidValue : (Guid?)null)
                                            .Select(x => x.HasValue ? RDCustomEntity.Find(x.Value) : null)
                                            .Where(x => x != null)
                                            .ToArray();
                                        // выполнить поиск по наименованию среди объектов с типом атрибута
                                        if (!entityValues.Any() && classInfo != null)
                                        {
                                            entityValues = value.Split(ArraySeparator)
                                                .Select(x => classInfo.GetInstances().FirstOrDefault(y => string.Equals(y.Caption, x, StringComparison.InvariantCultureIgnoreCase)))
                                                .OfType<RDCustomEntity>()
                                                .ToArray();
                                        }

                                        // оставить введенные значения, чтобы попытаться их распарсить на этапе установки в экземпляр (методы Assing и Equals)
                                        if (entityValues.Length == 0 && _entityAttributeRefNamesWithUnderlyingValues.Contains(attribute.RefName))
                                        {
                                            result = SetParseResultValue(attribute, value.Split(ArraySeparator), out resultValue);
                                        }
                                        else
                                        {
                                            result = entityValues.Length == 0 || SetParseResultValue(attribute, entityValues, out resultValue);
                                        }
                                    }
                                    break;
                                case AttributeValueType.Blob:
                                    byte[] blobValues = null;
                                    try
                                    {
                                        blobValues = value.Split(ArraySeparator)
                                            .Select(x => Convert.ToByte(x, 16))
                                            .ToArray();
                                    }
                                    catch
                                    {
                                        //
                                    }

                                    result = SetParseResultValue(attribute, blobValues, out resultValue);
                                    break;
                            }

                            if (!result)
                                errorText = string.Format(
                                    "Не удалось разобрать значение параметра \"{0}\"={1}. Ожидалось значение типа {2}",
                                    attributeCaption, value, attribute.AttributeType);
                            else
                                resultParams[attribute.RefName] = resultValue;

                            // удалить разобранное
                            otherParamsDic.Remove(attributeCaption);
                        }

                        // ошибка
                        if (!string.IsNullOrEmpty(errorText))
                            return false;
                    }

                    // не разобранные параметры
                    if (otherParamsDic.Count > 0)
                    {
                        errorText = string.Format("Ни одни из следующих параметров не соответствует атрибуту класса \"{0}\":\n{1}",
                            _classInfo.Caption, string.Join("\n", otherParamsDic.Select(x => x.Key)));
                        return false;
                    }
                }
                
                return true;
            }

            /// <summary>
            /// Разбор специфичных атрибутов
            /// </summary>
            /// <param name="resultParams"></param>
            /// <param name="excludeParams"></param>
            /// <param name="errorText"></param>
            /// <returns></returns>
            public bool CheckRequired(Dictionary<Guid, object> resultParams, IEnumerable<Guid> excludeParams, out string errorText)
            {
                errorText = null;
                
                // список исключений
                var excludeHash = new HashSet<Guid>(excludeParams ?? new Guid[0]);

                // проверить обязательные атрибуты
                foreach (var attribute in _classAttributesCache.Get(_classInfo)
                    .Where(x => x.Options.MinOccurs == 1)
                    .Where(x => !x.AttributeWebVisualizationRules().GetValues().Contains(AttributeWebVisualizationRuleItem.Instances.DisallowViewInPassport))
                    .Where(x => !excludeHash.Contains(x.RefName)))
                {
                    if (resultParams.ContainsKey(attribute.RefName))
                        continue;
                    errorText = string.Format(
                        "Не удалось найти значение обязательного параметра \"{0}\" класса \"{1}\"",
                        attribute.Caption, _classInfo.Caption);
                    return false;
                }

                return true;
            }

            /// <summary>
            /// Получить значение параметра, с возможными преобразованиями
            /// </summary>
            /// <returns></returns>
            private object GetUndelyingValue(RDAttribute attribute, object value)
            {
                if (value == null)
                    return null;

                // получить значение из Lazy (!!! на текущий момент более не используется)
                var type = value.GetType();
                if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Lazy<>))
                {
                    var prop = type.GetProperty("Value");
                    if (prop != null)
                        value = prop.GetValue(value);
                }

                // атрибуты из _entityAttributeRefNamesWithUnderlyingValues

                // серверный сокет
                // !! здесь можем оказаться также, если порт указали как строка - наименование и экземпляр отсутствует в системе
                // в этом случае, устанавливаемое значение будет null
                if (attribute.RefName == RDMetadataGuids.ITCPServerRoute.AttributeListeningTCPSocket && !(value is ListeningTCPSocket))
                {
                    // порт как целое число, будет создан, если отсутствует
                    var port = (value as int?) ?? (int.TryParse(value as string ?? "", out var intPort) ? intPort : default(int?));
                    if (port != null)
                    {
                        var resultSocket = ImportHelpers.FindOrCreate.Entity<ListeningTCPSocket>(socket =>
                            socket.AttributePort == port);
                        resultSocket.AttributePort = port;
                        value = resultSocket;
                    }
                    else
                    {
                        if (value != null)
                            ImportHelpers.AddWarn($"Не удалось разобрать значение для атрибута {attribute} не одним из возможных вариантов");
                        value = null;
                    }
                }
                // модемный пул
                else if (attribute.RefName == RDMetadataGuids.CustomModemRoute.InstanceAttributeModemsPool && !(value is PoolOfModems))
                {
                    if (value is string poolCaption && !string.IsNullOrEmpty(poolCaption))
                    {
                        var resultPool = ImportHelpers.FindOrCreate.Entity<PoolOfModems>(pool =>
                            pool.AttributeCaption == poolCaption);
                        resultPool.AttributeCaption = poolCaption;
                        value = resultPool;
                    }
                    else
                    {
                        if (value != null && !(value is string))
                            ImportHelpers.AddWarn($"Не удалось разобрать значение для атрибута {attribute} не одним из возможных вариантов");
                        value = null;
                    }
                }

                return value;
            }

            /// <summary>
            /// Заполнить параметры экземпляра
            /// </summary>
            public void Assign(RDInstance other, Dictionary<Guid, object> resultParams)
            {
                if (other == null || other.Class != _classInfo || resultParams == null)
                    return;

                // TODO: Внимание! resultParams уже должен содержать ссылку на экземпляр сущности, а не EntityLink, т.е должен быть подменен перед вызовом метода

                var classAttributes = _classAttributesCache.Get(_classInfo);
                foreach (var attribute in classAttributes)
                {
                    object paramValue;
                    if (!resultParams.TryGetValue(attribute.RefName, out paramValue))
                        continue;

                    paramValue = GetUndelyingValue(attribute, paramValue);

                    var values = other.GetAttributeValues(attribute) as RDAttributeValues;
                    if (values == null)
                        continue;

                    // массив
                    if (attribute.Options.MaxOccurs > 1)
                    {
                        var enumerable = paramValue as IEnumerable;
                        if (enumerable != null)
                            values.SetValues(enumerable.OfType<object>());
                    }
                    // не массив
                    else
                    {
                        values.SetValues(new[] { paramValue });
                    }
                }
            }

            /// <summary>
            /// Сравнение на равенство
            /// </summary>
            /// <param name="other"></param>
            /// <param name="resultParams"></param>
            /// <param name="excludeAtributes"></param>
            /// <returns></returns>
            public bool Equals(RDInstance other, Dictionary<Guid, object> resultParams, IEnumerable<RDAttribute> excludeAtributes)
            {
                if (other == null || other.Class != _classInfo || resultParams == null)
                    return false;

                bool? result = null;

                // TODO: Внимание! Params уже должен содержать ссылку на экземпляр сущности, а не EntityLink, т.е должен быть подменен перед вызовом метода

                var excludeAtributesHash = new HashSet<RDAttribute>(excludeAtributes ?? new RDAttribute[0]);
                var classAttributes = _classAttributesCache.Get(_classInfo);
                foreach (var attribute in classAttributes
                    .Where(x => !excludeAtributesHash.Contains(x)))
                {
                    object paramValue;
                    if (!resultParams.TryGetValue(attribute.RefName, out paramValue))
                        continue;

                    paramValue = GetUndelyingValue(attribute, paramValue);

                    var values = other.GetAttributeValues(attribute) as RDAttributeValues;
                    if (values == null)
                        continue;

                    // массив
                    if (attribute.Options.MaxOccurs > 1)
                    {
                        var enumerable = paramValue as IEnumerable;
                        if (enumerable != null)
                            result = values.GetValues()
                                .Select(x => x.Value)
                                .SequenceEqual(enumerable.OfType<object>());
                    }
                    // не массив
                    else
                    {
                        result = Equals(values.GetValues()
                                .Select(x => x.Value)
                                .FirstOrDefault(),
                            paramValue);
                    }

                    // значения не равны
                    if (!result.GetValueOrDefault(true))
                        break;
                }

                return result.GetValueOrDefault();
            }

            /// <summary>
            /// Сравнение на равенство
            /// </summary>
            /// <param name="other"></param>
            /// <param name="resultParams"></param>
            /// <returns></returns>
            public bool Equals(RDInstance other, Dictionary<Guid, object> resultParams)
            {
                return Equals(other, resultParams, null);
            }
        }
    }
    
    /// <summary>
    /// Ссылка на сущность
    /// </summary>
    internal class EntityLink
    {
        /// <summary>
        /// Серийный номер
        /// </summary>
        public string SerialNumber { get; set; }

        /// <summary>
        /// Тип сущности
        /// </summary>
        public string EntityType { get; set; }

        /// <summary>
        /// Другие параметры переданные в формате Ключ=Значение
        /// </summary>
        public Dictionary<string, string> OtherParams = new Dictionary<string, string>();
    }

    /// <summary>
    /// Маршрут
    /// </summary>
    internal class ImportRoute : IEquatable<Route>
    {
        /// <summary>
        /// Класс маршрута
        /// </summary>
        public RouteClassInfo ClassInfo { get; set; }
        
        /// <summary>
        /// Специфичные параметры
        /// </summary>
        public Dictionary<Guid, object> Params = new Dictionary<Guid, object>();

        /// <summary>
        /// Другие параметры переданные в формате Ключ=Значение
        /// </summary>
        public Dictionary<string, string> OtherParams = new Dictionary<string, string>();

        /// <summary>
        /// Конструктор
        /// </summary>
        public ImportRoute(RouteClassInfo classInfo)
        {
            ClassInfo = classInfo;
        }

        /// <summary>
        /// Распарсить из строки для уточнения по типу ПУ
        /// </summary>
        public bool ParseFor(EquipmentClassInfo equipmentClassInfo, EquipmentClassInfo channelizingEquipmentClassInfo, out string errorText)
        {
            errorText = null;
            // маршрут через УСПД
            // маршрут через промежуточное оборудование
            if (ClassInfo == NonDirectRouteClassInfo.Get() && channelizingEquipmentClassInfo != null)
            {
                // берем первый попавшийся тип маршрута для класса УСПД
                NonDirectRouteClassInfo[] routeClassInfos;
                if (Helpers.NonDirectRouteByRtuClassInfos.Value.TryGetValue(channelizingEquipmentClassInfo, out routeClassInfos) &&
                    routeClassInfos.Any())
                {
                    ClassInfo = routeClassInfos.First();
                }
            }

            return ParseFor(equipmentClassInfo, out errorText);
        }

        /// <summary>
        /// Распарсить из строки для уточнения по типу оборудования
        /// </summary>
        public bool ParseFor(EquipmentClassInfo equipmentClassInfo, out string errorText)
        {
            errorText = null;
            
            // явно определенный класс маршрута в параметрах
            string routeClassStringValue;
            var attributeCaption = Helpers.AttributesAliases[RDMetadataGuids.BaseClass.RefName];
            if (OtherParams.TryGetValue(attributeCaption, out routeClassStringValue))
            {
                // подмена заменителей, на реальные знаки
                routeClassStringValue = routeClassStringValue.WithoutStubs();
                RouteClassInfo routeClassInfo;
                if (Helpers.RouteClassInfos.Value.TryGetValue(routeClassStringValue, out routeClassInfo))
                {
                    ClassInfo = routeClassInfo;
                }
                else
                {
                    errorText = string.Format("Элемент {0} содержит точное указание типа маршрута {1}, " +
                        "которое не удалось найти в системе", equipmentClassInfo, routeClassStringValue);
                    return false;
                }
                OtherParams.Remove(attributeCaption);
            }

            // проверим может ли класс оборудования принимать указанный класс маршрута
            HashSet<RouteClassInfo> routeClassInfos;
            if (equipmentClassInfo != null &&
                ((routeClassInfos = Helpers.SupportedRouteClassInfosByEquipment.Get(equipmentClassInfo)) == null ||
                !routeClassInfos.Contains(ClassInfo)))
            {
                errorText = string.Format("Элемент {0} не поддерживает тип маршрута {1}", equipmentClassInfo, ClassInfo);
                return false;
            }

            // обработать другие параметры и преобразовать их в значения в зависимости от типа
            var parser = new Helpers.ClassAttributesParser(ClassInfo);
            return
                parser.Parse(OtherParams, Params, out errorText) &&
                parser.CheckRequired(Params, null, out errorText);
        }

        /// <summary>
        /// Заполнить параметры экземпляра
        /// </summary>
        public void Assign(Route route)
        {
            new Helpers.ClassAttributesParser(ClassInfo).Assign(route, Params);
        }

        /// <summary>
        /// Сравнение
        /// </summary>
        public bool Equals(Route other)
        {
            // UPD: атрибут приоритет не участвует в сравнении маршрутов
            return new Helpers.ClassAttributesParser(ClassInfo).Equals(other, Params, new[] { Route.InstanceAttributes.Priority });
        }

        /// <summary>
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int GetHashCode(ImportRoute obj)
        {
            return ClassInfo.GetHashCode();
        }
    }

    /// <summary>
    /// Компоратор для подсказки FIAS
    /// </summary>
    public class FiasSuggectionDataComparer : IEqualityComparer<FiasSuggestionData>
    {
        bool IEqualityComparer<FiasSuggestionData>.Equals(FiasSuggestionData x, FiasSuggestionData y)
        {
            return x != null && y != null &&
                x.Region == y.Region &&
                x.CityDistrict == y.CityDistrict &&
                x.Area == y.Area &&
                x.City == y.City &&
                x.Settlement == y.Settlement &&
                x.Street == y.Street &&
                x.House == y.House &&
                x.Block == y.Block;
        }

        int IEqualityComparer<FiasSuggestionData>.GetHashCode(FiasSuggestionData obj)
        {
            return new
            {
                obj.Region,
                obj.CityDistrict,
                obj.Area,
                obj.City,
                obj.Settlement,
                obj.Street,
                obj.House,
                obj.Block
            }.GetHashCode();
        }
    }

    /// <summary>
    /// Триггер, отслеживающий модификации пользовательских типов оборудования
    /// </summary>
    internal class TriggerEquipmentClass : ClassInstancesAndAllAttributesWatchTrigger<Equipment>
    {
        private readonly Action _resetAction;

        /// <summary>
        /// Удалили пользовательский класс
        /// </summary>
        /// <param name="ancestor"></param>
        protected override void InternalNotifyAfterClassRemove(RDClass? ancestor)
        {
            _resetAction();
        }

        /// <summary>
        /// Добавили пользовательский класс
        /// </summary>
        /// <param name="value"></param>
        protected override void InternalNotifyAfterClassCreate(RDClass? value)
        {
            _resetAction();
        }

        protected override Guid[]? EnumAttributesToWatch()
        {
            return EquipmentClassInfo.Get()
                .GetOwnAttributes(ClassAreaOfAction.ForClass, InheritanceLevelKind.CurrentLevelAncestorsAndDescendants, true)
                .Select(x => x.RefName)
                .ToArray();
        }

        protected override void OnAfterAppendAttributeValue(RDAttributeValue attributeValue)
        {
            _resetAction();
        }

        protected override void OnAfterRemoveAttributeValue(RDCustomAttributeValues attributeValues)
        {
            _resetAction();
        }

        protected override void OnAfterUpdateAttributeValue(RDAttributeValue attributeValue)
        {
            _resetAction();
        }

        public TriggerEquipmentClass(Action resetAction)
        {
            _resetAction = resetAction;
        }
    }

    #endregion Helpers
}