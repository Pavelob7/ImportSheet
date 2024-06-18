using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GemBox.Spreadsheet;
using ImportSheetConsole.Global;
using ObjStudioClasses;

namespace ImportSheetConsole.DefaultScheme
{
    #region CheckHelper
    /// <summary>
    /// Утилиты валидации (внутренние)
    /// </summary>
    internal static class CheckHelper
    {
        // проверка листа
        public static Dictionary<int, object> CheckMeterPointsRow(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var values = new Dictionary<int, object>();
            values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;
            
            for (var i = Sheet.COL_NUM + 1; i <= MeterPoints.COL_MAX_NUMBER; i++)
                values[i] = null;

            var possibleId = Helper.TryGetGuidIdentifier(sheet.Cells[row.Index, Sheet.COL_NUM].Value);
            var existingEntity = possibleId.HasValue ? MeterPoint.Find(possibleId.Value) : null;
            if (possibleId.HasValue && existingEntity == null)
                sheet.Cells[row.Index, Sheet.COL_NUM].SetError(result, string.Format("ТУ с идентификатором {0} не найдена", possibleId.Value));
                
            // Все, что касается расположения ТУ в классификаторах, подлежит рассмотрению (что при проверке, что при импорте)
            // только в случае, если идентификатор ТУ не указан явно
            // Если же ТУ указан, то все это игнорируется
            if (existingEntity == null)
            {
                // ПЭС
                values[MeterPoints.COL_PES] =
                    sheet.Cells[row.Index, MeterPoints.COL_PES].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                // РЭС
                values[MeterPoints.COL_RES] =
                    sheet.Cells[row.Index, MeterPoints.COL_RES].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (values[MeterPoints.COL_RES] != null)
                {
                    // ПЭС
                    if (values[MeterPoints.COL_PES] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_PES].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                }

                // ПС

                // ПС
                values[MeterPoints.COL_SUBSTATION_SUBSTATION] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_SUBSTATION].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                // Уровень напряжения ПС
                values[MeterPoints.COL_SUBSTATION_VOLTAGE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_VOLTAGE].GetWithCheck(result, Helpers.GetVoltage, CheckHelpers.DefaultCheck);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (values[MeterPoints.COL_SUBSTATION_VOLTAGE] != null)
                {
                    // ПС
                    if (values[MeterPoints.COL_SUBSTATION_SUBSTATION] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_SUBSTATION].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                }

                // ПС - Высокая сторона

                var substationHiVoltage = new List<object>();
                // Уровень напряжения РУ
                values[MeterPoints.COL_SUBSTATION_HI_VOLTAGE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_HI_VOLTAGE].GetWithCheck(result, Helpers.GetVoltage, CheckHelpers.DefaultCheck);
                substationHiVoltage.Add(values[MeterPoints.COL_SUBSTATION_HI_VOLTAGE]);
                // Секция шин
                values[MeterPoints.COL_SUBSTATION_HI_BUSBAR] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_HI_BUSBAR].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                substationHiVoltage.Add(values[MeterPoints.COL_SUBSTATION_HI_BUSBAR]);
                // Ячейка
                values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                substationHiVoltage.Add(values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME]);
                // Тип ячейки
                values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_TYPE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_HI_CUBICLE_TYPE].GetWithCheck(result, Helpers.GetCubicleClass, CheckHelpers.DefaultCheck);
                substationHiVoltage.Add(values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_TYPE]);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (substationHiVoltage.Any(x => x != null))
                {
                    // Уровнень напрядения РУ
                    if (values[MeterPoints.COL_SUBSTATION_HI_VOLTAGE] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_HI_VOLTAGE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // Ячейка
                    if (values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                }

                // ПС - Низкая сторона

                var substationLoVoltage = new List<object>();
                // Уровень напряжения РУ				
                values[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_VOLTAGE].GetWithCheck(result, Helpers.GetVoltage, CheckHelpers.DefaultCheck);
                substationLoVoltage.Add(values[MeterPoints.COL_SUBSTATION_LO_VOLTAGE]);
                // Секция шин
                values[MeterPoints.COL_SUBSTATION_LO_BUSBAR] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_BUSBAR].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                substationLoVoltage.Add(values[MeterPoints.COL_SUBSTATION_LO_BUSBAR]);
                // Ячейка
                values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                substationLoVoltage.Add(values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME]);
                // Тип ячейки
                values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_TYPE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_CUBICLE_TYPE].GetWithCheck(result, Helpers.GetCubicleClass, CheckHelpers.DefaultCheck);
                substationLoVoltage.Add(values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_TYPE]);
                // Линия/фидер
                values[MeterPoints.COL_SUBSTATION_LO_POWERLINE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_POWERLINE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                substationLoVoltage.Add(values[MeterPoints.COL_SUBSTATION_LO_POWERLINE]);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (substationLoVoltage.Any(x => x != null))
                {
                    // Уровнень напряжения РУ
                    if (values[MeterPoints.COL_SUBSTATION_LO_VOLTAGE] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_VOLTAGE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // Ячейка
                    if (values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // применимость линии к типу ячейки
                    CubicleClassInfo cubicleClassInfo;
                    if ((cubicleClassInfo = (CubicleClassInfo)values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_TYPE]) != null &&
                        values[MeterPoints.COL_SUBSTATION_LO_POWERLINE] != null)
                    {
                        var sourceClass = PowerLineClassInfo.Get();
                        if (!sourceClass.IsAppliedToClassifierAttributes(cubicleClassInfo))
                            sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_POWERLINE].SetError(result, 
                                string.Format(CheckHelpers.NOT_APPLIED_CELLVALUE_ERROR, sourceClass, cubicleClassInfo));
                    }
                }

                // проверить обязательные поля целиком ПС, если заполнено хотя бы одна ячейка
                if (substationHiVoltage.Any(x => x != null) || substationLoVoltage.Any(x => x != null))
                {
                    // ПС
                    if (values[MeterPoints.COL_SUBSTATION_SUBSTATION] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_SUBSTATION].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // Уровнень напряжения ПС
                    if (values[MeterPoints.COL_SUBSTATION_VOLTAGE] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_VOLTAGE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                }

                // РП

                var switchgearSubstation = new List<object>();
                // РП
                values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] =
                    sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                //switchgearSubstation.Add(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION]);
                // Секция шин
                values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR] =
                    sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                switchgearSubstation.Add(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR]);
                // Ячейка, входящая от ПС
                values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] =
                    sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                switchgearSubstation.Add(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN]);
                // Ячейка, отходящая в ТП
                values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] =
                    sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                switchgearSubstation.Add(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT]);
                // Линия/фидер
                values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] =
                    sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                switchgearSubstation.Add(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE]);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (switchgearSubstation.Any(x => x != null))
                {
                    // РП
                    if (values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // Ячейка, входящая от ПС
                    if (values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // Ячейка, отходящая в ТП
                    if (values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT] == null &&
                        values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] != null)
                        sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // ПС - Низкая сторона - Линия/фидер 
                    if (values[MeterPoints.COL_SUBSTATION_LO_POWERLINE] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_POWERLINE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                }

                // ТП - Высокая сторона

                var loSubstationHiVoltage = new List<object>();
                // ТП
                values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_SUBSTATION].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                //loSubstationHiVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION]);
                // Секция шин
                values[MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                loSubstationHiVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR]);
                // Ячейка
                values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                loSubstationHiVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME]);
                // Тип ячейки
                values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_TYPE] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_TYPE].GetWithCheck(result, Helpers.GetCubicleClass, CheckHelpers.DefaultCheck);
                loSubstationHiVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_TYPE]);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (loSubstationHiVoltage.Any(x => x != null))
                {
                    // ТП
                    if (values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_SUBSTATION].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // Ячейка
                    if (values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // РП - Линия/фидер
                    if (switchgearSubstation.Any(x => x != null) &&
                        values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SWITCHGEAR_SUBSTATION_POWERLINE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // или ПС - Низкая сторона - Линия/фидер, если РП отсутствует
                    else if (values[MeterPoints.COL_SUBSTATION_LO_POWERLINE] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_SUBSTATION_LO_POWERLINE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                }

                // ТП - Низкая сторона

                var loSubstationLoVoltage = new List<object>();
                // Секция шин
                values[MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                loSubstationLoVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR]);
                // Ячейка
                values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                loSubstationLoVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME]);
                // Тип ячейки
                values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_TYPE] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_TYPE].GetWithCheck(result, Helpers.GetCubicleClass, CheckHelpers.DefaultCheck);
                loSubstationLoVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_TYPE]);
                // Линия/фидер
                values[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] =
                    sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
                loSubstationLoVoltage.Add(values[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE]);
                // проверить обязательные поля, если заполнено хотя бы одна ячейка из группы
                if (loSubstationLoVoltage.Any(x => x != null))
                {
                    // Ячейка
                    if (values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME] == null)
                        sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                    // применимость линии к типу ячейки
                    CubicleClassInfo cubicleClassInfo;
                    if ((cubicleClassInfo = (CubicleClassInfo)values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_TYPE]) != null &&
                        values[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] != null)
                    {
                        var sourceClass = PowerLineClassInfo.Get();
                        if (!sourceClass.IsAppliedToClassifierAttributes(cubicleClassInfo))
                            sheet.Cells[row.Index, MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE].SetError(result, 
                                string.Format(CheckHelpers.NOT_APPLIED_CELLVALUE_ERROR, sourceClass, cubicleClassInfo));
                    }
                }

                // Адрес

                // Адрес ФИАС
                values[MeterPoints.COL_ADDRESS_FIAS] =
                    sheet.Cells[row.Index, MeterPoints.COL_ADDRESS_FIAS].GetWithCheck(result, Helpers.GetAddressFias, CheckHelpers.CheckAddressFiasDefault);

                // Квартира
                values[MeterPoints.COL_ADDRESS_FLAT] =
                    sheet.Cells[row.Index, MeterPoints.COL_ADDRESS_FLAT].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);

                // или Ввод в МКД
                var flatName = values[MeterPoints.COL_ADDRESS_FLAT] as string;
                if (flatName != null &&
                    flatName.IndexOf(Helper.TENEMENT_HOUSE_POWERLINE_METERPOINT_NAME, StringComparison.InvariantCultureIgnoreCase) != -1 &&
                    values[MeterPoints.COL_LO_SUBSTATION_LO_POWERLINE] == null)
                {
                    sheet.Cells[row.Index, MeterPoints.COL_ADDRESS_FLAT].SetError(result,
                        "Для вводной ячейки МКД должен быть указан Линия/фидер низкой стороны ТП");
                }

                // или Помещение
            }

            // ПУ

            var meter = new List<object>();

            // Тип ПУ
            values[MeterPoints.COL_METER_TYPE] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_TYPE].GetWithCheck(result, Helpers.GetMeterClass, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_TYPE]);
            // Серийный номер
            values[MeterPoints.COL_METER_SERIALNUMBER] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_SERIALNUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_SERIALNUMBER]);
            // Дата выпуска
            values[MeterPoints.COL_METER_CREATEDATE] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_CREATEDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_CREATEDATE]);
            // Дата установки
            values[MeterPoints.COL_METER_INSTALLDATE] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_INSTALLDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_INSTALLDATE]);
            // Дата последней поверки
            values[MeterPoints.COL_METER_CHECKDATE] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_CHECKDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_CHECKDATE]);
            // Класс точности
            values[MeterPoints.COL_METER_ACCURACYCLASS] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_ACCURACYCLASS].GetWithCheck(result, Helpers.GetAccuracyRating, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_ACCURACYCLASS]);
            // Часовой пояс
            values[MeterPoints.COL_METER_TIMEZONE] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_TIMEZONE].GetWithCheck(result, Helpers.GetTimeZone, CheckHelpers.DefaultCheck);
            meter.Add(values[MeterPoints.COL_METER_TIMEZONE]);
            // проверить обязательные поля
            if (meter.Any(x => x != null))
            {
                // Тип ПУ
                if (values[MeterPoints.COL_METER_TYPE] == null)
                    sheet.Cells[row.Index, MeterPoints.COL_METER_TYPE].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);
                // Серийный номер
                if (values[MeterPoints.COL_METER_SERIALNUMBER] == null)
                    sheet.Cells[row.Index, MeterPoints.COL_METER_SERIALNUMBER].SetError(result, CheckHelpers.EMPTY_CELLVALUE_ERROR);

                // наличие квартиры в адресе
                var address = values[MeterPoints.COL_ADDRESS_FIAS] as FiasSuggestionData;
                if (address != null && string.IsNullOrEmpty(address.House))
                    sheet.Cells[row.Index, MeterPoints.COL_ADDRESS_FIAS].SetError(result, "Адрес не содержит \"Номер дома\", необходимый для создания ТУ");
            }
            // Если ТУ уже существует, то ПУ д.б строго привязан к ТУ
            var meterClass = values[MeterPoints.COL_METER_TYPE] as ElectricityMeterClassInfo;
            var existingSerial = values[MeterPoints.COL_METER_SERIALNUMBER] as string;

            if (existingEntity != null)
            {
                var existingEntityMeter = existingEntity.AttributeElectricityMeter;
                if (existingEntityMeter == null && meterClass != null)
                    sheet.Cells[row.Index, MeterPoints.COL_METER_SERIALNUMBER].SetError(result, "Недопустима привязка другого ПУ к существующей ТУ");
                else if (existingEntityMeter != null && meterClass == null)
                    sheet.Cells[row.Index, MeterPoints.COL_METER_SERIALNUMBER].SetError(result, "Недопустима привязка другого ПУ к существующей ТУ");
                else if (existingEntityMeter != null &&
                    (!Equals(existingEntityMeter.AttributeSerialNumber, existingSerial) || existingEntityMeter.Class != meterClass))
                    sheet.Cells[row.Index, MeterPoints.COL_METER_SERIALNUMBER].SetError(result, "Недопустима привязка другого ПУ к существующей ТУ");
            }
            else
            {
                // иначе приверим ПУ на наличие его в другой ТУ
                if (meterClass != null && existingSerial != null)
                {
                    var existingEntityMeter = EquipmentSerialNumbersManager.Current.FindEquipmentBySerialNumber(meterClass, existingSerial, null)
                        .OfType<ElectricityMeter>()
                        .FirstOrDefault();
                    var meterPointPlacement = existingEntityMeter?.AttributeMeterPointPlacement;
                    if (meterPointPlacement != null)
                    {
                        string roomName = null;

                        // проверить наименование ТУ, без иерархии
                        var newMeterPointName = existingEntityMeter.Caption;
                        if (values[MeterPoints.COL_ADDRESS_FIAS] is FiasSuggestionData &&
                            values[MeterPoints.COL_ADDRESS_FLAT] is string flatName &&
                            !string.IsNullOrEmpty(flatName))
                        {
                            if (flatName.IndexOf(Helper.ROOM_PRIFIX_NAME, StringComparison.InvariantCultureIgnoreCase) != -1)
                                roomName = flatName;
                            else
                                newMeterPointName = string.Format("{0}, {1}", flatName, newMeterPointName);
                        }
                        var f = string.Equals(meterPointPlacement.AttributeCaption, newMeterPointName, StringComparison.InvariantCultureIgnoreCase);

                        // проверить иерархию, если необходимо
                        if (f)
                        {
                            static bool Check(object? schemaName, IEntityWithCaption? localInstance)
                            {
                                return schemaName as string == localInstance?.AttributeCaption;
                            }

                            var topology = meterPointPlacement.GetElectricityTopology();
                            // география
                            if (values[MeterPoints.COL_ADDRESS_FIAS] is FiasSuggestionData fias && topology.GeoPlacement?.GeoDivisions != null)
                            {
                                var divisions = topology.GeoPlacement.GeoDivisions;
                                f =
                                    Check(fias.Region, divisions.OfType<SubjectRF>().FirstOrDefault()) &&
                                    Check(fias.Area, divisions.OfType<District>().FirstOrDefault()) &&
                                    Check(fias.Street, divisions.OfType<Street>().FirstOrDefault()) &&
                                    Check(fias.House, divisions.OfType<CustomBuilding>().FirstOrDefault()) &&
                                    // учесть помещение
                                    (roomName == null || Check(roomName, divisions.OfType<Room>().FirstOrDefault()));
                            }
                            else if (topology.PowerNetworkConfiguration != null)
                            {
                                var cfg = topology.PowerNetworkConfiguration;

                                var pes = cfg.NetworkDivisions?.OfType<ElectricalNetworksSubsidiary>().FirstOrDefault();
                                var res = cfg.NetworkDivisions?.OfType<ElectricalNetworksDistrict>().FirstOrDefault();

                                f =
                                    Check(values[MeterPoints.COL_PES], pes) &&
                                    Check(values[MeterPoints.COL_RES], res) &&
                                    Check(ImportHelper.ENERGY_CLASSIFIER_CAPTION, pes?.RelationsClassifierOfMeterPointsByEnergyEntitiesAttributeEnergyManagementItems.FirstOrDefault()) &&
                                    Check(values[MeterPoints.COL_SUBSTATION_SUBSTATION], (cfg as PowerNetworkHighVoltageSubstationHighSideData)?.HighVoltageSubstation);

                                // ПС
                                if (f)
                                {
                                    f =
                                        //Check(values[MeterPoints.COL_SUBSTATION_HI_VOLTAGE], (cfg as PowerNetworkHighVoltageSubstationHighSideCubicleData)?.HighVoltageSubstationHighSideSwitchGear) &&
                                        Check(values[MeterPoints.COL_SUBSTATION_HI_BUSBAR], (cfg as PowerNetworkHighVoltageSubstationHighSideCubicleData)?.HighVoltageSubstationHighSideBusbar) &&
                                        Check(values[MeterPoints.COL_SUBSTATION_HI_CUBICLE_NAME], (cfg as PowerNetworkHighVoltageSubstationHighSideCubicleData)?.HighVoltageSubstationHighSideMeterPointCubicle) &&
                                        //Check(values[MeterPoints.COL_SUBSTATION_LO_VOLTAGE], (cfg as PowerNetworkHighVoltageSubstationLowSideData)?.HighVoltageSubstationLowSideSwitchGear) &&
                                        Check(values[MeterPoints.COL_SUBSTATION_LO_BUSBAR], (cfg as PowerNetworkHighVoltageSubstationLowSideData)?.HighVoltageSubstationLowSideBusbar) &&
                                        Check(values[MeterPoints.COL_SUBSTATION_LO_CUBICLE_NAME],
                                            (cfg as PowerNetworkHighVoltageSubstationLowSideCubicleData)?.HighVoltageSubstationLowSideMeterPointCubicle ??
                                            (cfg as PowerNetworkHighVoltageSubstationLowSideConnectionData)?.HighVoltageSubstationLowSidePowerLineCubicle);
                                }

                                // РП
                                if (f)
                                {
                                    f =
                                        Check(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_SUBSTATION], (cfg as PowerNetworkSwitchgearSubstationData)?.SwitchgearSubstation) &&
                                        Check(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_BUSBAR], (cfg as PowerNetworkSwitchgearSubstationData)?.SwitchgearSubstationLeftSideBusbar) &&
                                        Check(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_IN], (cfg as PowerNetworkSwitchgearSubstationData)?.SwitchgearSubstationLeftSideCubicle) &&
                                        Check(values[MeterPoints.COL_SWITCHGEAR_SUBSTATION_CUBICLE_NAME_OUT],
                                            (cfg as PowerNetworkSwitchgearSubstationCubicleData)?.SwitchgearSubstationMeterPointCubicle ??
                                            (cfg as PowerNetworkSwitchgearSubstationConnectionData)?.SwitchgearSubstationRightSidePowerLineCubicle);
                                }

                                // ТП
                                if (f)
                                {
                                    f =
                                        Check(values[MeterPoints.COL_LO_SUBSTATION_SUBSTATION],
                                            // через РП
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationHighSideData)?.LowVoltageSubstation ??
                                            // без РП
                                            (cfg as PowerNetworkLowVoltageSubstationHighSideData)?.LowVoltageSubstation) &&
                                        Check(values[MeterPoints.COL_LO_SUBSTATION_HI_BUSBAR],
                                            // через РП
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationHighSideData)?.LowVoltageSubstationHighSideBusbar ??
                                            // без РП
                                            (cfg as PowerNetworkLowVoltageSubstationHighSideData)?.LowVoltageSubstationHighSideBusbar) &&
                                        Check(values[MeterPoints.COL_LO_SUBSTATION_HI_CUBICLE_NAME],
                                            // через РП
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationHighSideCubicleData)?.LowVoltageSubstationHighSideMeterPointCubicle ??
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationHighSideConnectionData)?.LowVoltageSubstationHighSideRightPowerLineCubicle ??
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationHighSideData)?.LowVoltageSubstationHighSidePowerLineCubicle ??
                                            // без РП
                                            (cfg as PowerNetworkLowVoltageSubstationHighSideCubicleData)?.LowVoltageSubstationHighSideMeterPointCubicle ??
                                            (cfg as PowerNetworkLowVoltageSubstationHighSideConnectionData)?.LowVoltageSubstationHighSideRightPowerLineCubicle ??
                                            (cfg as PowerNetworkLowVoltageSubstationHighSideData)?.LowVoltageSubstationHighSidePowerLineCubicle) &&
                                        Check(values[MeterPoints.COL_LO_SUBSTATION_LO_BUSBAR],
                                            // через РП
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationLowSideData)?.LowVoltageSubstationLowSideBusbar ??
                                            // без РП
                                            (cfg as PowerNetworkLowVoltageSubstationLowSideData)?.LowVoltageSubstationLowSideBusbar) &&
                                        Check(values[MeterPoints.COL_LO_SUBSTATION_LO_CUBICLE_NAME],
                                            // через РП
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationLowSideCubicleData)?.LowVoltageSubstationLowSideMeterPointCubicle ??
                                            (cfg as PowerNetworkLowVoltageSubstationViaSwitchgearSubstationLowSideConnectionData)?.LowVoltageSubstationLowSidePowerLineCubicle ??
                                            // без РП
                                            (cfg as PowerNetworkLowVoltageSubstationLowSideCubicleData)?.LowVoltageSubstationLowSideMeterPointCubicle ??
                                            (cfg as PowerNetworkLowVoltageSubstationLowSideConnectionData)?.LowVoltageSubstationLowSidePowerLineCubicle);
                                }
                            }
                        }
                        if (!f)
                            sheet.Cells[row.Index, MeterPoints.COL_METER_SERIALNUMBER].SetError(result, $"Данный ПУ уже привязан к другой ТУ - \"{meterPointPlacement.Caption}\"");
                    }
                }
            }

            // проверить модель ПУ, относительно соответствия атрибутам модели
            if (!Equals(meterClass?.AttributeAccuracyClass_this_class, values[MeterPoints.COL_METER_ACCURACYCLASS]))
                sheet.Cells[row.Index, MeterPoints.COL_METER_ACCURACYCLASS].SetError(result, "Класс точности не соответствует указанной модели ПУ");

            // Трансформаторы

            // Ктт
            values[MeterPoints.COL_TRANS_KTT] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_KTT].GetWithCheck(result, Helpers.GetDouble, CheckHelpers.DefaultCheck);
            // Ктн
            values[MeterPoints.COL_TRANS_KTN] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_KTN].GetWithCheck(result, Helpers.GetDouble, CheckHelpers.DefaultCheck);

            // ТТ, фаза 1
            values[MeterPoints.COL_TRANS_CURRENT_PHASE1] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_CURRENT_PHASE1].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);
            // ТТ, фаза 2
            values[MeterPoints.COL_TRANS_CURRENT_PHASE2] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_CURRENT_PHASE2].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);
            // ТТ, фаза 3
            values[MeterPoints.COL_TRANS_CURRENT_PHASE3] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_CURRENT_PHASE3].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);

            // ТН, фаза 1
            values[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_VOLTAGE_PHASE1].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);
            // ТН, фаза 2
            values[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_VOLTAGE_PHASE2].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);
            // ТН, фаза 3
            values[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] =
                sheet.Cells[row.Index, MeterPoints.COL_TRANS_VOLTAGE_PHASE3].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);

            if ((values[MeterPoints.COL_TRANS_KTT] ?? values[MeterPoints.COL_TRANS_KTN]) != null &&
               ((values[MeterPoints.COL_TRANS_CURRENT_PHASE1] ?? values[MeterPoints.COL_TRANS_CURRENT_PHASE2] ?? values[MeterPoints.COL_TRANS_CURRENT_PHASE3]) != null ||
                (values[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] ?? values[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] ?? values[MeterPoints.COL_TRANS_VOLTAGE_PHASE3]) != null))
            {
                var errorText = "Одновременное указание Упрощенных и Расширенных настроек Ктт, Ктн не возможно";
                if (values[MeterPoints.COL_TRANS_KTT] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_KTT].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_KTN] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_KTN].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_CURRENT_PHASE1] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_CURRENT_PHASE1].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_CURRENT_PHASE2] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_CURRENT_PHASE2].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_CURRENT_PHASE3] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_CURRENT_PHASE3].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_VOLTAGE_PHASE1] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_VOLTAGE_PHASE1].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_VOLTAGE_PHASE2] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_VOLTAGE_PHASE2].SetError(result, errorText);
                if (values[MeterPoints.COL_TRANS_VOLTAGE_PHASE3] != null)
                    sheet.Cells[row.Index, MeterPoints.COL_TRANS_VOLTAGE_PHASE3].SetError(result, errorText);
            }

            // Связь с ПУ

            // Связной номер
            values[MeterPoints.COL_METER_NETWORKNUMBER] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_NETWORKNUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Пользователь
            values[MeterPoints.COL_METER_LOGIN] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_LOGIN].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Пароль
            values[MeterPoints.COL_METER_PASSWORD] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_PASSWORD].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Маршрут опроса
            values[MeterPoints.COL_METER_ROUTE] =
                sheet.Cells[row.Index, MeterPoints.COL_METER_ROUTE].GetWithCheck(result, Helpers.GetRoutes, CheckHelpers.DefaultCheck);
            //// Специальные параметры
            //var meterParams =
            //    sheet.Cells[row.Index, MeterPoints.COL_METER_PARAMS].GetWithCheck(result, Helpers.GetParams, CheckHelpers.DefaultCheck)
            //    ?? new Dictionary<string, string>();
            //// проверить параметры на соответствие классу ПУ
            //var meterClassInfo = values[MeterPoints.COL_METER_TYPE] as MeterEquipmentClassInfo;
            //if (meterClassInfo != null)
            //{
            //    var meterResultParams = new Dictionary<Guid, object>();
            //    string errorText;
            //    var parser = new Helpers.ClassAttributesParser(meterClassInfo);
            //    if (!parser.Parse(meterParams, meterResultParams, out errorText) || 
            //        !parser.CheckRequired(meterResultParams, MeterPoints.KnownMeterParams, out errorText))
            //    {
            //        sheet.Cells[row.Index, MeterPoints.COL_METER_PARAMS].SetError(result, errorText);
            //        meterResultParams = null;
            //    }
            //    values[MeterPoints.COL_METER_PARAMS] = meterResultParams;
            //}

            // УСПД

            values[MeterPoints.COL_RTU] =
                sheet.Cells[row.Index, MeterPoints.COL_RTU].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);
            
            // Абонент

            values[MeterPoints.COL_CONSUMER] =
                sheet.Cells[row.Index, MeterPoints.COL_CONSUMER].GetWithCheck(result, Helpers.GetEntityLink, CheckHelpers.DefaultCheck);

            // Область видимости
            values[MeterPoints.COL_ISOLATION_LEVEL] =
                sheet.Cells[row.Index, MeterPoints.COL_ISOLATION_LEVEL].GetWithCheck(result, Helpers.GetIsolationLevels, CheckHelpers.DefaultCheck);

            // Тариф
            values[MeterPoints.COL_TARIFF] =
                sheet.Cells[row.Index, MeterPoints.COL_TARIFF].GetWithCheck(result, Helpers.GetTariff, CheckHelpers.DefaultCheck);

            return values;
        }
        
        // проверка листа
        public static Dictionary<int, object> CheckRtus(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var possibleId = Helper.TryGetGuidIdentifier(sheet.Cells[row.Index, Sheet.COL_NUM].Value);
            var existingEntity = possibleId.HasValue ? ChannelizingEquipment.Find(possibleId.Value) : null;
            if (possibleId.HasValue && existingEntity == null)
                sheet.Cells[row.Index, Sheet.COL_NUM].SetError(result, string.Format("УСПД с идентификатором {0} не найден", possibleId.Value));
            
            var values = new Dictionary<int, object>();
            values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;
            // Место установки
            values[Rtus.COL_PLACEMENT] =
                sheet.Cells[row.Index, Rtus.COL_PLACEMENT].GetWithCheck(result, Helpers.GetStrings, CheckHelpers.DefaultCheck);
            // Тип УСПД
            values[Rtus.COL_TYPE] =
                sheet.Cells[row.Index, Rtus.COL_TYPE].GetWithCheck(result, Helpers.GetRtuClass, CheckHelpers.DefaultCheck, true);
            // Серийный номер
            values[Rtus.COL_SERIALNUMBER] =
                sheet.Cells[row.Index, Rtus.COL_SERIALNUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Дата выпуска
            values[Rtus.COL_CREATEDATE] =
                sheet.Cells[row.Index, Rtus.COL_CREATEDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Дата установки
            values[Rtus.COL_INSTALLDATE] =
                sheet.Cells[row.Index, Rtus.COL_INSTALLDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Дата последней поверки
            values[Rtus.COL_CHECKDATE] =
                sheet.Cells[row.Index, Rtus.COL_CHECKDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Часовой пояс
            values[Rtus.COL_TIMEZONE] =
                sheet.Cells[row.Index, Rtus.COL_TIMEZONE].GetWithCheck(result, Helpers.GetTimeZone, CheckHelpers.DefaultCheck);

            // Связь с УСПД

            // Связной номер
            values[Rtus.COL_NETWORKNUMBER] =
                sheet.Cells[row.Index, Rtus.COL_NETWORKNUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Пользователь
            values[Rtus.COL_LOGIN] =
                sheet.Cells[row.Index, Rtus.COL_LOGIN].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Пароль
            values[Rtus.COL_PASSWORD] =
                sheet.Cells[row.Index, Rtus.COL_PASSWORD].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Маршрут опроса
            values[Rtus.COL_ROUTE] =
                sheet.Cells[row.Index, Rtus.COL_ROUTE].GetWithCheck(result, Helpers.GetRoutes, CheckHelpers.DefaultCheck);
            //// Специальные параметры
            //var rtuParams =
            //    sheet.Cells[row.Index, Rtus.COL_PARAMS].GetWithCheck(result, Helpers.GetParams, CheckHelpers.DefaultCheck)
            //    ?? new Dictionary<string, string>();
            //// проверить параметры на соответствие классу
            //var classInfo = (ChannelizingEquipmentClassInfo)values[Rtus.COL_TYPE];
            //if (classInfo != null)
            //{
            //    var rtuResultParams = new Dictionary<Guid, object>();
            //    string errorText;
            //    var parser = new Helpers.ClassAttributesParser(classInfo);
            //    if (!parser.Parse(rtuParams, rtuResultParams, out errorText) ||
            //        !parser.CheckRequired(rtuResultParams, Rtus.KnownMeterParams, out errorText))
            //    {
            //        sheet.Cells[row.Index, Rtus.COL_PARAMS].SetError(result, errorText);
            //        rtuResultParams = null;
            //    }
            //    values[Rtus.COL_PARAMS] = rtuResultParams;
            //}

            return values;
        }

        // проверка листа
        public static Dictionary<int, object> CheckTransCurrents(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var possibleId = Helper.TryGetGuidIdentifier(sheet.Cells[row.Index, Sheet.COL_NUM].Value);
            var existingEntity = possibleId.HasValue ? CurrentTransformer.Find(possibleId.Value) : null;
            if (possibleId.HasValue && existingEntity == null)
                sheet.Cells[row.Index, Sheet.COL_NUM].SetError(result, string.Format("ТТ с идентификатором {0} не найден", possibleId.Value));

            var values = new Dictionary<int, object>();
            values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;

            // ТТ

            // Тип ТТ
            values[TransCurrents.COL_TYPE] =
                sheet.Cells[row.Index, TransCurrents.COL_TYPE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Серийный номер
            values[TransCurrents.COL_SERIALNUMBER] =
                sheet.Cells[row.Index, TransCurrents.COL_SERIALNUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Класс точности
            values[TransCurrents.COL_ACCURACYCLASS] =
                sheet.Cells[row.Index, TransCurrents.COL_ACCURACYCLASS].GetWithCheck(result, Helpers.GetAccuracyRating, CheckHelpers.DefaultCheck);
            // Дата выпуска
            values[TransCurrents.COL_CREATEDATE] =
                sheet.Cells[row.Index, TransCurrents.COL_CREATEDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Дата установки
            values[TransCurrents.COL_INSTALLDATE] =
                sheet.Cells[row.Index, TransCurrents.COL_INSTALLDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Дата последней поверки
            values[TransCurrents.COL_CHECKDATE] =
                sheet.Cells[row.Index, TransCurrents.COL_CHECKDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);

            // Параметры ТТ

            // I ном. перв, А
            values[TransCurrents.COL_PRIMARY_NOM] =
                sheet.Cells[row.Index, TransCurrents.COL_PRIMARY_NOM].GetWithCheck(result, Helpers.GetNominalCurrent, CheckHelpers.DefaultCheck);
            // I ном. втор, А
            values[TransCurrents.COL_SECONDARY_NOM] =
                sheet.Cells[row.Index, TransCurrents.COL_SECONDARY_NOM].GetWithCheck(result, Helpers.GetNominalMeasureCurrent, CheckHelpers.DefaultCheck);

            return values;
        }

        // проверка листа
        public static Dictionary<int, object> CheckTransVoltages(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var possibleId = Helper.TryGetGuidIdentifier(sheet.Cells[row.Index, Sheet.COL_NUM].Value);
            var existingEntity = possibleId.HasValue ? VoltageTransformer.Find(possibleId.Value) : null;
            if (possibleId.HasValue && existingEntity == null)
                sheet.Cells[row.Index, Sheet.COL_NUM].SetError(result, string.Format("ТН с идентификатором {0} не найден", possibleId.Value));

            var values = new Dictionary<int, object>();
            values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;

            // ТН

            // Тип ТН
            values[TransVoltages.COL_TYPE] =
                sheet.Cells[row.Index, TransVoltages.COL_TYPE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Серийный номер
            values[TransVoltages.COL_SERIALNUMBER] =
                sheet.Cells[row.Index, TransVoltages.COL_SERIALNUMBER].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Класс точности
            values[TransVoltages.COL_ACCURACYCLASS] =
                sheet.Cells[row.Index, TransVoltages.COL_ACCURACYCLASS].GetWithCheck(result, Helpers.GetAccuracyRating, CheckHelpers.DefaultCheck);
            // Дата выпуска
            values[TransVoltages.COL_CREATEDATE] =
                sheet.Cells[row.Index, TransVoltages.COL_CREATEDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Дата установки
            values[TransVoltages.COL_INSTALLDATE] =
                sheet.Cells[row.Index, TransVoltages.COL_INSTALLDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);
            // Дата последней поверки
            values[TransVoltages.COL_CHECKDATE] =
                sheet.Cells[row.Index, TransVoltages.COL_CHECKDATE].GetWithCheck(result, Helpers.GetDateTime, CheckHelpers.DefaultCheck);

            // Параметры ТН

            // U ном. перв, В
            values[TransVoltages.COL_PRIMARY_NOM] =
                sheet.Cells[row.Index, TransVoltages.COL_PRIMARY_NOM].GetWithCheck(result, Helpers.GetNominalVoltage, CheckHelpers.DefaultCheck);
            // U ном. втор, В
            values[TransVoltages.COL_SECONDARY_NOM] =
                sheet.Cells[row.Index, TransVoltages.COL_SECONDARY_NOM].GetWithCheck(result, Helpers.GetNominalMeasureVoltage, CheckHelpers.DefaultCheck);

            return values;
        }

        // проверка листа
        public static Dictionary<int, object> CheckNaturalPersons(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var possibleId = Helper.TryGetGuidIdentifier(sheet.Cells[row.Index, Sheet.COL_NUM].Value);
            var existingEntity = possibleId.HasValue ? NaturalPerson.Find(possibleId.Value) : null;
            if (possibleId.HasValue && existingEntity == null)
                sheet.Cells[row.Index, Sheet.COL_NUM].SetError(result, string.Format("ФЛ с идентификатором {0} не найдено", possibleId.Value));
            
            var values = new Dictionary<int, object>();
            values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;

            // ФИО

            // Фамилия
            values[NaturalPersons.COL_LASTNAME] =
                sheet.Cells[row.Index, NaturalPersons.COL_LASTNAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Имя
            values[NaturalPersons.COL_FIRSTNAME] =
                sheet.Cells[row.Index, NaturalPersons.COL_FIRSTNAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Отчество
            values[NaturalPersons.COL_MIDDLENAME] =
                sheet.Cells[row.Index, NaturalPersons.COL_MIDDLENAME].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);

            // Связь с абонентом

            // Эл.почта
            values[NaturalPersons.COL_EMAIL] =
                sheet.Cells[row.Index, NaturalPersons.COL_EMAIL].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Телефон
            values[NaturalPersons.COL_PHONE] =
                sheet.Cells[row.Index, NaturalPersons.COL_PHONE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);

            // Номер лицевого счета
            values[NaturalPersons.COL_CURRENTACCOUNT] =
                sheet.Cells[row.Index, NaturalPersons.COL_CURRENTACCOUNT].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);

            return values;
        }

        // проверка листа
        public static Dictionary<int, object> CheckLegalEntities(ImportSheetCheckResultData result, ExcelWorksheet sheet, ExcelRow row)
        {
            var possibleId = Helper.TryGetGuidIdentifier(sheet.Cells[row.Index, Sheet.COL_NUM].Value);
            var existingEntity = possibleId.HasValue ? LegalEntity.Find(possibleId.Value) : null;
            if (possibleId.HasValue && existingEntity == null)
                sheet.Cells[row.Index, Sheet.COL_NUM].SetError(result, string.Format("ЮЛ с идентификатором {0} не найдено", possibleId.Value));

            var values = new Dictionary<int, object>();
            values[Sheet.COL_NUM] = sheet.Cells[row.Index, Sheet.COL_NUM].Value;

            // Наименование
            values[LegalEntities.COL_CAPTION] =
                sheet.Cells[row.Index, LegalEntities.COL_CAPTION].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);
            // Юридический адрес
            values[LegalEntities.COL_ADDRESS] =
                sheet.Cells[row.Index, LegalEntities.COL_ADDRESS].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);

            // Связь с абонентом

            // Эл.почта
            values[LegalEntities.COL_EMAIL] =
                sheet.Cells[row.Index, LegalEntities.COL_EMAIL].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);
            // Телефон
            values[LegalEntities.COL_PHONE] =
                sheet.Cells[row.Index, LegalEntities.COL_PHONE].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck);

            // Номер лицевого счета
            values[LegalEntities.COL_CURRENTACCOUNT] =
                sheet.Cells[row.Index, LegalEntities.COL_CURRENTACCOUNT].GetWithCheck(result, Helpers.GetString, CheckHelpers.DefaultCheck, true);

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
    #endregion CheckHelper
}