using System;
using System.Collections.Generic;
using System.Linq;
using ObjStudioClasses;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Средства отслеживания уникальности серийных номеров оборудования 
    /// </summary>
    public class EquipmentSerialNumbersManager
    {
        /// <summary>
        /// Единственный экземпляр
        /// </summary>
        private static readonly EquipmentSerialNumbersManager _current = new EquipmentSerialNumbersManager();

        /// <summary>
        /// Единственный экземпляр
        /// </summary>
        public static EquipmentSerialNumbersManager Current
        {
            get
            {
                return _current;
            }
        }
        
        /// <summary>
        /// Найти оборудование по серийному номеру
        /// </summary>
        /// <param name="equipmentClass"></param>
        /// <param name="serialNumber"></param>
        /// <param name="callbackFilter"></param>
        /// <returns></returns>
        public IEnumerable<Equipment> FindEquipmentBySerialNumber(EquipmentClassInfo equipmentClass, 
            string serialNumber, Func<IEnumerable<Equipment>, IEnumerable<Equipment>> callbackFilter)
        {
            var equipmentHash = equipmentClass.GetInstances()
                .OfType<Equipment>()
                .Where(x => Equals(x.AttributeSerialNumber, serialNumber));
            return callbackFilter == null ? equipmentHash : callbackFilter(equipmentHash);
        }
    }
}