using ObjStudioClasses;
using System;

namespace ConsoleUsage.UserScripts.ImportSheets.Global
{
    /// <summary>
    /// Заглушка сервиса триггеров
    /// </summary>
    internal class CSServiceTriggers
    {
        public class CustomTrigger
        {
        }

        public abstract class ClassInstancesAndAllAttributesWatchTrigger<T> : CustomTrigger
        {
            protected abstract Guid[]? EnumAttributesToWatch();
            protected abstract void InternalNotifyAfterClassRemove(RDClass? ancestor);
            protected abstract void InternalNotifyAfterClassCreate(RDClass? value);
            protected abstract void OnAfterAppendAttributeValue(RDAttributeValue attributeValue);
            protected abstract void OnAfterRemoveAttributeValue(RDCustomAttributeValues attributeValues);
            protected abstract void OnAfterUpdateAttributeValue(RDAttributeValue attributeValue);
        }

        public class Implementation
        {
            public static void RegisterExternalTrigger(CustomTrigger trigger)
            {
            }
        }
    }
}