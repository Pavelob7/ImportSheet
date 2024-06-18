using GemBox.Spreadsheet;
using ObjStudioClasses;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Инфраструктура сценария импорта опросного листа на сервере
    /// </summary>
    public interface IScriptImpl
    {
        /// <summary>
        /// Вывод в журнал сообщения
        /// </summary>
        /// <param name="messageKind"></param>
        /// <param name="messageText"></param>
        /// <param name="formatArgs"></param>
        void AddLogMessage(DiagnosticMessageKindData messageKind, string messageText, params object[] formatArgs);

        /// <summary>
        /// Вывод в журнал сообщения
        /// </summary>
        /// <param name="messageText"></param>
        /// <param name="formatArgs"></param>
        void AddLogInfo(string messageText, params object[] formatArgs);

        /// <summary>
        /// Вывод в журнал сообщения
        /// </summary>
        /// <param name="messageText"></param>
        /// <param name="formatArgs"></param>
        void AddLogDebug(string messageText, params object[] formatArgs);

        /// <summary>
        /// Книга опросного листа
        /// </summary>
        ExcelFile WorkbookNonExcel { get; }
    }

    /// <summary>
    /// Класс описания инфраструктуры скрипта сервера
    /// </summary>
    public static class ScriptImpl
    {
        /// <summary>
        /// Экземпляр
        /// </summary>
        public static IScriptImpl Instance;
    }
}