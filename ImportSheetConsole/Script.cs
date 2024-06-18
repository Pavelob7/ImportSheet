using System;
using System.Diagnostics;
using System.IO;
using BridgeTypes;
using CSClient;
using GemBox.Spreadsheet;
using ImportSheetConsole.Global;
using ImportSheetConsole.Properties;
using ObjStudioClasses;
using ObjStudioClasses.Internal;

namespace ImportSheetConsole
{
    /// <summary>
    /// Базовый класс сценария
    /// </summary>
    internal abstract class Script : IServiceLoggerImplementation, IScriptImpl
    {
        /// <summary>
        /// Вывод в журнал сообщения
        /// </summary>
        /// <param name="messageKind"></param>
        /// <param name="messageText"></param>
        /// <param name="formatArgs"></param>
        public void AddLogMessage(DiagnosticMessageKindData messageKind, string messageText, params object[] formatArgs)
        {
            Logger.AddInfo(string.Format(messageText, formatArgs));
        }

        /// <summary>
        /// Вывод в журнал сообщения
        /// </summary>
        /// <param name="messageText"></param>
        /// <param name="formatArgs"></param>
        public void AddLogInfo(string messageText, params object[] formatArgs)
        {
            AddLogMessage(DiagnosticMessageKindData.Information, messageText, formatArgs);
        }

        /// <summary>
        /// Вывод в журнал сообщения
        /// </summary>
        /// <param name="messageText"></param>
        /// <param name="formatArgs"></param>
        public void AddLogDebug(string messageText, params object[] formatArgs)
        {
            AddLogMessage(DiagnosticMessageKindData.Debug, messageText, formatArgs);
        }

        /// <summary>
        /// Книга опросного листа
        /// </summary>
        public ExcelFile WorkbookNonExcel { get; private set; }
        
        /// <summary>
        /// Конструктор
        /// </summary>
        protected Script()
        {
            var connect = ControlServiceClient.CreateConnection(
                Settings.Default.Login,
                Settings.Default.Password,
                new BridgeConnectionParameters(Settings.Default.Host, Settings.Default.Port));
            ConnectInterface.AssignThreadAccessor(connect.CreateAccessor());
            ConnectInterface.AssignLoggerImplementation(this);
            ScriptImpl.Instance = this;
        }

        /// <summary>
        /// Метод выполнения
        /// </summary>
        /// <returns></returns>
        public void Execute(string fileName)
        {
            try
            {
                SpreadsheetInfo.SetLicense("SN-2020Dec23-2eruvV6RS4r9X9Fllv15nZDnsAdSJaWVDAEv5QAkiLINH0cqUbOPtPb7wflFeIncdFfKNm/pD2yQdcM357j9ePVOEaA==A"); // 4.5

                WorkbookNonExcel = ExcelFile.Load(fileName, LoadOptions.XlsxDefault);

                // проверка
                var checkResult = Check();
                if (checkResult.TotalErrorsCount > 0 || checkResult.TotalWarningsCount > 0)
                    return;

                // импорт
                Import();
            }
            finally
            {
                var checkedFilename = Path.Combine(
                    Path.GetDirectoryName(fileName) ?? string.Empty,
                    Path.GetFileName(fileName) + "_checked" + Path.GetExtension(fileName));

                WorkbookNonExcel.Save(checkedFilename);
                Process.Start(checkedFilename);
            }
        }

        /// <summary>
        /// Выполнить проверку опросного листа (метод вызываемый по кнопке "Проверить" в Веб)
        /// </summary>
        /// <returns></returns>
        protected abstract ImportSheetCheckResultData Check();

        /// <summary>
        /// Выполнить проверку опросного листа (метод вызываемый по кнопке "Импортировать" в Веб)
        /// </summary>
        /// <returns></returns>
        protected abstract ImportSheetProcessedResultData Import();

        /// <inheritdoc />
        /// <summary>
        /// </summary>
        /// <param name="kind"></param>
        /// <param name="message"></param>
        public void NotifyMessage(ServiceLoggerMessageKind kind, string message)
        {
            Console.WriteLine(message);
        }
    }
}