using System;
using System.Collections.Generic;
using ObjStudioClasses;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Менеджер
    /// </summary>
    public class PreloadManager
    {
        private class DisposableDummy : IDisposable
        {
            /// <summary>
            /// С бездействующим методом
            /// </summary>
            void IDisposable.Dispose()
            {

            }
        }

        public static PreloadManager Current = new PreloadManager();

        public IDisposable RegisterCache(Func<DataSourceReceiveDataBatchCache> action)
        {
            return new DisposableDummy();
        }
    }

    /// <summary>
    /// Кэш ReceiveData
    /// </summary>
    public class DataSourceReceiveDataBatchCache
    {
        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="sources"></param>
        /// <param name="parameters"></param>
        /// <param name="interval"></param>
        public DataSourceReceiveDataBatchCache(IEnumerable<IDataSource> sources, IEnumerable<Parameter> parameters, DayIntervalData interval)
        {
        }
    }
}