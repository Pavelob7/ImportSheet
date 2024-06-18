using System;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Взаимодействие с Асинхронным заданием (заглушка серверного интерфейса IAsyncTaskWrapperContextStub)
    /// </summary>
    public class AsyncTaskWrapperContextStub
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

        /// <summary>
        /// Уведомление о прогрессе выполнения
        /// </summary>
        /// <param name="value"></param>
        public void NotifyPercent(double value)
        {
        }

        /// <summary>
        /// Обозначить виртуальные границы процентного диапазона, внутри которого процент будет меняться от 0 до 100, 
        /// а на деле будет заполнять текущий переданный диапазон.
        /// Разрешается любой уровень вложенности таких относительных границ
        /// Вызов Dispose() возвращенного объекта закрывает область действия переданных относительных границ
        /// </summary>
        /// <param name="lowPercents"></param>
        /// <param name="highPercents"></param>
        /// <returns></returns>
        public IDisposable CreateRelativePercentBorders(double lowPercents, double highPercents)
        {
            return new DisposableDummy();
        }
    }

    /// <summary>
    /// Помощник для асинхронных заданий (заглушка серверного класса)
    /// </summary>
    public static class AsyncTasksHelper
    {
        /// <summary>
        /// Контекст асинхронного задания текущего потока
        /// </summary>
        public static AsyncTaskWrapperContextStub PublishedTaskContext = new AsyncTaskWrapperContextStub();
    }
}
