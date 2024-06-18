using ObjStudioClasses;
using JetBrains.Annotations;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Заглушка серверных средств работы с контекстом безопасности
    /// </summary>
    public static class RDClassesAndInstances
    {
        /// <summary>
        /// Заглушка подсистемы безопасности
        /// </summary>
        public class SecurityManagerStub
        {
            /// <summary>
            /// Временно заместить текущий маркер безопасности системным (контекст сменяется на ControlService или указаный)
            /// </summary>
            public void PushSecurityContext()
            {
                // Ничего не делаем
            }

            /// <summary>
            /// Восстановить контекст безопасности предыдущим
            /// </summary>
            public void PopSecurityContext()
            {
                // Ничего не делаем
            }

            /// <summary>
            /// Текущий персонализированный источник (описывает пару "Пользователь-сервис")
            /// </summary>
            public RDInstance? CurrentSource { get; set; }
        }

        /// <summary>
        /// Заглушка
        /// </summary>
        [NotNull]
        public static readonly SecurityManagerStub SecurityManager = new SecurityManagerStub();
    }
}