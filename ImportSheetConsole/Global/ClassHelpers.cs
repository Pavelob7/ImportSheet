using System.Collections.Generic;
using System.Linq;
using ObjStudioClasses;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Утилиты для RDClass
    /// </summary>
    public static class ClassHelpers
    {
        /// <summary>
        /// Получить текущий класс и всех наследников этого класса на всю глубину
        /// </summary>
        /// <param name="rootClass"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetThisAndAllDescendants<T>(this T rootClass) where T : RDClass
        {
            yield return rootClass;
            foreach (var subClass in rootClass.GetDescendants().SelectMany(GetThisAndAllDescendants).Where(x => x != null))
                yield return (T)subClass;
        }

        /// <summary>
        /// Получить всех наследников класса
        /// </summary>
        /// <param name="rootClass"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetAllDescendants<T>(this T rootClass) where T : RDClass
        {
            return rootClass.GetThisAndAllDescendants().Skip(1);
        }

        /// <summary>
        /// Получить текущий класс и всех наследников этого класса на всю глубину
        /// </summary>
        /// <param name="rootInterface"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetThisAndAllDescendantsInf<T>(this T rootInterface) where T : RDInterface
        {
            yield return rootInterface;
            foreach (var subInterface in rootInterface.GetDescendants().SelectMany(GetThisAndAllDescendantsInf).Where(x => x != null))
                yield return (T)subInterface;
        }

        /// <summary>
        /// Получить всех наследников класса
        /// </summary>
        /// <param name="rootClass"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetAllDescendantsInf<T>(this T rootClass) where T : RDInterface
        {
            return rootClass.GetThisAndAllDescendantsInf().Skip(1);
        }

        /// <summary>
        /// Перечислить все вышестоящие классы
        /// </summary>
        /// <param name="aClass"></param>
        /// <returns></returns>
        public static IEnumerable<RDClass> GetAllAccessors(this RDClass aClass)
        {
            var ancestor = aClass.GetAncestor();
            while (ancestor != null)
            {
                yield return ancestor;
                ancestor = ancestor.GetAncestor();
            }
        }

        /// <summary>
        /// Перечислить все вышестоящие интерфейсы
        /// </summary>
        /// <param name="anInterface"></param>
        /// <returns></returns>
        public static IEnumerable<RDInterface> GetAllAccessors(this RDInterface anInterface)
        {
            var ancestor = anInterface.GetAncestor();
            while (ancestor != null)
            {
                yield return ancestor;
                ancestor = ancestor.GetAncestor();
            }
        }

        /// <summary>
        /// Перечислить все вышестоящие классы
        /// </summary>
        /// <param name="aClass"></param>
        /// <returns></returns>
        public static IEnumerable<RDClass> GetThisAndAllAccessors(this RDClass aClass)
        {
            yield return aClass;
            var ancestor = aClass.GetAncestor();
            while (ancestor != null)
            {
                yield return ancestor;
                ancestor = ancestor.GetAncestor();
            }
        }

        /// <summary>
        /// Перечислить все вышестоящие интерфейсы
        /// </summary>
        /// <param name="anInterface"></param>
        /// <returns></returns>
        public static IEnumerable<RDInterface> GetThisAndAllAccessors(this RDInterface anInterface)
        {
            yield return anInterface;
            var ancestor = anInterface.GetAncestor();
            while (ancestor != null)
            {
                yield return ancestor;
                ancestor = ancestor.GetAncestor();
            }
        }
    }
}