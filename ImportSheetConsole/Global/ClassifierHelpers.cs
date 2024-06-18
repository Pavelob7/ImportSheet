using System;
using System.Collections.Generic;
using System.Linq;
using CommonTools;
using JetBrains.Annotations;
using ObjStudioClasses;

namespace ImportSheetConsole.Global
{
    /// <summary>
    /// Утилиты работы с элементами классификатора
    /// </summary>
    public static class ClassifierHelper
    {
        /// <summary>
        /// Кэш атрибутов, пригодных для построения узла классификатора из экземпляра указанного класса
        /// </summary>
        [NotNull]
        private static readonly Dictionary<RDClass, HashSet<RDAttribute>> _attributesCache = new Dictionary<RDClass, HashSet<RDAttribute>>();

        /// <summary>
        /// Заполучить атрибуты, пригодные для построения следующего уровня классификатора
        /// </summary>
        /// <param name="classifierClass"></param>
        /// <returns></returns>
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<RDAttribute> GetClassifierAttributesFast([NotNull] this CustomClassifierNodeClassInfo classifierClass)
        {
            lock (_attributesCache)
            {
                HashSet<RDAttribute> attributes;
                if (!_attributesCache.TryGetValue(classifierClass, out attributes) || attributes == null)
                {
                    attributes = new HashSet<RDAttribute>(classifierClass.GetOwnAttributes(ClassAreaOfAction.ForInstance, InheritanceLevelKind.CurrentLevelAndAncestors, true)
                        .Where(IsClassifierAttribute));
                    _attributesCache[classifierClass] = attributes;
                }

                return attributes;
            }
        }

        /// <summary>
        /// Получить перечисление ВСЕХ вышестоящих элементов
        /// </summary>
        /// <param name="lowerItem"></param>
        /// <returns></returns>
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<CustomClassifierNode> AllUpperItems([NotNull] this ClassifierItem lowerItem)
        {
            foreach (var upperLevel in lowerItem.UpperItems())
            {
                yield return upperLevel;
                var upperLevelAsClassifierItem = upperLevel as ClassifierItem;
                if (upperLevelAsClassifierItem != null)
                {
                    foreach (var upper2 in upperLevelAsClassifierItem.AllUpperItems())
                    {
                        yield return upper2;
                    }
                }
            }
        }

        ///// <summary>
        ///// Получить перечисление ВСЕХ вышестоящих элементов с учетом зависимостей
        ///// </summary>
        ///// <param name="lowerItem"></param>
        ///// <param name="appendRelations"></param>
        ///// <param name="distinctProcessor"></param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //private static IEnumerable<CustomClassifierNode> AllUpperItemsIncludeRelationsInternal([NotNull] this ClassifierItem lowerItem, bool appendRelations,
        //    [NotNull] HashSet<CustomClassifierNode> distinctProcessor)
        //{
        //    if (distinctProcessor.Add(lowerItem))
        //    {
        //        if (appendRelations)
        //        {
        //            CustomClassifierNode[] upperrelations = lowerItem.UpperRelations().ToArray();
        //            foreach (CustomClassifierNode upperRelation in upperrelations)
        //            {
        //                if (distinctProcessor.Add(upperRelation))
        //                    yield return upperRelation;
        //            }
        //        }

        //        CustomClassifierNode[] upperItems = lowerItem.UpperItems().ToArray();
        //        foreach (CustomClassifierNode upperLevel in upperItems)
        //        {
        //            yield return upperLevel;
        //            var upperLevelAsClassifierItem = upperLevel as ClassifierItem;
        //            if (upperLevelAsClassifierItem != null)
        //            {
        //                CustomClassifierNode[] upperLevelItems = upperLevelAsClassifierItem.AllUpperItemsIncludeRelationsInternal(false, distinctProcessor).ToArray();
        //                foreach (var upper2 in upperLevelItems)
        //                {
        //                    yield return upper2;
        //                }
        //            }
        //        }
        //    }
        //}

        ///// <summary>
        ///// Получить перечисление ВСЕХ вышестоящих элементов с учетом зависимостей
        ///// </summary>
        ///// <param name="lowerItem"></param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //public static IEnumerable<CustomClassifierNode> AllUpperItemsIncludeRelations([NotNull] this ClassifierItem lowerItem)
        //{
        //    return AllUpperItemsIncludeRelationsInternal(lowerItem, true, new HashSet<CustomClassifierNode>());
        //}

        /// <summary>
        /// Получить перечисление вышестоящих элементов классификации (их может быть несколько, т.к. одна и та же сущность
        /// может присутствовать в разных классификаторах)
        /// </summary>
        /// <param name="lowerItem"></param>
        /// <returns></returns>
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<CustomClassifierNode> UpperItems([NotNull] this ClassifierItem lowerItem)
        {
            return lowerItem.UpperItemsEx().Select(x => x.Key);
        }

        /// <summary>
        /// Получить перечисление вышестоящих элементов классификации (их может быть несколько, т.к. одна и та же сущность
        /// может присутствовать в разных классификаторах)
        /// Расширенный метод с указанием связывающих атрибутов, если таковые есть.
        /// Связывающего атрибута может не быть, если связь искусственная (н-р, абонентские точки на фидере)
        /// </summary>
        /// <param name="lowerItem"></param>
        /// <returns></returns>
        [NotNull]
        public static IEnumerable<KeyValuePair<CustomClassifierNode, RDAttribute>> UpperItemsEx([NotNull] this ClassifierItem lowerItem)
        {
            return lowerItem.GetRelations()
                .Where(x => IsClassifierAttribute(x.Owner.Descriptor) && x.Owner.Owner is CustomClassifierNode)
                .Select(x => new KeyValuePair<CustomClassifierNode, RDAttribute>((CustomClassifierNode)x.Owner.Owner, x.Owner.Descriptor));
        }

        ///// <summary>
        ///// Получить связи элемента, по которым можно пройти снизу вверх по классификатору
        ///// </summary>
        ///// <param name="lowerItem"></param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //public static IEnumerable<CustomClassifierNode> UpperRelations([NotNull]this ClassifierItem lowerItem)
        //{
        //    return lowerItem.GetUpperRelations();
        //}

        /// <summary>
        /// Получить нижесстоящие элементы классифиатора
        /// </summary>
        /// <param name="upperItem"></param>
        /// <returns></returns>
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<ClassifierItem> LowerItems([NotNull] this CustomClassifierNode upperItem)
        {
            return upperItem.Class.GetClassifierAttributesFast()
                .SelectMany(x => upperItem.GetAttributeValues(x).GetValues<ClassifierItem>());
        }

        ///// <summary>
        ///// Получить связи вниз по классификатору (ссылочные элементы под элементом классификатора)
        ///// </summary>
        ///// <param name="upperItem"></param>
        ///// <param name="contextItem"></param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //public static IEnumerable<CustomClassifierNode> LowerRelations([NotNull]this ClassifierItem upperItem, CustomClassifierNode contextItem)
        //{
        //    return upperItem.GetContextRelations(contextItem)
        //        .Where(x => x != null);
        //}

        ///// <summary>
        ///// Получить нижестоящие элементы классифиатора, включая связи
        ///// </summary>
        ///// <param name="upperItem"></param>
        ///// <param name="contextItem">Вышестоящий родитель, относительно которого определяется содержимое связей</param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //public static IEnumerable<ClassifierItem> LowerItemsIncludeRelations([NotNull]this CustomClassifierNode upperItem, CustomClassifierNode contextItem)
        //{
        //    var upperItemAsClassifierItem = upperItem as ClassifierItem;
        //    if (upperItemAsClassifierItem == null)
        //        return upperItem.LowerItems();
        //    return upperItem.LowerItems().Concat(LowerRelations(upperItemAsClassifierItem, contextItem).OfType<ClassifierItem>());
        //}

        /// <summary>
        /// Получить все элементы ниже текущего (без связей)
        /// </summary>
        /// <param name="node"></param>
        /// <param name="distinctProcessor"></param>
        /// <returns></returns>
        [NotNull]
        [ItemNotNull]
        private static IEnumerable<CustomClassifierNode> GetThisAndAllLowerNodesInternal([NotNull] CustomClassifierNode node, [NotNull] HashSet<CustomClassifierNode> distinctProcessor)
        {
            if (distinctProcessor.Add(node))
            {
                yield return node;
                foreach (var subNode in node.LowerItems().SelectMany(x => GetThisAndAllLowerNodesInternal(x, distinctProcessor.AssertNull())))
                    yield return subNode;
            }
        }

        /// <summary>
        /// Получить все элементы ниже текущего (без связей)
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<CustomClassifierNode> GetThisAndAllLowerNodes([NotNull] this CustomClassifierNode node)
        {
            return GetThisAndAllLowerNodesInternal(node, new HashSet<CustomClassifierNode>());
        }

        ///// <summary>
        ///// Получить себя и все элементы ниже текущего, включая связи
        ///// </summary>
        ///// <param name="node"></param>
        ///// <param name="contextNode"></param>
        ///// <param name="antiRecursive"></param>
        ///// <param name="distProcessor"></param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //private static IEnumerable<CustomClassifierNode> GetThisAndAllLowerNodesIncludeRelationsInternal([NotNull]CustomClassifierNode node,
        //    CustomClassifierNode contextNode, [NotNull]HashSet<Tuple<CustomClassifierNode, CustomClassifierNode>> antiRecursive,
        //    [NotNull]HashSet<CustomClassifierNode> distProcessor)
        //{
        //    if (!antiRecursive.Add(new Tuple<CustomClassifierNode, CustomClassifierNode>(node, contextNode)))
        //        yield break;
        //    if (distProcessor.Add(node))
        //        yield return node;
        //    // в зависимости от разных контекстов зависимости могут быть разными
        //    var nodeAsClassifierItem = node as ClassifierItem;
        //    if (nodeAsClassifierItem != null)
        //    {
        //        foreach (var relatedNode in nodeAsClassifierItem.LowerRelations(contextNode))
        //        {
        //            if (distProcessor.Add(relatedNode))
        //                yield return relatedNode;
        //        }
        //    }
        //    foreach (var subNode in node.LowerItems()
        //        .SelectMany(x => GetThisAndAllLowerNodesIncludeRelationsInternal(x, node, antiRecursive, distProcessor)))
        //        yield return subNode.AssertNull();
        //}

        ///// <summary>
        ///// Получить себя и все элементы ниже текущего, включая связи
        ///// </summary>
        ///// <param name="node"></param>
        ///// <param name="contextNode"></param>
        ///// <returns></returns>
        //[NotNull]
        //[ItemNotNull]
        //public static IEnumerable<CustomClassifierNode> GetThisAndAllLowerNodesIncludeRelations([NotNull]this CustomClassifierNode node, CustomClassifierNode contextNode)
        //{
        //    return GetThisAndAllLowerNodesIncludeRelationsInternal(node, contextNode, new HashSet<Tuple<CustomClassifierNode, CustomClassifierNode>>(),
        //        new HashSet<CustomClassifierNode>());
        //}

        /// <summary>
        /// Кэширование признака сокрытия атрибута в web
        /// </summary>
        [NotNull]
        private static readonly ThreadSafeCache<RDAttribute, bool> _hiddenAttributesCache =
            new ThreadSafeCache<RDAttribute, bool>(x =>
            {
                var rules = x.AttributeWebVisualizationRules().GetValues().Select(y => y.ToEnumItem()).ToArray();
                return rules.Contains(AttributeWebVisualizationRuleData.DisallowViewInPassport) && !rules.Contains(AttributeWebVisualizationRuleData.AllowViewInDictionary);
            });

        /// <summary>
        /// Атрибут не редактиурется в паспорте
        /// </summary>
        /// <param name="attribute"></param>
        /// <returns></returns>
        public static bool IsHidenInWebPassport([NotNull] this RDAttribute attribute)
        {
            return _hiddenAttributesCache.Get(attribute);
        }

        /// <summary>
        /// Определение принадлежности атрибута к числу тех, на которых формируется классификатор
        /// </summary>
        /// <param name="attribute"></param>
        /// <returns></returns>
        public static bool IsClassifierAttribute([NotNull] RDAttribute attribute)
        {
            // критерии определения атрибута как образующего следующий уровень классификатора:
            // - ссылочная связь
            if (attribute.DataType != AttributeValueType.Entity)
                return false;
            var options = attribute.Options;
            // - с указанием на экземпляр
            if (options.LinkTargetKind != AttributeLinkTargetKind.Instance)
                return false;
            /*
            // - поддержка множества значений
            if (options.MaxOccurs <= 1)
                return false;
             */

            // - ссылка не композиционная (объект не включен в состав другого родительского)
            if (options.KindOfLink == LinkKind.Composition)
                return false;
            // - задан ограничивающий содержимое класс
            if (options.RequiredTypeRefName == null)
                return false;
            var requredClass = RDCustomEntity.Find(options.RequiredTypeRefName.Value) as RDClass;
            // и этот класс наследуется от объекта классификатора
            return requredClass is ClassifierItemClassInfo;
        }

        /// <summary>
        /// Уведомление о модификации состава атрибутов в Runtime
        /// Вызывается из триггера по факту добавления/удаления пользовательских атрибутов
        /// </summary>
        public static void NotifyClassifierAttributesChangesDetected()
        {
            lock (_attributesCache)
                _attributesCache.Clear();
        }

        /// <summary>
        /// Средства синхронизации инициализации _classifierItemClassRefNames
        /// </summary>
        [NotNull]
        private static readonly object _classifierItemClassRefNamesSync = new object();

        /// <summary>
        /// Кэшированный список типов элементов классификатора
        /// Нужен для максимальной оптимизации обращения
        /// </summary>
        private static HashSet<Guid> _classifierItemClassRefNames;

        /// <summary>
        /// Является ли указанный идентификатор идентификатором класса элемента классификатора
        /// </summary>
        /// <param name="refName"></param>
        /// <returns></returns>
        private static bool IsClassifierItemClassRefName(Guid refName)
        {
            lock (_classifierItemClassRefNamesSync)
            {
                if (_classifierItemClassRefNames == null)
                {
                    _classifierItemClassRefNames = new HashSet<Guid>();
                    var rootClassInfo = ClassifierItem.GetClassInfo();
                    _classifierItemClassRefNames.Add(rootClassInfo.RefName);
                    foreach (var classInfo in rootClassInfo.GetAllDescendants())
                    {
                        _classifierItemClassRefNames.Add(classInfo.RefName);
                    }
                }
            }

            return _classifierItemClassRefNames.Contains(refName);
        }
    }
}