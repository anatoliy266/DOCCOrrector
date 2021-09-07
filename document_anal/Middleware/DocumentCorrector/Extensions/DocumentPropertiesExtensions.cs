using document_anal.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace document_anal.Middleware.WordCorrector.Extensions
{
    public static class DocumentPropertiesExtensions
    {
        /// <summary>
        /// Добавляет свойства предназначенные для документа
        /// </summary>
        public static void AddProperties(this SectionProperties sectionProperties, IGenegalRules genegalRules)
        {
            if (genegalRules == null)
                return;

            sectionProperties.AddMargins(genegalRules);
        }

        /// <summary>
        /// Добавляет свойство "Отступы документа справа, слева, сверху и снизу"
        /// </summary>
        private static void AddMargins(this SectionProperties sectionProperties, IGenegalRules genegalRules)
        {
            sectionProperties.AppendChild(new PageMargin()
            {
                Left = (uint)genegalRules.Fields.Left,
                Right = (uint)genegalRules.Fields.Right,
                Bottom = (int)genegalRules.Fields.Down,
                Top = (int)genegalRules.Fields.Up
            }); // лучше изменить типы в Fields на соответствующие
        }
    }
}
