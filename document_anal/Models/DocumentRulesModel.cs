using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace document_anal.Models
{
    public class DocumentField
    {
        public float Left { get; set; }
        public float Right { get; set; }
        public float Up { get; set; }
        public float Down { get; set; }
    }

    public class FootNote
    {
        public float Size { get; set; }
        public string Style { get; set; }
    }

    public interface IGenegalRules
    {
        /// <summary>
        /// Поля документа
        /// </summary>
        DocumentField Fields { get; set; }
        /// <summary>
        /// Шрифт (Arial, Times New Roman, ...)
        /// </summary>
        string Font { get; set; }
        /// <summary>
        /// Размер шрифта * 2
        /// </summary>
        float FontSize { get; set; }
        /// <summary>
        /// Отступ первой строки параграфа
        /// </summary>
        float ParagraphIndent { get; set; }
        /// <summary>
        /// Отступ между параграфами
        /// </summary>
        float ParagraphSpacing { get; set; }
        /// <summary>
        /// Межстрочный интервал
        /// </summary>
        float LineSpacing { get; set; }
        /// <summary>
        /// формат сноски
        /// </summary>
        FootNote FootNote { get; set; }
        /// <summary>
        /// стиль шрифта (Жирный, курсив)
        /// </summary>
        OnOffType FontStyle { get; set; }
        /// <summary>
        /// Центрирование параграфа
        /// </summary>
        JustificationValues Justification { get; set; }
        /// <summary>
        /// Отступы текста параграфа
        /// </summary>
        DocumentField ParagraphIntendation { get; set; }
    }

    public class Rules : IGenegalRules
    {
        public DocumentField Fields { get; set; }
        public string Font { get; set; }
        public float FontSize { get; set; }
        public float ParagraphIndent { get; set; }
        public float ParagraphSpacing { get; set; }
        public float LineSpacing { get; set; }
        public FootNote FootNote { get; set; }
        public OnOffType FontStyle { get; set; }
        public JustificationValues Justification { get; set; }
        public DocumentField ParagraphIntendation { get; set; }
    }
}