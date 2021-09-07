using System.Collections.Generic;
using document_anal.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace document_anal.Middleware.WordCorrector.Extensions
{
    public static class TextPropertiesExtensions
    {
        public static void AddTextProperties(this RunProperties runProperties, IGenegalRules genegalRules)
        {
            if (genegalRules.FontStyle != null)
                runProperties.AddFontStyle(genegalRules.FontStyle);

            if (genegalRules.FontSize != 0)
                runProperties.AddFontSize(genegalRules.FontSize);

            if (genegalRules.Font != null)
                runProperties.AddFont(genegalRules.Font);
        }

        public static void AddFootNoteTextProperties(this RunProperties runProperties, IGenegalRules genegalRules)
        {
            //if (genegalRules.FootNote.FontStyle != null)
            //    runProperties.AddFontStyle(genegalRules.FootNote.FontStyle);
            if (genegalRules.FootNote != null)
            {
                if (genegalRules.FootNote.Style != null)
                    runProperties.AddFont(genegalRules.FootNote.Style);
                if (genegalRules.FootNote.Size != 0)
                    runProperties.AddFontSize(genegalRules.FootNote.Size);
            }
        }

        private static void AddFontStyle(this RunProperties runProperties, OnOffType type)
        {
            if (type is Italic)
            {
                runProperties.Italic = new Italic();
            } 
            if (type is Bold)
            {
                runProperties.Bold = new Bold();
            }
        }

        /// <summary>
        /// Размер шрифта будет в два раза меньше
        /// </summary>
        private static void AddFontSize(this RunProperties runProperties, float size)
        {
            runProperties.AppendChild(new FontSize() { Val = size.ToString()});
        }

        private static void AddFont(this RunProperties runProperties, string font)
        {
            runProperties.AppendChild(new RunFonts() { Ascii = font });
        }
    }
}
