using System.Collections.Generic;
using document_anal.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace document_anal.Middleware.WordCorrector.Extensions
{
    public static class ParagraphPropertiesExtensions
    {
        static int defaultSize = 240;
        public static void AddParagraphProperties(this ParagraphProperties paragraphProperties, IGenegalRules genegalRules)
        {
            if (genegalRules.LineSpacing != 0)
                paragraphProperties.AddParagraphAndLineSpacing(genegalRules.ParagraphSpacing, genegalRules.LineSpacing);

            if (genegalRules.Justification != null)
                paragraphProperties.AddJustification(genegalRules.Justification);

            if (genegalRules.ParagraphIndent != 0) //табы
                paragraphProperties.AddIndentention(genegalRules.ParagraphIndent);

            if (genegalRules.ParagraphIntendation != null)
                paragraphProperties.AddFieldIndentation(genegalRules.ParagraphIntendation);

            //if (genegalRules.ParagraphSpacing != 0)
            //    paragraphProperties.AddParagraphSpacing(genegalRules.ParagraphSpacing);

        }

        private static void AddJustification(this ParagraphProperties runProperties, JustificationValues value)
        { 
            runProperties.AppendChild(new Justification() { Val = value });
        }

        private static void AddIndentention(this ParagraphProperties runProperties, float value)
        {
            runProperties.AppendChild(new Indentation() { FirstLine = (value * defaultSize).ToString() }); //табы
        }

        private static void AddFieldIndentation(this ParagraphProperties runProperties, DocumentField documentField) //отступы справа и слева
        {
            var indent = new Indentation();
            if (documentField.Left != 0)
                indent.Left = (documentField.Left * defaultSize).ToString();
            if (documentField.Right != 0)
                indent.Right = (documentField.Right * defaultSize).ToString();
            runProperties.AppendChild(indent);
        }
        private static void AddParagraphAndLineSpacing(this ParagraphProperties runProperties, float paragraphSpacing = 0, float lineSpacing = 0) //отступы справа и слева
        {
            var intend = new SpacingBetweenLines();
            if (paragraphSpacing > 0)
            {
                intend.After = new StringValue((paragraphSpacing*240).ToString());
                intend.Before = new StringValue((paragraphSpacing * 240).ToString());
            }
            if (lineSpacing > 0)
            {
                intend.Line = new StringValue((lineSpacing * 240).ToString());
            }

            runProperties.AppendChild(intend);
        }
    }
}
