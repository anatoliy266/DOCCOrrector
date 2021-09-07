using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using System.Collections.Generic;
using System.Linq;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using System;
using document_anal.Models;
using document_anal.Middleware.WordCorrector.Extensions;
using System.IO;

namespace document_anal.Middleware.WordCorrector
{
    public class DocumentCorrector
    {
        public string author = "Сбербанк";

        public WordprocessingDocument _wordprocessingDocument { get; set; }
        public DocumentCorrector(WordprocessingDocument wordprocessingDocument)
        {
            _wordprocessingDocument = wordprocessingDocument;
            TurnOnTrackRevesions();
        }

        public DocumentCorrector()
        {
        }

        private void TurnOnTrackRevesions()
        {
            DocumentSettingsPart documentSettingsPart =
                _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
            TrackRevisions trackRevisions = new TrackRevisions() { Val = true };
            documentSettingsPart.Settings.Append(trackRevisions);
        }
        private void TurnOnTrackRevesions(WordprocessingDocument doc)
        {
            DocumentSettingsPart documentSettingsPart =
                doc.MainDocumentPart.DocumentSettingsPart;
            TrackRevisions trackRevisions = new TrackRevisions() { Val = true };
            documentSettingsPart.Settings.Append(trackRevisions);
        }

        public void CorrectProperties(HexBinaryValue paragraphIndex, IGenegalRules genegalRules)
        {
            ChangeParagraphProperties(paragraphIndex, genegalRules);
            ChangeTextProperties(paragraphIndex, genegalRules);
        }

        public byte[] CorrectProperties(byte[] bytes, HexBinaryValue paragraphIndex, IGenegalRules genegalRules)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(bytes, 0, bytes.Length);
                ms.Position = 0;
                using (var document = WordprocessingDocument.Open(ms, true))
                {
                    TurnOnTrackRevesions(document);
                    ChangeFootnotesProperties(document, genegalRules);
                    ChangeParagraphProperties(document, paragraphIndex, genegalRules);
                    ChangeTextProperties(document, paragraphIndex, genegalRules);
                    ChangeFootnotesProperties(document, genegalRules);
                }
                return ms.ToArray();
            }
        }

        public void ChangeFootnotesProperties(WordprocessingDocument document, IGenegalRules genegalRules)
        {

            document.MainDocumentPart.FootnotesPart.Footnotes.ChildElements.OfType<Footnote>().ToList().ForEach(x =>
            {
                x.ChildElements.OfType<Paragraph>().ToList().ForEach(y => ChangeParagraphProperties(y, genegalRules));
            });
        }

        /// <summary>
        ///  Изменяет свойства документа
        /// </summary>
        private void ChangeDocumentProperties(WordprocessingDocument document, IGenegalRules genegalRules)
        {
            SectionProperties oldSectionProperties = document.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
            if (oldSectionProperties == null)
                oldSectionProperties = document.MainDocumentPart.Document.Body.AppendChild(new SectionProperties());

            SectionProperties newSectionProperties = new SectionProperties();

            // сначала добавляем новые свойства для документа
            newSectionProperties.AddProperties(genegalRules);
            // потом помечаем старые свойства, как изменённые (создаём копию)
            newSectionProperties.AppendChild(new SectionPropertiesChange(oldSectionProperties.CloneNode(true))
            {
                Author = author,
                Date = DateTime.Now
            });
            // и обновляем текущие свойства документа
            document.MainDocumentPart.Document.Body.ReplaceChild(newSectionProperties, oldSectionProperties);
        }
        /// <summary>
        /// Изменяет свойства текста.
        /// </summary>
        /// 
        private void ChangeTextProperties(HexBinaryValue paragraphIndex, IGenegalRules genegalRules)
        {
            //Run originalRun = GetParagraph(paragraphIndex).GetFirstChild<Run>();
            foreach (var originalRun in GetParagraph(paragraphIndex).ChildElements.OfType<Run>())
            {
                if (originalRun.RunProperties == null)
                {
                    var rp = new RunProperties();
                    originalRun.Append(rp);
                }

                RunProperties oldRunProperties = originalRun.RunProperties;
                RunProperties newRunProperties = new RunProperties();

                // сначала добавляем новые свойства для текста
                newRunProperties.AddTextProperties(genegalRules);
                // потом помечаем старые свойства, как изменённые (создаём копию)
                newRunProperties.AppendChild(new RunPropertiesChange(oldRunProperties.CloneNode(true)) { Author = author, Date = DateTime.Now });
                // и обновляем текущие свойства текста
                originalRun.ReplaceChild(newRunProperties, oldRunProperties);
            }
        }

        private void ChangeTextProperties(WordprocessingDocument doc, HexBinaryValue paragraphIndex, IGenegalRules genegalRules)
        {
            foreach (var originalRun in GetParagraph(doc, paragraphIndex).ChildElements.OfType<Run>())
            {

                if (originalRun.RunProperties == null)
                {
                    var rp = new RunProperties();
                    originalRun.Append(rp);
                }
                RunProperties oldRunProperties = originalRun.RunProperties;
                RunProperties newRunProperties = new RunProperties();

                // сначала добавляем новые свойства для текста
                newRunProperties.AddTextProperties(genegalRules);
                // потом помечаем старые свойства, как изменённые (создаём копию)
                newRunProperties.AppendChild(new RunPropertiesChange(oldRunProperties.CloneNode(true)) { Author = author, Date = DateTime.Now });
                // и обновляем текущие свойства текста
                originalRun.ReplaceChild(newRunProperties, oldRunProperties);
            }
            //Run originalRun = GetParagraph(doc, paragraphIndex).GetFirstChild<Run>();
        }

        /// <summary>
        /// Изменяет свойства параграфа
        /// </summary>
        private void ChangeParagraphProperties(WordprocessingDocument doc, HexBinaryValue paragraphIndex, IGenegalRules genegalRules)
        {
            Paragraph paragraph = GetParagraph(doc, paragraphIndex);
            if (paragraph.ParagraphProperties == null)
            {
                var pp = new ParagraphProperties();
                paragraph.Append(pp);
            }
            ParagraphProperties oldParagraphProperties = paragraph.ParagraphProperties;
            ParagraphProperties newParagraphProperties = new ParagraphProperties();

            if (oldParagraphProperties.NumberingProperties != null)
            {
                newParagraphProperties.NumberingProperties = new NumberingProperties() { NumberingId = new NumberingId() { Val = oldParagraphProperties.NumberingProperties.NumberingId.Val } };
            }

            // сначала добавляем новые свойства для параграфа
            newParagraphProperties.AddParagraphProperties(genegalRules);
            // потом помечаем старые свойства, как изменённые (создаём копию)
            newParagraphProperties.AppendChild(new ParagraphPropertiesChange(oldParagraphProperties.CloneNode(true)) { Author = author, Date = DateTime.Now });
            // и обновляем текущие свойства параграфа
            paragraph.ReplaceChild(newParagraphProperties, oldParagraphProperties);

        }

        private void ChangeParagraphProperties(HexBinaryValue paragraphIndex, IGenegalRules genegalRules)
        {
            Paragraph paragraph = GetParagraph(paragraphIndex);

            ParagraphProperties oldParagraphProperties = paragraph.ParagraphProperties;
            ParagraphProperties newParagraphProperties = new ParagraphProperties();

            // сначала добавляем новые свойства для параграфа
            newParagraphProperties.AddParagraphProperties(genegalRules);
            // потом помечаем старые свойства, как изменённые (создаём копию)
            newParagraphProperties.AppendChild(new ParagraphPropertiesChange(oldParagraphProperties.CloneNode(true)) { Author = author, Date = DateTime.Now });
            // и обновляем текущие свойства параграфа
            paragraph.ReplaceChild(newParagraphProperties, oldParagraphProperties);

        }

        private void ChangeParagraphProperties(Paragraph paragraph, IGenegalRules genegalRules)
        {
            foreach (var originalRun in paragraph.ChildElements.OfType<Run>())
            {

                if (originalRun.RunProperties == null)
                {
                    var runProperties = new RunProperties();
                    Text text = originalRun.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null)
                        text.InsertBeforeSelf(runProperties);
                    else
                        originalRun.Append(runProperties);
                }
                RunProperties oldRunProperties = originalRun.RunProperties;
                RunProperties newRunProperties = new RunProperties();

                // сначала добавляем новые свойства для текста
                newRunProperties.AddFootNoteTextProperties(genegalRules);
                // потом помечаем старые свойства, как изменённые (создаём копию)
                newRunProperties.AppendChild(new RunPropertiesChange(oldRunProperties.CloneNode(true))
                {
                    Author = author,
                    Date = DateTime.Now
                });
                // и обновляем текущие свойства текста
                originalRun.ReplaceChild(newRunProperties, oldRunProperties);
            }
        }

        private Paragraph GetParagraph(HexBinaryValue index)
        {
            OpenXmlElementList openXmlElementList =
                _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements;
            Paragraph paragraph =
                openXmlElementList.OfType<Paragraph>().Where(x => x.ParagraphId == index).FirstOrDefault();
            return paragraph;
        }

        private Paragraph GetParagraph(WordprocessingDocument doc, HexBinaryValue index)
        {
            OpenXmlElementList openXmlElementList =
                doc.MainDocumentPart.Document.Body.ChildElements;
            Paragraph paragraph =
                openXmlElementList.OfType<Paragraph>().Where(x => x.ParagraphId == index).FirstOrDefault();
            return paragraph;
        }
    }
}



// ВСЕ СВОЙСТВА ПОМЕЩАЮТСЯ В RUNPROPERTYCHANGED, для ТЕКСТА
// ПЕРЕД RUNPROPERTYCHANGED ПОМЕЩАЮТСЯ ТЕКУЩИЕ СВОЙСТВА ПАРАГРАФА

// CHANGED FONTSIZE AND FONT

//<w:rPr>
//  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
//  <w:b/>
//  <w:bCs/>
//  <w:sz w:val="20"/>
//  <w:szCs w:val="20"/>
//  <w:lang w:val="en-US"/>
//  <w:rPrChange w:id="1" w:author="ярослав волков" w:date="2021-05-25T12:43:00Z">
//    <w:rPr>
//      <w:rFonts w:cstheme="minorHAnsi"/>
//      <w:b/>
//      <w:bCs/>
//      <w:sz w:val="32"/>
//      <w:szCs w:val="32"/>
//      <w:lang w:val="en-US"/>
//    </w:rPr>
//  </w:rPrChange>
//</w:rPr>

// ORIGINAL

//<w:r w:rsidRPr="00A867CB">
//  <w:rPr>
//    <w:rFonts w:cstheme="minorHAnsi"/>
//    <w:b/>
//    <w:bCs/>
//    <w:sz w:val="32"/>
//    <w:szCs w:val="32"/>
//    <w:lang w:val="en-US"/>
//  </w:rPr>
//  <w:t>One</w:t>
//</w:r>