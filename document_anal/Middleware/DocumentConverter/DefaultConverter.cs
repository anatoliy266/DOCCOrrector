using document_anal.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using OpenXmlPowerTools.HtmlToWml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using HtmlAgilityPack;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using document_anal.Middleware.Enums;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using DocumentFormat.OpenXml;
using document_anal.Middleware.WordCorrector;

namespace document_anal.Middleware.DocumentConverter
{
    public class DefaultConverter : IConverter
    {
        public DocumentViewModel Fix(string name, List<DocumentValidationError> errors)
        {
            using (var document = WordprocessingDocument.Open(name, true))
            {
                ///применение изменений к документу
                ///
                var corrector = new DocumentCorrector(document);
                foreach (var err in errors)
                {
                    corrector.CorrectProperties(err.ParagraphId, new Rules()
                    {
                        Fields = new DocumentField() { Left = 3.0f, Right = 1.5f, Up = 2.0f, Down = 2.0f },
                        Font = "Times New Roman",
                        FontSize = 24.0f,
                        ParagraphIndent = 0.0f,
                        ParagraphSpacing = 1.0f,
                        LineSpacing = 1.5f,
                        FootNote = new FootNote() { Size = 12, Style = "Arabic" },
                        FontStyle = null,
                    });
                }
                corrector._wordprocessingDocument.SaveAs(name);
                //document.Close();
            }
            ///преобразование документа в HTML для отображения на сайте
            var htmlDoc = ToHtml(name, Enums.DocumentType.Распоряжение);
            return htmlDoc;
        }

        public DocumentViewModel Fix(string name, byte[] bytes, List<DocumentValidationError> errors, out byte[] stream)
        {
            var corrector = new DocumentCorrector();
            foreach (var err in errors)
            {
                bytes = corrector.CorrectProperties(bytes, err.ParagraphId, new Rules()
                {
                    Fields = new DocumentField() { Left = 3.0f, Right = 1.5f, Up = 2.0f, Down = 2.0f },
                    Font = "Times New Roman",
                    FontSize = 24.0f,
                    ParagraphIndent = 0.0f,
                    ParagraphSpacing = 1.0f,
                    LineSpacing = 1.5f,
                    FootNote = new FootNote() { Size = 12, Style = "Arabic" },
                    FontStyle = null,
                });
            }
            ///преобразование документа в HTML для отображения на сайте
            ///
            var htmlDoc = ToHtml(name, bytes, Enums.DocumentType.Распоряжение, out var memstream);
            stream = memstream;
            return htmlDoc;
        }

        public DocumentViewModel Fix(string name, MemoryStream ms, List<DocumentValidationError> errors, out byte[] stream)
        {
            var bytes = ms.ToArray();
            using (var _ms = new MemoryStream())
            {
                _ms.Write(bytes, 0, bytes.Length);
                _ms.Position = 0;
                using (var document = WordprocessingDocument.Open(ms, true))
                {
                    ///применение изменений к документу
                    ///
                    var corrector = new DocumentCorrector(document);
                    foreach (var err in errors)
                    {
                        corrector.CorrectProperties(err.ParagraphId, new Rules()
                        {
                            Fields = new DocumentField() { Left = 3.0f, Right = 1.5f, Up = 2.0f, Down = 2.0f },
                            Font = "Times New Roman",
                            FontSize = 24.0f,
                            ParagraphIndent = 0.0f,
                            ParagraphSpacing = 1.0f,
                            LineSpacing = 1.5f,
                            FootNote = new FootNote() { Size = 12, Style = "Arabic" },
                            FontStyle = null,
                        });
                    }
                    document.MainDocumentPart.Document = corrector._wordprocessingDocument.MainDocumentPart.Document;
                    //document.Close();

                    ///преобразование документа в HTML для отображения на сайте
                    ///
                    var htmlDoc = ToHtml(name, _ms, Enums.DocumentType.Распоряжение, out var memstream);
                    stream = memstream;
                    return htmlDoc;
                }
            }
        }

        public DocumentViewModel ToHtml(string name, Enums.DocumentType type)
        {
            ///преобразование документа в HTML для отображения на сайте 
            ///в класс DocumentViewModel уходят стили, контент сайта и обнаруженные ошибки стилей и форматирования
            ///
            var htmlDoc = new DocumentViewModel { Name = name, DocumentType = type };
            ///открытие загруженного документа
            using (var document = WordprocessingDocument.Open(name, true))
            {
                ///добавление ID парахрафам с отсутствующим идентификатором
                ///
                foreach (var element in document.MainDocumentPart.Document.Body.ChildElements)
                {
                    if (element is Paragraph)
                    {
                        if ((element as Paragraph).ParagraphId == null)
                            (element as Paragraph).ParagraphId = HexBinaryValue.FromString(element.GetHashCode().ToString());
                    }
                }
                ///проверка на наличие ошибок стиля и форматирования
                ///
                htmlDoc.ValidationErrors = Validate(document);

                ///преобразование в HTML, дополнительных параметров не нужно, поэтому создается пустой класс HtmlConverterSettings
                ///
                var settings = new HtmlConverterSettings();

                ///функция библиотеки openxml-powertools, конвертирует документ ворд в html
                ///
                var d = HtmlConverter.ConvertToHtml(document, settings);

                ///под тегом style генерируются стили документа
                ///под body - контент
                ///
                foreach (var node in d.DescendantNodes())
                {
                    ///записываем стили и контент в свойства класса DocumentViewModel для дальнейшей передачи на форму
                    if (node is XElement)
                    {
                        if (((XElement)node).Name.LocalName == "style") htmlDoc.Style = ((XElement)node).Value.Replace("\r", " ").Replace("\n", " ");
                        if (((XElement)node).Name.LocalName == "body") htmlDoc.Content = ((XElement)node).ToString();
                    }
                }

                ///освобождаем документ
                document.Close();
            }
            return htmlDoc;
        }

        public DocumentViewModel ToHtml(string name, MemoryStream ms, Enums.DocumentType type, out byte[] currentStream)
        {
            byte[] bytes;
            ///преобразование документа в HTML для отображения на сайте 
            ///в класс DocumentViewModel уходят стили, контент сайта и обнаруженные ошибки стилей и форматирования
            ///
            var htmlDoc = new DocumentViewModel { Name = name, DocumentType = type, };

            using (var _ms = new MemoryStream())
            {
                _ms.Write(ms.ToArray(), 0, ms.ToArray().Length);
                _ms.Position = 0;
                ///открытие загруженного документа
                using (var document = WordprocessingDocument.Open(_ms, true))
                {

                    ///проверка на наличие ошибок стиля и форматирования
                    ///
                    htmlDoc.ValidationErrors = Validate(document);

                    ///чтобы не хранить файлы на сервере файл храним в модели представления в byte массиве
                    ///
                    //htmlDoc.MemoryStream = ms.ToArray();
                    ///преобразование в HTML, дополнительных параметров не нужно, поэтому создается пустой класс HtmlConverterSettings
                    ///
                    var settings = new HtmlConverterSettings();

                    ///функция библиотеки openxml-powertools, конвертирует документ ворд в html
                    ///
                    var d = HtmlConverter.ConvertToHtml(document, settings);

                    ///под тегом style генерируются стили документа
                    ///под body - контент
                    ///
                    foreach (var node in d.DescendantNodes())
                    {
                        ///записываем стили и контент в свойства класса DocumentViewModel для дальнейшей передачи на форму
                        if (node is XElement)
                        {
                            if (((XElement)node).Name.LocalName == "style") htmlDoc.Style = ((XElement)node).Value.Replace("\r", " ").Replace("\n", " ");
                            if (((XElement)node).Name.LocalName == "body") htmlDoc.Content = ((XElement)node).ToString();
                        }
                    }
                }
                bytes = _ms.ToArray();
            }
            currentStream = bytes;
            return htmlDoc;
        }

        public DocumentViewModel ToHtml(string name, byte[] bytes, Enums.DocumentType type, out byte[] currentStream)
        {
            ///преобразование документа в HTML для отображения на сайте 
            ///в класс DocumentViewModel уходят стили, контент сайта и обнаруженные ошибки стилей и форматирования
            ///
            var htmlDoc = new DocumentViewModel { Name = name, DocumentType = type, };

            ///добавление ID парахрафам с отсутствующим идентификатором
            ///
            byte[] b;
            using (var ms = new MemoryStream())
            {
                ms.Write(bytes, 0, bytes.Length);
                ms.Position = 0;
                ///открытие загруженного документа
                using (var document = WordprocessingDocument.Open(ms, true))
                {

                    ///проверка на наличие ошибок стиля и форматирования
                    ///
                    htmlDoc.ValidationErrors = Validate(document);

                    ///чтобы не хранить файлы на сервере файл храним в модели представления в byte массиве
                    ///


                    ///преобразование в HTML, дополнительных параметров не нужно, поэтому создается пустой класс HtmlConverterSettings
                    ///
                    var settings = new HtmlConverterSettings();

                    ///функция библиотеки openxml-powertools, конвертирует документ ворд в html
                    ///
                    var d = HtmlConverter.ConvertToHtml(document, settings);

                    ///под тегом style генерируются стили документа
                    ///под body - контент
                    ///
                    foreach (var node in d.DescendantNodes())
                    {
                        ///записываем стили и контент в свойства класса DocumentViewModel для дальнейшей передачи на форму
                        if (node is XElement)
                        {
                            if (((XElement)node).Name.LocalName == "style") htmlDoc.Style = ((XElement)node).Value.Replace("\r", " ").Replace("\n", " ");
                            if (((XElement)node).Name.LocalName == "body") htmlDoc.Content = ((XElement)node).ToString();
                        }
                    }

                    ///освобождаем документ
                    ///
                    ///document.Close();

                }
                b = ms.ToArray();
                //htmlDoc.MemoryStream = ms.ToArray();

            }
            currentStream = b;
            return htmlDoc;
        }

        public List<DocumentValidationError> Validate(WordprocessingDocument doc)
        {
            return new List<DocumentValidationError>() { new DocumentValidationError() { Position = -1, ErrorType = ErrorType.GridError, Description = "Правила для документа не определены" } };
        }
    }
}