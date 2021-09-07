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
    public class ServiceNoteConverter : IConverter
    {
        
        public DocumentViewModel Fix(string name, byte[] bytes, List<DocumentValidationError> errors, out byte[] stream)
        {
            var corrector = new DocumentCorrector();
            foreach (var err in errors)
            {
                ///для каждого параграфа применяется стиль форматирования
                ///
                bytes = corrector.CorrectProperties(bytes, err.ParagraphId, new Rules()
                {
                    Fields = new DocumentField() { Left = 3.0f, Right = 1.5f, Up = 2.0f, Down = 2.0f },
                    Font = "Times New Roman",
                    FontSize = 24.0f,
                    ParagraphSpacing = 1.0f,
                    LineSpacing = 1.5f,
                    FootNote = new FootNote() { Size = 20, Style = "Times New Roman" },
                    FontStyle = null,
                });
            }
            ///применение кастомных стилей для заголовков и элементов шаблона
            ///
            using (var _ms = new MemoryStream())
            {
                _ms.Write(bytes, 0, bytes.Length);
                _ms.Position = 0;
                using (var doc = WordprocessingDocument.Open(_ms, true))
                {
                    var paragraphs = doc.MainDocumentPart.Document.Body.OfType<Paragraph>().Where(x => x.InnerText.Trim() != "").ToList();
                    ///
                    bytes = corrector.CorrectProperties(bytes, paragraphs[0].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 24.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, ParagraphIntendation = new DocumentField() { Left = 21.0f }, });
                    bytes = corrector.CorrectProperties(bytes, paragraphs[1].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 24.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, ParagraphIntendation = new DocumentField() { Left = 21.0f }, });
                    bytes = corrector.CorrectProperties(bytes, paragraphs[2].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 24.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, ParagraphIntendation = new DocumentField() { Left = 21.0f }, });
                    bytes = corrector.CorrectProperties(bytes, paragraphs[3].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 24.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, ParagraphIntendation = new DocumentField() { Left = 21.0f }, });

                    bytes = corrector.CorrectProperties(bytes, paragraphs[5].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 24.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, FontStyle = new Italic(), Justification = JustificationValues.Left });
                    bytes = corrector.CorrectProperties(bytes, paragraphs[6].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 24.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, Justification = JustificationValues.Center });

                    bytes = corrector.CorrectProperties(bytes, paragraphs[paragraphs.Count - 2].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 20.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, Justification = JustificationValues.Left, FontStyle = new Italic() });
                    bytes = corrector.CorrectProperties(bytes, paragraphs[paragraphs.Count - 1].ParagraphId, new Rules() { Font = "Times New Roman", FontSize = 20.0f, LineSpacing = 1.5f, ParagraphSpacing = 1.0f, Justification = JustificationValues.Left, FontStyle = new Italic() });
                }
            }
            ///преобразование документа в HTML для отображения на сайте
            ///
            var htmlDoc = ToHtml(name, bytes, Enums.DocumentType.Распоряжение, out var memstream);
            stream = memstream;
            return htmlDoc;
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
            ///инициализируем класс свойств документа 
            ///
            var rule = new Rules()
            {
                Fields = new DocumentField() { Left = 3.0f, Right = 1.5f, Up = 2.0f, Down = 2.0f },
                Font = "Times New Roman",
                FontSize = 24.0f,
                ParagraphIndent = 0.0f,
                ParagraphSpacing = 1.0f,
                LineSpacing = 1.5f,
                FootNote = new FootNote() { Size = 20, Style = "Times New Roman" },
                FontStyle = null,
            };
            ///список ошибок форматирования документа
            ///
            var result = new List<DocumentValidationError>();
            ///проверка на путой документ
            ///
            if (doc.MainDocumentPart.Document.Body.ChildElements == null
                || doc.MainDocumentPart.Document.Body.ChildElements.Count == 0
                || doc.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().Count() == 0
                || doc.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().Where(x => x.InnerText != "").Count() == 0)
            {
                result.Add(new DocumentValidationError() { ParagraphId = new HexBinaryValue("-1"), ErrorType = ErrorType.GridError, Position = -1, Description = "Пустой документ" });
                return result;
            }

            var paragraphs = doc.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().Where(x => x.InnerText.Trim() != "").ToList();



            ///параграфы 1,2,4,5 без учета пустых строк должны быть отформатированы по праваому краю
            ///свойство Indentation значение 5103
            ///если нет - добавить к ошибкам
            ///
            if (paragraphs[0].ParagraphProperties == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[0].ParagraphId, Position = 0, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[0].ParagraphProperties.Indentation == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[0].ParagraphId, Position = 0, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[0].ParagraphProperties.Indentation.Left != "5040")
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[0].ParagraphId, Position = 0, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }

            if (paragraphs[1].ParagraphProperties == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[1].ParagraphId, Position = 1, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[1].ParagraphProperties.Indentation == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[1].ParagraphId, Position = 1, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[1].ParagraphProperties.Indentation.Left != "5040")
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[1].ParagraphId, Position = 1, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }

            if (paragraphs[2].ParagraphProperties == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[2].ParagraphId, Position = 0, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[2].ParagraphProperties.Indentation == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[2].ParagraphId, Position = 0, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[2].ParagraphProperties.Indentation.Left != "5040")
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[2].ParagraphId, Position = 0, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }

            if (paragraphs[3].ParagraphProperties == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[3].ParagraphId, Position = 1, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[3].ParagraphProperties.Indentation == null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[3].ParagraphId, Position = 1, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }
            else if (paragraphs[3].ParagraphProperties.Indentation.Left != "5040")
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[3].ParagraphId, Position = 1, Description = "Некорректная шапка служебной записки, необходимо отформатировать по правому краю" });
            }

            
            ///проверка даты в параграфе
            ///если даты нет - добавить к ошибкам
            ///
            if (!Regex.IsMatch(paragraphs[4].InnerText, @"На\s№\s(.)*\sот\s(0[1-9]|[12][0-9]|3[01])[-/.](0[1-9]|1[12])[-/.](19|20)\d\d[-/.]", RegexOptions.IgnoreCase))
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[4].ParagraphId, Position = 2, Description = "Некорректная дата служебной записки" });
            }

            ///параграф 4 - заголовок, курсивом
            ///если не курсив - добавить к ошибкам
            ///если форматирование не по правому краю - добавить к ошибкам
            ///
            ///проверка на форматирование
            if (paragraphs[5].ParagraphProperties != null && paragraphs[5].ParagraphProperties.Indentation != null && paragraphs[5].ParagraphProperties.Indentation.Left != null)
            {
                result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[5].ParagraphId, Position = 0, Description = "Некорректный заголовок, необходимо отформатировать по правому краю" });
            }

            ///проверка на курсивный шрифт
            ///должен быть в каждом Run в параграфе
            ///
            foreach (var run in paragraphs[5].ChildElements.OfType<Run>())
            {
                if (run.RunProperties != null) result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[3].ParagraphId, Position = 3, Description = "Некорректный заголовок, необходим курсивный шрифт" });
                {
                    if (run.RunProperties?.Italic != new Italic())
                        result.Add(new DocumentValidationError() { ErrorType = ErrorType.GridError, ParagraphId = paragraphs[3].ParagraphId, Position = 3, Description = "Некорректный заголовок, необходим курсивный шрифт" });
                    break;
                }
            }

            ///Обращение в служебной записке должно быть по центру
            ///Начинается с Уважаем(.)*
            ///
            ///проверка на форматирование по центру
            ///
            if (paragraphs[6].ParagraphProperties == null)
            {
                result.Add(new DocumentValidationError() { ParagraphId = paragraphs[6].ParagraphId, ErrorType = ErrorType.GridError, Position = 4, Description = "Некорректный формат обращения, необходимо форматирование по центру" });
            }
            else if (paragraphs[6].ParagraphProperties.Justification.Val != JustificationValues.Center)
            {
                result.Add(new DocumentValidationError() { ParagraphId = paragraphs[6].ParagraphId, ErrorType = ErrorType.GridError, Position = 4, Description = "Некорректный формат обращения, необходимо форматирование по центру" });
            }
            else
            {
                ///проверка текста
                ///должно начинаться с Уважаем(.)*
                if (!Regex.IsMatch(paragraphs[6].InnerText, @"^уважаем", RegexOptions.IgnoreCase))
                {
                    result.Add(new DocumentValidationError() { ParagraphId = paragraphs[6].ParagraphId, ErrorType = ErrorType.GridError, Position = 4, Description = "Некорректный текст обращения" });
                }
            }

            ///последние 2 параграфа - исполнитель и телефон исполнителя - курсивным шрифтом
            foreach (var run in paragraphs[paragraphs.Count - 2].ChildElements.OfType<Run>())
            {
                if (run.RunProperties != null && run.RunProperties?.Italic != new Italic())
                {
                    result.Add(new DocumentValidationError() { ParagraphId = paragraphs[paragraphs.Count - 2].ParagraphId, ErrorType = ErrorType.GridError, Position = paragraphs.Count - 2, Description = "Некорректный заголовок, необходим курсивный шрифт" });
                    break;
                }
            }
            foreach (var run in paragraphs[paragraphs.Count - 1].ChildElements.OfType<Run>())
            {
                if (run.RunProperties != null && run.RunProperties?.Italic != new Italic())
                {
                    result.Add(new DocumentValidationError() { ParagraphId = paragraphs[paragraphs.Count - 1].ParagraphId, ErrorType = ErrorType.GridError, Position = paragraphs.Count - 1, Description = "Некорректный заголовок, необходим курсивный шрифт" });
                    break;
                }
            }

            ///проверка форматирования текста параграфа
            ///
            foreach (var paragraph in paragraphs)
            {
                ///свойства параграфа должны быть прописаны в ParagraphProperties
                ///
                if (paragraph.ParagraphProperties == null)
                    result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}" });
                else
                {
                    ///в параграфе свойства для проверки - отступы текста и межстрочный интервал
                    ///
                    ///межстрочный интервал
                    ///если свойство не определено - ошибка
                    ///
                    if (paragraph.ParagraphProperties.SpacingBetweenLines == null)
                        result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, межстрочный интервал" });
                    else
                    {
                        ///проверка значения свойства, должна совпадать с записанной в классе распоряжения rule, свойства ParagraphIndent, ParagraphSpacing, LineSpacing, Fields
                        ///если не совпадает - добавить в список ошибок
                        ///
                        ///межстрочный интервал
                        ///
                        if (paragraph.ParagraphProperties.SpacingBetweenLines.Line != new StringValue((rule.LineSpacing * 240).ToString()))
                            result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, межстрочный интервал, ожидаемое значение {rule.LineSpacing * 240}" });
                        ///расстояние между параграфами
                        ///отступ от вышестоящего параграфа
                        ///
                        if (paragraph.ParagraphProperties.SpacingBetweenLines.Before != new StringValue((rule.ParagraphSpacing * 240).ToString()))
                            result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, межстрочный интервал, ожидаемое значение {rule.ParagraphSpacing * 240}" });
                        ///отступ от следующего параграфа
                        ///
                        if (paragraph.ParagraphProperties.SpacingBetweenLines.After != new StringValue((rule.ParagraphSpacing * 240).ToString()))
                            result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, межстрочный интервал, ожидаемое значение {rule.ParagraphSpacing * 240}" });
                    }

                    ///отступы от полей документа
                    ///если свойство не определено - ошибка
                    ///
                    
                }


                
                


                foreach (var run in paragraph.ChildElements.OfType<Run>().ToList())
                {
                    ///формат сносок
                    ///ищем в документе ссылки на сноски, по id сноски выбираем элемент и проверяем его на соответствие форматированию
                    ///
                    if (run.ChildElements.OfType<FootnoteReference>().Count() > 0)
                    {
                        foreach (var noteref in run.ChildElements.OfType<FootnoteReference>().ToList())
                        {
                            if (doc.GetPartsOfType<FootnotesPart>().Count() > 0)
                            {
                                ///список всех сносок документа
                                ///
                                var footnotes = doc.GetPartsOfType<FootnotesPart>().First().Footnotes.ChildElements.OfType<Footnote>();

                                if (footnotes.Count() > 0)
                                {
                                    ///ищем сноску с нужным id
                                    ///
                                    var note = footnotes.Where(x => x.Id == noteref.Id).FirstOrDefault();
                                    if (note != null)
                                    {
                                        ///получаем список параграфов в сноске
                                        ///
                                        var noteparagraphs = note.ChildElements.OfType<Paragraph>();
                                        if (noteparagraphs.Count() > 0)
                                        {
                                            foreach (var par in noteparagraphs)
                                            {
                                                ///ищем все Run в параграфе
                                                ///
                                                var runs = par.ChildElements.OfType<Run>().Count() > 0 ? par.ChildElements.OfType<Run>() : null;
                                                if (runs != null)
                                                {
                                                    foreach (var r in runs)
                                                    {
                                                        ///если свойства Run не заданы - ошибка
                                                        ///
                                                        if (r.RunProperties == null)
                                                            result.Add(new DocumentValidationError() { ParagraphId = par.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, формат сноски" });
                                                        {
                                                            ///размер шрифта
                                                            ///
                                                            if (r.RunProperties.FontSize.Val != new StringValue(rule.FootNote.Size.ToString()))
                                                                result.Add(new DocumentValidationError() { ParagraphId = par.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, формат сноски, ожидаемое значение {rule.FootNote.Size}" });
                                                            ///стиль шрифта
                                                            ///
                                                            if (r.RunProperties.RunFonts.Ascii != new StringValue(rule.FootNote.Style))
                                                                result.Add(new DocumentValidationError() { ParagraphId = par.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, формат сноски, ожидаемое значение {rule.FootNote.Style}" });
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (paragraph.ParagraphId != paragraphs[0].ParagraphId
                    && paragraph.ParagraphId != paragraphs[1].ParagraphId
                    && paragraph.ParagraphId != paragraphs[2].ParagraphId
                    && paragraph.ParagraphId != paragraphs[3].ParagraphId
                    && paragraph.ParagraphId != paragraphs[paragraphs.Count - 2].ParagraphId
                    && paragraph.ParagraphId != paragraphs[paragraphs.Count - 1].ParagraphId)
                    {
                        if (run.RunProperties == null)
                            result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, неверное форматирование текста" });
                        else
                        {
                            ///размер шрифта
                            ///
                            if (run.RunProperties.FontSize == null) result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, неверное форматирование текста" });
                            else
                            {
                                if (run.RunProperties.FontSize?.Val != new StringValue(rule.FontSize.ToString()))
                                    result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, формат сноски, ожидаемое значение {rule.FontSize}" });
                            }

                            ///стиль шрифта
                            ///
                            if (run.RunProperties.RunFonts?.Ascii == null) result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, неверное форматирование текста" });
                            else
                            {
                                if (run.RunProperties.RunFonts?.Ascii != new StringValue(rule.Font))
                                    result.Add(new DocumentValidationError() { ParagraphId = paragraph.ParagraphId, ErrorType = ErrorType.StyleError, Position = -1, Description = $"Отсутствует форматирование текста параграфа {paragraph.ParagraphId}, формат сноски, ожидаемое значение {rule.Font}" });
                            }

                        }
                    }
                    ///формат текста
                    ///для каждого Run в параграфе проверить наличие свойств и их соответствие свойствам класса rule
                    
                }
            }

            return result;
        }
    }
}