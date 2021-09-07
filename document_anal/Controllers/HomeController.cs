using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using document_anal.Middleware.DocumentConverter;
using document_anal.Middleware.Enums;
using document_anal.Middleware.Models;
using document_anal.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace document_anal.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// Стартовая страница
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Предварительная обработка документа, добавление идентификаторов для параграфов
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        private byte[] Preprocess(byte[] bytes)
        {
            byte[] b = new byte[0];
            using (var ms = new MemoryStream())
            {

                ms.Write(bytes, 0, bytes.Length);
                ms.Position = 0;
                using (var doc = WordprocessingDocument.Open(ms, true))
                {
                    ///для каждого параграфа, если отстутсвует ParagraphId - присваиваем параграфу хэшкод текущего класса
                    foreach (var element in doc.MainDocumentPart.Document.Body.ChildElements)
                    {
                        if (element is Paragraph)
                        {
                            if ((element as Paragraph).ParagraphId == null)
                                (element as Paragraph).ParagraphId = HexBinaryValue.FromString(element.GetHashCode().ToString());
                        }
                    }
                }
                ///конвертируем в массив байт для дальнейшей обработки
                b = ms.ToArray();
            }
            return b;
        }
        /// <summary>
        /// Загрузка проверка и отображение файла на форме
        /// </summary>
        /// <param name="model"></param>
        /// <param name="upload"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult LoadFile(DocumentViewModel model, HttpPostedFileBase upload)
        {
            try
            {
                ///Загрузка 
                ///если файл не загружен возвращаем на форрму с сообщением
                if (upload != null)
                {
                    byte[] bytes;
                    var buffer = new byte[4 * 1024];
                    using (var ms = new MemoryStream())
                    {
                        int read;
                        while ((read = upload.InputStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            ms.Write(buffer, 0, read);
                        }
                        bytes = ms.ToArray();
                    }

                    if (bytes.Length > 0)
                    {
                        ///добавление идентификаторов параграфам
                        ///
                        bytes = Preprocess(bytes);

                        ///выбор типа документа для обработки
                        ///
                        switch (model.DocumentType)
                        {
                            case Middleware.Enums.DocumentType.Распоряжение:
                                {
                                    ///обработка документа
                                    ///
                                    var converter = new OrderConverter();
                                    ///конвертируем в HTML код
                                    ///
                                    var content = converter.ToHtml(upload.FileName, bytes, model.DocumentType, out var memstream);
                                    ///присваиваем документу уникальный идентификатор
                                    ///
                                    var docGuid = Guid.NewGuid();
                                    ///забираем ранее загруженные документы из модели
                                    ///
                                    content.Documents = model.Documents;
                                    if (content.Documents == null)
                                        content.Documents = new List<DocumentMemoryStream>();
                                    ///добавляем к ранее загружнным новый документ
                                    ///
                                    content.Documents.Add(new DocumentMemoryStream() { Guid = docGuid, MemoryStream = memstream, FileName = upload.FileName });
                                    ///устанавливаем загруженный документ текущим
                                    ///
                                    content.CurrentGuid = docGuid;
                                    content.Name = upload.FileName;
                                    return View("Index", content);
                                }
                            case Middleware.Enums.DocumentType.Служебная_записка:
                                {
                                    ///обработка документа
                                    ///
                                    var converter = new ServiceNoteConverter();
                                    ///конвертируем в HTML код
                                    ///
                                    var content = converter.ToHtml(upload.FileName, bytes, model.DocumentType, out var memstream);
                                    ///присваиваем документу уникальный идентификатор
                                    ///
                                    var docGuid = Guid.NewGuid();
                                    ///забираем ранее загруженные документы из модели
                                    ///
                                    content.Documents = model.Documents;
                                    if (content.Documents == null)
                                        content.Documents = new List<DocumentMemoryStream>();
                                    ///добавляем к ранее загружнным новый документ
                                    ///
                                    content.Documents.Add(new DocumentMemoryStream() { Guid = docGuid, MemoryStream = memstream, FileName = upload.FileName });
                                    ///устанавливаем загруженный документ текущим
                                    ///
                                    content.CurrentGuid = docGuid;
                                    content.Name = upload.FileName;
                                    return View("Index", content);
                                }
                            case Middleware.Enums.DocumentType.Письмо:
                                {
                                    ///обработка документа
                                    ///
                                    var converter = new MailConverter();
                                    ///конвертируем в HTML код
                                    ///
                                    var content = converter.ToHtml(upload.FileName, bytes, model.DocumentType, out var memstream);
                                    ///присваиваем документу уникальный идентификатор
                                    ///
                                    var docGuid = Guid.NewGuid();
                                    ///забираем ранее загруженные документы из модели
                                    ///
                                    content.Documents = model.Documents;
                                    if (content.Documents == null)
                                        content.Documents = new List<DocumentMemoryStream>();
                                    ///добавляем к ранее загружнным новый документ
                                    ///
                                    content.Documents.Add(new DocumentMemoryStream() { Guid = docGuid, MemoryStream = memstream, FileName = upload.FileName });
                                    ///устанавливаем загруженный документ текущим
                                    ///
                                    content.CurrentGuid = docGuid;
                                    content.Name = upload.FileName;
                                    return View("Index", content);
                                }
                            default:
                                {
                                    ///обработка документа
                                    ///
                                    var converter = new DefaultConverter();
                                    ///конвертируем в HTML код
                                    ///
                                    var content = converter.ToHtml(upload.FileName, bytes, model.DocumentType, out var memstream);
                                    ///присваиваем документу уникальный идентификатор
                                    ///
                                    var docGuid = Guid.NewGuid();
                                    ///забираем ранее загруженные документы из модели
                                    ///
                                    content.Documents = model.Documents;
                                    if (content.Documents == null)
                                        content.Documents = new List<DocumentMemoryStream>();
                                    ///добавляем к ранее загружнным новый документ
                                    ///
                                    content.Documents.Add(new DocumentMemoryStream() { Guid = docGuid, MemoryStream = memstream, FileName = upload.FileName });
                                    ///устанавливаем загруженный документ текущим
                                    ///
                                    content.CurrentGuid = docGuid;
                                    content.Name = upload.FileName;
                                    return View("Index", content);
                                }
                        }
                    } else
                    {
                        ViewBag.Message = "Файл не загружен, попробуйте еще раз";

                        return View("Index", model);
                    }
                }
                else
                {
                    ViewBag.Message = "Файл не загружен, попробуйте еще раз";
                    return View("Index");
                }
            } catch (Exception e)
            {
                return View("Error", new DocumentViewErrorModel() { StackTrace = e.StackTrace, Description = e.Message, Position = e.Source });
            }
        }

        /// <summary>
        /// Обработка ошибок документа
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Correction(DocumentViewModel model)
        {
            try
            {
                ///получаем текущий документ из можели
                ///
                var memStream = model.Documents.Where(x => x.Guid == model.CurrentGuid).First().MemoryStream;
                ///выбор типа документа для обработки
                ///
                switch (model.DocumentType)
                {
                    case Middleware.Enums.DocumentType.Распоряжение:
                        {
                            ///обработка ошибок
                            ///
                            var converter = new OrderConverter();
                            var content = converter.Fix(model.Name, memStream, model.ValidationErrors, out var stream);
                            ///заполняем модель ранее загруженными документами
                            ///
                            content.Documents = model.Documents;
                            ///подменяем документ на измененный
                            ///
                            content.Documents.Where(x => x.Guid == model.CurrentGuid).First().MemoryStream = stream;
                            content.CurrentGuid = model.CurrentGuid;
                            content.Name = model.Name;
                            return View("Index", content);
                        }
                    case Middleware.Enums.DocumentType.Служебная_записка:
                        {
                            ///обработка ошибок
                            ///
                            var converter = new ServiceNoteConverter();
                            var content = converter.Fix(model.Name, memStream, model.ValidationErrors, out var stream);
                            ///заполняем модель ранее загруженными документами
                            ///
                            content.Documents = model.Documents;
                            ///подменяем документ на измененный
                            ///
                            content.Documents.Where(x => x.Guid == model.CurrentGuid).First().MemoryStream = stream;
                            content.CurrentGuid = model.CurrentGuid;
                            content.Name = model.Name;
                            return View("Index", content);
                        }
                    case Middleware.Enums.DocumentType.Письмо:
                        {
                            ///обработка ошибок
                            ///
                            var converter = new MailConverter();
                            var content = converter.Fix(model.Name, memStream, model.ValidationErrors, out var stream);
                            ///заполняем модель ранее загруженными документами
                            ///
                            content.Documents = model.Documents;
                            ///подменяем документ на измененный
                            ///
                            content.Documents.Where(x => x.Guid == model.CurrentGuid).First().MemoryStream = stream;
                            content.CurrentGuid = model.CurrentGuid;
                            content.Name = model.Name;
                            return View("Index", content);
                        }
                    default:
                        {
                            ///обработка ошибок
                            ///
                            var converter = new DefaultConverter();
                            var content = converter.Fix(model.Name, memStream, model.ValidationErrors, out var stream);
                            ///заполняем модель ранее загруженными документами
                            ///
                            content.Documents = model.Documents;
                            ///подменяем документ на измененный
                            ///
                            content.Documents.Where(x => x.Guid == model.CurrentGuid).First().MemoryStream = stream;
                            content.CurrentGuid = model.CurrentGuid;
                            content.Name = model.Name;
                            return View("Index", content);
                        }
                }
            } catch (Exception e)
            {
                return View("Error", new DocumentViewErrorModel() { StackTrace = e.StackTrace, Description = e.Message, Position = e.Source });
            }
        }

        /// <summary>
        /// Изменение контекста страницы
        /// </summary>
        /// <param name="model"></param>
        /// <param name="guid"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ChangeCurrentDocument(DocumentViewModel model, Guid guid)
        {
            ///выбор типа документа для обработки
            ///
            switch (model.DocumentType)
            {
                case Middleware.Enums.DocumentType.Распоряжение:
                    {
                        var converter = new OrderConverter();
                        ///получение документа и названия документа из модели
                        ///
                        var currentMemStream = model.Documents.Where(x => x.Guid == guid).First().MemoryStream;
                        var currentDocName = model.Documents.Where(x => x.Guid == guid).First().FileName;
                        ///Конвертация документа в html
                        ///
                        var content = converter.ToHtml(currentDocName, currentMemStream, model.DocumentType, out var memstream);
                        ///заполнение ранее загруженных документов
                        ///
                        content.Documents = model.Documents;
                        ///замена документа на измененный
                        ///
                        content.Documents.Where(x => x.Guid == guid).First().MemoryStream = memstream;
                        ///устанавливаем измененный документ текущим
                        ///
                        content.CurrentGuid = guid;
                        return View("Index", content);
                    }
                case Middleware.Enums.DocumentType.Служебная_записка:
                    {
                        var converter = new ServiceNoteConverter();
                        ///получение документа и названия документа из модели
                        ///
                        var currentMemStream = model.Documents.Where(x => x.Guid == guid).First().MemoryStream;
                        var currentDocName = model.Documents.Where(x => x.Guid == guid).First().FileName;
                        ///Конвертация документа в html
                        ///
                        var content = converter.ToHtml(currentDocName, currentMemStream, model.DocumentType, out var memstream);
                        ///заполнение ранее загруженных документов
                        ///
                        content.Documents = model.Documents;
                        ///замена документа на измененный
                        ///
                        content.Documents.Where(x => x.Guid == guid).First().MemoryStream = memstream;
                        ///устанавливаем измененный документ текущим
                        ///
                        content.CurrentGuid = guid;
                        return View("Index", content);
                    }
                case Middleware.Enums.DocumentType.Письмо:
                    {
                        var converter = new MailConverter();
                        ///получение документа и названия документа из модели
                        ///
                        var currentMemStream = model.Documents.Where(x => x.Guid == guid).First().MemoryStream;
                        var currentDocName = model.Documents.Where(x => x.Guid == guid).First().FileName;
                        ///Конвертация документа в html
                        ///
                        var content = converter.ToHtml(currentDocName, currentMemStream, model.DocumentType, out var memstream);
                        ///заполнение ранее загруженных документов
                        ///
                        content.Documents = model.Documents;
                        ///замена документа на измененный
                        ///
                        content.Documents.Where(x => x.Guid == guid).First().MemoryStream = memstream;
                        ///устанавливаем измененный документ текущим
                        ///
                        content.CurrentGuid = guid;
                        return View("Index", content);
                    }
                default:
                    {
                        var converter = new DefaultConverter();
                        ///получение документа и названия документа из модели
                        ///
                        var currentMemStream = model.Documents.Where(x => x.Guid == guid).First().MemoryStream;
                        var currentDocName = model.Documents.Where(x => x.Guid == guid).First().FileName;
                        ///Конвертация документа в html
                        ///
                        var content = converter.ToHtml(currentDocName, currentMemStream, model.DocumentType, out var memstream);
                        ///заполнение ранее загруженных документов
                        ///
                        content.Documents = model.Documents;
                        ///замена документа на измененный
                        ///+
                        content.Documents.Where(x => x.Guid == guid).First().MemoryStream = memstream;
                        ///устанавливаем измененный документ текущим
                        ///
                        content.CurrentGuid = guid;
                        return View("Index", content);
                    }
            }
        }

        /// <summary>
        /// Сохранение документа 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public FileResult Save(DocumentViewModel model)
        {
            return File(model.Documents.Where(x => x.Guid == model.CurrentGuid).First().MemoryStream, "application/ostet-stream", System.IO.Path.GetFileName(model.Name) + "_Проверено.docx");
        }
    }
}