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
    public interface IConverter
    {
        /// <summary>
        /// Конвертация документа в HTML с проверкой содержимого на соответствие шаблону
        /// </summary>
        /// <param name="path">Путь к файлу документа</param>
        /// <param name="type">Тип документа</param>
        /// <returns></returns>
        DocumentViewModel ToHtml(string name, Enums.DocumentType type);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="ms"></param>
        /// <param name="type"></param>
        /// <param name="currentStream"></param>
        /// <returns></returns>
        DocumentViewModel ToHtml(string name, MemoryStream ms, Enums.DocumentType type, out byte[] currentStream);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="bytes"></param>
        /// <param name="type"></param>
        /// <param name="currentStream"></param>
        /// <returns></returns>
        DocumentViewModel ToHtml(string name, byte[] bytes, Enums.DocumentType type, out byte[] currentStream);
        /// <summary>
        /// Проверка документа на соответствие шаблону и форматирование
        /// </summary>
        /// <param name="doc">пакет документа</param>
        /// <returns></returns>
        List<DocumentValidationError> Validate(WordprocessingDocument doc);
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="bytes"></param>
        /// <param name="errors"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        DocumentViewModel Fix(string name, byte[] bytes, List<DocumentValidationError> errors, out byte[] stream);

    }

}