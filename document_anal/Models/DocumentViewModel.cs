using document_anal.Middleware.Enums;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;

namespace document_anal.Models
{
    public class DocumentValidationError
    {
        public string ParagraphId { get; set; }
        public int Position { get; set; }
        public ErrorType ErrorType { get; set; }
        public string Description { get; set; }
    }
    public class DocumentMemoryStream
    {
        public byte[] MemoryStream { get; set; }
        public Guid Guid { get; set; }
        public string FileName { get; set; }
    }
    public class DocumentViewModel
    {
        public DocumentType DocumentType { get; set; }
        public string Name { get; set; }
        [AllowHtml]
        public string Content { get; set; }
        public string Style { get; set; }
        public List<DocumentValidationError> ValidationErrors { get; set; }
        public Guid CurrentGuid { get; set; }
        public List<DocumentMemoryStream> Documents { get; set; }
    }
}