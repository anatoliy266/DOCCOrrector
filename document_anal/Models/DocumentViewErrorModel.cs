using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace document_anal.Models
{
    public class DocumentViewErrorModel
    {
        public string Position { get; set; }
        public string Description { get; set; }
        public string StackTrace { get; set; }
    }
}