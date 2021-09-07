using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace document_anal.Middleware.Models
{
    public interface IDocumentStyleModel
    {

    }
    public class DocumentStyleModel: IDocumentStyleModel
    {
        /// <summary>
        /// Parse string-style styles from document
        /// </summary>
        /// <param name="docStyle">string-style document styles</param>
        /// <returns>DocumentStyleModel object</returns>
        public DocumentStyleModel FromString(string docStyle)
        {
            return this;
        }

    }
}