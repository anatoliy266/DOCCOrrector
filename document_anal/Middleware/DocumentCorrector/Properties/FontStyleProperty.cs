using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace document_anal.Middleware.WordCorrector.Properties
{
    public class FontStyleProperty
    {
        public void Add(RunProperties runProperties, OnOffType type)
        {
            runProperties.AppendChild(type);
            //runProperties.AppendChild(typeComplexScripts[value]);
        }
    }


}
