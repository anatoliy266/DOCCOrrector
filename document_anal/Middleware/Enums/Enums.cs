using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace document_anal.Middleware.Enums
{
    public enum ErrorType
    {
        LogoError = 0,
        TextError,
        StyleError,
        GridError,
    }

    public enum DocumentType
    {
        Распоряжение,
        Служебная_записка,
        Письмо,
        Другое,
    }
}