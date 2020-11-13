using System;
using System.Collections.Generic;
using System.Text;

namespace KABService.Helper
{
    class UtilHelper
    {
        public static string ConvertToLatin1(string _source)
        {
            Encoding iso = Encoding.GetEncoding("ISO-8859-1"); // Latin 1 code page for Danish language
            Encoding utf8 = Encoding.UTF8;
            byte[] utfBytes = utf8.GetBytes(_source);
            byte[] isoBytes = Encoding.Convert(utf8, iso, utfBytes);
            return iso.GetString(isoBytes);
        }
    }
}
