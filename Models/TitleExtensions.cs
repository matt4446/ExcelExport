using System;
using System.Linq;
using System.Text;

namespace ExcelExport.Models
{
    public static class TitleExtensions
    { 
        public static string CleanTitle(this string s, int? maxLength)
        {
            StringBuilder sb = new StringBuilder(s);

            sb.Replace(@"\", "");
            sb.Replace("/", "");
            sb.Replace("?", " ");
            sb.Replace("*", "-");
            sb.Replace("[", "");
            sb.Replace("]", "");

            return (sb.Length >= 20) ? sb.ToString().Substring(0, maxLength.GetValueOrDefault(20)) : 
                   sb.ToString();
        }
    }
}