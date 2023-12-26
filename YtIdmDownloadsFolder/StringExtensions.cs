using System.Text.RegularExpressions;

namespace YtIdmDownloadsFolder
{
    public static class StringExtensions
    {
        public static string RemoveSpecialCharacters(this string str)
        {
            return Regex.Replace(str, @"[^\w\d\s]|_|-", "");
        }        
        
        public static string ValidateFileName(this string str)
        {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                str = str.Replace(c, '_');
            }

            return str;
        }
    }
}