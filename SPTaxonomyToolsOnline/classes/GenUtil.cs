using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Xml.Serialization;
using System.Xml;
using System.Security;

namespace BandR
{

    public static class ObjectExtensions
    {

        public static bool IsNull(this object o)
        {
            return o == null ? true : (o.ToString().Trim().Length <= 0 ? true : false);
        }

        public static string SafeTrim(this object o)
        {
            return o.IsNull() ? "" : o.ToString().Trim();
        }

        public static string SafeToUpper(this object o)
        {
            return o.IsNull() ? "" : o.ToString().Trim().ToUpper();
        }

        public static bool IsEqual(this object o, object o2)
        {
            if (o == null || o2 == null)
            {
                return false;
            }
            else
            {
                return o.SafeToUpper() == o2.SafeToUpper();
            }
        }

        public static string CombineFS(this string s1, string s2)
        {
            return CombineFS(s1, s2, "\\");
        }

        public static string CombineWeb(this string s1, string s2)
        {
            return CombineFS(s1, s2, "/");
        }

        private static string CombineFS(string s1, string s2, string separator)
        {
            if (s1.IsNull())
                return s2.SafeTrim();
            else if (s2.IsNull())
                return s1.SafeTrim();
            else
            {
                return s1.SafeTrim().TrimEnd(separator.ToCharArray()) + separator + s2.SafeTrim().TrimStart(separator.ToCharArray());
            }
        }

    }

    public static class JsonExtensionMethod
    {

        /// <summary>
        /// </summary>
        public static string ToJson(this object o)
        {
            System.Web.Script.Serialization.JavaScriptSerializer jsonSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();

            return jsonSerializer.Serialize(o);
        }

        /// <summary>
        /// </summary>
        public static T FromJson<T>(this string s)
        {
            System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();

            return serializer.Deserialize<T>(s);
        }

    }

    public class GenUtil
    {

        public class XmlSerialization
        {

            /// <summary>
            /// </summary>
            public static string Serialize<T>(T value)
            {
                if (value == null)
                {
                    return null;
                }

                var serializer = new XmlSerializer(typeof(T));

                var settings = new XmlWriterSettings();
                settings.Encoding = new UnicodeEncoding(false, false); // no BOM in a .NET string
                settings.Indent = false;
                settings.OmitXmlDeclaration = false;

                using (var textWriter = new StringWriter())
                {
                    using (var xmlWriter = XmlWriter.Create(textWriter, settings))
                    {
                        serializer.Serialize(xmlWriter, value);
                    }
                    return textWriter.ToString();
                }
            }

            /// <summary>
            /// </summary>
            public static T Deserialize<T>(string xml)
            {
                if (string.IsNullOrEmpty(xml))
                {
                    return default(T);
                }

                var serializer = new XmlSerializer(typeof(T));

                var settings = new XmlReaderSettings();

                using (var textReader = new StringReader(xml))
                {
                    using (var xmlReader = XmlReader.Create(textReader, settings))
                    {
                        return (T)serializer.Deserialize(xmlReader);
                    }
                }
            }

        }

        /// <summary>
        /// </summary>
        public static string CleanFilenameForFS(string s)
        {
            var invalids = System.IO.Path.GetInvalidFileNameChars();
            var newName = String.Join("", s.Split(invalids, StringSplitOptions.RemoveEmptyEntries));
            return newName;
        }

        /// <summary>
        /// </summary>
        public static string CleanFilenameForSP(string s, string r)
        {
            var pattern = string.Concat("[", @"\~\""\#\%\&\*\:\<\>\?\/\\\{\|\}", "]"); // ~ " # % & * : < > ? / \ { | }
            return Regex.Replace(s, pattern, r);
        }

        /// <summary>
        /// </summary>
        public static SecureString BuildSecureString(string s)
        {
            var securePassword = new SecureString();
            foreach (char c in s)
            {
                securePassword.AppendChar(c);
            }
            return securePassword;
        }

        /// <summary>
        /// </summary>
        public static string CombineFileSysPaths(object path1, object path2)
        {
            if (IsNull(path1) && IsNull(path2))
            {
                return "";
            }
            else if (IsNull(path1))
            {
                return SafeTrim(path2);
            }
            else if (IsNull(path2))
            {
                return SafeTrim(path1);
            }
            else
            {
                return string.Concat(SafeTrim(path1).TrimEnd(new char[] { '\\' }), "\\", SafeTrim(path2).TrimStart(new char[] { '\\' }));
            }
        }

        /// <summary>
        /// </summary>
        public static string CombinePaths(object path1, object path2)
        {
            if (IsNull(path1) && IsNull(path2))
            {
                return "";
            }
            else if (IsNull(path1))
            {
                return SafeTrim(path2);
            }
            else if (IsNull(path2))
            {
                return SafeTrim(path1);
            }
            else
            {
                return string.Concat(SafeTrim(path1).TrimEnd(new char[] { '/' }), "/", SafeTrim(path2).TrimStart(new char[] { '/' }));
            }
        }

        /// <summary>
        /// Use this before saving to MMD, these chars are illegal in MMD, must be removed/replaced.
        /// </summary>
        public static string MmdRemoveIllegalChars(string s)
        {
            s = s.Replace('"', '\'');
            s = s.Replace(';', ',');
            s = s.Replace('<', '(');
            s = s.Replace('>', ')');
            s = s.Replace('|', '/');

            return s;
        }

        /// <summary>
        /// Only use this for comparisons, never for normalizing data to save in MMD (this is done automatically).
        /// </summary>
        //public static string MmdNormalizeForComparison(object o)
        //{
        //    // SharePoint internally removes multiple spaces, and replaces & and " with unicode equivalents.
        //    return Term.NormalizeName(MmdRemoveIllegalChars(SafeTrim(o)));
        //}

        /// <summary>
        /// Not used.
        /// </summary>
        public static string MmdDenormalize(object o)
        {
            return GenUtil.SafeTrim(o)
                .Replace(Convert.ToChar(char.ConvertFromUtf32(65286)), '&')
                .Replace(Convert.ToChar(char.ConvertFromUtf32(65282)), '"');
        }

        /// <summary>
        /// </summary>
        public static string CleanUsername(object username)
        {
            var un = SafeTrim(username);

            un = un.Substring(un.LastIndexOf('#') + 1);
            un = un.Substring(un.LastIndexOf('|') + 1);

            return un;
        }

        /// <summary>
        /// </summary>
        public static string SafeTrimLookupFieldValue(object o)
        {
            if (IsNull(o))
            {
                return "";
            }
            else
            {
                return SafeTrim(o).Substring(SafeTrim(o).IndexOf('#') + 1);
            }
        }

        /// <summary>
        /// </summary>
        public static string EnsureStartsWithForwardSlash(string s)
        {
            return string.Concat("/", s.SafeTrim().TrimStart("/".ToCharArray()));
        }

        /// <summary>
        /// </summary>
        public static List<string> ConvertStringToList(string str)
        {
            var lst = new List<string>();

            // normalize delimiters, shoul be ";"
            str = GenUtil.SafeTrim(str).Replace(",", ";");

            return ConvertStringToList(str, ";");
        }

        /// <summary>
        /// </summary>
        public static List<string> ConvertStringToList(string str, string delimiter)
        {
            var lst = new List<string>();

            str = GenUtil.SafeTrim(str);

            if (str.Contains(delimiter))
            {
                lst.AddRange(str.Split(new char[] { Convert.ToChar(delimiter) }));
            }
            else
            {
                lst.Add(str);
            }

            return lst.Where(x => x.Trim().Length > 0).Select(x => x.Trim()).Distinct().ToList();
        }

        /// <summary>
        /// </summary>
        public static string NVL(object a, object b)
        {
            if (!IsNull(a))
            {
                return SafeTrim(a);
            }
            else if (!IsNull(b))
            {
                return SafeTrim(b);
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// </summary>
        public static string ToNull(object x)
        {
            if ((x == null)
                || (Convert.IsDBNull(x))
                || x.ToString().Trim().Length == 0)
                return null;
            else
                return x.ToString();
        }

        /// <summary>
        /// </summary>
        public static bool IsNull(object x)
        {
            if ((x == null)
                || (Convert.IsDBNull(x))
                || x.ToString().Trim().Length == 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// </summary>
        public static string SafeTrim(object x)
        {
            if (IsNull(x))
                return "";
            else
                return x.ToString().Trim();
        }

        /// <summary>
        /// Case insensitive comparison.
        /// </summary>
        public static bool IsEqual(object o1, object o2)
        {
            return string.Compare(SafeTrim(o1), SafeTrim(o2), true) == 0;
        }

        /// <summary>
        /// If not valid returns 0.
        /// </summary>
        public static int SafeToInt(object o)
        {
            if (IsNull(o))
                return 0;
            else
            {
                if (IsInt(o))
                    return int.Parse(o.ToString());
                else
                    return 0;
            }
        }

        /// <summary>
        /// Return 0 if null, or the int.
        /// </summary>
        public static int SafeNullIntToInt(object o)
        {
            return ((int?)o).HasValue ? ((int?)o).Value : 0;
        }

        /// <summary>
        /// If not valid returns 0.
        /// </summary>
        public static double SafeToDouble(object o)
        {
            if (IsNull(o))
                return 0;
            else
            {
                double test;
                if (!double.TryParse(o.ToString(), out test))
                    return 0;
                else
                    return test;
            }
        }

        /// <summary>
        /// If not valid returns 0.
        /// </summary>
        public static decimal SafeToDecimal(object o)
        {
            if (IsNull(o))
                return 0;
            else
            {
                decimal test;
                if (!decimal.TryParse(o.ToString(), out test))
                    return 0;
                else
                    return test;
            }
        }

        /// <summary>
        /// If not valid returns false.
        /// </summary>
        public static bool SafeToBool(object o)
        {
            if (SafeToUpper(o) == "1" ||
                SafeToUpper(o) == "YES" ||
                SafeToUpper(o) == "Y" ||
                SafeToUpper(o) == "TRUE")
                return true;
            else
                return false;
        }

        /// <summary>
        /// </summary>
        public static bool IsBool(object o)
        {
            o = SafeToUpper(o);
            return
                (o.ToString() == "1" || o.ToString() == "0" ||
                o.ToString() == "YES" || o.ToString() == "NO" ||
                o.ToString() == "Y" || o.ToString() == "N" ||
                o.ToString() == "TRUE" || o.ToString() == "FALSE");
        }

        /// <summary>
        /// If not valid returns 01/01/1900 12:00:00 AM.
        /// </summary>
        public static DateTime SafeToDateTime(object o)
        {
            if (IsNull(o))
                return DateTime.Parse("01/01/1900 12:00:00 AM");
            else
            {
                DateTime dummy;

                if (IsInt(o))
                    return new DateTime(Convert.ToInt64(o)); // use ticks
                else
                {
                    if (DateTime.TryParse(o.ToString(), out dummy))
                        return dummy;
                    else
                        return DateTime.Parse("01/01/1900 12:00:00 AM");
                }
            }
        }

        /// <summary>
        /// </summary>
        public static bool IsInt(object o)
        {
            if (IsNull(o))
                return false;

            Int64 dummy = 0;
            return Int64.TryParse(o.ToString(), out dummy);
        }

        /// <summary>
        /// </summary>
        public static bool IsDouble(object o)
        {
            if (IsNull(o))
                return false;

            Double dummy = 0;
            return Double.TryParse(o.ToString(), out dummy);
        }

        /// <summary>
        /// Trims and converts to upper case.
        /// </summary>
        public static string SafeToUpper(object o)
        {
            if (IsNull(o))
                return "";
            else
                return SafeTrim(o).ToUpper();
        }

        /// <summary>
        /// </summary>
        public static string SafeToProperCase(object o)
        {
            if (IsNull(o))
                return "";
            else
            {
                if (o.ToString().Trim().Length == 1)
                    return o.ToString().Trim().ToUpper();
                else
                    return o.ToString().Trim().Substring(0, 1).ToUpper() + o.ToString().Trim().Substring(1).ToLower();
            }
        }

        /// <summary>
        /// </summary>
        public static string SafeGetArrayVal(string[] list, int index)
        {
            return index < list.Length ? SafeTrim(list[index]) : "";
        }

        /// <summary>
        /// </summary>
        public static string SafeGetArrayVal(List<string> list, int index)
        {
            return index < list.Count ? SafeTrim(list[index]) : "";
        }

        /// <summary>
        /// </summary>
        public static Guid? SafeToGuid(object o)
        {
            if (IsGuid(o))
                return new Guid(SafeTrim(o));
            else
                return null;
        }

        /// <summary>
        /// </summary>
        public static bool IsGuid(object o)
        {
            if (IsNull(o))
                return false;

            try
            {
                var tmp = new Guid(SafeTrim(o));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// </summary>
        public static string NormalizeEol(string s)
        {
            return Regex.Replace(SafeTrim(s), @"\r\n|\n\r|\n|\r", "\r\n");
        }

        /// <summary>
        /// Converts null to "n/a", YES/1/TRUE to "Yes", otherwise "No".
        /// </summary>
        public static string TransformToYes(object o)
        {
            if (IsNull(o))
                return "n/a";
            else
            {
                if (SafeToUpper(o) == "YES" || SafeToUpper(o) == "Y" || SafeToUpper(o) == "1" || SafeToUpper(o) == "TRUE" || SafeToUpper(o) == "T")
                    return "Yes";
                else
                    return "No";
            }
        }

        /// <summary>
        /// </summary>
        public static string TransformToYesFlipped(object o)
        {
            if (TransformToYes(o) == "Yes")
                return "No";
            else if (TransformToYes(o) == "No")
                return "Yes";
            else
                return "n/a";
        }

        /// <summary>
        /// If null, shows "n/a".
        /// </summary>
        public static string TransformToNA(object o)
        {
            if (IsNull(o))
                return "n/a";
            else
                return SafeTrim(o);
        }

        /// <summary>
        /// </summary>
        public static string TransformToNADate(object o)
        {
            if (SafeToDateTime(o).Year == 1900)
                return "n/a";
            else
                return SafeToDateTime(o).ToShortDateString();
        }

        /// <summary>
        /// </summary>
        public static string TransformToNADateTime(object o)
        {
            if (SafeToDateTime(o).Year == 1900)
                return "n/a";
            else
                return SafeToDateTime(o).ToString();
        }

        /// <summary>
        /// </summary>
        public static string TransformToDoubleFormat(object o)
        {
            return SafeToDouble(o).ToString("###,###,###,###.###");
        }

        /// <summary>
        /// </summary>
        public static string TransformToBr(string val)
        {
            if (IsNull(val))
                return "";

            return Regex.Replace(NormalizeEol(SafeTrim(val)), "\r\n", "<br/>", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Convert string to base64, then rot13 it.
        /// </summary>
        public static string SecureBase64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return Rot13Transform(System.Convert.ToBase64String(plainTextBytes));
        }

        /// <summary>
        /// Decode rot13 base64 string.
        /// </summary>
        public static string SecureBase64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(Rot13Transform(base64EncodedData));
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        /// <summary>
        /// </summary>
        private static string Rot13Transform(string value)
        {
            char[] array = value.ToCharArray();
            for (int i = 0; i < array.Length; i++)
            {
                int number = (int)array[i];

                if (number >= 'a' && number <= 'z')
                {
                    if (number > 'm')
                    {
                        number -= 13;
                    }
                    else
                    {
                        number += 13;
                    }
                }
                else if (number >= 'A' && number <= 'Z')
                {
                    if (number > 'M')
                    {
                        number -= 13;
                    }
                    else
                    {
                        number += 13;
                    }
                }

                array[i] = (char)number;
            }
            return new string(array);
        }

    }
}
