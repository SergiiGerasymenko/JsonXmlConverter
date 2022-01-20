using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Xml;

namespace JSONlib
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("4ec26533-0f05-4eb1-bcb4-4d42410a2521")]
    [ComVisible(true)]
    public interface IJsonXmlConverter
    {
        string GetJsonTxt(object xmlObj, bool includeRoot);
        object GetJsonStream(object xmlObj, bool includeRoot);
        string GetXmlTxt(object jsonObj, string root);
        object GetXmlStream(object jsonObj, string root);
    }


    [ClassInterface(ClassInterfaceType.None)]
    [Guid("fc0f4e8e-5668-459b-ae69-9e90881cb8b2")]
    [ComVisible(true)]
    public class JsonXmlConverter : IJsonXmlConverter
    {
        public string GetJsonTxt(object xmlObj, bool includeRoot)
        {
            var xmlTxt = GetTextFromObject(xmlObj);
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlTxt);

            var jObject = JObject.FromObject(xmlDoc);

            if (!includeRoot)
                return jObject.First.First.ToString();

            return jObject.ToString();
        }

        public object GetJsonStream(object xmlObj, bool includeRoot) => CreateComStream(GetJsonTxt(xmlObj, includeRoot));

        public string GetXmlTxt(object jsonObj, string root)
        {
            string jsonTxt = GetTextFromObject(jsonObj);
            return JsonConvert.DeserializeXNode(jsonTxt, root).ToString();
        }

        public object GetXmlStream(object jsonObj, string root) => CreateComStream(GetXmlTxt(jsonObj, root));


        private string GetTextFromObject(object sourceObj)
        {
            if (sourceObj.GetType() == typeof(string))
                return (string)sourceObj;

            var comStream = sourceObj as IStream;
            if (comStream != null)
            {
                string result;
                using (StreamReader streamReader = new StreamReader(comStream.ReadToMemoryStream()))
                {
                    result = streamReader.ReadToEnd();
                }
                return result;
            }

            var xmlNode = sourceObj as MSXML2.IXMLDOMNode;
            if (xmlNode != null)
                return (sourceObj as MSXML2.IXMLDOMNode).xml;

            throw new NotImplementedException(sourceObj.GetType().ToString());
        }

        [DllImport("ole32.dll")]
        private static extern int CreateStreamOnHGlobal(
          IntPtr hGlobal,
          bool fDeleteOnRelease,
          out IStream ppstm);

        private object CreateComStream(string str)
        {
            MemoryStream memoryStream = new MemoryStream();
            TextWriter textWriter = (TextWriter)new StreamWriter((Stream)memoryStream);
            textWriter.Write(str);
            textWriter.Flush();
            byte[] numArray = new byte[memoryStream.Length];
            memoryStream.Seek(0L, SeekOrigin.Begin);
            memoryStream.Read(numArray, 0, numArray.Length);
            textWriter.Close();
            memoryStream.Close();
            IntPtr num = Marshal.AllocHGlobal(numArray.Length);
            Marshal.Copy(numArray, 0, num, numArray.Length);
            IStream ppstm;
            CreateStreamOnHGlobal(num, true, out ppstm);
            return (object)ppstm;
        }
    }
}
