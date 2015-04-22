using Maximis.Toolkit.Caching;
using Maximis.Toolkit.IO;
using System;
using System.IO;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Serialization;

namespace Maximis.Toolkit.Xml
{
    public enum SerialisationMethod { Xml, DataContract }

    public static class SerialisationHelper
    {
        private static CacheHelper<DataContractSerializer> dcSerCache = new CacheHelper<DataContractSerializer>();
        private static CacheHelper<XmlSerializer> xmlSerCache = new CacheHelper<XmlSerializer>();

        public static T DeserialiseFromFile<T>(string path, SerialisationMethod method = SerialisationMethod.Xml)
        {
            return (T)DeserialiseFromFile(path, typeof(T), method);
        }

        public static object DeserialiseFromFile(string path, Type t, SerialisationMethod method = SerialisationMethod.Xml)
        {
            if (string.IsNullOrEmpty(path)) return null;

            using (FileStream fs = File.OpenRead(path))
            {
                if (method == SerialisationMethod.Xml)
                {
                    return GetXmlSerializer(t).Deserialize(fs);
                }
                else
                {
                    return GetDataContractSerializer(t).ReadObject(fs);
                }
            }
        }

        public static T DeserialiseFromString<T>(string xml, SerialisationMethod method = SerialisationMethod.Xml)
        {
            return (T)DeserialiseFromString(xml, typeof(T), method);
        }

        public static object DeserialiseFromString(string xml, Type t, SerialisationMethod method = SerialisationMethod.Xml)
        {
            if (string.IsNullOrEmpty(xml)) return null;

            using (StringReader sr = new StringReader(xml))
            {
                if (method == SerialisationMethod.Xml)
                {
                    return GetXmlSerializer(t).Deserialize(sr);
                }
                else
                {
                    using (XmlTextReader xtr = new XmlTextReader(sr))
                    {
                        return GetDataContractSerializer(t).ReadObject(xtr);
                    }
                }
            }
        }

        public static void SerialiseToFile<T>(T obj, string path, SerialisationMethod method = SerialisationMethod.Xml)
        {
            SerialiseToFile(obj, path, typeof(T), method);
        }

        public static void SerialiseToFile(object obj, string path, SerialisationMethod method = SerialisationMethod.Xml)
        {
            SerialiseToFile(obj, path, obj.GetType(), method);
        }

        public static void SerialiseToFile(object obj, string path, Type t, SerialisationMethod method = SerialisationMethod.Xml)
        {
            using (FileStream fs = File.Create(path))
            {
                if (method == SerialisationMethod.Xml)
                {
                    GetXmlSerializer(t).Serialize(fs, obj);
                }
                else
                {
                    GetDataContractSerializer(t).WriteObject(fs, obj);
                }
                fs.Flush();
            }
        }

        public static string SerialiseToString<T>(T obj, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString(obj, typeof(T), method, extraTypes);
        }

        public static string SerialiseToString(object obj, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString(obj, obj.GetType(), method, extraTypes);
        }

        public static string SerialiseToString(object obj, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            using (StringWriterUtf8 sw = new StringWriterUtf8())
            {
                if (method == SerialisationMethod.Xml)
                {
                    GetXmlSerializer(t, extraTypes).Serialize(sw, obj);
                }
                else
                {
                    using (XmlTextWriter xtw = new XmlTextWriter(sw))
                    {
                        GetDataContractSerializer(t, extraTypes).WriteObject(xtw, obj);
                    }
                }
                sw.Flush();
                return sw.ToString();
            }
        }

        private static DataContractSerializer GetDataContractSerializer(Type t, params Type[] extraTypes)
        {
            string typeName = t.AssemblyQualifiedName;
            if (!dcSerCache.Dictionary.ContainsKey(typeName))
            {
                dcSerCache.Dictionary.Add(typeName, new DataContractSerializer(t, extraTypes));
            }
            return dcSerCache.Dictionary[typeName];
        }

        private static XmlSerializer GetXmlSerializer(Type t, params Type[] extraTypes)
        {
            string typeName = t.AssemblyQualifiedName;
            if (!xmlSerCache.Dictionary.ContainsKey(typeName))
            {
                xmlSerCache.Dictionary.Add(typeName, new XmlSerializer(t, extraTypes));
            }
            return xmlSerCache.Dictionary[typeName];
        }
    }
}