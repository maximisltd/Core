using Maximis.Toolkit.Caching;
using Maximis.Toolkit.IO;
using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml.Serialization;

namespace Maximis.Toolkit.Xml
{
    public enum SerialisationMethod { Xml, DataContract, Json }

    public static class SerialisationHelper
    {
        /// <summary>
        /// Uses Serialisation to create a copy of an object
        /// </summary>
        public static T CloneObject<T>(CacheManager cache, T obj, params Type[] extraTypes)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                SerialiseToStream<T>(cache, obj, ms, SerialisationMethod.DataContract, extraTypes);
                ms.Position = 0;
                return DeserialiseFromStream<T>(cache, ms, SerialisationMethod.DataContract, extraTypes);
            }
        }

        public static DataContractSerializer GetDataContractSerializer(CacheManager cache, Type t, params Type[] extraTypes)
        {
            DataContractSerializer dcSer = cache.Get<DataContractSerializer>(t.AssemblyQualifiedName);
            if (dcSer == null)
            {
                dcSer = new DataContractSerializer(t, extraTypes);
                cache.Set<DataContractSerializer>(t.AssemblyQualifiedName, dcSer);
            }
            return dcSer;
        }

        public static DataContractJsonSerializer GetJsonSerializer(CacheManager cache, Type t, params Type[] extraTypes)
        {
            DataContractJsonSerializer jsonSer = cache.Get<DataContractJsonSerializer>(t.AssemblyQualifiedName);
            if (jsonSer == null)
            {
                jsonSer = new DataContractJsonSerializer(t, extraTypes);
                cache.Set<DataContractJsonSerializer>(t.AssemblyQualifiedName, jsonSer);
            }
            return jsonSer;
        }

        public static XmlSerializer GetXmlSerializer(CacheManager cache, Type t, params Type[] extraTypes)
        {
            XmlSerializer xmlSer = cache.Get<XmlSerializer>(t.AssemblyQualifiedName);
            if (xmlSer == null)
            {
                xmlSer = new XmlSerializer(t, extraTypes);
                cache.Set<XmlSerializer>(t.AssemblyQualifiedName, xmlSer);
            }
            return xmlSer;
        }

        private static void RemoveSerialiserFromCache<T>(CacheManager cache, SerialisationMethod method)
        {
            Type t = typeof(T);
            cache.Remove(t.AssemblyQualifiedName);
        }

        #region Stream

        public static T DeserialiseFromStream<T>(CacheManager cache, Stream s, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            Type t = typeof(T);
            try
            {
                return (T)DeserialiseFromStream(cache, s, t, method, extraTypes);
            }
            catch (InvalidCastException)
            {
                RemoveSerialiserFromCache<T>(cache, method);
                return (T)DeserialiseFromStream(cache, s, t, method, extraTypes);
            }
        }

        public static object DeserialiseFromStream(CacheManager cache, Stream s, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            switch (method)
            {
                case SerialisationMethod.Xml:
                    return GetXmlSerializer(cache, t, extraTypes).Deserialize(s);

                case SerialisationMethod.DataContract:
                    return GetDataContractSerializer(cache, t, extraTypes).ReadObject(s);

                case SerialisationMethod.Json:
                    return GetJsonSerializer(cache, t, extraTypes).ReadObject(s);
            }
            return null;
        }

        public static T DeserialiseFromStream<T>(Stream s, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return DeserialiseFromStream<T>(CacheManager.Default, s, method, extraTypes);
        }

        public static object DeserialiseFromStream(Stream s, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return DeserialiseFromStream(CacheManager.Default, s, t, method, extraTypes);
        }

        public static void SerialiseToStream<T>(CacheManager cache, T obj, Stream s, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToStream(cache, obj, s, typeof(T), method, extraTypes);
        }

        public static void SerialiseToStream(CacheManager cache, object obj, Stream s, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToStream(cache, obj, s, obj.GetType(), method, extraTypes);
        }

        public static void SerialiseToStream(CacheManager cache, object obj, Stream s, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            if (method == SerialisationMethod.Xml)
            {
                GetXmlSerializer(cache, t, extraTypes).Serialize(s, obj);
            }
            else
            {
                GetDataContractSerializer(cache, t, extraTypes).WriteObject(s, obj);
            }
            s.Flush();
        }

        public static void SerialiseToStream<T>(T obj, Stream s, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToStream<T>(CacheManager.Default, obj, s, method, extraTypes);
        }

        public static void SerialiseToStream(object obj, Stream s, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToStream(CacheManager.Default, obj, s, method, extraTypes);
        }

        public static void SerialiseToStream(object obj, Stream s, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToStream(CacheManager.Default, obj, s, t, method, extraTypes);
        }

        #endregion Stream

        #region File

        public static T DeserialiseFromFile<T>(CacheManager cache, string path, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            Type t = typeof(T);
            try
            {
                return (T)DeserialiseFromFile(cache, path, t, method, extraTypes);
            }
            catch (InvalidCastException)
            {
                RemoveSerialiserFromCache<T>(cache, method);
                return (T)DeserialiseFromFile(cache, path, t, method, extraTypes);
            }
        }

        public static object DeserialiseFromFile(CacheManager cache, string path, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            if (string.IsNullOrEmpty(path)) return null;

            using (FileStream fs = File.OpenRead(path))
            {
                return DeserialiseFromStream(cache, fs, t, method);
            }
        }

        public static T DeserialiseFromFile<T>(string path, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return DeserialiseFromFile<T>(CacheManager.Default, path, method, extraTypes);
        }

        public static object DeserialiseFromFile(string path, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return DeserialiseFromFile(CacheManager.Default, path, t, method, extraTypes);
        }

        public static void SerialiseToFile<T>(CacheManager cache, T obj, string path, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToFile(cache, obj, path, typeof(T), method);
        }

        public static void SerialiseToFile(CacheManager cache, object obj, string path, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToFile(cache, obj, path, obj.GetType(), method);
        }

        public static void SerialiseToFile(CacheManager cache, object obj, string path, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            FileHelper.EnsureDirectoryExists(path, PathType.File);
            using (FileStream fs = File.Create(path))
            {
                SerialiseToStream(cache, obj, fs, t, method);
            }
        }

        public static void SerialiseToFile<T>(T obj, string path, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToFile<T>(CacheManager.Default, obj, path, method, extraTypes);
        }

        public static void SerialiseToFile(object obj, string path, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToFile(CacheManager.Default, obj, path, method, extraTypes);
        }

        public static void SerialiseToFile(object obj, string path, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            SerialiseToFile(CacheManager.Default, obj, path, t, method, extraTypes);
        }

        #endregion File

        #region String

        public static T DeserialiseFromString<T>(CacheManager cache, string xml, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            Type t = typeof(T);
            try
            {
                return (T)DeserialiseFromString(cache, xml, t, method, extraTypes);
            }
            catch (InvalidCastException)
            {
                RemoveSerialiserFromCache<T>(cache, method);
                return (T)DeserialiseFromString(cache, xml, t, method, extraTypes);
            }
        }

        public static object DeserialiseFromString(CacheManager cache, string xml, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            if (string.IsNullOrEmpty(xml)) return null;

            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                return DeserialiseFromStream(cache, ms, t, method);
            }
        }

        public static T DeserialiseFromString<T>(string xml, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return DeserialiseFromString<T>(CacheManager.Default, xml, method, extraTypes);
        }

        public static object DeserialiseFromString(string xml, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return DeserialiseFromString(CacheManager.Default, xml, t, method, extraTypes);
        }

        public static string SerialiseToString<T>(CacheManager cache, T obj, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString(cache, obj, typeof(T), method, extraTypes);
        }

        public static string SerialiseToString(CacheManager cache, object obj, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString(cache, obj, obj.GetType(), method, extraTypes);
        }

        public static string SerialiseToString(CacheManager cache, object obj, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                SerialiseToStream(cache, obj, ms, t, method, extraTypes);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static string SerialiseToString<T>(T obj, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString<T>(CacheManager.Default, obj, method, extraTypes);
        }

        public static string SerialiseToString(object obj, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString(CacheManager.Default, obj, method, extraTypes);
        }

        public static string SerialiseToString(object obj, Type t, SerialisationMethod method = SerialisationMethod.Xml, params Type[] extraTypes)
        {
            return SerialiseToString(CacheManager.Default, obj, t, method, extraTypes);
        }

        #endregion String
    }
}