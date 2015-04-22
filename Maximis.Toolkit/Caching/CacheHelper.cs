using System;
using System.Collections.Generic;
using System.Runtime.Caching;

namespace Maximis.Toolkit.Caching
{
    public class CacheHelper<T>
    {
        private Dictionary<string, T> dictionary;

        public CacheHelper(int expireMinutes = 30)
        {
            string typeName = typeof(T).FullName;
            if (MemoryCache.Default.Contains(typeName))
            {
                dictionary = (Dictionary<string, T>)MemoryCache.Default[typeName];
            }
            else
            {
                dictionary = new Dictionary<string, T>();
                MemoryCache.Default.Add(typeName, dictionary, new CacheItemPolicy { SlidingExpiration = TimeSpan.FromMinutes(expireMinutes) });
            }
        }

        public Dictionary<string, T> Dictionary { get { return dictionary; } }
    }
}