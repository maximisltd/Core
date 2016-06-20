using System;
using System.Runtime.Caching;

namespace Maximis.Toolkit.Caching
{
    public class CacheManager
    {
        private static CacheManager defaultCacheManager = new CacheManager("DefaultCacheManager");
        private static CacheItemPolicy policy = new CacheItemPolicy { SlidingExpiration = TimeSpan.FromHours(3) };

        public CacheManager(string uniqueId)
        {
            this.UniqueId = uniqueId;
        }

        public static CacheManager Default { get { return defaultCacheManager; } }
        public string UniqueId { get; set; }

        /// <summary>
        /// Tests if a key exists in the cache
        /// </summary>
        public bool Contains(string key)
        {
            return MemoryCache.Default.Contains(GetFullKey(key));
        }

        /// <summary>
        /// Retrieve an object from the cache
        /// </summary>
        public T Get<T>(string key)
        {
            string fullKey = GetFullKey(key);
            if (!MemoryCache.Default.Contains(fullKey)) return default(T);
            object result = MemoryCache.Default.Get(fullKey);
            try { return result == null ? default(T) : (T)result; }
            catch { return default(T); }
        }

        /// <summary>
        /// Remove an object from the cache
        /// </summary>
        public void Remove(string key)
        {
            MemoryCache.Default.Remove(GetFullKey(key));
        }

        /// <summary>
        /// Add an object to the cache
        /// </summary>
        public void Set<T>(string key, T val)
        {
            try
            {
                MemoryCache.Default.Set(GetFullKey(key), val, policy);
            }
            catch (ArgumentNullException ex)
            {
                throw new NullReferenceException(string.Format("Attempting to set Cache Key '{0}' to a null value", key), ex);
            }
        }

        private string GetFullKey(string key)
        {
            return string.Format("{0}_{1}", UniqueId, key);
        }
    }
}