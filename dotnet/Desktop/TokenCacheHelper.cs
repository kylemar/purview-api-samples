﻿// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IO;
using System.Security.Cryptography;
using Microsoft.Identity.Client;

namespace PurviewAPIExp
{
    internal static class TokenCacheHelper
    {
        /// <summary>
        /// Path to the token cache
        /// </summary>
        public static readonly string CacheFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location + ".msalcache.bin3";

        private static readonly object FileLock = new object();

        public static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                try
                {
                    args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath)
                            ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath),
                                                     null,
                                                     DataProtectionScope.CurrentUser)
                            : null);
                }
                catch
                {
                    File.Delete(CacheFilePath);
                    args.TokenCache.DeserializeMsalV3(null);
                }
            }
        }

        public static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                lock (FileLock)
                {
                    // reflect changesgs in the persistent store
                    File.WriteAllBytes(CacheFilePath,
                                       ProtectedData.Protect(args.TokenCache.SerializeMsalV3(),
                                                             null,
                                                             DataProtectionScope.CurrentUser)
                                      );
                }
            }
        }

        internal static void EnableSerialization(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
        }

        public static void ClearCache()
        {
            lock (FileLock)
            {
                File.Delete(CacheFilePath);
            }
        }

    }
}
