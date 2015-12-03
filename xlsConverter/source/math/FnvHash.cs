using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace xlsConverter
{
    class FnvHash
    {
        /// <summary>
        /// 32bit fnv-1 ハッシュを取得する
        /// データが32bit以上であればこちらの方が推称されています。
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static int getFNV_1_32(string source)
        {
            const int fnv_prime = 16777619;
            const int offset_basis = 0xCE942FA;   // 2166136261
            int hash = offset_basis;
            char[] work = source.ToCharArray();
            int length = work.GetLength(0);
            for (int i = 0; i < length; ++i)
            {
                hash *= fnv_prime;
                hash ^= work[i];
            }
            
            return hash;
        }

        /// <summary>
        /// 32bit fnv-1a ハッシュを取得する
        /// データが32bit以下であればこちらの方が推称されています。
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static int getFNV_1a_32(string source)
        {
            const int fnv_prime = 16777619;
            const int offset_basis = 0xCE942FA;   // 2166136261
            int hash = offset_basis;
            char[] work = source.ToCharArray();
            int length = work.GetLength(0);
            for (int i = 0; i < length; ++i)
            {
                hash ^= work[i];
                hash *= fnv_prime;
            }

            return hash;
        }

        /// <summary>
        /// 64bit fnv-1 ハッシュを取得する
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static Int64 getFNV_1_64(string source)
        {
            Int64 fnv_prime = 1099511628211;
            Int64 offset_basis = 0x57984997;
            Int64 hash = offset_basis;
            char[] work = source.ToCharArray();
            int length = work.GetLength(0);
            for (int i = 0; i < length; ++i)
            {
                hash *= fnv_prime;
                hash ^= work[i];
            }

            return hash;
        }

        /// <summary>
        /// 64bit fnv-1a ハッシュを取得する
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static Int64 getFNV_1a_64(string source)
        {
            Int64 fnv_prime = 1099511628211;
            Int64 offset_basis = 0x57984997;
            Int64 hash = offset_basis;
            char[] work = source.ToCharArray();
            int length = work.GetLength(0);
            for (int i = 0; i < length; ++i)
            {
                hash ^= work[i];
                hash *= fnv_prime;
            }

            return hash;
        }

    }
}
