using System;
using System.Collections.Generic;
using System.Text;

namespace RFIDDemoCS
{
    public class Utils
    {       
        public static void Byte2Hex(byte[] epc, ref String strEPC)
        {
            for (int i = 0; i < epc.Length; i++)
            {
                strEPC += epc[i].ToString("x2");
            }
        }

        public static int GetHex(char bChar)
        {
            String strHex = "0123456789abcdef";

            return strHex.IndexOf(bChar);
        }

        public static int Chars2Hexs(String strScr, ref byte[] pszDest)
        {
            int length = strScr.Length;
            int nTemp, nMod;

            if (length % 2 != 0)
            {
                strScr += "0";
                length++;
            }
            nMod = length / 2;
            pszDest = new Byte[nMod];

            String strEPC = strScr.ToLower();
            for (int i = 0; i < nMod; i++)
            {
                nTemp = GetHex(strEPC[i * 2]) * 16;
                nTemp += GetHex(strEPC[i * 2 + 1]);
                pszDest[i] = (Byte)nTemp;
            }

            return nMod;
        }

        public static void Copy(byte[] src, int srcIndex, byte[] dst, int dstIndex, int count)
        {
            if (src == null || srcIndex < 0 ||
                dst == null || dstIndex < 0 || count < 0)
            {
                throw new System.ArgumentException();
            }

            int srcLen = src.Length;
            int dstLen = dst.Length;
            if (srcLen - srcIndex < count || dstLen - dstIndex < count)
            {
                throw new System.ArgumentException();
            }

            for (int i = 0; i < count; i++)
            {
                dst[dstIndex++] = src[srcIndex++];
            }
        }      
    }
}
