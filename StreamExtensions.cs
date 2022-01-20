using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace JSONlib
{
    public static class StreamExtensions
    {
        private const int bufferSize = 8192;

        public static MemoryStream ReadToMemoryStream(this IStream comStream)
        {            
            MemoryStream memoryStream = new MemoryStream();
            IntPtr num = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            Marshal.WriteInt32(num, bufferSize);
            byte[] numArray = new byte[bufferSize];            
            while (Marshal.ReadInt32(num) > 0)
            {
                comStream.Read(numArray, numArray.Length, num);
                memoryStream.Write(numArray, 0, Marshal.ReadInt32(num));                
            }            
            memoryStream.Position = 0L;
            //throw new Exception("Stop after position");
            //comStream.Seek(0L, 0, IntPtr.Zero);            
            return memoryStream;
        }
    }
}

