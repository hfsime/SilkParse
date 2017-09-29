using System;
using System.Runtime.InteropServices;

namespace XunfeiVoiceParse
{
    class MscDll
    {
        [DllImport("msc.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int MSPLogin(string user, string password, string param);

        [DllImport("msc.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int MSPLogout();

        [DllImport("msc.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr QISRSessionBegin(string grammarList, string _params, ref int errorCode);

        [DllImport("msc.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int QISRSessionEnd(string sessionID, string hints);

        [DllImport("msc.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int QISRAudioWrite(string sessionID, byte[] waveData, uint waveLen, AudioStatus audioStatus, ref EpStatus epStatus, ref RecogStatus recogStatus);

        [DllImport("msc.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr QISRGetResult(string sessionID, ref RecogStatus rsltStatus, int waitTime, ref int errorCode);
    }
}
