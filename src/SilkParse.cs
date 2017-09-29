using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace XunfeiVoiceParse
{
    /// <summary>
    /// QQ/微信语音文件解析
    /// </summary>
    class SilkParse
    {
        /// <summary>
        /// 原始语音文件转成wav波形音频文件
        /// </summary>
        /// <param name="path">原始语音文件路径</param>
        /// <returns>转成wav文件后的路径</returns>
        public static string Parse(string path)
        {
            if (!path.EndsWith(".amr"))
            {
                return string.Empty;
            }
            if (!File.Exists(path))
            {
                return string.Empty;
            }
            var buffer = File.ReadAllBytes(path);
            if (Encoding.UTF8.GetString(buffer, 1, 9).Equals("#!SILK_V3"))
            {
                var silkPath = path.Replace(".amr", ".silk");
                using (var silkStream = new FileStream(silkPath, FileMode.Create))
                {
                    silkStream.Write(buffer, 1, buffer.Length - 1);
                }
                var process = new Process();
                process.StartInfo.FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "decoder.exe");
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.UseShellExecute = false;
                var pcmPath = silkPath.Replace(".silk", ".pcm");
                process.StartInfo.Arguments = string.Format(" {0} {1}", new FileInfo(silkPath).Name, new FileInfo(pcmPath).Name);
                process.Start();
                process.WaitForExit();
                if (!process.HasExited)
                {
                    process.Kill();
                }
                File.Delete(silkPath);

                //return Transfer2Wav(pcmPath);
                return Transfer2Wav2(pcmPath);
            }

            return string.Empty;
        }

        /// <summary>  
        /// PCM to wav  
        /// 添加Wav头文件  
        /// 参考资料：http://blog.csdn.net/bluesoal/article/details/932395  
        /// </summary>  
        private static string Transfer2Wav(string pcmPath)
        {
            if (!File.Exists(pcmPath))
            {
                return string.Empty;
            }
            var wavPath = pcmPath.Replace(".pcm", ".wav");
            var pcmBuffer = File.ReadAllBytes(pcmPath);
            using (var stream = new MemoryStream())
            using (var binaryWriter = new BinaryWriter(stream))
            {
                binaryWriter.Write("RIFF".ToCharArray());
                binaryWriter.Write(pcmBuffer.Length + 36);
                binaryWriter.Write("WAVE".ToCharArray());

                binaryWriter.Write("fmt ".ToCharArray());
                binaryWriter.Write(0x10);
                binaryWriter.Write((short)1);

                binaryWriter.Write((short)1);   // Mono,声道数目，1-- 单声道；2-- 双声道  
                binaryWriter.Write(22050);       // 16KHz 采样频率                     
                binaryWriter.Write(22050 * 2);       // 每秒所需字节数

                binaryWriter.Write((short)2);   // 数据块对齐单位(每个采样需要的字节数)  
                binaryWriter.Write((short)16);   // 16Bit,每个采样需要的bit数    

                // 数据块  
                binaryWriter.Write("data".ToCharArray());
                binaryWriter.Write(pcmBuffer.Length);
                binaryWriter.Write(pcmBuffer);
                binaryWriter.Flush();

                File.WriteAllBytes(wavPath, stream.ToArray());
            }
            File.Delete(pcmPath);

            return wavPath;
        }

        private static string Transfer2Wav2(string pcmPath)
        {
            var wavPath = pcmPath.Replace(".pcm", ".wav");

            var process = new Process();
            process.StartInfo.FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe");
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.Arguments = string.Format(" -f s16le -ar 22050 -ac 1 -i {0} -f wav -ar 16000 -ac 1 {1}", 
                new FileInfo(pcmPath).Name, new FileInfo(wavPath).Name);
            process.Start();
            process.WaitForExit();
            if (!process.HasExited)
            {
                process.Kill();
            }
            try
            {
                File.Delete(pcmPath);
            }
            catch { }

            return wavPath;
        }

        /// <summary>
        /// 原始语音文件转文本
        /// </summary>
        /// <param name="wavPath"></param>
        /// <returns></returns>
        public static string Transfer2Text(string wavPath)
        {
            var text = string.Empty;

            MspLogin();

            var param = "sub = iat, domain = iat, language = zh_cn, accent = mandarin, sample_rate = 16000, result_type = plain, result_encoding = gb2312";
            var sessionResult = 0;
            var ptr = MscDll.QISRSessionBegin(null, param, ref sessionResult);
            if (sessionResult != (int)ErrorCode.MSP_SUCCESS)
            {
                MspLogout();
            }
            var sessionId = Marshal.PtrToStringAnsi(ptr);
            var audioStatus = AudioStatus.ISR_AUDIO_SAMPLE_FIRST;
            var epStatus = EpStatus.ISR_EP_LOOKING_FOR_SPEECH;
            var recStatus = RecogStatus.ISR_REC_STATUS_SUCCESS;

            try
            {
                using (var fs = new FileStream(wavPath, FileMode.Open))
                using (var br = new BinaryReader(fs))
                {
                    // 每次写入200ms音频(16k，16bit)：1帧音频20ms，10帧=200ms。16k采样率的16位音频，一帧的大小为640Byte
                    while (true)
                    {
                        uint writeLen = 10 * 640;
                        if ((fs.Length - fs.Position) < 2 * writeLen)
                        {
                            writeLen = (uint)(fs.Length - fs.Position);
                        }
                        var buffer = br.ReadBytes((int)writeLen);
                        string data;
                        if (TryReadBlock(sessionId, buffer, audioStatus, epStatus, recStatus, out data))
                        {
                            text = string.Concat(text, data);
                            if (epStatus == EpStatus.ISR_EP_AFTER_SPEECH)
                            {
                                break;
                            }
                            audioStatus = AudioStatus.ISR_AUDIO_SAMPLE_CONTINUE;
                            Thread.Sleep(200);
                        }
                        else
                        {
                            break;
                        }
                    }
                    var writeResult = MscDll.QISRAudioWrite(sessionId, null, 0, AudioStatus.ISR_AUDIO_SAMPLE_LAST, ref epStatus, ref recStatus);
                    if (writeResult != (int)ErrorCode.MSP_SUCCESS)
                    {
                        throw new Exception("写入失败");
                    }
                    while (recStatus != RecogStatus.ISR_REC_STATUS_SPEECH_COMPLETE)
                    {
                        //读取数据
                        var readResult = 0;
                        var dataPtr = MscDll.QISRGetResult(sessionId, ref recStatus, 0, ref readResult);
                        if (readResult != (int)ErrorCode.MSP_SUCCESS)
                        {
                            break;
                        }
                        text += Marshal.PtrToStringAnsi(dataPtr);
                        Thread.Sleep(150);
                    }
                }
            }
            catch(Exception e)
            {
                var s = e.Message;
            }

            MscDll.QISRSessionEnd(sessionId, null);

            MspLogout();
            return text;
        }

        private static bool TryReadBlock(string sessionId, byte[] buffer, 
            AudioStatus audioStatus, EpStatus epStatus, RecogStatus recStatus, out string text)
        {
            text = string.Empty;
            var audioWriteResult = MscDll.QISRAudioWrite(sessionId, buffer, (uint)buffer.Length, audioStatus, ref epStatus, ref recStatus);
            if (audioWriteResult != (int)ErrorCode.MSP_SUCCESS)
            {
                return false;
            }
            if (recStatus == RecogStatus.ISR_REC_STATUS_SUCCESS)
            {
                //读取数据
                var readResult = 0;
                var dataPtr = MscDll.QISRGetResult(sessionId, ref recStatus, 0, ref readResult);
                if (readResult != (int)ErrorCode.MSP_SUCCESS)
                {
                    return false;
                }
                text += Marshal.PtrToStringAnsi(dataPtr);
            }

            return true;
        }

        private static void MspLogin()
        {
            string loginParam = "appid = 59c8ae72, work_dir = .";
            var loginResult = MscDll.MSPLogin(null, null, loginParam);
            if (loginResult != (int)ErrorCode.MSP_SUCCESS)
            {
                throw new Exception("登录失败");
            }
        }

        private static void MspLogout()
        {
            var logoutResult = MscDll.MSPLogout();
            if (logoutResult != (int)ErrorCode.MSP_SUCCESS)
            {
                throw new Exception("注销失败");
            }
        }
    }
}
