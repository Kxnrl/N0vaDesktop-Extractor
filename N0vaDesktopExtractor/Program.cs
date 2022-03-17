using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;

namespace N0vaDesktopExtractor
{
    internal class Program
    {
        private static readonly string Worker = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "n0va_output");

        static void Main(string[] args)
        {
            Console.Title = "人工桌面 N0vaDesktop Extractor by Kyle";

            // SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\人工桌面 || InstallPath
            // SYSTEM\ControlSet001\Services\N0vaDesktop Service || ImagePath

            var regKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\ControlSet001\Services\N0vaDesktop Service");
            if (regKey == null)
            {
                Fail("您尚未安装人工桌面");
                return;
            }

            var ImagePath = regKey.GetValue("ImagePath");
            if (ImagePath == null || string.IsNullOrEmpty(ImagePath.ToString()))
            {
                Fail("您的程序已损坏");
                return;
            }

            var InstallPath = Path.GetDirectoryName(ImagePath.ToString());
            if (string.IsNullOrEmpty(InstallPath) || !Directory.Exists(InstallPath))
            {
                Fail("找不到资源文件目录");
                return;
            }

            var ResourcePath = Path.Combine(InstallPath, "N0vaDesktopCache", "game");
            if (!Directory.Exists(ResourcePath))
            {
                Fail("找不到资源文件目录");
                return;
            }

            if (Directory.Exists(Worker))
                Directory.Delete(Worker, true);

            Directory.CreateDirectory(Worker);

            var processed = 0;
            var images = 0;
            var videos = 0;
            var fails = 0;

            foreach (var file in Directory.GetFiles(ResourcePath))
            {
                try
                {
                    var bytes = File.ReadAllBytes(file);
                    var oPath = Path.Combine(Worker, Path.GetFileNameWithoutExtension(file));
                    var nFile = Path.GetFileNameWithoutExtension(file);

                    //PNG
                    if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                    {
                        var info = ReadPngInfo(bytes);

                        if (info.Width == 480 && info.Height == 270)
                        {
                            // skipped thumbnail
                            Info($"[PNG] Skipped {nFile} by thumbnail.");
                            continue;
                        }

                        File.WriteAllBytes($"{oPath}.png", bytes);
                        images++;
                        Success($"Processed {nFile} as PNG ({info.Width}*{info.Height}).");
                    }
                    // JPG
                    // https://en.wikipedia.org/wiki/JPEG
                    else if ((bytes[0] == 0xFF && bytes[1] == 0xD8) &&  //SOI
                             (bytes[bytes.Length - 2] == 0xFF && bytes[bytes.Length - 1] == 0xD9))   //EOI
                    {
                        var info = ReadJepgInfo(bytes);

                        if (info.Width == 480 && info.Height == 270)
                        {
                            // skipped thumbnail
                            Info($"[JPG] Skipped {nFile} by thumbnail.");
                            continue;
                        }

                        File.WriteAllBytes($"{oPath}.jpg", bytes);
                        images++;
                        Success($"Processed {nFile} as JPG ({info.Width}*{info.Height}).");
                    }
                    //MP4 with offset
                    else if (bytes[0] == 0x00 || bytes[1] == 0x00 || bytes[2] == 0x00 || bytes[3] == 0x00 || bytes[4] == 0x00 || bytes[5] == 0x00)
                    {
                        var result = new List<byte>(bytes);
                        result.RemoveRange(0, 2);
                        File.WriteAllBytes($"{oPath}.mp4", result.ToArray());
                        videos++;
                        Success($"Processed {nFile} as MP4.");
                    }
                    else
                    {
                        throw new InvalidDataException("无法解析的文件格式");
                    }

                    processed++;
                }
                catch (Exception e)
                {
                    fails++;
                    Error($"Failed to process {Path.GetFileName(file)}: {e.Message}.");
                }
            }

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"Processed {processed} files: {images} images, {videos} videos.");
            Console.ReadKey(true);
        }

        private static ImageInfo ReadPngInfo(byte[] data)
        {
            ImageInfo info;
            info.Width = (data[16] << 24) + (data[17] << 16) + (data[18] << 8) + (data[19] << 0);
            info.Height = (data[20] << 24) + (data[21] << 16) + (data[22] << 8) + (data[23] << 0);
            return info;
        }

        private static ImageInfo ReadJepgInfo(byte[] data)
        {
            ImageInfo info;
            info.Width = 9999;
            info.Height = 9999;

            var off = 0;
            while (off < data.Length)
            {
                while (data[off] == 0xff)
                {
                    off++;
                }

                var cdpr = data[off++];

                if (cdpr == 0xd8) continue;
                if (cdpr == 0xd9) break;
                if (0xd0 <= cdpr && cdpr <= 0xd7) continue;
                if (cdpr == 0x01) continue;

                var len = (data[off] << 8) | data[off + 1];
                off += 2;

                if (cdpr == 0xc0)
                {
                    info.Height = (data[off + 1] << 8) | data[off + 2];
                    info.Width = (data[off + 3] << 8) | data[off + 4];
                    return info;
                }
                off += len - 2;
            }

            return info;
        }

        private static void Fail(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ReadKey(true);
            Environment.Exit(1);
        }

        private static void Error(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(message);
        }

        private static void Info(string message)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(message);
        }

        private static void Success(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
        }

        struct ImageInfo
        {
            public int Height;
            public int Width;
            //public int Bpc;
            //public int Cps;
        }
    }
}
