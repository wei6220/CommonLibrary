using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary
{
    public class ZipHelper
    {
        public enum CompressTypeEnum
        {
            zip,
            gz,
        }

        public byte[] Compress(Dictionary<string, byte[]> files)
        {
            byte[] zipBytes;
            using (var stream = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Update))
                {
                    foreach (var fileName in files.Keys)
                    {
                        var zipArchiveEntry = archive.CreateEntry(fileName, CompressionLevel.Fastest);
                        using (var zipArchiveEntryStream = zipArchiveEntry.Open())
                        {
                            zipArchiveEntryStream.Write(files[fileName], 0, files[fileName].Length);
                        }
                    }
                }
                zipBytes = stream.ToArray();
            }
            return zipBytes;
        }

        public byte[] Compress(string targetPath)
        {
            var dirPath = (targetPath + "\\").Replace("\\\\", "\\");
            var dir = new DirectoryInfo(dirPath);
            var files = GetUnderFiles(dir);
            byte[] zipBytes;
            using (var stream = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Update))
                {
                    foreach (FileInfo file in files)
                    {
                        string path = file.FullName.Substring(dirPath.Length);
                        var zipArchiveEntry = archive.CreateEntry(path, CompressionLevel.Fastest);
                        using (var zipArchiveEntryStream = zipArchiveEntry.Open())
                        {
                            using (var fs = file.OpenRead())
                                fs.CopyTo(zipArchiveEntryStream);
                        }
                    }
                }
                zipBytes = stream.ToArray();
            }
            return zipBytes;
        }

        public List<FileInfo> GetUnderFiles(DirectoryInfo dir)
        {
            var files = new List<FileInfo>();
            files.AddRange(dir.GetFiles().AsEnumerable());
            var folders = dir.GetDirectories();
            if (folders.Count() > 0)
            {
                foreach (var subDir in folders)
                {
                    files.AddRange(GetUnderFiles(subDir));
                }
            }
            return files;
        }
    }
}
