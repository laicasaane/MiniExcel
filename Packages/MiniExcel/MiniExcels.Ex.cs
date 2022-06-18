namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Utils;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;

    public static partial class MiniExcel
    {
        public static void GetSheetNames<T>(string path, T output) where T : ICollection<string>
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                GetSheetNames(stream, output);
        }

        public static void GetSheetNames<T>(this Stream stream, T output) where T : ICollection<string>
        {
            var archive = new ExcelOpenXmlZip(stream);
            var sheetNames = new ExcelOpenXmlSheetReader(stream, null).GetWorkbookRels(archive.entries).Select(s => s.Name);

            foreach (var sheetName in sheetNames)
                output.Add(sheetName);
        }

        public static void GetSheetNames<T>(string path, T output, Func<string, bool> predicate) where T : ICollection<string>
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                GetSheetNames(stream, output, predicate);
        }

        public static void GetSheetNames<T>(this Stream stream, T output, Func<string, bool> predicate) where T : ICollection<string>
        {
            var archive = new ExcelOpenXmlZip(stream);
            var sheetNames = new ExcelOpenXmlSheetReader(stream, null).GetWorkbookRels(archive.entries).Select(s => s.Name).Where(predicate);

            foreach (var sheetName in sheetNames)
                output.Add(sheetName);
        }
    }
}
