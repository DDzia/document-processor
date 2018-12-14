using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using NPOI.XSSF.Streaming;

namespace DocumentProcessor.Spreadsheets
{
    public sealed class Spreadsheet
    {
        internal SXSSFWorkbook _wb;

        public static Spreadsheet OpenRead(string filePath)
        {
            return new Spreadsheet();
        }

        private readonly Lazy<SheetSet> _sheetSet = new Lazy<SheetSet>(() => new SheetSet());
        public SheetSet Sheets() => _sheetSet.Value;

        //public IEnumerable<Sheet> Sheets[string name]
        //{
        //    get
        //    {
        //        yield return new Sheet()
        //    }
        //}
    }
}
