using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace DocumentProcessor.Spreadsheets
{
    public class SheetSet: IEnumerable<Sheet>
    {
        private Spreadsheet _spreadsheet;

        private int _currentIndex = 0;

        public IEnumerator<Sheet> GetEnumerator()
        {
            if (_currentIndex < Count)
            {
                yield return this[_currentIndex++];
            }
            else
            {
                _currentIndex = 0;
                yield break;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Count => _spreadsheet._wb.NumberOfSheets;

        public Sheet this[int index]
        {
            get
            {
                var poiSheet = _spreadsheet._wb.GetSheetAt(index);
                return new Sheet(poiSheet);
            }
        }

        public Sheet this[string sheetName]
        {
            get
            {
                var poiSheet = _spreadsheet._wb.GetSheet(sheetName);
                return new Sheet(poiSheet);
            }
        }
    }

    //internal class SheetEnumerator
    //{

    //}
}
