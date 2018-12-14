using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace DocumentProcessor.Spreadsheets
{
    internal class SheetDataReader
    {
        internal SheetDataReader(string fileName)
        {


            var package = OPCPackage.Create(fileName);

            var xssfWb = new XSSFWorkbook(package);

            var sxssfWb = new SXSSFWorkbook(xssfWb);
            // sxssfWb.
            // new HSSFRequest().
        }
    }
}