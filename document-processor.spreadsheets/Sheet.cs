using NPOI.SS.UserModel;

namespace DocumentProcessor.Spreadsheets
{
    public class Sheet
    {
        internal readonly ISheet _poiSheet;

        internal Sheet(ISheet poiSheet)
        {
            _poiSheet = poiSheet;
        }
    }
}
