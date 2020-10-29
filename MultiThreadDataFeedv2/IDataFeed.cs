using System.Runtime.InteropServices;

namespace MultiThreadDataFeedv2
{
    [ComVisible(true)]
    public interface IDataFeed
    {
        void CreateExportThread(string workbookFullPath, bool active, bool closed, string scenario);
        void CreateExportThreadTest();
    }
}
