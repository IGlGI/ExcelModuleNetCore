using System.Data;
using System.Threading.Tasks;

namespace IExcelDocumentsModule
{
    public interface IExcelModule
    {
        Task WriteDocument(DataSet document, string outputPath);

        Task<DataSet> ReadDocument(string fileName, bool isFirstRowHead = false);
    }
}
