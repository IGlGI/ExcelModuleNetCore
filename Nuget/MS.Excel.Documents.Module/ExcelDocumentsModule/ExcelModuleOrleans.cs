using IExcelDocumentsModule;
using System.Data;
using System.Threading.Tasks;

namespace ExcelDocumentsModule
{
    public class ExcelModuleOrleans : Orleans.Grain, IExcelModuleOrleans
    {
        private ExcelModule excelModule;

        public ExcelModuleOrleans()
        {
            this.excelModule = new ExcelModule();
        }

        public Task<DataSet> ReadDocument(string fileName, bool isFirstRowHead = false) => this.excelModule.ReadDocument(fileName, isFirstRowHead);

        public Task WriteDocument(DataSet document, string outputPath) => this.excelModule.WriteDocument(document, outputPath);
    }
}
