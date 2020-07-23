using ExcelDocumentsModule;
using FluentAssertions;
using System;
using System.Data;
using System.IO;
using Xunit;

namespace ExcelDocumentsModuleTests
{
    public class ExcelModuleTests
    {
        private ExcelModule excelModule;

        private string filesPath;

        public ExcelModuleTests()
        {
            this.excelModule = new ExcelModule();
            this.filesPath = $@"{Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName}{Path.DirectorySeparatorChar}SourceFiles";
        }

        [Fact]
        public void ReadDocumentShouldReturnDataSetFromCSVExcelFile()
        {
            //Arrange
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}CSVExample.csv";

            //Act
            var result = this.excelModule.ReadDocument(sourceFile).GetAwaiter().GetResult();
            var table = result.Tables[0];

            //Assert
            result.Should().NotBeNull();
            result.Should().BeOfType(typeof(DataSet));
            table.Should().NotBeNull();
            table.Rows.Count.Should().Be(9);
        }

        [Fact]
        public void ReadDocumentShouldReturnDataSetFromExcelFile()
        {
            //Arrange
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithHeader.xlsx";

            //Act
            var result = this.excelModule.ReadDocument(sourceFile).GetAwaiter().GetResult();
            var table = result.Tables[0];

            //Assert
            result.Should().NotBeNull();
            result.Should().BeOfType(typeof(DataSet));
            table.Should().NotBeNull();
            table.Rows.Count.Should().Be(10);
        }

        [Fact]
        public void ReadDocumentShouldSetFirstRowAsHeaderInDataSetFromExcelFile()
        {
            //Arrange
            var isFirstRowHeader = true;
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithHeader.xlsx";
            var firstColumnName = "Header 1";
            var secondColumnName = "Header 2";

            //Act
            var result = this.excelModule.ReadDocument(sourceFile, isFirstRowHeader).GetAwaiter().GetResult();
            var table = result.Tables[0];

            //Assert
            table.Should().NotBeNull();
            table.Rows.Count.Should().Be(9);
            table.Columns[0].ColumnName.Trim().Should().Be(firstColumnName);
            table.Columns[1].ColumnName.Trim().Should().Be(secondColumnName);
        }

        [Fact]
        public void ReadDocumentShouldSetDefaultHeaderInDataSetFromExcelFile()
        {
            //Arrange
            var isFirstRowHeader = false;
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithoutHeader.xlsx";

            //Act
            var result = this.excelModule.ReadDocument(sourceFile, isFirstRowHeader).GetAwaiter().GetResult();
            var table = result.Tables[0];

            //Assert
            table.Should().NotBeNull();
            table.Rows.Count.Should().Be(16);
            table.Columns[0].ColumnName.Trim().Should().Be("A");
            table.Columns[1].ColumnName.Trim().Should().Be("B");
        }

        [Fact]
        public void ReadDocumentShouldReadDataWithHeaderFromExcelFile()
        {
            //Arrange
            var isFirstRowHeader = true;
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithHeader.xlsx";

            //Act
            var result = this.excelModule.ReadDocument(sourceFile, isFirstRowHeader).GetAwaiter().GetResult();
            var table = result.Tables[0];

            //Assert
            table.Should().NotBeNull();
            table.Rows.Count.Should().Be(9);

            table.Columns[0].ColumnName.Trim().Should().Be("Header 1");
            table.Columns[1].ColumnName.Trim().Should().Be("Header 2");

            table.Rows[0].ItemArray[0].Should().Be("some data h1 1");
            table.Rows[0].ItemArray[1].Should().Be("some data h2 1");
            table.Rows[1].ItemArray[0].Should().Be("some data h1 2");
            table.Rows[1].ItemArray[1].Should().Be("some data h2 2");
            table.Rows[2].ItemArray[0].Should().Be("some data h1 3");
            table.Rows[2].ItemArray[1].Should().Be("some data h2 3");
            table.Rows[3].ItemArray[0].Should().Be("some data h1 4");
            table.Rows[3].ItemArray[1].Should().Be("some data h2 4");
            table.Rows[4].ItemArray[0].Should().Be("some data h1 5");
            table.Rows[4].ItemArray[1].Should().Be("some data h2 5");
            table.Rows[5].ItemArray[0].Should().Be("some data h1 6");
            table.Rows[5].ItemArray[1].Should().Be("some data h2 6");
            table.Rows[6].ItemArray[0].Should().Be("some data h1 7");
            table.Rows[6].ItemArray[1].Should().Be("some data h2 7");
            table.Rows[7].ItemArray[0].Should().Be("some data h1 8");
            table.Rows[7].ItemArray[1].Should().Be("some data h2 8");
            table.Rows[8].ItemArray[0].Should().Be("some data h1 9");
            table.Rows[8].ItemArray[1].Should().Be("some data h2 9");
        }

        [Fact]
        public void ReadDocumentShouldReadDataWithoutHeaderFromExcelFile()
        {
            //Arrange
            var isFirstRowHeader = false;
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithoutHeader.xlsx";

            //Act
            var result = this.excelModule.ReadDocument(sourceFile, isFirstRowHeader).GetAwaiter().GetResult();
            var table = result.Tables[0];

            //Assert
            table.Should().NotBeNull();
            table.Rows.Count.Should().Be(16);

            table.Columns[0].ColumnName.Trim().Should().Be("A");
            table.Columns[1].ColumnName.Trim().Should().Be("B");
            table.Columns[2].ColumnName.Trim().Should().Be("E");
            table.Columns[3].ColumnName.Trim().Should().Be("F");

            table.Rows[0].ItemArray[0].ToString().Should().Be(string.Empty);
            table.Rows[1].ItemArray[1].ToString().Trim().Should().Be("Test");
            table.Rows[2].ItemArray[0].ToString().Should().Be(string.Empty);
            table.Rows[3].ItemArray[0].ToString().Should().Be(string.Empty);
            table.Rows[4].ItemArray[0].ToString().Should().Be(string.Empty);
            table.Rows[5].ItemArray[0].ToString().Should().Be(string.Empty);
            table.Rows[6].ItemArray[0].ToString().Trim().Should().Be("SomeData");
            table.Rows[6].ItemArray[1].ToString().Should().Be(string.Empty);
            table.Rows[6].ItemArray[2].ToString().Trim().Should().Be("Header 1");
            table.Rows[6].ItemArray[3].ToString().Trim().Should().Be("Header 2");
            table.Rows[7].ItemArray[2].ToString().Trim().Should().Be("some data h1 1");
            table.Rows[7].ItemArray[3].ToString().Trim().Should().Be("some data h2 1");
            table.Rows[8].ItemArray[2].ToString().Trim().Should().Be("some data h1 2");
            table.Rows[8].ItemArray[3].ToString().Trim().Should().Be("some data h2 2");
            table.Rows[9].ItemArray[2].ToString().Trim().Should().Be("some data h1 3");
            table.Rows[9].ItemArray[3].ToString().Trim().Should().Be("some data h2 3");
            table.Rows[10].ItemArray[2].ToString().Trim().Should().Be("some data h1 4");
            table.Rows[10].ItemArray[3].ToString().Trim().Should().Be("some data h2 4");
            table.Rows[11].ItemArray[2].ToString().Trim().Should().Be("some data h1 5");
            table.Rows[11].ItemArray[3].ToString().Trim().Should().Be("some data h2 5");
            table.Rows[12].ItemArray[2].ToString().Trim().Should().Be("some data h1 6");
            table.Rows[12].ItemArray[3].ToString().Trim().Should().Be("some data h2 6");
            table.Rows[13].ItemArray[2].ToString().Trim().Should().Be("some data h1 7");
            table.Rows[13].ItemArray[3].ToString().Trim().Should().Be("some data h2 7");
            table.Rows[14].ItemArray[2].ToString().Trim().Should().Be("some data h1 8");
            table.Rows[14].ItemArray[3].ToString().Trim().Should().Be("some data h2 8");
            table.Rows[15].ItemArray[2].ToString().Trim().Should().Be("some data h1 9");
            table.Rows[15].ItemArray[3].ToString().Trim().Should().Be("some data h2 9");
        }

        [Fact]
        public void ReadDocumentShouldThrowExceptionIfFileFormatIsNotSupports()
        {
            //Arrange
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithoutHeader.xlsm";

            //Act-Assert
            Assert.Throws<FileFormatException>(() => this.excelModule.ReadDocument(sourceFile).GetAwaiter().GetResult());
        }

        [Fact]
        public void ReadDocumentShouldThrowExceptionIfFirstRowWasNotFound()
        {
            //Arrange
            var isFirstRowHeader = true;
            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithoutHeader.xlsx";

            //Act-Assert
            Assert.Throws<MissingFieldException>(() => this.excelModule.ReadDocument(sourceFile, isFirstRowHeader).GetAwaiter().GetResult());
        }

        [Fact]
        public void WriteDocumentShouldWriteDataSetToExcelFile()
        {
            //Arrange
            var isFirstRowHeader = true;
            var fileName = "TestExcel.xlsx";
            var outputPath = CreateDirectory($@"{Directory.GetParent(Environment.CurrentDirectory).FullName}{Path.DirectorySeparatorChar}{Guid.NewGuid()}");
            var fullPath = $@"{outputPath}{Path.DirectorySeparatorChar}{fileName}";

            var sourceFile = $@"{filesPath}{Path.DirectorySeparatorChar}ExampleWithHeader.xlsx";
            var dataSet = this.excelModule.ReadDocument(sourceFile, isFirstRowHeader).GetAwaiter().GetResult();
            var expectedTable = dataSet.Tables[0];

            //Act
            this.excelModule.WriteDocument(dataSet, fullPath).GetAwaiter().GetResult();
            var result = this.excelModule.ReadDocument(fullPath, isFirstRowHeader).GetAwaiter().GetResult();

            //Assert
            result.Should().NotBeNull();
            var resultTable = result.Tables[0];
            resultTable.Columns[0].ColumnName.Trim().Should().Be(expectedTable.Columns[0].ColumnName.Trim());
            resultTable.Columns[1].ColumnName.Trim().Should().Be(expectedTable.Columns[1].ColumnName.Trim());
            resultTable.Rows[0].ItemArray[0].ToString().Trim().Should().Be(expectedTable.Rows[0].ItemArray[0].ToString().Trim());
            resultTable.Rows[1].ItemArray[1].ToString().Trim().Should().Be(expectedTable.Rows[1].ItemArray[1].ToString().Trim());
            resultTable.Rows[2].ItemArray[0].ToString().Trim().Should().Be(expectedTable.Rows[2].ItemArray[0].ToString().Trim());
            resultTable.Rows[3].ItemArray[1].ToString().Trim().Should().Be(expectedTable.Rows[3].ItemArray[1].ToString().Trim());
            resultTable.Rows[4].ItemArray[0].ToString().Trim().Should().Be(expectedTable.Rows[4].ItemArray[0].ToString().Trim());
            resultTable.Rows[4].ItemArray[1].ToString().Trim().Should().Be(expectedTable.Rows[4].ItemArray[1].ToString().Trim());

            //Dispose
            DeleteDirectory(outputPath);
        }

        private void DeleteDirectory(string path)
        {
            if (Directory.Exists(path))
                Directory.Delete(path, true);
        }

        private string CreateDirectory(string path)
        {
            if (Directory.Exists(path))
                Directory.Delete(path, true);

            Directory.CreateDirectory(path);
            return path;
        }
    }
}
