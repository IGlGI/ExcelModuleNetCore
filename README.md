<h1>ExcelDocumentsModule</h1>
<h3>Supports .NET Core 3.xx</h3>

This module allows easy to write/read excel files to DataSet using <a href="https://github.com/OfficeDev/Open-XML-SDK" target="_blank">OpenXml</a> under the hood and supports <a href="https://dotnet.github.io/orleans/" target="_blank">MS Orleans</a> ver. 3.2 and higher.

Please feel free to fork and submit pull requests to the develop branch.

Supported file formats and versions

<table>
<thead>
<tr>
<th>Extension</th>
<th>Excel Version(s)</th>
</tr>
</thead>
<tbody>
<tr>
<td>.xlsx</td>
<td>2007 and newer</td>
</tr>
<tr>
<td>.xls</td>
<td> 2003, and newer</td>
</tr>
<tr>
<td>.csv</td>
<td>Only with the comma separator</td>
</tr>
</tbody>
</table>

<h3>Finding the binaries:</h3>
It's recommended to use NuGet through the VS Package Manager Console Install-Package <package name>.

<h3>Example:</h3>

<b>Install-Package ExcelDocumentsModule -Version 1.0.0</b>

<h3>How to use read method without MS Orleans:</h3>

    /// "ReadDocument" method has bool parameter "isFirstRowHeader" which works only with .xls and .xlsx files.
    /// CSV files always use the comma separator to define the first row as a header inside the DataSet. 

    var excelModule = new ExcelModule();
    var filePath = "../Some/Directory/FileName.xlsx";
    var dataSet = excelModule.ReadDocument(filePath);
    
<h3>How to use write method without MS Orleans:</h3>

    var excelModule = new ExcelModule();
    var outputFilePath = "../Some/Directory/FileName.xlsx";
    var dataSetToFile = new DataSet();
    //...
    // Adding some data to dataSetToFile...
    //...
    
    excelModule.WriteDocument(dataSetToFile, outputFilePath);

<h3>How to use ExcelDocumentsModule with MS Orleans:</h3>

Just replace <b>new ExcelModule(); or container.Resolve\<IExcelModule\>();</b> to <b>client.GetGrain\<IExcelModuleOrleans\>(Guid.NewGuid());</b> where the client is <b>IClusterClient</b>.

For more information look at the tests projects.
