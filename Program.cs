using ExcelAnalyzer;

Console.WriteLine("Starting Excel Walker...");
if (args.Length != 1)
{
    Console.WriteLine("Syntax:   ExcelAnalyzer.exe <path to traverse>");
    Console.WriteLine("Example:  ExcelAnalyzer.exe C:\\Data");
    Console.WriteLine("If no path is supplied, default directory C:\\Data will be used");
    Console.WriteLine("Additional arguments will be ignored.");
}
//https://stackoverflow.com/questions/5181405/best-way-to-iterate-folders-and-subfolders
//https://stackoverflow.com/questions/26633810/how-to-decompress-a-single-file-from-an-zipfile-using-shfileopstruct
//https://stackoverflow.com/questions/13003555/c-sharp-how-to-extract-the-file-name-and-extension-from-a-path
//https://stackoverflow.com/questions/364253/how-to-deserialize-xml-document
//http://www.hanselman.com/blog/?date=2010-02-04
//https://stackoverflow.com/questions/36327240/create-directory-if-not-exists
//https://stackoverflow.com/questions/56638375/changing-excel-power-query-connection-string-in-c-sharp  (Major thanks here)
//https://bengribaudo.com/blog/2020/04/22/5198/data-mashup-binary-stream#:~:text=Custom%20parts%20are%20stored%20in%20the%20Excel%20OPC,the%20part%20containing%20Power%20Queries%20is%20in%20item1.xml.

var directoryToProcess = args.Length > 0 ? args[0] : "C:\\Data";

var connections = DirectoryWalker.ProcessDirectory(directoryToProcess);

Directory.CreateDirectory(@"C:\Temp\xl");
File.WriteAllText(@"C:\Temp\xl\output.csv", connections.Where(c => c != null).ToCsv());
