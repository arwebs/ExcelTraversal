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

var directoryToProcess = args.Length > 0 ? args[0] : "C:\\Data";

var conns = DirectoryWalker.ProcessDirectory(directoryToProcess);
var flats = new List<Flattened>();
foreach (var conn in conns.Where(c => c != null))
{
    flats.AddRange(conn.connection.Select(c =>
    {
        var kvpList = c.dbPr?.connection.Split(';')
            .Where(s => s.Contains('='))
            .Select(k => new KeyValuePair<string, string>(k.Split('=')[0], k.Split('=')[1]));

        return new Flattened
        {
            filename = conn.filename,
            id = c.id,
            name = c.name,
            type = c.type,
            description = c.description,
            connection = c.dbPr?.connection ?? "NA",
            provider = kvpList?.FirstOrDefault(k => k.Key == "Provider").Value,
            dataSource = kvpList?.FirstOrDefault(k => k.Key == "Data Source").Value,
            location = kvpList?.FirstOrDefault(k => k.Key == "Location").Value,
            extendedProperties = kvpList?.FirstOrDefault(k => k.Key == "Extended Properties").Value,

            command = c.dbPr?.command ?? "NA",
            commandType = c.dbPr?.commandType ?? 0,
            refreshedVersion = c.refreshedVersion,
            minRefreshableVersion = c.minRefreshableVersion
        };
    }));
}

Directory.CreateDirectory(@"C:\Temp\xl");
File.WriteAllText(@"C:\Temp\xl\output.csv", flats.ToCsv());


public class Flattened
{
    public string filename { get; set; }
    public int id { get; set; }
    public string name { get; set; }
    public string description { get; set; }
    public int type { get; set; }
    public int refreshedVersion { get; set; }
    public int minRefreshableVersion { get; set; }
    public string connection { get; set; }
    public string command { get; set; }
    public int commandType { get; set; }
    public string provider { get; set; }
    public string dataSource { get; set; }
    public string location { get; set; }
    public string extendedProperties { get; set; }
}