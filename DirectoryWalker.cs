using System.IO.Compression;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExcelAnalyzer
{
    public static partial class DirectoryWalker
    {
        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static List<MData> ProcessDirectory(string targetDirectory)
        {
            try
            {
                var fileEntries = Directory.GetFiles(targetDirectory);
                var listOfConnections = fileEntries.Where(f => f.EndsWith(".xlsx")).Select(ProcessFile).ToList();
                Console.WriteLine($"Processing {targetDirectory}... Total Files: {fileEntries.Length} Eligible Files: {listOfConnections.Count}");
                // Recurse into subdirectories of this directory.
                var subDirectoryEntries = Directory.GetDirectories(targetDirectory);
                foreach (var subDirectory in subDirectoryEntries)
                {
                    listOfConnections.AddRange(ProcessDirectory(subDirectory));
                }
                return listOfConnections;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not process folder {targetDirectory} - {ex.Message}");
                return new List<MData>();
            }
        }

        public static MData ProcessFile(string path)
        {
            try
            {
                const int fieldsLength = 4;
                using var fileStream = File.Open(path, FileMode.Open);
                using var archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
                
                var entry = archive.GetEntry("customXml/item1.xml");
                if (entry == null) { return null; }

                using var entryStream = entry.Open();
                var doc = XDocument.Load(entryStream);
                var dataMashup = Convert.FromBase64String(doc.Root.Value);
                int packagePartsLength = BitConverter.ToUInt16(dataMashup.Skip(fieldsLength).Take(fieldsLength).ToArray(), 0);
                var packageParts = dataMashup.Skip(fieldsLength * 2).Take(packagePartsLength).ToArray();

                using var packagePartsStream = new MemoryStream(packageParts);
                using var package = Package.Open(packagePartsStream, FileMode.Open, FileAccess.Read);
                var section = package.GetParts().FirstOrDefault(x => x.Uri.OriginalString == "/Formulas/Section1.m");

                using var reader = new StreamReader(section.GetStream());
                var query = reader.ReadToEnd();
                return new MData
                {
                    Filename = path,
                    QueryCount = Regex.Matches(query, Regex.Escape("shared #")).Count,
                    Referenced = string.Join(",", GetReferencedThings(query)),
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not process {path} -- {ex.Message}");
                return null;
            }
        }
        
        // Hackily kinda parses through .M file to extract query sources
        // Returns string list of referenced sources
        private static IEnumerable<string> GetReferencedThings(string query)
        {
            var references = new List<string>();
            var spl = query.Split("shared ").Where(q => q.Length > 0);
            foreach (var part in spl)
            {
                if (part.StartsWith("section Section1;"))
                {
                    continue;
                }
                if (part.Contains("Source = Sql.Database"))
                {
                    references.Add("Direct DB Query");
                    //TODO: parse to figure out which DB we're querying
                } 
                else if (part.Contains("Source = Excel.CurrentWorkbook()"))
                {
                    references.Add("This Workbook");
                }
                else if (part.Contains("Source = Xml.Tables") || part.Contains("Source = Excel.Workbook") || part.Contains("Source = Csv.Document"))
                {
                    var tempString = part.Substring(part.IndexOf("File.Contents(\"") + 15);
                    references.Add(tempString.Substring(0, tempString.IndexOf("\"")));
                }
                else
                {
                    references.Add("Unknown");
                    var tempString = part.Substring(part.IndexOf("Source = ") + 9);
                    Console.WriteLine("Unknown: " + tempString.Substring(0, tempString.IndexOf("(")));
                }
            }
            return references;
        }
    }
}