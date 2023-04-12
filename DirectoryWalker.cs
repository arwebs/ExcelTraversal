using System.IO.Compression;
using System.Xml.Serialization;
using System.Xml;
using System.ComponentModel;

namespace ExcelAnalyzer
{
    public static class DirectoryWalker
    {
        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static List<connections> ProcessDirectory(string targetDirectory)
        {
            try
            {
                // Process the list of files found in the directory.
                var fileEntries = Directory.GetFiles(targetDirectory);
                var listOfConnections = fileEntries.Where(f => f.EndsWith(".xlsx")).Select(ProcessFile).ToList();
                Console.WriteLine(
                    $"Processing {targetDirectory}... Total Files: {fileEntries.Length} Eligible Files: {listOfConnections.Count}");
                // Recurse into subdirectories of this directory.
                var subdirectoryEntries = Directory.GetDirectories(targetDirectory);
                foreach (var subdirectory in subdirectoryEntries)
                {
                    listOfConnections.AddRange(ProcessDirectory(subdirectory));
                }

                return listOfConnections;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not process folder {targetDirectory} - {ex.Message}");
                return new List<connections>();
            }
        }

        // Insert logic for processing found files here.
        public static connections ProcessFile(string path)
        {
            var myId = Guid.NewGuid();
            try
            {
                using (var archive = ZipFile.OpenRead(path))
                {
                    archive.Entries
                        .FirstOrDefault(e => e.FullName.Contains("connections.xml"))?.ExtractToFile(
                            Path.Combine(@"C:\Temp\xl", $"{Path.GetFileNameWithoutExtension(path)}-{myId}.xml"));
                    Console.WriteLine($"Processed file '{path}'.");
                    if (archive.Entries
                        .Any(e => e.FullName.Contains("connections.xml")))
                    {
                        var foo = DeserializeXMLFile<connections>(
                            @$"C:\Temp\xl\{Path.GetFileNameWithoutExtension(path)}-{myId}.xml");
                        foo.filename = path;
                        return foo;
                    }

                    return null;

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not process {path} -- {ex.Message}");
                return null;
            }
        }

        public static T DeserializeXMLFile<T>(string file) where T : class
        {
            var ser = new XmlSerializer(typeof(T));
            using (var stream = new FileStream(file, FileMode.Open))
            using (var reader = XmlReader.Create(stream))
            {
                return (T)ser.Deserialize(reader);
            }
        }


        // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
        [Serializable]
        [DesignerCategory("code")]
        [XmlType(AnonymousType = true, Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
        public class connections
        {
            public string filename { get; set; }
            [XmlElement("connection")]
            public connectionsConnection[] connection { get; set; }
        }

        
        [Serializable]
        [DesignerCategory("code")]
        [XmlType(AnonymousType = true, Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public class connectionsConnection
        {
            
            [System.Xml.Serialization.XmlAttribute]
            public byte id { get; set; }

            
            [XmlAttribute]
            public string name { get; set; }

            
            [XmlAttribute]
            public string description { get; set; }

            
            [XmlAttribute]
            public int type { get; set; }

            
            [XmlAttribute]
            public int refreshedVersion { get; set; }

            
            [XmlAttribute]
            public int minRefreshableVersion { get; set; }

            public connectionsConnectionDbPr dbPr { get; set; }
        }

        
        [Serializable]
        [DesignerCategory("code")]
        [XmlType(AnonymousType = true, Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public class connectionsConnectionDbPr
        {
            
            [XmlAttribute]
            public string connection { get; set; }

            
            [XmlAttribute]
            public string command { get; set; }

            
            [XmlAttribute]
            public int commandType { get; set; }

        }

    }
}
