using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Linq;
using TCatSysManagerLib;

class TwinCATScript
{
    static async Task Main(string[] args)
    {
        string projectPath = @"C:\Path\To\Your\TwinCATProject.tsproj";
        string xmlExportPath = @"C:\Path\To\ExportedConfiguration.xml";
        string apiUrl = "https://api.example.com/udt-modifications"; // Example API URL

        try
        {
            // Step 1: Export TwinCAT Project to XML
            Console.WriteLine("Exporting TwinCAT project to XML...");
            TcSysManagerClass systemManager = new TcSysManagerClass();
            systemManager.LoadProject(projectPath);
            systemManager.ExportConfiguration(xmlExportPath);
            Console.WriteLine($"Exported XML saved to: {xmlExportPath}");

            // Step 2: Fetch Data from API
            Console.WriteLine("Fetching UDT modification data from API...");
            var udtUpdates = await FetchUdtModifications(apiUrl);
            Console.WriteLine($"API data fetched: {udtUpdates}");

            // Step 3: Modify the XML
            Console.WriteLine("Modifying UDTs in the XML...");
            ModifyUdtInXml(xmlExportPath, udtUpdates);
            Console.WriteLine("XML modification complete.");

            // Step 4: Re-Import the Modified XML into TwinCAT
            Console.WriteLine("Re-importing the modified XML back into the TwinCAT project...");
            systemManager.ImportConfiguration(xmlExportPath);
            systemManager.SaveProject(true);
            Console.WriteLine("TwinCAT project updated successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Fetch UDT modification data from the API
    static async Task<string> FetchUdtModifications(string apiUrl)
    {
        using HttpClient client = new HttpClient();
        string response = await client.GetStringAsync(apiUrl);
        return response;
    }

    // Modify UDT structures in the XML
    static void ModifyUdtInXml(string xmlPath, string udtUpdates)
    {
        // Parse the XML
        XDocument xmlDoc = XDocument.Load(xmlPath);

        // Example: Find and update UDTs
        var udtNode = xmlDoc.Descendants("UserDefinedType")
                            .FirstOrDefault(x => (string)x.Attribute("Name") == "MyUDT");

        if (udtNode != null)
        {
            // Example update: Add a new member using data from the API
            udtNode.Add(new XElement("Member",
                new XAttribute("Name", "NewFieldFromAPI"),
                new XAttribute("Type", "STRING")));
        }

        // Save the modified XML
        xmlDoc.Save(xmlPath);
    }
}