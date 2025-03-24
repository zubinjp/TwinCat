using System;
using TCatSysManagerLib;

class TwinCATExportTest
{
    static void Main(string[] args)
    {
        string projectPath = @"C:\Users\zubinjp\source\Repos\TwinCAT_API\Projects\p24_telemetry\P24_Telemetry_C6030lab.tsproj";  // Path to your TwinCAT project
        string xmlExportPath = @"C:\Users\zubinjp\source\Repos\TwinCAT_API\Codes\ExportedConfiguration.xml"; // Path to save the exported XML

        try
        {
            Console.WriteLine("Starting TwinCAT project export test...");

            // Step 1: Load the TwinCAT project
            Console.WriteLine($"Loading TwinCAT project from: {projectPath}");
            TcSysManagerClass systemManager = new TcSysManagerClass();
            systemManager.LoadProject(projectPath);

            // Step 2: Export the project to XML
            Console.WriteLine($"Exporting TwinCAT project to XML at: {xmlExportPath}");
            systemManager.ExportConfiguration(xmlExportPath);

            // Step 3: Confirm the export operation
            Console.WriteLine("Export successful! XML has been saved.");
        }
        catch (Exception ex)
        {
            // Handle any errors during the export process
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }
}