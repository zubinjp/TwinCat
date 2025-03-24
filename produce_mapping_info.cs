using System;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;

class TwinCATAutomation
{
    static void Main()
    {
        try
        {
            // Attach to the active TwinCAT XAE (VS) instance
            DTE2 dte = (DTE2)Marshal.GetActiveObject("TcXaeShell.DTE.15.0");
            Console.WriteLine("Successfully attached to the active TwinCAT XAE instance.");

            // Validate the TwinCAT solution is open
            if (dte.Solution == null || !dte.Solution.IsOpen)
            {
                Console.WriteLine("Error: No TwinCAT solution is currently open.");
                return;
            }
            Console.WriteLine($"Successfully connected to the open solution: {dte.Solution.FullName}");

            // Access the first project in the solution
            Project project = dte.Solution.Projects.Item(1);
            if (project == null)
            {
                Console.WriteLine("Error: Unable to retrieve the first project in the solution.");
                return;
            }
            Console.WriteLine($"Debug: Project Name - {project.Name}");

            // Access the ITC System Manager
            dynamic systemManager = project.Object;
            if (systemManager == null)
            {
                Console.WriteLine($"Error: Unable to retrieve the ITC System Manager for project: {project.Name}");
                return;
            }
            Console.WriteLine($"Successfully accessed the ITC System Manager for project: {project.Name}");

            // Retrieve Target Net ID
            string targetNetId = systemManager.GetTargetNetId();
            Console.WriteLine($"Target Net ID: {targetNetId}");

            // Activate Configuration and Restart TwinCAT
            systemManager.ActivateConfiguration();
            systemManager.StartRestartTwinCAT();
            Console.WriteLine("TwinCAT configuration activated and restarted successfully.");

            // Export mapping data to an XML file
            Console.WriteLine("Generating mapping information...");
            string mappingInfoXml = systemManager.ProduceMappingInfo();

            if (string.IsNullOrEmpty(mappingInfoXml))
            {
                Console.WriteLine("Error: ProduceMappingInfo returned null data. No mapping information generated.");
                return;
            }

            // Specify the output file path
            string outputFilePath = @"C:\Users\zubinjp\source\Repos\TwinCAT_API\Codes\TwinCAT_Mapping_Info.xml";
            File.WriteAllText(outputFilePath, mappingInfoXml);
            Console.WriteLine($"Mapping information successfully exported to: {outputFilePath}");

            // Browse the TwinCAT configuration tree
            string rootItemName = "TIPC"; // Replace with your actual tree item name
            dynamic rootTreeItem = systemManager.LookupTreeItem(rootItemName);
            if (rootTreeItem == null)
            {
                Console.WriteLine($"Error: Tree item '{rootItemName}' not found in the ITC System Manager.");
                return;
            }
            Console.WriteLine($"Successfully found tree item: {rootItemName}");

            // Iterate through child nodes in the tree item
            foreach (var child in rootTreeItem)
            {
                Console.WriteLine($"Found Tree Item: {child.Name}");

                // Access variables within this tree item, if applicable
                if (child.VarCount(0) > 0)
                {
                    Console.WriteLine("  Variables:");
                    for (int i = 0; i < child.VarCount(0); i++)
                    {
                        dynamic variable = child.Var(0, i);
                        Console.WriteLine($"    - {variable.Name}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
