using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using EnvDTE;
using TCatSysManagerLib;

namespace TwinCATXmlGenerator
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Attach to active TwinCAT XAE instance
                DTE dte = AttachToTwinCAT();
                if (dte == null)
                {
                    Console.WriteLine("No active TwinCAT instance found. Ensure that a TwinCAT solution is open.");
                    return;
                }

                // Validate and retrieve the solution/project
                Project project = ValidateSolution(dte);
                if (project == null)
                {
                    Console.WriteLine("Failed to retrieve the active TwinCAT project.");
                    return;
                }

                // Access the System Manager
                ITcSysManager4 systemManager = project.Object as ITcSysManager4;
                if (systemManager == null)
                {
                    Console.WriteLine("Error: Could not retrieve the System Manager from the project.");
                    return;
                }

                // Prepare an XML document
                XmlDocument xmlDocument = new XmlDocument();
                XmlElement rootElement = xmlDocument.CreateElement("TwinCATConfiguration");
                xmlDocument.AppendChild(rootElement);

                // Traverse the tree and produce XML
                string[] rootPaths = { "TIIC", "TIID", "TIRC", "TIRR", "TIRT", "TIRS", "TIPC", "TINC", "TICC", "TIAC" };
                foreach (string path in rootPaths)
                {
                    ITcSmTreeItem rootTreeItem = systemManager.LookupTreeItem(path);
                    if (rootTreeItem != null)
                    {
                        Console.WriteLine($"Starting traversal from root: {path}");
                        TraverseTree(rootTreeItem, rootElement, xmlDocument);
                    }
                }

                // Save the XML to file
                string filePath = @"C:\Users\zubinjp\source\Repos\TwinCAT_API\Codes\TwinCAT_Project_Config.xml";
                SaveXmlToFile(xmlDocument, filePath);
                Console.WriteLine($"XML successfully generated at: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Attach to the active TwinCAT XAE instance
        static DTE AttachToTwinCAT()
        {
            try
            {
                DTE dte = (DTE)Marshal.GetActiveObject("TcXaeShell.DTE.15.0");
                Console.WriteLine("Successfully attached to the active TwinCAT XAE instance.");
                return dte;
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to attach to the active TwinCAT XAE instance.");
                return null;
            }
        }

        // Validate the TwinCAT solution and retrieve the active project
        static Project ValidateSolution(DTE dte)
        {
            try
            {
                if (dte.Solution == null || !dte.Solution.IsOpen)
                {
                    Console.WriteLine("No TwinCAT solution is open.");
                    return null;
                }

                Project project = dte.Solution.Projects.Item(1);
                Console.WriteLine($"Connected to solution: {dte.Solution.FullName}");
                Console.WriteLine($"Active project: {project.Name}");
                return project;
            }
            catch
            {
                Console.WriteLine("Error validating the TwinCAT solution.");
                return null;
            }
        }

        // Recursively traverse the tree and add items to XML
        static void TraverseTree(ITcSmTreeItem treeItem, XmlElement parentElement, XmlDocument xmlDocument)
        {
            if (treeItem == null)
            {
                Console.WriteLine("[DEBUG] Encountered null tree item. Skipping...");
                return;
            }

            Console.WriteLine($"[DEBUG] Processing Tree Item: Name='{treeItem.Name}', Path='{treeItem.PathName}', ChildCount={treeItem.ChildCount}");

            // Create an XML element for the current tree item
            XmlElement currentElement = xmlDocument.CreateElement("TreeItem");
            currentElement.SetAttribute("PathName", treeItem.PathName);
            parentElement.AppendChild(currentElement);

            // Add fields based on sample XML
            currentElement.AppendChild(CreateTextElement(xmlDocument, "ItemName", treeItem.Name));
            currentElement.AppendChild(CreateTextElement(xmlDocument, "PathName", treeItem.PathName));
            currentElement.AppendChild(CreateTextElement(xmlDocument, "ItemType", treeItem.ItemType.ToString())); // Adjust as needed
            currentElement.AppendChild(CreateTextElement(xmlDocument, "ChildCount", treeItem.ChildCount.ToString()));

            // Process children and find deepest child
            foreach (ITcSmTreeItem child in treeItem)
            {
                try
                {
                    if (child == null || string.IsNullOrEmpty(child.Name))
                    {
                        Console.WriteLine("[DEBUG] Skipping invalid or unnamed child node.");
                        continue;
                    }

                    // Recursive traversal
                    TraverseTree(child, currentElement, xmlDocument);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERROR] Exception while processing '{treeItem.Name}': {ex.Message}");
                }
            }
        }

        // Create an XML text element
        static XmlElement CreateTextElement(XmlDocument doc, string name, string value)
        {
            XmlElement element = doc.CreateElement(name);
            element.InnerText = value ?? string.Empty;
            return element;
        }

        // Save the generated XML to a file
        static void SaveXmlToFile(XmlDocument xmlDocument, string filePath)
        {
            try
            {
                xmlDocument.Save(filePath);
                Console.WriteLine($"[INFO] XML saved successfully to {filePath}. Size: {new FileInfo(filePath).Length} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Failed to save XML file: {ex.Message}");
            }
        }
    }
}