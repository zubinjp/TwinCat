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
                string[] rootPaths = { "TIPC", "TIRT", "TINC", "TIRC" };
                foreach (string path in rootPaths)
                {
                    ITcSmTreeItem rootTreeItem = systemManager.LookupTreeItem(path);
                    if (rootTreeItem != null)
                    {
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
                return null;
            }
        }

        // Recursively traverse the tree and add items to XML
        static void TraverseTree(ITcSmTreeItem treeItem, XmlElement parentElement, XmlDocument xmlDocument)
        {
            foreach (ITcSmTreeItem child in treeItem)
            {
                try
                {
                    if (child == null || string.IsNullOrEmpty(child.Name))
                        continue;

                    // Create an XML element for the current tree item
                    XmlElement childElement = xmlDocument.CreateElement(SanitizeXmlName(child.Name));
                    childElement.SetAttribute("PathName", child.PathName);
                    parentElement.AppendChild(childElement);

                    // Use ProduceXml for granular data when available
                    try
                    {
                        string itemXml = child.ProduceXml();
                        XmlDocument itemDoc = new XmlDocument();
                        itemDoc.LoadXml(itemXml);
                        XmlNode parametersNode = itemDoc.SelectSingleNode("TreeItem/Parameters");
                        if (parametersNode != null)
                        {
                            XmlNode importedNode = xmlDocument.ImportNode(parametersNode, true);
                            childElement.AppendChild(importedNode);
                        }
                    }
                    catch
                    {
                        Console.WriteLine($"Warning: ProduceXml failed for {child.Name}. Skipping detailed parameters.");
                    }

                    // Recursive traversal
                    if (child.ChildCount > 0)
                    {
                        TraverseTree(child, childElement, xmlDocument);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing tree item '{treeItem.Name}': {ex.Message}");
                }
            }
        }

        // Save the generated XML to a file
        static void SaveXmlToFile(XmlDocument xmlDocument, string filePath)
        {
            try
            {
                xmlDocument.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving XML file: {ex.Message}");
            }
        }

        // Sanitize XML element names
        static string SanitizeXmlName(string name)
        {
            return System.Text.RegularExpressions.Regex.Replace(name, @"[^a-zA-Z0-9_.-]", "_");
        }
    }
}