using System;
using System.Xml;
using System.Runtime.InteropServices;
using EnvDTE;
using TCatSysManagerLib;

namespace TwinCATProjectXMLGenerator
{
    class Program
    {
        static void Main()
        {
            // Attach to TwinCAT instance
            DTE dte = AttachToTwinCAT();
            if (dte == null)
            {
                Console.WriteLine("No active TwinCAT instance found. Ensure the solution is open in TwinCAT XAE.");
                return;
            }

            // Validate and retrieve the solution and project
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
                Console.WriteLine("Error: System Manager object not found.");
                return;
            }

            // Prepare XML document
            XmlDocument xmlDocument = new XmlDocument();
            XmlElement rootElement = xmlDocument.CreateElement("TwinCATConfiguration");
            xmlDocument.AppendChild(rootElement);

            // Traverse tree and generate XML
            string[] rootItems = { "TIPC", "TIIC", "TIID", "TIRT" };
            foreach (string path in rootItems)
            {
                ITcSmTreeItem rootTreeItem = GetTreeItem(systemManager, path);
                if (rootTreeItem != null)
                {
                    TraverseTree(rootTreeItem, rootElement, xmlDocument);
                }
            }

            // Save XML to file
            string filePath = @"C:\Users\zubinjp\source\Repos\TwinCAT_API\Codes\TwinCAT_Project_Config.xml";
            SaveXmlToFile(xmlDocument, filePath);
        }

        // Attach to the active TwinCAT instance
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
                Console.WriteLine("Error: No active TwinCAT instance found.");
                return null;
            }
        }

        // Validate the TwinCAT solution and retrieve the project
        static Project ValidateSolution(DTE dte)
        {
            try
            {
                if (dte.Solution == null || !dte.Solution.IsOpen)
                {
                    Console.WriteLine("Error: No TwinCAT solution is currently open.");
                    return null;
                }

                Console.WriteLine("Successfully connected to the open solution: " + dte.Solution.FullName);
                Project project = dte.Solution.Projects.Item(1);
                Console.WriteLine("Active Project Name: " + project.Name);
                return project;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to validate the TwinCAT solution state. " + ex.Message);
                return null;
            }
        }

        // Retrieve tree items by path
        static ITcSmTreeItem GetTreeItem(ITcSysManager4 systemManager, string path)
        {
            try
            {
                ITcSmTreeItem treeItem = systemManager.LookupTreeItem(path);
                if (treeItem != null)
                {
                    Console.WriteLine("Successfully retrieved tree item: " + path);
                }
                else
                {
                    Console.WriteLine("Warning: Tree item '" + path + "' not found.");
                }
                return treeItem;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error retrieving tree item: " + ex.Message);
                return null;
            }
        }

        // Recursive tree traversal
        static void TraverseTree(ITcSmTreeItem treeItem, XmlElement xmlNode, XmlDocument xmlDocument)
        {
            foreach (ITcSmTreeItem child in treeItem)
            {
                try
                {
                    if (child == null || string.IsNullOrEmpty(child.Name))
                    {
                        Console.WriteLine("Skipping null or invalid tree item.");
                        continue;
                    }

                    Console.WriteLine("Processing Tree Item: Name=" + child.Name + ", PathName=" + child.PathName);
                    XmlElement childElement = xmlDocument.CreateElement(SanitizeXmlName(child.Name));
                    xmlNode.AppendChild(childElement);

                    foreach (dynamic property in child.Properties)
                    {
                        if (property != null && property != "System.__ComObject")
                        {
                            childElement.SetAttribute(property.Name, property.Value.ToString());
                        }
                    }

                    if (child.ChildCount > 0)
                    {
                        TraverseTree(child.Children, childElement, xmlDocument);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error processing tree item: " + ex.Message);
                }
            }
        }

        // Save XML to file
        static void SaveXmlToFile(XmlDocument xmlDocument, string filePath)
        {
            try
            {
                xmlDocument.Save(filePath);
                Console.WriteLine("XML saved successfully to: " + filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving XML to file: " + ex.Message);
            }
        }

        // Sanitize XML element names
        static string SanitizeXmlName(string name)
        {
            return System.Text.RegularExpressions.Regex.Replace(name, @"[^a-zA-Z0-9_.-]", "_").Replace(" ", "_");
        }
    }
}