using System;
using System.Xml;
using System.Runtime.InteropServices; // For handling COM objects
using TCatSysManagerLib; // TwinCAT COM library

namespace TwinCATXmlGenerator.Services
{
    public class XmlProducerService : IXmlProducerService
    {
        /// <summary>
        /// Produces XML from the active TwinCAT project.
        /// </summary>
        /// <param name="rootPaths">The root paths in the TwinCAT project configuration.</param>
        /// <param name="systemManager">The TwinCAT System Manager instance.</param>
        /// <returns>The generated XML document.</returns>
        public XmlDocument GenerateXml(string[] rootPaths, ITcSysManager4 systemManager)
        {
            XmlDocument xmlDocument = new XmlDocument();
            XmlElement rootElement = xmlDocument.CreateElement("TwinCATProject");
            xmlDocument.AppendChild(rootElement);

            foreach (var path in rootPaths)
            {
                ITcSmTreeItem treeItem = systemManager.LookupTreeItem(path);
                if (treeItem != null)
                {
                    XmlElement element = xmlDocument.CreateElement(treeItem.Name);
                    element.SetAttribute("Path", treeItem.PathName);
                    rootElement.AppendChild(element);

                    // Traverse child items if they exist
                    if (treeItem.ChildCount > 0)
                    {
                        TraverseTree(treeItem, element, xmlDocument);
                    }
                }
                else
                {
                    Console.WriteLine($"Root path not found: {path}");
                }
            }

            Console.WriteLine("XML successfully generated from the TwinCAT project.");
            return xmlDocument;
        }

        /// <summary>
        /// Recursively traverses a TwinCAT tree item and appends it to the XML document.
        /// </summary>
        /// <param name="treeItem">The tree item to traverse.</param>
        /// <param name="parentElement">The parent XML element.</param>
        /// <param name="xmlDocument">The XML document being generated.</param>
        private void TraverseTree(ITcSmTreeItem treeItem, XmlElement parentElement, XmlDocument xmlDocument)
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
                    childElement.SetAttribute("ItemType", child.ItemType.ToString());
                    parentElement.AppendChild(childElement);

                    // Recurse for children
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

        /// <summary>
        /// Saves the generated XML document to a file.
        /// </summary>
        /// <param name="xmlDocument">The XML document to save.</param>
        /// <param name="filePath">The file path where the document should be saved.</param>
        public void SaveXmlToFile(XmlDocument xmlDocument, string filePath)
        {
            xmlDocument.Save(filePath);
            Console.WriteLine($"XML saved to: {filePath}");
        }

        /// <summary>
        /// Sanitizes an XML element name to ensure it is valid.
        /// </summary>
        /// <param name="name">The name to sanitize.</param>
        /// <returns>A valid XML element name.</returns>
        private string SanitizeXmlName(string name)
        {
            return System.Text.RegularExpressions.Regex.Replace(name, @"[^a-zA-Z0-9_.-]", "_");
        }
    }
}