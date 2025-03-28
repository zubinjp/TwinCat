using System;
using System.Xml;
using System.Runtime.InteropServices; // Needed for COMException
using TCatSysManagerLib;

namespace TwinCATXmlGenerator.Services
{
    public interface ITwinCATXmlGenerator
    {
        XmlDocument GenerateXmlForTree(ITcSysManager systemManager, string[] rootPaths);
    }

    public class TwinCATXmlGenerator : ITwinCATXmlGenerator
    {
        public XmlDocument GenerateXmlForTree(ITcSysManager systemManager, string[] rootPaths)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                XmlElement rootElement = document.CreateElement("TwinCATConfig");
                document.AppendChild(rootElement);

                foreach (string path in rootPaths)
                {
                    try
                    {
                        ITcSmTreeItem treeItem = systemManager.LookupTreeItem(path);
                        if (treeItem != null)
                        {
                            try
                            {
                                // Attempt to produce XML for the current tree item.
                                string itemXml = treeItem.ProduceXml();
                                if (!string.IsNullOrEmpty(itemXml))
                                {
                                    XmlDocument itemDoc = new XmlDocument();
                                    itemDoc.LoadXml(itemXml);
                                    XmlNode importedNode = document.ImportNode(itemDoc.DocumentElement, true);
                                    rootElement.AppendChild(importedNode);
                                }
                                else
                                {
                                    Console.WriteLine($"No XML content returned for tree item '{path}'.");
                                }
                            }
                            catch (COMException comEx)
                            {
                                int hr = comEx.ErrorCode;
                                Console.WriteLine($"COMException for tree item '{path}': Error Code: 0x{hr:X8}, Message: {comEx.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Tree item '{path}' not found.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing tree item '{path}': {ex.Message}");
                    }
                }
                return document;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error in GenerateXmlForTree: {ex.Message}");
                throw;
            }
        }
    }
}