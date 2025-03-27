using System;
using System.Xml;
using TwinCATXmlGenerator.Services;
using EnvDTE; // Correct case-sensitive namespace for DTE
using System.Runtime.InteropServices; // Required for Marshal
using TCatSysManagerLib; // Required for ITcSysManager4 interface

namespace TwinCATXmlGenerator
{
    class Program
    {
        /// <summary>
        /// Attaches to the active TwinCAT XAE instance.
        /// </summary>
        static DTE AttachToTwinCAT()
        {
            try
            {
                // Retrieve the active TwinCAT XAE Shell instance
                DTE dte = (DTE)Marshal.GetActiveObject("TcXaeShell.DTE.15.0");
                Console.WriteLine("Successfully attached to the active TwinCAT XAE instance.");
                return dte;
            }
            catch (COMException ex)
            {
                Console.WriteLine($"Failed to attach to TwinCAT XAE instance: {ex.Message}");
                return null;
            }
        }

        static void Main()
        {
            try
            {
                // Attach to the TwinCAT XAE instance via DTE
                DTE twinCATInstance = AttachToTwinCAT();
                if (twinCATInstance == null)
                {
                    Console.WriteLine("TwinCAT XAE Shell is not running or inaccessible. Please open the TwinCAT solution.");
                    return;
                }

                // Retrieve the System Manager object from the TwinCAT instance
                ITcSysManager4 systemManager = (ITcSysManager4)twinCATInstance.GetObject("SystemManager");
                if (systemManager == null)
                {
                    Console.WriteLine("Failed to attach to System Manager. Ensure the TwinCAT solution is loaded.");
                    return;
                }

                // Initialize XML Producer Service
                IXmlProducerService xmlProducerService = new XmlProducerService();

                // Define the root paths for XML traversal
                string[] rootPaths = { "TIPC", "TIRT", "TINC", "TIRC" };

                // Generate the XML document using the System Manager
                XmlDocument xmlDocument = xmlProducerService.GenerateXml(rootPaths, systemManager);
                if (xmlDocument == null)
                {
                    Console.WriteLine("Failed to generate the XML document. Check the root paths and system configuration.");
                    return;
                }

                // Save the generated XML to a file
                string filePath = @"C:\Path\To\TwinCAT_Project_Config.xml";
                xmlProducerService.SaveXmlToFile(xmlDocument, filePath);

                Console.WriteLine($"XML generation and save completed successfully.\nFile saved to: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}