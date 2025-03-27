using System;
using System.Web.Http;
using System.Xml;
using TwinCATXmlGenerator.Services;
using TCatSysManagerLib; // Import TwinCAT SysManager library

namespace TwinCATConfigExplorer.Controllers
{
    [RoutePrefix("api/xml")]
    public class XmlController : ApiController
    {
        private readonly ITwinCATService _twinCatService;

        public XmlController(ITwinCATService twinCatService)
        {
            _twinCatService = twinCatService;
        }

        [HttpGet]
        [Route("generate")]
        public IHttpActionResult GenerateXml()
        {
            try
            {
                // Attach to the TwinCAT SysManager
                var systemManager = _twinCatService.AttachToTwinCAT();
                if (systemManager == null)
                {
                    return NotFound(); // Failed to attach
                }

                // Define the root paths to retrieve XML from
                string[] rootPaths =
                {
                    "TIIC", "TIID", "TIRC", "TIRR", "TIRT",
                    "TIRS", "TIPC", "TINC", "TICC", "TIAC"
                };

                // Generate the XML document using the root paths
                XmlDocument xmlDocument = GenerateXmlFromPaths(rootPaths, systemManager);
                if (xmlDocument == null)
                {
                    return BadRequest("Failed to generate XML from TwinCAT project.");
                }

                // Save the XML document to a file
                string filePath = @"C:\Path\To\TwinCAT_Config.xml";
                xmlDocument.Save(filePath);

                return Ok(new
                {
                    Status = "success",
                    Message = "XML generated successfully.",
                    FilePath = filePath
                });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex); // Handle unexpected exceptions
            }
        }

        /// <summary>
        /// Generates an XML document based on the specified root paths and TwinCAT SysManager instance.
        /// </summary>
        /// <param name="rootPaths">The root paths to generate XML from.</param>
        /// <param name="systemManager">The TwinCAT SysManager instance.</param>
        /// <returns>Generated XML document.</returns>
        private XmlDocument GenerateXmlFromPaths(string[] rootPaths, ITcSysManager systemManager)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                XmlElement rootElement = document.CreateElement("TwinCATConfig");
                document.AppendChild(rootElement);

                // Iterate through each root path to extract data from TwinCAT configuration
                foreach (string path in rootPaths)
                {
                    ITcSmTreeItem treeItem = systemManager.LookupTreeItem(path);
                    if (treeItem != null)
                    {
                        // Produce XML for the current tree item and add it to the document
                        string itemXml = treeItem.ProduceXml();
                        XmlDocument itemDoc = new XmlDocument();
                        itemDoc.LoadXml(itemXml);

                        XmlNode importedNode = document.ImportNode(itemDoc.DocumentElement, true);
                        rootElement.AppendChild(importedNode);
                    }
                }

                return document;
            }
            catch (Exception ex)
            {
                // Log the error if needed; here we output to console.
                Console.WriteLine($"Error generating XML from paths: {ex.Message}");
                return null;
            }
        }
    }
}