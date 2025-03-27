using System.Xml; // For XML handling
using TCatSysManagerLib; // For TwinCAT System Manager

namespace TwinCATXmlGenerator.Services
{
    public interface IXmlProducerService
    {
        /// <summary>
        /// Generates an XML document based on specified root paths and the TwinCAT System Manager instance.
        /// </summary>
        /// <param name="rootPaths">The root paths in the TwinCAT project configuration.</param>
        /// <param name="systemManager">The TwinCAT System Manager instance.</param>
        /// <returns>An XML document containing the configuration.</returns>
        XmlDocument GenerateXml(string[] rootPaths, ITcSysManager4 systemManager);

        /// <summary>
        /// Saves the XML document to the specified file path.
        /// </summary>
        /// <param name="xmlDocument">The XML document to save.</param>
        /// <param name="filePath">The file path to save the document.</param>
        void SaveXmlToFile(XmlDocument xmlDocument, string filePath);
    }
}