using System.Web.Http;
using System.Xml;
using System;

public class XmlRetrievalController : ApiController
{
    private readonly string _xmlFilePath = @"C:\Path\To\TwinCAT_Config.xml";

    /// <summary>
    /// Retrieve the entire XML content from the specified file.
    /// </summary>
    [HttpGet]
    [Route("retrieve")]
    public IHttpActionResult RetrieveXml()
    {
        try
        {
            // Validate file path
            if (string.IsNullOrWhiteSpace(_xmlFilePath) || !System.IO.File.Exists(_xmlFilePath))
            {
                return BadRequest("XML file not found or invalid file path.");
            }

            // Load the XML document
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(_xmlFilePath);

            // Return the XML content
            return Ok(xmlDocument.OuterXml);
        }
        catch (XmlException ex)
        {
            // Handle XML-specific errors
            return InternalServerError(new Exception($"Error reading the XML file: {ex.Message}"));
        }
        catch (Exception ex)
        {
            // Handle general errors
            return InternalServerError(ex);
        }
    }
}