using System.Web.Http;
using System.Xml;
using System;

public class XmlModificationController : ApiController
{
    private readonly string _xmlFilePath = @"C:\Path\To\TwinCAT_Config.xml";

    [HttpPost]
    [Route("modify")]
    public IHttpActionResult ModifyXml([FromBody] XmlModificationRequest request)
    {
        if (request == null || string.IsNullOrWhiteSpace(request.XPath) || request.NewValue == null)
        {
            return BadRequest("Invalid request. Please provide a valid XPath and NewValue.");
        }

        try
        {
            // Load the XML document from a file or memory
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(_xmlFilePath);

            // Locate and modify the requested node
            XmlNode nodeToModify = xmlDocument.SelectSingleNode(request.XPath);
            if (nodeToModify != null)
            {
                nodeToModify.InnerText = request.NewValue;
                xmlDocument.Save(_xmlFilePath); // Save changes

                return Ok("XML updated successfully.");
            }
            else
            {
                // Use the Content method to return a NotFound message.
                return Content(System.Net.HttpStatusCode.NotFound,
                               $"Node not found at XPath '{request.XPath}'.");
            }
        }
        catch (Exception ex)
        {
            return InternalServerError(ex);
        }
    }
}

public class XmlModificationRequest
{
    public string XPath { get; set; }
    public string NewValue { get; set; }
}