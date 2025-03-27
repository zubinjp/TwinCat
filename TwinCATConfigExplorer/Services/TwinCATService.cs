using System; // Core functionality
using System.Runtime.InteropServices; // For handling COM objects
using System.Xml;
using TCatSysManagerLib; // TwinCAT COM library

namespace TwinCATXmlGenerator.Services
{
        public interface ITwinCATService
    {
        ITcSysManager AttachToTwinCAT(); // Return SysManager instance instead of DTE
        string GetActiveProjectName(ITcSysManager systemManager); // Retrieve active project name
    }

    public class TwinCATService : ITwinCATService
    {
        // Attach to the active TwinCAT XAE instance
        public ITcSysManager AttachToTwinCAT()
        {
            try
            {
                Console.WriteLine("Connecting to TwinCAT SysManager...");

                // Replace DTE retrieval with SysManager initialization
                ITcSysManager4 systemManager = (ITcSysManager4)Activator.CreateInstance(Type.GetTypeFromProgID("TcSysManager"));
                if (systemManager == null)
                {
                    throw new InvalidOperationException("Failed to initialize the TwinCAT System Manager. Ensure the COM library is registered.");
                }
                Console.WriteLine("Successfully connected to TwinCAT SysManager.");

                return systemManager;
            }
            catch (COMException ex)
            {
                // Handle errors during COM object retrieval
                Console.WriteLine($"Failed to connect to TwinCAT SysManager. Error: {ex.Message}");
                return null;
            }
        }

        // Get the active project name from the TwinCAT SysManager instance
        public string GetActiveProjectName(ITcSysManager systemManager)
        {
            try
            {
                if (systemManager == null)
                {
                    Console.WriteLine("SysManager instance is null.");
                    return null;
                }

                // Use TwinCAT SysManager methods to access active project information
                ITcSmTreeItem solutionTreeItem = systemManager.LookupTreeItem("TIRS"); // Example: Top-level solution tree
                Console.WriteLine($"Connected to solution: {solutionTreeItem.Name}");
                return solutionTreeItem.Name;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving active project: {ex.Message}");
                return null;
            }
        }
    }
}