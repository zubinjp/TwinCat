# Load the Visual Studio DTE COM object
$visualStudioVersion = "TcXaeShell.DTE.15.0"  # Adjust version based on your setup (e.g., VisualStudio.DTE.17.0)
$dte = new-object -com $visualStudioVersion

# Define source and destination paths
$sourceSolutionPath = "C:\Users\zubinjp\source\Repos\TwinCAT_API\Projects\p24_telemetry\P24_Telemetry_C6030lab.sln"
$destinationSolutionPath = "C:\Users\zubinjp\source\Repos\TwinCAT_API\Projects\phm_project\phm_project.sln"

# Open the source solution
$dte.Solution.Open($sourceSolutionPath)

# Save the solution to a new location (copy and rename)
$dte.Solution.SaveAs($destinationSolutionPath)

# Close the solution to release resources
$dte.Solution.Close()

Write-Host "The solution has been copied and renamed using the DTE object."