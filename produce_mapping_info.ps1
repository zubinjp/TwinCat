# Attach to the TwinCAT XAE (eXtended Automation Engineering) environment
try {
    $dte = [System.Runtime.Interopservices.Marshal]::GetActiveObject("TcXaeShell.DTE.15.0")
    Write-Host "Successfully attached to the active TwinCAT XAE instance."
} catch {
    Write-Host "Error: No active TwinCAT instance found. Ensure that the solution is open in TwinCAT XAE."
    exit
}

# Validate that the TwinCAT solution is open
try {
    if ($dte.Solution -eq $null -or -not $dte.Solution.IsOpen) {
        Write-Host "Error: No TwinCAT solution is currently open."
        exit
    }
    Write-Host "Successfully connected to the open solution: $($dte.Solution.FullName)"
} catch {
    Write-Host "Error: Failed to validate the TwinCAT solution state: $_"
    exit
}

# Access the ITC System Manager for the first project
try {
    $project = $dte.Solution.Projects.Item(1)

    if ($project -eq $null) {
        Write-Host "Error: Unable to retrieve the first project in the solution."
        exit
    }

    # Debug: Check the project object
    Write-Host "Debug: Project Name - $($project.Name)"

    $systemManager = $project.Object
    if ($systemManager -eq $null) {
        Write-Host "Error: Unable to retrieve the ITC System Manager for project: $($project.Name)."
        exit
    }
    Write-Host "Successfully accessed the ITC System Manager for project: $($project.Name)."

    # Retrieve Target Net ID
    $targetNetId = $systemManager.GetTargetNetId()
    Write-Host "Target Net ID: $targetNetId"

    # Activate Configuration and Restart TwinCAT
    $systemManager.ActivateConfiguration()
    $systemManager.StartRestartTwinCAT()
    Write-Host "TwinCAT configuration activated and restarted successfully."
} catch {
    Write-Host "Error: Failed to access the ITC System Manager or extend functionality: $_"
    exit
}

# Use the ProduceMappingInfo method to export mapping data to an XML file
try {
    Write-Host "Generating mapping information..."
    $mappingInfoXml = $systemManager.ProduceMappingInfo()

    if ($mappingInfoXml -eq $null) {
        Write-Host "Error: ProduceMappingInfo returned null data. No mapping information generated."
        exit
    }

    # Specify the output file path
    $outputFilePath = "C:\Users\zubinjp\source\Repos\TwinCAT_API\Codes\TwinCAT_Mapping_Info.xml"
    [System.IO.File]::WriteAllText($outputFilePath, $mappingInfoXml)
    Write-Host "Mapping information successfully exported to: $outputFilePath"
} catch {
    Write-Host "Error: Failed to generate or save mapping information: $_"
    exit
}

# Browse the TwinCAT configuration tree
$rootItemName = "TIPC" # Replace with your tree item name
try {
    $rootTreeItem = $systemManager.LookupTreeItem($rootItemName)
    if ($rootTreeItem -eq $null) {
        Write-Host "Error: Tree item '$rootItemName' not found in the ITC System Manager."
        exit
    }
    Write-Host "Successfully found tree item: $rootItemName"

    # Iterate through all child nodes in the tree item
    foreach ($child in $rootTreeItem) {
        Write-Host "Found Tree Item: $($child.Name)"

        # Access variables within this tree item, if applicable
        if ($child.VarCount(0) -gt 0) {
            Write-Host "  Variables:"
            for ($i = 0; $i -lt $child.VarCount(0); $i++) {
                $variable = $child.Var(0, $i)
                Write-Host "    - $($variable.Name)"
            }
        }
    }
} catch {
    Write-Host "Error: Failed to lookup tree item '$rootItemName' or iterate through child nodes: $_"
}
