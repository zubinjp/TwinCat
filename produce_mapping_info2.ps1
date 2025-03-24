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
    Write-Host "Debug: Project Name - $($project.Name)"
    $systemManager = $project.Object
    if ($systemManager -eq $null) {
        Write-Host "Error: Unable to retrieve the ITC System Manager for project: $($project.Name)."
        exit
    }
    Write-Host "Successfully accessed the ITC System Manager for project: $($project.Name)."
} catch {
    Write-Host "Error: Failed to access the ITC System Manager or extend functionality: $_"
    exit
}

# Sanitize tree item names to conform to XML standards
function SanitizeXmlName($name) {
    return $name -replace '\s', '_' -replace '[^a-zA-Z0-9_.-]', ''
}

# Function to recursively traverse the tree and build the XML hierarchy
function TraverseTree($treeItem, $xmlNode) {
    foreach ($child in $treeItem) {
        try {
            # Validate child node
            if ($child -eq $null -or -not $child.Name) {
                Write-Host "Warning: Encountered a null or invalid tree item. Skipping..."
                continue
            }

            # Sanitize the child name for XML
            $sanitizedName = SanitizeXmlName($child.Name)

            # Create XML node for the child tree item
            $childElement = $xmlNode.OwnerDocument.CreateElement($sanitizedName)
            $xmlNode.AppendChild($childElement)

            # Add variables to XML if applicable
            if ($child.PSObject.Methods.Name -contains "VarCount" -and $child.VarCount(0) -gt 0) {
                for ($i = 0; $i -lt $child.VarCount(0); $i++) {
                    $variable = $child.Var(0, $i)
                    $variableElement = $xmlNode.OwnerDocument.CreateElement("Variable")
                    $variableElement.SetAttribute("Name", $variable.Name)
                    $variableElement.SetAttribute("Type", $variable.Type)
                    $childElement.AppendChild($variableElement)
                }
            }

            # Recursively process child tree items
            TraverseTree $child $childElement
        } catch {
            Write-Host "Error processing tree item '$($child.Name)': $_"
            continue
        }
    }
}

# Retrieve the root tree item and begin recursive traversal
$rootItemName = "TIPC"  # Replace with your desired root tree item name
try {
    $rootTreeItem = $systemManager.LookupTreeItem($rootItemName)
    if ($rootTreeItem -eq $null) {
        Write-Host "Error: Tree item '$rootItemName' not found in the ITC System Manager."
        exit
    }
    Write-Host "Successfully found root tree item: $rootItemName"

    # Create an XML document for the hierarchy
    $xmlDocument = New-Object System.Xml.XmlDocument
    $rootElement = $xmlDocument.CreateElement((SanitizeXmlName $rootTreeItem.Name))
    $xmlDocument.AppendChild($rootElement)

    # Traverse the entire tree hierarchy starting from the root
    TraverseTree $rootTreeItem $rootElement

    # Specify the output file path
    $outputFilePath = "C:\Users\zubinjp\source\Repos\TwinCAT_API\Codes\$($project.Name)_TreeHierarchy.xml"
    $xmlDocument.Save($outputFilePath)
    Write-Host "Tree hierarchy successfully exported to: $outputFilePath"
} catch {
    Write-Host "Error: Failed to export tree hierarchy: $_"
    exit
}