param(
    [string]$PublishDir,
    [string]$OutputWxs,
    [string]$RootDirId = "AddinFolder",
    [string]$RegistryRoot = "HKCU",
    [string]$RegKeyBase = "Software\MyCompany\MyProduct\ExtraFiles"
)

# === Stable GUID helper ===
function Get-StableGuid {
    param([string]$InputString)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $hash = $md5.ComputeHash($bytes)
    $hex = [System.BitConverter]::ToString($hash).Replace("-", "")
    "{0}-{1}-{2}-{3}-{4}" -f `
        $hex.Substring(0, 8),
        $hex.Substring(8, 4),
        $hex.Substring(12, 4),
        $hex.Substring(16, 4),
        $hex.Substring(20, 12)
}

# === Setup XML ===
$namespace = "http://wixtoolset.org/schemas/v4/wxs"
$xml = New-Object System.Xml.XmlDocument
$xml.AppendChild($xml.CreateXmlDeclaration("1.0","UTF-8",$null)) | Out-Null
$wix = $xml.CreateElement("Wix", $namespace)
$xml.AppendChild($wix) | Out-Null

# === Directories ===
$fragment1 = $xml.CreateElement("Fragment", $namespace)
$wix.AppendChild($fragment1) | Out-Null
$dirRef = $xml.CreateElement("DirectoryRef", $namespace)
$dirRef.SetAttribute("Id", $RootDirId)
$fragment1.AppendChild($dirRef) | Out-Null

function New-DirectoryNode {
    param(
        [System.IO.DirectoryInfo]$Directory,
        [System.Xml.XmlElement]$ParentXmlElement
    )
    foreach ($subDir in $Directory.GetDirectories()) {
        $relPath = $subDir.FullName.Substring($PublishDir.Length).Trim("\")
        $dirId = ($relPath -replace '[\\\/]', '_') + "_DIR"
        $dirName = $subDir.Name

        $dirElement = $xml.CreateElement("Directory", $namespace)
        $dirElement.SetAttribute("Id", $dirId)
        $dirElement.SetAttribute("Name", $dirName)
        $ParentXmlElement.AppendChild($dirElement) | Out-Null

        New-DirectoryNode -Directory $subDir -ParentXmlElement $dirElement
    }
}

New-DirectoryNode -Directory (Get-Item $PublishDir) -ParentXmlElement $dirRef

# === Components ===
$fragment2 = $xml.CreateElement("Fragment", $namespace)
$wix.AppendChild($fragment2) | Out-Null
$components = $xml.CreateElement("ComponentGroup", $namespace)
$components.SetAttribute("Id", "ExtraDistributables")
$fragment2.AppendChild($components) | Out-Null

# === Get directories and files ===
$dirGroups = @{}

# Group all files by directory
Get-ChildItem -Path $PublishDir -Recurse -File | Where-Object { $_.Extension -ne ".xll" } | ForEach-Object {
    $relDir = if ($_.Directory.FullName -eq $PublishDir) { "" } else { $_.DirectoryName.Substring($PublishDir.Length).Trim("\") }
    if (-not $dirGroups.ContainsKey($relDir)) { $dirGroups[$relDir] = @() }
    $dirGroups[$relDir] += $_
}

# Ensure ALL directories get an entry, even if empty
Get-ChildItem -Path $PublishDir -Recurse -Directory | ForEach-Object {
    $relDir = $_.FullName.Substring($PublishDir.Length).Trim("\")
    if (-not $dirGroups.ContainsKey($relDir)) { $dirGroups[$relDir] = @() }
}
# Also add root directory explicitly
if (-not $dirGroups.ContainsKey("")) { $dirGroups[""] = @() }

# === Emit Components ===
foreach ($kvp in $dirGroups.GetEnumerator()) {
    $relDirPath = $kvp.Key
    $filesInDir = $kvp.Value

    $dirId = if ($relDirPath -eq "") { $RootDirId } else { ($relDirPath -replace '[\\\/]', '_') + "_DIR" }
    $stableGuidInput = if ($relDirPath -ne "") { $relDirPath } else { "_root_" }
    $componentGuid = Get-StableGuid $stableGuidInput

    $compElement = $xml.CreateElement("Component", $namespace)
    $compElement.SetAttribute("Guid", "{$componentGuid}")
    $compElement.SetAttribute("Directory", $dirId)

    # Always include a RegistryValue for KeyPath
    $reg = $xml.CreateElement("RegistryValue", $namespace)
    $regKey = $RegKeyBase
    if ($relDirPath -ne "") {
        $relClean = ($relDirPath -replace '^[\\\/]+', '').Replace('/', '\').Replace('\\', '\')
        $regKey += "\" + $relClean
    }
    $reg.SetAttribute("Root", $RegistryRoot)
    $reg.SetAttribute("Key", $regKey)
    $reg.SetAttribute("Name", "Installed")
    $reg.SetAttribute("Type", "string")
    $reg.SetAttribute("Value", "1")
    $reg.SetAttribute("KeyPath", "yes")
    $compElement.AppendChild($reg) | Out-Null

    # Add files and RemoveFile for each file with unique ID
    foreach ($file in $filesInDir) {
        $relPath = $file.FullName.Substring($PublishDir.Length).Trim("\")
        $safeRelPath = ($relPath -replace '[\\\/]', '_')

        $fileElem = $xml.CreateElement("File", $namespace)
        $fileElem.SetAttribute("Id", "Fil_" + $safeRelPath)
        $fileElem.SetAttribute("Source", $file.FullName)
        $compElement.AppendChild($fileElem) | Out-Null

        $removeElem = $xml.CreateElement("RemoveFile", $namespace)
        $removeElem.SetAttribute("Id", "Remove_" + $safeRelPath)  # Use rel path for uniqueness
        $removeElem.SetAttribute("Name", $file.Name)
        $removeElem.SetAttribute("On", "uninstall")
        $compElement.AppendChild($removeElem) | Out-Null
    }

    # If root dir: do not add RemoveFolder for root
    if ($relDirPath -ne "") {
        # Subdirectories only get RemoveFolder
        $removeFolder = $xml.CreateElement("RemoveFolder", $namespace)
        $safeDirId = if ($relDirPath -eq "") { "_root_" } else { ($relDirPath -replace '[\\\/]', '_') }
        $removeFolder.SetAttribute("Id", "RemoveFolder_" + $safeDirId)
        $removeFolder.SetAttribute("Directory", $dirId)
        $removeFolder.SetAttribute("On", "uninstall")
        $compElement.AppendChild($removeFolder) | Out-Null
    }

    $components.AppendChild($compElement) | Out-Null
}

$xml.Save($OutputWxs)
Write-Host "✅ Wrote $OutputWxs to $OutputWxs"
