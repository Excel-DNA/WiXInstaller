param(
    [string]$PublishDir,
    [string]$OutputWxs,
    [string]$RootDirId = "AddinFolder",
    [string]$RegistryRoot = "HKCU",
    [string]$RegKeyBase = "Software\MyCompany\MyProduct\ExtraFiles"
)

# === Helper: make stable GUID from any string ===
function Get-StableGuid {
    param([string]$InputString)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $hash = $md5.ComputeHash($bytes)

    # Format: 8-4-4-4-12 hex digits
    $hex = [System.BitConverter]::ToString($hash).Replace("-", "")
    $guidStr = "{0}-{1}-{2}-{3}-{4}" -f `
        $hex.Substring(0, 8),
        $hex.Substring(8, 4),
        $hex.Substring(12, 4),
        $hex.Substring(16, 4),
        $hex.Substring(20, 12)

    return $guidStr
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

# === Group files by directory ===
$dirGroups = @{}
Get-ChildItem -Path $PublishDir -Recurse -File | Where-Object { $_.Extension -ne ".xll" } | ForEach-Object {
    $relDir = if ($_.Directory.FullName -eq $PublishDir) { "" } else { $_.DirectoryName.Substring($PublishDir.Length).Trim("\") }
    if (-not $dirGroups.ContainsKey($relDir)) { $dirGroups[$relDir] = @() }
    $dirGroups[$relDir] += $_
}

# === Emit one component per directory ===
foreach ($kvp in $dirGroups.GetEnumerator()) {
    $relDirPath = $kvp.Key
    $filesInDir = $kvp.Value

    $dirId = if ($relDirPath -eq "") { $RootDirId } else { ($relDirPath -replace '[\\\/]', '_') + "_DIR" }
    $stableGuidInput = if ($relDirPath -ne "") { $relDirPath } else { "_root_" }
    $componentGuid = Get-StableGuid $stableGuidInput

    $compElement = $xml.CreateElement("Component", $namespace)
    $compElement.SetAttribute("Guid", "{$componentGuid}")
    $compElement.SetAttribute("Directory", $dirId)

    # Always create a valid RegistryValue once per component
    $reg = $xml.CreateElement("RegistryValue", $namespace)

    # Valid Key: subkey per directory
    $regKey = $RegKeyBase
    if ($relDirPath -ne "") {
        $relClean = ($relDirPath -replace '^[\\\/]+', '').Replace('/', '\').Replace('\\', '\')
        if ($relClean -ne "") {
            $regKey += "\" + $relClean
        }
    }
    $reg.SetAttribute("Root", $RegistryRoot)
    $reg.SetAttribute("Key", $regKey)
    $reg.SetAttribute("Name", "Installed")
    $reg.SetAttribute("Type", "string")
    $reg.SetAttribute("Value", "1")
    $reg.SetAttribute("KeyPath", "yes")
    $compElement.AppendChild($reg) | Out-Null

    # Files in this directory
    foreach ($file in $filesInDir) {
        $relPath = $file.FullName.Substring($PublishDir.Length).Trim("\")
        $fileElem = $xml.CreateElement("File", $namespace)
        $fileElem.SetAttribute("Id", "Fil_" + ($relPath -replace '[\\\/]', '_'))
        $fileElem.SetAttribute("Source", $file.FullName)
        $compElement.AppendChild($fileElem) | Out-Null

        $removeElem = $xml.CreateElement("RemoveFile", $namespace)
        $removeElem.SetAttribute("Id", "Remove_" + $file.Name)
        $removeElem.SetAttribute("Name", $file.Name)
        $removeElem.SetAttribute("On", "uninstall")
        $compElement.AppendChild($removeElem) | Out-Null
    }

    if ($relDirPath -ne "") {
        $removeFolder = $xml.CreateElement("RemoveFolder", $namespace)
        $removeFolder.SetAttribute("Id", "Remove_" + ($relDirPath -replace '[\\\/]', '_'))
        $removeFolder.SetAttribute("Directory", $dirId)
        $removeFolder.SetAttribute("On", "uninstall")
        $compElement.AppendChild($removeFolder) | Out-Null
    }

    $components.AppendChild($compElement) | Out-Null
}

$xml.Save($OutputWxs)
Write-Host "✅ Wrote $OutputWxs to $OutputWxs"
