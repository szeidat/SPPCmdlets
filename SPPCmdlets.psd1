#
# Module manifest for module 'SPPCmdlets'
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'SPPCmdlets.psm1'

# Version number of this module.
ModuleVersion = '2.0.2'

# ID used to uniquely identify this module
GUID = '23c34f9c-8d5f-47d1-ac56-74c48c896ad0'

# Author of this module
Author = 'Sameer Zeidat'

# Description of the functionality provided by this module
Description = 'Cmdlets for processing HPE Service Pack for ProLiant (SPP) contents'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '3.0'

# Minimum version of the .NET Framework required by this module
DotNetFrameworkVersion = '4.0'

# Minimum version of the common language runtime (CLR) required by this module
CLRVersion = '4.0'

# Formats to process from this module
FormatsToProcess = 'SPPCmdlets.ps1xml'

# Functions to export from this module
FunctionsToExport = @(
    'Add-SPPBundle',
    'Get-SPPBundle',
    'Remove-SPPBundle'
    'ConvertTo-SPPBundleCsv',
    'ConvertTo-SPPBundleHtml',
    'Get-SPPCategory',
    'Get-SPPDevice',
    'Get-SPPOperatingSystem',
    'Get-SPPSystem',
    'Get-SPPType',
    'Get-SPPComponent',
    'Copy-SPPComponent',
    'ConvertTo-SPPComponentCsv',
    'ConvertTo-SPPComponentHtml'
    )

# Private data
PrivateData = @{
PSData = @{
# Tags applied to this module. These help with module discovery in online galleries.
Tags = @('HPE', 'ProLiant', 'SPP')

# A URL to the license for this module.
LicenseUri = 'https://github.com/szeidat/SPPCmdlets/blob/master/LICENSE'

# A URL to the main website for this project.
ProjectUri = 'https://github.com/szeidat/SPPCmdlets'

# ReleaseNotes of this module
ReleaseNotes = 'Version 2.0.2
Removed extra files'
}
}
}
