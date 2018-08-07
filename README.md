# SPPCmdlets

PowerShell module for querying HPE Service Pack for ProLiant (SPP) bundles

## Description:

This is a basic PowerShell module for processing HPE Service Pack for ProLiant (SPP) contents.
It contains commands to parse manifest xml files in SPP folders to query, search, filter, and report on SPP bundles and components.

## Usage:

Copy the module folder to the powershell modules folder (e.g. %USERPROFILE%\Documents\WindowsPowerShell\Modules).
Import the module using the command "Import-Module SPPCmdlets".
Use the "Add-SPPBundle" command to add SPP bundles for processing.
Use the "Get-SPP*" commands to retreive bundles and components.
Use the "ConvertTo-SPP*" commands to produce HTML or CSV reports of bundles and components.
