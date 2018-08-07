## 2.0.0
Re-wrote the module to support querying multiple bundles.
Added support for CSV output formatting.
Updated HTML output formatting to support comparing component versions.

Added commands:
  - Add-SPPBundle
  - Remove-SPPBundle
  - ConvertTo-SPPBundleHtml
  - ConvertTo-SPPBundleCsv
  - ConvertTo-SPPComponentHtml
  - ConvertTo-SPPComponentCsv

Changed commands:
  - Get-SPPSystem
  - Get-SPPOperatingSystem
  - Get-SPPCategory
  - Get-SPPDevice
  - Get-SPPType

Removed commands
  - Set-SPPFolderPath
  - Get-SPPFolderPath
  - Get-SPPBundleHtml
  - Get-SPPComponentHtml

## 1.3.0
Added support for new SPP structure starting with SPP 2017.07.0

## 1.2.2
Renamed module and commands from HPSPP* to SPP*

## 1.2.1
Modified Get-HPSPPComponentHtml to use streams
Modified Get-HPSPPComponentHtml to automatically open html results
Modified Get-HPSPPBundleHtml to use streams
Modified Get-HPSPPBundleHtml to automatically open html results

## 1.2.0
Added commands
  - Get-HPSPPBundle
  - Get-HPSPPBundleHtml

Added SPPVersion property to components

## 1.1.0
Added commands:
  - Get-HPSPPComponentHtml

Removed regex name matching

## 1.0.0
Initial release. Implemented commands:
  - Set-HPSPPFolderPath
  - Get-HPSPPFolderPath
  - Get-HPSPPSystemFilter
  - Get-HPSPPOperatingSystemFilter
  - Get-HPSPPCategoryFilter
  - Get-HPSPPDeviceFilter
  - Get-HPSPPTypeFilter
  - Get-HPSPPComponent
  - Copy-HPSPPComponent