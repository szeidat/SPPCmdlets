<#
    .SYNOPSIS
        Set path to the SPP folder.

    .DESCRIPTION
        The Set-HPSPPFolderPath command sets the path to the folder where the SPP contents are located.

    .PARAMETER Path
        SPP folder path.

    .EXAMPLE
        Set-HPSPPFolderPath E:\
        Set SPP folder path to e:\ drive.

    .INPUTS
        None

    .OUTPUTS
        None
#>
Function Set-HPSPPFolderPath {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="SPP folder path", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $Path
    )

    # Check SPP folder path
    if (!(Test-Path -PathType Container $Path)) {
        Write-Error "SPP folder '$Path' not found"
        Return
    }

    # Set SPP folder path
    $Script:HPSPPPath = $Path
}

<#
    .SYNOPSIS
        Get path to the SPP folder.

    .DESCRIPTION
        The Get-HPSPPFolderPath command gets the path set for the folder where the SPP contents are located.

    .EXAMPLE
        Get-HPSPPFolderPath
        Get path set for SPP folder.

    .INPUTS
        None

    .OUTPUTS
        String
#>
Function Get-HPSPPFolderPath {

    # Get SPP folder path
    if ($Script:HPSPPPath) {
        Get-ChildItem $Script:HPSPPPath
    }
}

<#
    .SYNOPSIS
        Get SPP component system filter objects.

    .DESCRIPTION
        The Get-HPSPPSystemFilter command gets one or more SPP component filter objects based on the system name provided. These objects can be used to filter the output of Get-HPSPPComponent command.

    .PARAMETER Name
        System name. Wildcards are accepted (e.g. *DL360*). If omitted the command will get all system filter objects. 

    .EXAMPLE
        Get-HPSPPSystemFilter *DL380*
        Get component filter objects for all systems that contain 'DL360' in their name.

    .EXAMPLE
        Get-HPSPPSystemFilter *G5*,*G6*
        Get component filter objects for all systems that contain 'G5' or 'G6' in their name.

    .INPUTS
        String[]

    .OUTPUTS
        HPSPPFilter[]
#>
Function Get-HPSPPSystemFilter {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="System name", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check SPP system manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\system.xml"
        if (!(Test-Path -PathType Leaf $Manifest)) {
            Write-Error "SPP system manifest file '$Manifest' not found"
            Return
        }

        # Get system xml nodes
        $Nodes = Select-Xml -XPath "/hp_manifest/systems/system" -Path $Manifest

        # Create system objects
        $Systems = @()
        foreach ($node in $Nodes) {
            $System = New-Object PSObject -Property ([ordered]@{
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                ID = $node.Node.id
                SystemID = $node.Node.systemid
                FirmwareID = $node.Node.firmwareid
                Node = $node.Node
            })

            $System.PSObject.TypeNames.Insert(0,'HPSPPFilter')
            $System.PSObject.TypeNames.Insert(1,'HPSPPSystemFilter')
            $Systems += $System
        }

        # Systems by name
        $SystemsByName = @()

        # Process names from variable
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($system in $Systems) {
                if ($system -notin $SystemsByName) {
                    if (($system.Name -like $value) -or ($system.Name -match $value_escaped)) {
                        $SystemsByName += $system
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($system in $Systems) {
                if ($system -notin $SystemsByName) {
                    if (($system.Name -like $value) -or ($system.Name -match $value_escaped)) {
                        $SystemsByName += $system
                    }
                }
            }
        }
    }

    END {
        # Output systems
        if (!$Name) {
            Write-Output $Systems
        } elseif ($SystemsByName) {
            Write-Output $SystemsByName
        }
    }
}

<#
    .SYNOPSIS
        Get SPP component operating system filter objects.

    .DESCRIPTION
        The Get-HPSPPOperatingSystemFilter command gets one or more SPP component filter objects based on the operating system name provided. These objects can be used to filter the output of Get-HPSPPComponent command.

    .PARAMETER Name
        Operating system name. Wildcards are accepted (e.g. *Windows*). If omitted the command will get all operating system filter objects.

    .EXAMPLE
        Get-HPSPPOperatingSystemFilter *Windows*
        Get component filter objects for all operating systems that contain 'Windows' in their name.

    .EXAMPLE
        Get-HPSPPSystemFilter *VMware*,*Linux*
        Get component filter objects for all operating systems that contain 'VMware' or 'Linux' in their name.

    .INPUTS
        String[]

    .OUTPUTS
        HPSPPFilter[]
#>
Function Get-HPSPPOperatingSystemFilter {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Operating system name", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check SPP operating system manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\os.xml"
        if (!(Test-Path -PathType Leaf $Manifest)) {
            Write-Error "SPP operating system manifest file '$Manifest' not found"
            Return
        }

        # Get operating system xml nodes
        $Nodes = Select-Xml -XPath "/hp_manifest/operating_systems/operating_system" -Path $Manifest

        # Create operating system objects
        $OperatingSystems = @()
        foreach ($node in $Nodes) {
            $OperatingSystem = New-Object PSObject -Property ([ordered]@{
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                ID = $node.Node.id
                Node = $node.Node
            })

            $OperatingSystem.PSObject.TypeNames.Insert(0,'HPSPPFilter')
            $OperatingSystem.PSObject.TypeNames.Insert(1,'HPSPPOperatingSystemFilter')
            $OperatingSystems += $OperatingSystem
        }

        # Operating systems by name
        $OperatingSystemsByName = @()

        # Process names from variable
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($os in $OperatingSystems) {
                if ($os -notin $OperatingSystemsByName) {
                    if (($os.Name -like $value) -or ($os.Name -match $value_escaped)) {
                        $OperatingSystemsByName += $os
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($os in $OperatingSystems) {
                if ($os -notin $OperatingSystemsByName) {
                    if (($os.Name -like $value) -or ($os.Name -match $value_escaped)) {
                        $OperatingSystemsByName += $os
                    }
                }
            }
        }
    }

    END {
        # Output operating systems
        if (!$Name) {
            Write-Output $OperatingSystems
        } elseif ($OperatingSystemsByName) {
            Write-Output $OperatingSystemsByName
        }
    }
}

<#
    .SYNOPSIS
        Get SPP component category filter objects.

    .DESCRIPTION
        The Get-HPSPPCategoryFilter command gets one or more SPP component filter objects based on the category name provided. These objects can be used to filter the output of Get-HPSPPComponent command.

    .PARAMETER Name
        Category name. Wildcards are accepted (e.g. *Driver*). If omitted the command will get all category filter objects.

    .EXAMPLE
        Get-HPSPPCategoryFilter *Driver*
        Get component filter objects for all categories that contain 'Driver' in their name.

    .EXAMPLE
        Get-HPSPPCategoryFilter *Bios*,*Firmware*
        Get component filter objects for all categories that contain 'Bios' or 'Firmware' in their name.

    .INPUTS
        String[]

    .OUTPUTS
        HPSPPFilter[]
#>
Function Get-HPSPPCategoryFilter {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Category name", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check SPP category manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\category.xml"
        if (!(Test-Path -PathType Leaf $Manifest)) {
            Write-Error "SPP category manifest file '$Manifest' not found"
            Return
        }

        # Get category xml nodes
        $Nodes = Select-Xml -XPath "/hp_manifest/categories/category" -Path $Manifest

        # Create category objects
        $Categories = @()
        foreach ($node in $Nodes) {
            $Category = New-Object PSObject -Property ([ordered]@{
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                ID = $node.Node.id
                Node = $node.Node
            })
            
            $Category.PSObject.TypeNames.Insert(0,'HPSPPFilter')
            $Category.PSObject.TypeNames.Insert(1,'HPSPPCategoryFilter')
            $Categories += $Category
        }

        # Categories by name
        $CategoriesByName = @()

        # Process names from variable
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($category in $Categories) {
                if ($category -notin $CategoriesByName) {
                    if (($category.Name -like $value) -or ($category.Name -match $value_escaped)) {
                        $CategoriesByName += $category
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($category in $Categories) {
                if ($category -notin $CategoriesByName) {
                    if (($category.Name -like $value) -or ($category.Name -match $value_escaped)) {
                        $CategoriesByName += $category
                    }
                }
            }
        }
    }

    END {
        # Output categories
        if (!$Name) {
            Write-Output $Categories
        } elseif ($CategoriesByName) {
            Write-Output $CategoriesByName
        }
    }
}

<#
    .SYNOPSIS
        Get SPP component device filter objects.

    .DESCRIPTION
        The Get-HPSPPDeviceFilter command gets one or more SPP component filter objects based on the device name provided. These objects can be used to filter the output of Get-HPSPPComponent command.

    .PARAMETER Name
        Device name. Wildcards are accepted (e.g. *iLO*). If omitted the command will get all device filter objects.

    .EXAMPLE
        Get-HPSPPDeviceFilter *iLO*
        Get component filter objects for all devices that contain 'iLO' in their name.

    .EXAMPLE
        Get-HPSPPDeviceFilter *iLO*,*Array*
        Get component filter objects for all devices that contain 'iLO' or 'Array' in their name.

    .INPUTS
        String[]

    .OUTPUTS
        HPSPPFilter[]
#>
Function Get-HPSPPDeviceFilter {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Device name", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check SPP device manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\device.xml"
        if (!(Test-Path -PathType Leaf $Manifest)) {
            Write-Error "SPP device manifest file '$Manifest' not found"
            Return
        }

        # Get device xml nodes
        $Nodes = Select-Xml -XPath "/hp_manifest/devices/device" -Path $Manifest

        # Create device objects
        $Devices = @()
        foreach ($node in $Nodes) {
            $Device = New-Object PSObject -Property ([ordered]@{
                Name = $node.Node.name.SelectSingleNode("name_xlate[@lang='en']").InnerText
                ID = $node.Node.id
                DeviceType = $node.Node.device_type
                DeviceIDString = $node.Node.deviceidstring
                Node = $node.Node
            })

            $Device.PSObject.TypeNames.Insert(0,'HPSPPFilter')
            $Device.PSObject.TypeNames.Insert(1,'HPSPPDeviceFilter')
            $Devices += $Device
        }

        # Devices by name
        $DevicesByName = @()

        # Process names from variable
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($device in $Devices) {
                if ($device -notin $DevicesByName) {
                    if (($device.Name -like $value) -or ($device.Name -match $value_escaped)) {
                        $DevicesByName += $device
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($device in $Devices) {
                if ($device -notin $DevicesByName) {
                    if (($device.Name -like $value) -or ($device.Name -match $value_escaped)) {
                        $DevicesByName += $device
                    }
                }
            }
        }
    }

    END {
        # Output devices
        if (!$Name) {
            Write-Output $Devices
        } elseif ($DevicesByName) {
            Write-Output $DevicesByName
        }
    }
}

<#
    .SYNOPSIS
        Get SPP component type filter objects.

    .DESCRIPTION
        The Get-HPSPPTypeFilter command gets one or more SPP component filter objects based on the type name provided. These objects can be used to filter the output of Get-HPSPPComponent command.

    .PARAMETER Name
        Type name. Wildcards are accepted (e.g. *software*). If omitted the command will get all type filter objects.

    .EXAMPLE
        Get-HPSPPTypeFilter *software*
        Get component filter objects for all types that contain 'software' in their name.

    .EXAMPLE
        Get-HPSPPTypeFilter *software*,*firmware*
        Get component filter objects for all types that contain 'software' or 'firmware' in their name.

    .INPUTS
        String[]

    .OUTPUTS
        HPSPPFilter[]
#>
Function Get-HPSPPTypeFilter {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Type name", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check SPP type manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\type.xml"
        if (!(Test-Path -PathType Leaf $Manifest)) {
            Write-Error "SPP type manifest file '$Manifest' not found"
            Return
        }

        # Get type xml nodes
        $Nodes = Select-Xml -XPath "/hp_manifest/types/type" -Path $Manifest

        # Create type objects
        $Types = @()
        foreach ($node in $Nodes) {
            $Type = New-Object PSObject -Property ([ordered]@{
                Name = $node.Node.id
                ID = $node.Node.id
                Node = $node.Node
            })

            $Type.PSObject.TypeNames.Insert(0,'HPSPPFilter')
            $Type.PSObject.TypeNames.Insert(1,'HPSPPTypeFilter')
            $Types += $Type
        }

        # Types by name
        $TypesByName = @()

        # Process names from variable
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($type in $Types) {
                if ($type -notin $TypesByName) {
                    if (($type.Name -like $value) -or ($type.Name -match $value_escaped)) {
                        $TypesByName += $type
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            $value_escaped = [Regex]::Escape($value)
            foreach ($type in $Types) {
                if ($type -notin $TypesByName) {
                    if (($type.Name -like $value) -or ($type.Name -match $value_escaped)) {
                        $TypesByName += $type
                    }
                }
            }
        }
    }

    END {
        # Output types
        if (!$Name) {
            Write-Output $Types
        } elseif ($TypesByName) {
            Write-Output $TypesByName
        }
    }
}

<#
    .SYNOPSIS
        Get SPP component objects.

    .DESCRIPTION
        The Get-HPSPPComponent command gets one or more SPP component objects based on the filter provided. Filter objects based on system, operating system, category, device, or type can be used.

    .PARAMETER Filter
        Component filter objects. If omitted the command will get all components.

    .EXAMPLE
        $filter = Get-HPSPPSystemFilter *DL360*; Get-HPSPPComponent $filter
        Get component objects filtered by systems containing 'DL360' in their name.

    .EXAMPLE
        $filter = @(); $filter += Get-HPSPPOperatingSystemFilter *Windows*; $filter += Get-HPSPPCategory *Driver*; Get-HPSPPComponent $filter
        Get component objects filtered by operating system names containing 'Windows' and category names containing 'Driver'.

    .INPUTS
        HPSPPFilter[]

    .OUTPUTS
        HPSPPComponent[]
#>
Function Get-HPSPPComponent {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Component filter", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Filter
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check SPP meta manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\meta.xml"
        if (!(Test-Path -PathType Leaf $Manifest)) {
            Write-Error "SPP meta manifest file '$Manifest' not found"
            Return
        }

        # Get component xml nodes
        $Nodes = Select-Xml -XPath "/hp_manifest/meta/product/product_version/id/.." -Path $Manifest

        # Create component objects
        $Components = @()
        foreach ($node in $Nodes) {
            $Component = New-Object PSObject -Property ([ordered]@{
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                ProductID = $node.Node.id.product
                VersionID = $node.Node.id.version
                Node = $node.Node
            })

            $Component.PSObject.TypeNames.Insert(0,'HPSPPComponent')
            $Components += $Component
        }
        
        # Filters
        $SystemFilters = @()
        $OperatingSystemFilters = @()
        $CategoryFilters = @()
        $DeviceFilters = @()
        $TypeFilters = @()

        # Process filters from variable
        foreach ($value in $Filter) {
            if ($value.PSObject.TypeNames -contains "HPSPPSystemFilter") {
                $SystemFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPOperatingSystemFilter") {
                $OperatingSystemFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPCategoryFilter") {
                $CategoryFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPDeviceFilter") {
                $DeviceFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPTypeFilter") {
                $TypeFilters += $value
            }
        }
    }

    PROCESS {
        # Process filters from pipeline
        foreach ($value in $Filter) {
            if ($value.PSObject.TypeNames -contains "HPSPPSystemFilter") {
                $SystemFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPOperatingSystemFilter") {
                $OperatingSystemFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPCategoryFilter") {
                $CategoryFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPDeviceFilter") {
                $DeviceFilters += $value
            }
            elseif ($value.PSObject.TypeNames -contains "HPSPPTypeFilter") {
                $TypeFilters += $value
            }
        }
    }

    END {
        # Filter components
        if (!$Filter) {
            $ComponentsByFilter = $Components
        }
        else {
            $ComponentsByFilter = @()
            foreach ($component in $Components) {
                if ($SystemFilters) {
                    $keep = $false
                    foreach ($value in $SystemFilters) {
                        $keep = Select-Xml -XPath "product_version/id[@product='$($component.ProductID)'][@version='$($component.VersionID)']" -Xml $value.Node
                        if ($keep) {break}
                    }
                    if (!$keep) {continue}
                }

                if ($OperatingSystemFilters) {
                    $keep = $false
                    foreach ($value in $OperatingSystemFilters) {
                        $keep = Select-Xml -XPath "product_version/id[@product='$($component.ProductID)'][@version='$($component.VersionID)']" -Xml $value.Node
                        if ($keep) {break}
                    }
                    if (!$keep) {continue}
                }

                if ($CategoryFilters) {
                    $keep = $false
                    foreach ($value in $CategoryFilters) {
                        $keep = Select-Xml -XPath "product_version/id[@product='$($component.ProductID)'][@version='$($component.VersionID)']" -Xml $value.Node
                        if ($keep) {break}
                    }
                    if (!$keep) {continue}
                }

                if ($DeviceFilters) {
                    $keep = $false
                    foreach ($value in $DeviceFilters) {
                        $keep = Select-Xml -XPath "product_version/id[@product='$($component.ProductID)'][@version='$($component.VersionID)']" -Xml $value.Node
                        if ($keep) {break}
                    }
                    if (!$keep) {continue}
                }

                if ($TypeFilters) {
                    $keep = $false
                    foreach ($value in $TypeFilters) {
                        $keep = Select-Xml -XPath "product_version/id[@product='$($component.ProductID)'][@version='$($component.VersionID)']" -Xml $value.Node
                        if ($keep) {break}
                    }
                    if (!$keep) {continue}
                }

                if ($keep) {
                    $ComponentsByFilter += $component
                }
            }
        }

        # Output components
        foreach ($component in $ComponentsByFilter) {
            $component | Add-Member -NotePropertyMembers ([ordered]@{
                Description = $component.Node.SelectSingleNode("description/description_xlate[@lang='en']").InnerText
                AltName = $component.Node.SelectSingleNode("alt_name/alt_name_xlate[@lang='en']").InnerText
                FileName = $component.Node.filename
                Version = $component.Node.version.value
                Revision = $component.Node.version.revision
                UpgradeRequirement = $component.Node.upgrade_requirement.value
                Different = $component.Node.different
                BuildNumber = $component.Node.build.number
                TypeID = $component.Node.type.id
                Category = $component.Node.SelectSingleNode("category/category_xlate[@lang='en']").InnerText
                Manufacturer = $component.Node.SelectSingleNode("manufacturer_name/manufacturer_name_xlate[@lang='en']").InnerText
                TypeOfChange = $component.Node.version.type_of_change
                ReleaseDate = "$($component.Node.release_date.year)-$($component.Node.release_date.month)-$($component.Node.release_date.day) $($component.Node.release_date.hour):$($component.Node.release_date.minute):$($component.Node.release_date.second)"
                RequiredDiskSpaceKB = $component.Node.prerequisites.required_diskspace.size_kb
                OperatingSystems = @()
                Divisions = @()
                Files = @()
            })

            $component.Node | Select-Xml -XPath "operating_systems/operating_system" | ForEach-Object {
                $operatingsystem = $_
                $component.OperatingSystems += $operatingsystem.Node.operating_system_xlate
            }

            $component.Node | Select-Xml -XPath "divisions/division" | ForEach-Object {
                $division = $_
                $component.Divisions += $division.Node.SelectSingleNode("division_xlate[@lang='en']").InnerText
            }

            $component.Node | Select-Xml -XPath "files/file" | ForEach-Object {
                $file = $_
                $component.Files += ([ordered]@{
                    Name = $file.Node.name
                    FtpUrl = $file.Node.SelectSingleNode("url[@id='http://ftp.hp.com']").InnerText
                    FileUrl = Join-Path $Script:HPSPPPath ($file.Node.SelectSingleNode("url[@id='file://.']").InnerText -replace "file://./","")
                    Size = $file.Node.size
                    DateModified = $file.Node.date_modified
                    Md5Sum = $file.Node.md5sum
                    Sha1Sum = $file.Node.sha1sum
                })
            }

            Write-Output $component
        }
    }
}

<#
    .SYNOPSIS
        Copy SPP component files to a destination folder.

    .DESCRIPTION
        The Copy-HPSPPComponent command copies SPP component files to the destination folder provided.

    .PARAMETER Component
        Source components.

    .PARAMETER DestinationFolder
        Destination folder. Folder must exist. If a file to be copied already exists in the destination folder it will be overwritten.

    .EXAMPLE
        Get-HPSPPSystemFilter *DL360* | Get-HPSPPComponent | Copy-HPSPPComponent -DestinationFolder C:\custom
        Get component objects filtered by systems containing 'DL360' in their name then copy component files to destination folder 'C:\custom'.

    .EXAMPLE
        $filter = @(); $filter += Get-HPSPPOperatingSystemFilter *Windows*; $filter += Get-HPSPPCategory *Driver*; $components = Get-HPSPPComponent $filter; Copy-HPSPPComponent $components C:\custom
        Get component objects filtered by the operating systems and categories provided then copy component files to destination folder 'C:\custom'.

    .INPUTS
        HPSPPComponent[]

    .OUTPUTS
        None
#>
Function Copy-HPSPPComponent {
    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Component to copy", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Component,

    [Parameter(Mandatory=$true, HelpMessage="Destination folder", Position=1)]
    [ValidateNotNullOrEmpty()]
    [String]
    $DestinationFolder
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Write-Error "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
            Return
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container $Script:HPSPPPath)) {
            Write-Error "SPP folder '$Script:HPSPPPath' not found"
            Return
        }

        # Check destination folder path
        if (!(Test-Path -PathType Container $DestinationFolder)) {
            Write-Error "Destination folder '$DestinationFolder' not found"
            Return
        }

        # Components to copy
        $Components = @()

        # Process components from variable
        foreach ($value in $Component) {
            if ($value.PSObject.TypeNames -contains "HPSPPComponent") {
                if ($value -notin $Components) {
                    $Components += $value
                }
            }
        }
    }

    PROCESS {
        # Process components from pipeline
        foreach ($value in $Component) {
            if ($value.PSObject.TypeNames -contains "HPSPPComponent") {
                if ($value -notin $Components) {
                    $Components += $value
                }
            }
        }
    }

    END {
        # Copy components
        $copied = 0
        foreach ($component in $Components) {
            foreach ($file in $component.Files) {
                Write-Progress -Activity "Copying file" -Status "$($file.FileUrl) to $DestinationFolder" -PercentComplete ([int]($copied++ / $Components.Count * 100))
                Copy-Item -Path $file.FileUrl -Destination $DestinationFolder -Force
            }
        }
    }
}
