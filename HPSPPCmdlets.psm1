Function Set-HPSPPFolderPath {
<#
    .SYNOPSIS
        Set path to the SPP folder.

    .DESCRIPTION
        The Set-HPSPPFolderPath command sets the path to the folder where the SPP contents are located.

    .PARAMETER Path
        SPP folder path.

    .EXAMPLE
        Set-HPSPPFolderPath E:\
        Set SPP folder path to e:\.

    .INPUTS
        None

    .OUTPUTS
        None
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="SPP folder path", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $Path
    )

    # Check SPP folder path
    if (!(Test-Path -PathType Container -LiteralPath $Path)) {
        Throw "SPP folder '$Path' not found"
    }

    # Set SPP folder path
    $Script:HPSPPPath = $Path

    # Set SPP version
    $Bundles = Join-Path $Path "hp\swpackages\bp*.xml"
    if (Test-Path -PathType Leaf -Path $Bundles) {
        $Bundle = Resolve-Path -Path $Bundles
        if ($Bundle -isnot [Array]) {
            $Node = Select-Xml -XPath "/cpq_bundle" -Path $Bundle
            if ($Node) {
                $Script:HPSPPVersion = $Node.Node.version.value + $Node.Node.version.revision
            }
        }
    }
}

Function Get-HPSPPFolderPath {
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
        System.IO.DirectoryInfo
#>

    [CmdletBinding()]
    Param ()

    # Get SPP folder path
    if ($Script:HPSPPPath) {
        Get-ChildItem $Script:HPSPPPath
    }
}

Function Get-HPSPPSystemFilter {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP system manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\system.xml"
        if (!(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP system manifest file '$Manifest' not found"
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
            foreach ($system in $Systems) {
                if ($system -notin $SystemsByName) {
                    if ($system.Name -like $value) {
                        $SystemsByName += $system
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            foreach ($system in $Systems) {
                if ($system -notin $SystemsByName) {
                    if ($system.Name -like $value) {
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

Function Get-HPSPPOperatingSystemFilter {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP operating system manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\os.xml"
        if (!(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP operating system manifest file '$Manifest' not found"
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
            foreach ($os in $OperatingSystems) {
                if ($os -notin $OperatingSystemsByName) {
                    if ($os.Name -like $value) {
                        $OperatingSystemsByName += $os
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            foreach ($os in $OperatingSystems) {
                if ($os -notin $OperatingSystemsByName) {
                    if ($os.Name -like $value) {
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

Function Get-HPSPPCategoryFilter {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP category manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\category.xml"
        if (!(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP category manifest file '$Manifest' not found"
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
            foreach ($category in $Categories) {
                if ($category -notin $CategoriesByName) {
                    if ($category.Name -like $value) {
                        $CategoriesByName += $category
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            foreach ($category in $Categories) {
                if ($category -notin $CategoriesByName) {
                    if ($category.Name -like $value) {
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

Function Get-HPSPPDeviceFilter {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP device manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\device.xml"
        if (!(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP device manifest file '$Manifest' not found"
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
            foreach ($device in $Devices) {
                if ($device -notin $DevicesByName) {
                    if ($device.Name -like $value) {
                        $DevicesByName += $device
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            foreach ($device in $Devices) {
                if ($device -notin $DevicesByName) {
                    if ($device.Name -like $value) {
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

Function Get-HPSPPTypeFilter {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP type manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\type.xml"
        if (!(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP type manifest file '$Manifest' not found"
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
            foreach ($type in $Types) {
                if ($type -notin $TypesByName) {
                    if ($type.Name -like $value) {
                        $TypesByName += $type
                    }
                }
            }
        }
    }

    PROCESS {
        # Process names from pipeline
        foreach ($value in $Name) {
            foreach ($type in $Types) {
                if ($type -notin $TypesByName) {
                    if ($type.Name -like $value) {
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

Function Get-HPSPPComponent {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP meta manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\meta.xml"
        if (!(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP meta manifest file '$Manifest' not found"
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
                SPPVersion = $Script:HPSPPVersion
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

function Get-HPSPPComponentHtml {
<#
    .SYNOPSIS
        Get SPP component details in html format.

    .DESCRIPTION
        The Get-HPSPPComponentHtml command gets SPP component details in html format.

    .PARAMETER Component
        Component objects.

    .PARAMETER Full
        Get full details (includes notes on prerequisites, installation, availability, and documentation as well as revision history).

    .EXAMPLE
        Get-HPSPPComponent | Get-HPSPPComponentHtml | Set-Content 'components.html'
        Get components details in html format and save the results to file 'components.html'.

    .EXAMPLE
        Get-HPSPPComponent | Get-HPSPPComponentHtml -Full | Set-Content 'components_full.html'
        Get full components details in html format and save the results to file 'components_full.html'.

    .INPUTS
        HPSPPComponent[]

    .OUTPUTS
        String[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Component objects", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Component,

    [Parameter(Mandatory=$false, HelpMessage="Full component details")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $Full
    )

    BEGIN {
        # Check SPP path variable
        if (!($Script:HPSPPPath)) {
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check SPP revision history manifest file
        $Manifest = Join-Path $Script:HPSPPPath "hp_manifest\revision_history.xml"
        if ($Full -and !(Test-Path -PathType Leaf -LiteralPath $Manifest)) {
            Throw "SPP revision history manifest file '$Manifest' not found"
        }

        # Components to report
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
        if ($Components) {
            # Write output
            Write-Output "<html>"
            Write-Output "  <head>"
            Write-Output "    <style>"
            Write-Output "      a {color: teal; text-decoration: none;}"
            Write-Output "      a:hover {text-decoration: underline;}"
            Write-Output "      .dimmed {color: gray;}"
            Write-Output "      .caps {text-transform: capitalize;}"
            Write-Output "      .title {width: 98%; margin: 1em auto; border-bottom: 2px solid skyblue; padding-bottom: 0.5em; font-size: 95%; font-family: verdana;color: chocolate;}"
            Write-Output "      .titlevalue {width: 100%;}"
            Write-Output "      .properties {width: 98%; margin: 1em auto 2em; padding-bottom: 0.5em; font-size: 75%; font-family: verdana;}"
            Write-Output "      .propname {width: 20%; vertical-align: top; padding-bottom: 1em;}"
            Write-Output "      .propvalue {width: 80%; vertical-align: top; padding-bottom: 1em;}"
            Write-Output "      .note p, .note ul, .note ol {margin: 0 !important; padding: 0 !important;}"
            Write-Output "      .note li {list-style-position: inside !important;}"
            Write-Output "      .note li ul, .note li ol {margin-left: 1em !important;}"
            Write-Output "      .note blockquote {margin: 0 !important; padding: 0 !important;}"
            Write-Output "      .note table {margin: 1em 0 !important; padding: 0 !important; font-style: normal !important; font-size: 100% !important; font-family: verdana !important; border-collapse: collapse !important;}"
            Write-Output "      .note strong, .note b {font-weight: normal !important;}"
            Write-Output "      .note u {text-decoration: none !important;}"
            Write-Output "      .note br {display: inline !important; line-height: 0 !important;}"
            Write-Output "      .note em, .note i {font-style: normal !important;}"
            Write-Output "    </style>" 
            Write-Output "  </head>" 
            Write-Output "  <body>" 

            foreach ($component in $Components) {
            # Write name
            Write-Output "  <table class=`"title`">" 
            Write-Output "    <tr>"
            Write-Output "      <td class=`"titlevalue`">" 
            Write-Output "        $($component.Name)" 
            Write-Output "      </td>" 
            Write-Output "    </tr>" 
            Write-Output "  </table>"

            # Write version
            Write-Output "  <table class=`"properties`">"
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Version"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">" 
            Write-Output "        $($component.Version)"
            Write-Output "      </td>" 
            Write-Output "    </tr>" 

            # Write update recommendation
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Update"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue caps`">" 
            Write-Output "        $($component.UpgradeRequirement)"
            Write-Output "      </td>" 
            Write-Output "    </tr>" 

            # Write category
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Category"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">" 
            Write-Output "        $($component.Category)"
            Write-Output "      </td>" 
            Write-Output "    </tr>" 

            # Write service pack version
            if ($component.SPPVersion) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Service Pack Version"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue caps`">" 
            Write-Output "        $($component.SPPVersion)"
            Write-Output "      </td>" 
            Write-Output "    </tr>"
            }

            # Write description
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Description"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">" 
            Write-Output "        $($component.Description)"
            Write-Output "      </td>" 
            Write-Output "    </tr>" 

            # Write release details
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Release Date"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">" 
            Write-Output "        <p>"
            Write-Output "          $($component.ReleaseDate)<br/>"
            Write-Output "          <span class=`"dimmed`">Type: $($component.TypeID)</span><br/>"
            if ($component.Revision) {
            Write-Output "          <span class=`"dimmed`">Revision: $($component.Revision)</span><br/>"
            }
            if ($component.BuildNumber) {
            Write-Output "          <span class=`"dimmed`">Build Number: $($component.BuildNumber)</span><br/>"
            }
            if ($component.Manufacturer) {
            Write-Output "          <span class=`"dimmed`">Manufacturer: $($component.Manufacturer)</span><br/>"
            }
            if ($component.Different) {
            Write-Output "          <span class=`"dimmed`">State: </span><span class=`"dimmed caps`">$($component.Different)</span><br/>"
            }
            Write-Output "        </p>"
            Write-Output "      </td>"
            Write-Output "    </tr>"

            # Write operating systems
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Operating Systems"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            Write-Output "        <p>"
            foreach ($os in ($component.OperatingSystems | Sort-Object)) {
            Write-Output "          $os<br/>"
            }
            Write-Output "        </p>"
            Write-Output "      </td>" 
            Write-Output "    </tr>" 

            # Write files
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Files"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">" 
            foreach ($file in $component.Files) {
            Write-Output "        <p>"
            Write-Output "          $($file.Name)<br/>"
            Write-Output "          <span class=`"dimmed`">Size: $($file.Size)</span><br/>"
            Write-Output "          <span class=`"dimmed`">Date: $($file.DateModified)</span><br/>"
            Write-Output "          <span class=`"dimmed`">Download: </span><a href=`"$($file.FtpUrl)`">Ftp</a> <a href=`"$($file.FileUrl)`">Local</a><br/>"
            Write-Output "          <span class=`"dimmed`">Md5sum: $($file.Md5Sum)</span><br/>"
            Write-Output "        </p>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"

            # Check full details requested
            if ($Full) {
            # Write prerequisite notes
            $PrerequisiteNotes = @()
            $component.Node | Select-Xml -XPath "prerequisite_notes/prerequisite_notes_xlate[@lang='en']/prerequisite_notes_xlate_part" | ForEach-Object {
                $note = $_
                $PrerequisiteNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
            }
            if ($PrerequisiteNotes) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Prerequisites"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            foreach ($note in $PrerequisiteNotes) {
            Write-Output "        <div class=`"note`">"
            Write-Output "          $note"
            Write-Output "        </div>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"
            }

            # Write installation notes
            $InstallationNotes = @()
            $component.Node | Select-Xml -XPath "installation_notes/installation_notes_xlate[@lang='en']/installation_notes_xlate_part" | ForEach-Object {
                $note = $_
                $InstallationNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
            }
            if ($InstallationNotes) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Installation"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            foreach ($note in $InstallationNotes) {
            Write-Output "        <div class=`"note`">"
            Write-Output "          $note"
            Write-Output "        </div>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"
            }

            # Write availability notes
            $AvailabilityNotes = @()
            $component.Node | Select-Xml -XPath "availability_notes/availability_notes_xlate[@lang='en']/availability_notes_xlate_part" | ForEach-Object {
                $note = $_
                $AvailabilityNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
            }
            if ($AvailabilityNotes) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Availability"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            foreach ($note in $AvailabilityNotes) {
            Write-Output "        <div class=`"note`">"
            Write-Output "          $note"
            Write-Output "        </div>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"
            }

            # Write documentation notes
            $DocumentationNotes = @()
            $component.Node | Select-Xml -XPath "documentation_notes/documentation_notes_xlate[@lang='en']/documentation_notes_xlate_part" | ForEach-Object {
                $note = $_
                $DocumentationNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
            }
            if ($DocumentationNotes) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Documentation"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            foreach ($note in $DocumentationNotes) {
            Write-Output "        <div class=`"note`">"
            Write-Output "          $note"
            Write-Output "        </div>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"
            }

            # Write revision history
            $Revisions = @(Select-Xml -XPath "hp_manifest/revision_history/product_version/id[@product='$($component.ProductID)'][@version='$($component.VersionID)']/../revision_history/revision" -Path $Manifest)
            foreach ($revision in $Revisions) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Revision History"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">" 
            Write-Output "        $($revision.Node.version.value)<br/>"
            if ($revision.Node.version.revision) {
            Write-Output "          <span class=`"dimmed`">Revision: $($revision.Node.version.revision)</span><br/>"
            }
            switch ($revision.Node.version.type_of_change) {
            0 {
            Write-Output "          <span class=`"dimmed`">Update: Optional</span><br/>"
            }
            1 {
            Write-Output "          <span class=`"dimmed`">Update: Recommended</span><br/>"
            }
            2 {
            Write-Output "          <span class=`"dimmed`">Update: Critical</span><br/>"
            }
            }
            Write-Output "      </td>" 
            Write-Output "    </tr>"

            # Write enhancements
            $EnhancementNotes = @()
            $revision.Node | Select-Xml -XPath "revision_enhancements_xlate[@lang='en']/revision_enhancements_xlate_part" | ForEach-Object {
                $note = $_
                $EnhancementNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
            }
            if ($EnhancementNotes) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Enhancements"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            foreach ($note in $EnhancementNotes) {
            Write-Output "        <div class=`"note`">"
            Write-Output "          $note"
            Write-Output "        </div>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"
            }

            # Write fixes
            $FixNotes = @()
            $revision.Node | Select-Xml -XPath "revision_fixes_xlate[@lang='en']/revision_fixes_xlate_part" | ForEach-Object {
                $note = $_
                $FixNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
            }
            if ($FixNotes) {
            Write-Output "    <tr>"
            Write-Output "      <td class=`"propname`">"
            Write-Output "        Fixes"
            Write-Output "      </td>"
            Write-Output "      <td class=`"propvalue`">"
            foreach ($note in $FixNotes) {
            Write-Output "        <div class=`"note`">"
            Write-Output "          $note"
            Write-Output "        </div>"
            }
            Write-Output "      </td>"
            Write-Output "    </tr>"
            }

            }
            }
            Write-Output "  </table>"
            }

            # Write html closing
            Write-Output "  </body>"
            Write-Output "</html>"
        }
    }
}

Function Copy-HPSPPComponent {
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
            Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
        }

        # Check SPP folder path
        if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
            Throw "SPP folder '$Script:HPSPPPath' not found"
        }

        # Check destination folder path
        if (!(Test-Path -PathType Container -LiteralPath $DestinationFolder)) {
            Throw "Destination folder '$DestinationFolder' not found"
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

Function Get-HPSPPBundle {
<#
    .SYNOPSIS
        Get SPP bundle object.

    .DESCRIPTION
        The Get-HPSPPBundle command gets the SPP bundle object.

    .EXAMPLE
        Get-HPSPPBundle
        Get SPP bundle object.

    .INPUTS
        None

    .OUTPUTS
        HPSPPBundle
#>

    [CmdletBinding()]
    Param ()

    # Check SPP path variable
    if (!($Script:HPSPPPath)) {
        Throw "SPP folder path not defined (use Set-HPSPPFolderPath to define it)"
    }

    # Check SPP folder path
    if (!(Test-Path -PathType Container -LiteralPath $Script:HPSPPPath)) {
        Throw "SPP folder '$Script:HPSPPPath' not found"
    }

    # Check SPP bundle path
    $BundlePath = Join-Path $Script:HPSPPPath "hp\swpackages"
    if (!(Test-Path -PathType Container -LiteralPath $BundlePath)) {
        Throw "SPP bundle folder '$BundlePath' not found"
    }

    # Check SPP bundle file
    $BundleFiles = Join-Path $BundlePath "bp*.xml"
    if (!(Test-Path -PathType Leaf -Path $BundleFiles)) {
        Throw "SPP bundle file not found"
    }
    else {
        $BundleFile = Resolve-Path -Path $BundleFiles
        if ($BundleFile -is [Array]) {
            Throw "SPP bundle folder is not valid"
        }
    }

    # Create SPP bundle object
    $Node = Select-Xml -XPath "/cpq_bundle" -Path $BundleFile
    $Bundle = New-Object PSObject -Property ([ordered]@{
        Name = $Node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
        Description = $Node.Node.SelectSingleNode("description/description_xlate[@lang='en']").InnerText
        ProductID = $Node.Node.id.product
        VersionID = $Node.Node.id.version
        Category = $Node.Node.SelectSingleNode("category/category_xlate[@lang='en']").InnerText
        Version = $Node.Node.version.value
        Revision = $Node.Node.version.revision
        TypeOfChange = $Node.Node.version.type_of_change
        ReleaseDate = "$($Node.Node.release_date.year)-$($Node.Node.release_date.month)-$($Node.Node.release_date.day) $($Node.Node.release_date.hour):$($Node.Node.release_date.minute):$($Node.Node.release_date.second)"

        Node = $Node.Node
        Divisions = @()
        OperatingSystems = @()
        Packages = @()
    })

    $Bundle.Node | Select-Xml -XPath "divisions/division" | ForEach-Object {
        $division = $_
        $Bundle.Divisions += $division.Node.SelectSingleNode("division_xlate[@lang='en']").InnerText
    }

    $Bundle.Node | Select-Xml -XPath "operating_systems/operating_system" | ForEach-Object {
        $operatingsystem = $_
        $Bundle.OperatingSystems += $operatingsystem.Node.operating_system_xlate
    }

    $Bundle.Node | Select-Xml -XPath "contents/package" | ForEach-Object {
        $package = $_
        $Bundle.Packages += $package.Node.InnerText
    }

    $Bundle.PSObject.TypeNames.Insert(0,'HPSPPBundle')

    # Write output
    Write-Output $Bundle
}

function Get-HPSPPBundleHtml {
<#
    .SYNOPSIS
        Get SPP bundle details in html format.

    .DESCRIPTION
        The Get-HPSPPBundleHtml command gets SPP bundle details in html format.

    .PARAMETER Bundle
        Bundle objects.

    .PARAMETER Full
        Get full details (includes notes on prerequisites, installation, availability, and documentation as well as revision history and packages).

    .EXAMPLE
        Get-HPSPPBundle | Get-HPSPPBundleHtml | Set-Content 'bundle.html'
        Get bundle details in html format and save the results to file 'bundle.html'.

    .EXAMPLE
        Get-HPSPPBundle | Get-HPSPPBundleHtml -Full | Set-Content 'bundle_full.html'
        Get full bundle details in html format and save the results to file 'bundle_full.html'.

    .INPUTS
        HPSPPBundle

    .OUTPUTS
        String[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Bundle object", Position=0, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Full bundle details")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $Full
    )

    if ($Bundle.PSObject.TypeNames -contains "HPSPPBundle") {
        # Write output
        Write-Output "<html>"
        Write-Output "  <head>"
        Write-Output "    <style>"
        Write-Output "      a {color: teal; text-decoration: none;}"
        Write-Output "      a:hover {text-decoration: underline;}"
        Write-Output "      .dimmed {color: gray;}"
        Write-Output "      .caps {text-transform: capitalize;}"
        Write-Output "      .title {width: 98%; margin: 1em auto; border-bottom: 2px solid skyblue; padding-bottom: 0.5em; font-size: 95%; font-family: verdana;color: chocolate;}"
        Write-Output "      .titlevalue {width: 100%;}"
        Write-Output "      .properties {width: 98%; margin: 1em auto 2em; padding-bottom: 0.5em; font-size: 75%; font-family: verdana;}"
        Write-Output "      .propname {width: 20%; vertical-align: top; padding-bottom: 1em;}"
        Write-Output "      .propvalue {width: 80%; vertical-align: top; padding-bottom: 1em;}"
        Write-Output "      .note p, .note ul, .note ol {margin: 0 !important; padding: 0 !important;}"
        Write-Output "      .note li {list-style-position: inside !important;}"
        Write-Output "      .note li ul, .note li ol {margin-left: 1em !important;}"
        Write-Output "      .note blockquote {margin: 0 !important; padding: 0 !important;}"
        Write-Output "      .note table {margin: 1em 0 !important; padding: 0 !important; font-style: normal !important; font-size: 100% !important; font-family: verdana !important; border-collapse: collapse !important;}"
        Write-Output "      .note strong, .note b {font-weight: normal !important;}"
        Write-Output "      .note u {text-decoration: none !important;}"
        Write-Output "      .note br {display: inline !important; line-height: 0 !important;}"
        Write-Output "      .note em, .note i {font-style: normal !important;}"
        Write-Output "    </style>" 
        Write-Output "  </head>" 
        Write-Output "  <body>" 

        # Write name
        Write-Output "  <table class=`"title`">" 
        Write-Output "    <tr>"
        Write-Output "      <td class=`"titlevalue`">" 
        Write-Output "        $($Bundle.Name)" 
        Write-Output "      </td>" 
        Write-Output "    </tr>" 
        Write-Output "  </table>"

        # Write version
        Write-Output "  <table class=`"properties`">"
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Version"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">" 
        Write-Output "        $($Bundle.Version)"
        Write-Output "      </td>" 
        Write-Output "    </tr>" 

        # Write revision
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Revision"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue caps`">" 
        Write-Output "        $($Bundle.Revision)"
        Write-Output "      </td>" 
        Write-Output "    </tr>" 

        # Write category
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Category"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">" 
        Write-Output "        $($Bundle.Category)"
        Write-Output "      </td>" 
        Write-Output "    </tr>" 

        # Write description
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Description"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">" 
        Write-Output "        $($Bundle.Description)"
        Write-Output "      </td>" 
        Write-Output "    </tr>" 

        # Write release date
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Release Date"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">" 
        Write-Output "        $($Bundle.ReleaseDate)"
        Write-Output "      </td>" 
        Write-Output "    </tr>" 

        # Write operating systems
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Operating Systems"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        Write-Output "        <p>"
        foreach ($os in ($Bundle.OperatingSystems | Sort-Object)) {
        Write-Output "          $os<br/>"
        }
        Write-Output "        </p>"
        Write-Output "      </td>" 
        Write-Output "    </tr>" 

        # Check full details requested
        if ($Full) {

        # Write prerequisite notes
        $PrerequisiteNotes = @()
        $Bundle.Node | Select-Xml -XPath "prerequisite_notes/prerequisite_notes_xlate[@lang='en']/prerequisite_notes_xlate_part" | ForEach-Object {
            $note = $_
            $PrerequisiteNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
        }
        if ($PrerequisiteNotes) {
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Prerequisites"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        foreach ($note in $PrerequisiteNotes) {
        Write-Output "        <div class=`"note`">"
        Write-Output "          $note"
        Write-Output "        </div>"
        }
        Write-Output "      </td>"
        Write-Output "    </tr>"
        }

        # Write installation notes
        $InstallationNotes = @()
        $Bundle.Node | Select-Xml -XPath "installation_notes/installation_notes_xlate[@lang='en']/installation_notes_xlate_part" | ForEach-Object {
            $note = $_
            $InstallationNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
        }
        if ($InstallationNotes) {
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Installation"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        foreach ($note in $InstallationNotes) {
        Write-Output "        <div class=`"note`">"
        Write-Output "          $note"
        Write-Output "        </div>"
        }
        Write-Output "      </td>"
        Write-Output "    </tr>"
        }

        # Write availability notes
        $AvailabilityNotes = @()
        $Bundle.Node | Select-Xml -XPath "availability_notes/availability_notes_xlate[@lang='en']/availability_notes_xlate_part" | ForEach-Object {
            $note = $_
            $AvailabilityNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
        }
        if ($AvailabilityNotes) {
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Availability"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        foreach ($note in $AvailabilityNotes) {
        Write-Output "        <div class=`"note`">"
        Write-Output "          $note"
        Write-Output "        </div>"
        }
        Write-Output "      </td>"
        Write-Output "    </tr>"
        }

        # Write documentation notes
        $DocumentationNotes = @()
        $Bundle.Node | Select-Xml -XPath "documentation_notes/documentation_notes_xlate[@lang='en']/documentation_notes_xlate_part" | ForEach-Object {
            $note = $_
            $DocumentationNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
        }
        if ($DocumentationNotes) {
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Documentation"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        foreach ($note in $DocumentationNotes) {
        Write-Output "        <div class=`"note`">"
        Write-Output "          $note"
        Write-Output "        </div>"
        }
        Write-Output "      </td>"
        Write-Output "    </tr>"
        }

        # Write packages
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Packages"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        Write-Output "        <p>"
        foreach ($package in ($Bundle.Packages | Sort-Object)) {
        Write-Output "          $package<br/>"
        }
        Write-Output "        </p>"
        Write-Output "      </td>" 
        Write-Output "    </tr>"

        # Write revision history
        $Bundle.Node | Select-Xml -XPath "revision_history/revision" | ForEach-Object {
        $revision = $_
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Revision History"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">" 
        Write-Output "        $($revision.Node.version.value)<br/>"
        if ($revision.Node.version.revision) {
        Write-Output "          <span class=`"dimmed`">Revision: $($revision.Node.version.revision)</span><br/>"
        }
        Write-Output "      </td>" 
        Write-Output "    </tr>"

        # Write enhancements
        $EnhancementNotes = @()
        $revision.Node | Select-Xml -XPath "revision_enhancements_xlate[@lang='en']/revision_enhancements_xlate_part" | ForEach-Object {
            $note = $_
            $EnhancementNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
        }
        if ($EnhancementNotes) {
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Enhancements"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        foreach ($note in $EnhancementNotes) {
        Write-Output "        <div class=`"note`">"
        Write-Output "          $note"
        Write-Output "        </div>"
        }
        Write-Output "      </td>"
        Write-Output "    </tr>"
        }

        # Write fixes
        $FixNotes = @()
        $revision.Node | Select-Xml -XPath "revision_fixes_xlate[@lang='en']/revision_fixes_xlate_part" | ForEach-Object {
            $note = $_
            $FixNotes += $note.Node.InnerText.Replace("&nbsp;"," ")
        }
        if ($FixNotes) {
        Write-Output "    <tr>"
        Write-Output "      <td class=`"propname`">"
        Write-Output "        Fixes"
        Write-Output "      </td>"
        Write-Output "      <td class=`"propvalue`">"
        foreach ($note in $FixNotes) {
        Write-Output "        <div class=`"note`">"
        Write-Output "          $note"
        Write-Output "        </div>"
        }
        Write-Output "      </td>"
        Write-Output "    </tr>"
        }
        }
        }
        Write-Output "  </table>"
        
        # Write html closing
        Write-Output "  </body>"
        Write-Output "</html>"
    }
}
