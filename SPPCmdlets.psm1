Function Add-SPPBundle {
<#
    .SYNOPSIS
        Add SPP bundle.

    .DESCRIPTION
        The Add-SPPBundle command adds a Service Pack for ProLiant bundle.

    .PARAMETER Directory
        Bundle directory.

    .PARAMETER File
        Bundle file.

    .PARAMETER Manifest
        Bundle manifest directory.

    .PARAMETER Packages
        Bundle packages directory.

    .EXAMPLE
        Add-SPPBundle E:\
        Add SPP bundle contained in the specified directory. Use default locations for the bundle file, manifest and packages directories.

    .EXAMPLE
        Add-SPPBundle -File E:\packages\bp003135.xml -Manifest E:\manifest -Packages E:\packages
        Add SPP bundle using the specified file. Use the specified locations for the manifest and packages directories.

    .INPUTS
        None

    .OUTPUTS
        None
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Bundle directory", Position=0, ParameterSetName="Directory")]
    [ValidateNotNullOrEmpty()]
    [String]
    $Directory,

    [Parameter(Mandatory=$true, HelpMessage="Bundle file", ParameterSetName="File")]
    [ValidateNotNullOrEmpty()]
    [String]
    $File,

    [Parameter(Mandatory=$true, HelpMessage="Bundle manifest folder", ParameterSetName="File")]
    [ValidateNotNullOrEmpty()]
    [String]
    $Manifest,

    [Parameter(Mandatory=$true, HelpMessage="Bundle packages folder", ParameterSetName="File")]
    [ValidateNotNullOrEmpty()]
    [String]
    $Packages
    )

    # Set bundles
    if (!$Script:Bundles) {
        $Script:Bundles = @{}
    }

    # Check bundle directory
    if ($Directory) {
        if (Test-Path -PathType Leaf -Path (Join-Path $Directory "packages\bp*.xml")) {
            # Set bundle file
            $BundleFile = (Resolve-Path -Path (Join-Path $Directory "packages\bp*.xml") | Select-Object -First 1).Path.ToLower()

            # Check bundle added
            if ($Script:Bundles.ContainsKey($BundleFile)) {
                Throw "Bundle '$($BundleFile)' already added"
            }

            # Set manifest directory
            $Manifest = Join-Path $Directory "manifest"

            # Check manifest directory
            if (Test-Path -PathType Container -Path $Manifest) {
                # Set manifest directory full path
                $ManifestDirectory = (Resolve-Path -Path $Manifest).Path
            } else {
                Throw "Manifest directory '$Manifest' not found"
            }

            # Set packages directory
            $Packages = Join-Path $Directory "packages"

            # Check packages directory
            if (Test-Path -PathType Container -Path $Packages) {
                # Set packages directory full path
                $PackagesDirectory = (Resolve-Path -Path $Packages).Path
            } else {
                Throw "Packages directory '$Packages' not found"
            }
        } elseif (Test-Path -PathType Leaf -Path (Join-Path $Directory "hp\swpackages\bp*.xml")) {
            # Set bundle file
            $BundleFile = (Resolve-Path -Path (Join-Path $Directory "hp\swpackages\bp*.xml") | Select-Object -First 1).Path.ToLower()

            # Check bundle added
            if ($Script:Bundles.ContainsKey($BundleFile)) {
                Throw "Bundle '$($BundleFile)' already added"
            }

            # Set manifest directory
            $Manifest = Join-Path $Directory "hp_manifest"

            # Check manifest directory
            if (Test-Path -PathType Container -Path $Manifest) {
                # Set manifest directory full path
                $ManifestDirectory = (Resolve-Path -Path $Manifest).Path
            } else {
                Throw "Manifest directory '$Manifest' not found"
            }

            # Set packages directory
            $Packages = Join-Path $Directory "hp\swpackages"

            # Check packages directory
            if (Test-Path -PathType Container -Path $Packages) {
                # Set packages directory full path
                $PackagesDirectory = (Resolve-Path -Path $Packages).Path
            } else {
                Throw "Packages directory '$Packages' not found"
            }
        } else {
            Throw "Bundle directory '$Directory' not valid"
        }
    }

    # Check bundle file
    if ($File) {
        # Check bundle file
        if (Test-Path -PathType Leaf -Path $File) {
            # Set full bundle file path
            $BundleFile = (Resolve-Path -Path $File).Path.ToLower()
        } else {
            Throw "Bundle file '$File' not found"
        }

        # Check bundle added
        if ($Script:Bundles.ContainsKey($BundleFile)) {
            Throw "Bundle '$($BundleFile)' already added"
        }

        # Check manifest directory
        if (Test-Path -PathType Container -Path $Manifest) {
            # Set full manifest directory path
            $ManifestDirectory = (Resolve-Path -Path $Manifest).Path
        } else {
            Throw "Manifest directory '$Manifest' not found"
        }

        # Check packages directory
        if (Test-Path -PathType Container -Path $Packages) {
            # Set full packages directory path
            $PackagesDirectory = (Resolve-Path -Path $Packages).Path
        } else {
            Throw "Packages directory '$Packages' not found"
        }
    }

    # Check system file
    $SystemFile = Join-Path $ManifestDirectory "system.xml"
    if (!(Test-Path -PathType Leaf -Path $SystemFile)) {
        Throw "Manifest file '$SystemFile' not found"
    }

    # Check operating system file
    $OperatingSystemFile = Join-Path $ManifestDirectory "os.xml"
    if (!(Test-Path -PathType Leaf -Path $OperatingSystemFile)) {
        Throw "Manifest file '$OperatingSystemFile' not found"
    }

    # Check category file
    $CategoryFile = Join-Path $ManifestDirectory "category.xml"
    if (!(Test-Path -PathType Leaf -Path $CategoryFile)) {
        Throw "Manifest file '$CategoryFile' not found"
    }

    # Check device file
    $DeviceFile = Join-Path $ManifestDirectory "device.xml"
    if (!(Test-Path -PathType Leaf -Path $DeviceFile)) {
        Throw "Manifest file '$DeviceFile' not found"
    }

    # Check type file
    $TypeFile = Join-Path $ManifestDirectory "type.xml"
    if (!(Test-Path -PathType Leaf -Path $TypeFile)) {
        Throw "Manifest file '$TypeFile' not found"
    }

    # Check component file
    $ComponentFile = Join-Path $ManifestDirectory "meta.xml"
    if (!(Test-Path -PathType Leaf -Path $ComponentFile)) {
        Throw "Manifest file '$ComponentFile' not found"
    }

    # Check revision history file
    $RevisionHistoryFile = Join-Path $ManifestDirectory "revision_history.xml"
    if (!(Test-Path -PathType Leaf -Path $RevisionHistoryFile)) {
        Throw "Manifest file '$RevisionHistoryFile' not found"
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading bundle" -PercentComplete 0

    # Read bundle xml
    $BundleXML = Select-Xml -XPath "/cpq_bundle" -Path $BundleFile

    # Create bundle object
    $Bundle = New-Object PSObject -Property ([ordered]@{
        File = $BundleFile
        Name = $BundleXML.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
        Description = $BundleXML.Node.SelectSingleNode("description/description_xlate[@lang='en']").InnerText
        Tag = $BundleXML.Node.tag
        ProductID = $BundleXML.Node.id.product
        VersionID = $BundleXML.Node.id.version
        Category = $BundleXML.Node.SelectSingleNode("category/category_xlate[@lang='en']").InnerText
        Version = $BundleXML.Node.version.value
        Revision = $BundleXML.Node.version.revision
        FullVersion = $BundleXML.Node.version.value + $BundleXML.Node.version.revision
        TypeOfChange = @{'0'='Optional';'1'='Recommended';'2'='Critical'}[$BundleXML.Node.version.type_of_change]
        ReleaseYear = $BundleXML.Node.release_date.year
        ReleaseMonth = $BundleXML.Node.release_date.month
        ReleaseDay = $BundleXML.Node.release_date.day
        ReleaseHour = $BundleXML.Node.release_date.hour
        ReleaseMinute = $BundleXML.Node.release_date.minute
        ReleaseSecond = $BundleXML.Node.release_date.second
        Divisions = @()
        OperatingSystems = @()
        PrerequisiteNotes = @()
        InstallationNotes = @()
        AvailabilityNotes = @()
        DocumentationNotes = @()
        RevisionHistory = @()
        Contents = @()
        Components = @()
    })

    # Set bundle object type
    $Bundle.PSObject.TypeNames.Insert(0,'SPPBundle')

    # Add divisions
    foreach ($node in Select-Xml -XPath "divisions/division/division_xlate[@lang='en']" -Xml $BundleXML.Node) {
        $Bundle.Divisions += $node.Node.InnerText
    }

    # Add operating systems
    foreach ($node in Select-Xml -XPath "operating_systems/operating_system/operating_system_xlate" -Xml $BundleXML.Node) {
        $Bundle.OperatingSystems += $node.Node.InnerText
    }

    # Add prerequisite notes
    foreach ($node in Select-Xml -XPath "prerequisite_notes/prerequisite_notes_xlate[@lang='en']/prerequisite_notes_xlate_part" -Xml $BundleXML.Node) {
        $Bundle.PrerequisiteNotes += $node.Node.InnerText.Replace("&nbsp;", " ")
    }

    # Add installation notes
    foreach ($node in Select-Xml -XPath "installation_notes/installation_notes_xlate[@lang='en']/installation_notes_xlate_part" -Xml $BundleXML.Node) {
        $Bundle.InstallationNotes += $node.Node.InnerText.Replace("&nbsp;", " ")
    }

    # Add availability notes
    foreach ($node in Select-Xml -XPath "availability_notes/availability_notes_xlate[@lang='en']/availability_notes_xlate_part" -Xml $BundleXML.Node) {
        $Bundle.AvailabilityNotes += $node.Node.InnerText.Replace("&nbsp;", " ")
    }

    # Add documentation notes
    foreach ($node in Select-Xml -XPath "documentation_notes/documentation_notes_xlate[@lang='en']/documentation_notes_xlate_part" -Xml $BundleXML.Node) {
        $Bundle.DocumentationNotes += $node.Node.InnerText.Replace("&nbsp;", " ")
    }

    # Add revision history
    foreach ($node in Select-Xml -XPath "revision_history/revision/version[@value]/parent::revision" -Xml $BundleXML.Node) {
        # Create revision
        $Revision = New-Object PSObject -Property ([ordered]@{
            Version = $node.Node.version.value
            Revision = $node.Node.version.revision
            FullVersion = $node.Node.version.value + $node.Node.version.revision
            TypeOfChange = @{'0'='Optional';'1'='Recommended';'2'='Critical'}[$node.Node.version.type_of_change]
            Enhancements = @()
            Fixes = @()
        })

        # Add enhancements
        foreach ($enhancement in Select-Xml -XPath "revision_enhancements_xlate[@lang='en']/revision_enhancements_xlate_part" -Xml $node.Node) {
            if ($enhancement.Node.InnerText) {
                $Revision.Enhancements += $enhancement.Node.InnerText.Replace("&nbsp;", " ")
            }
        }

        # Add fixes
        foreach ($fix in Select-Xml -XPath "revision_fixes_xlate[@lang='en']/revision_fixes_xlate_part" -Xml $node.Node) {
            if ($fix.Node.InnerText) {
                $Revision.fixes += $fix.Node.InnerText.Replace("&nbsp;", " ")
            }
        }

        # Add revision
        $Bundle.RevisionHistory += $Revision
    }

    # Add contents
    foreach ($node in Select-Xml -XPath "contents/package" -Xml $BundleXML.Node) {
        $Bundle.Contents += $node.Node.InnerText
    }

    # Add bundle
    $Script:Bundles.Add($Bundle.File, $Bundle)

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading systems" -PercentComplete 10

    # Set systems
    if (!$Script:Systems) {
        $Script:Systems = @{}
    }

    # Read system xml
    $SystemXML = Select-Xml -XPath "//systems/system" -Path $SystemFile

    # Add systems
    foreach ($node in $SystemXML) {
        # Read system key
        $SystemKey = $node.Node.id

        # Check system key
        if ($Script:Systems.ContainsKey($SystemKey)) {
            # Select system
            $System = $Script:Systems[$SystemKey]
        } else {
            # Create system object
            $System = New-Object PSObject -Property ([ordered]@{
                Key = $node.Node.id
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                SystemID = $node.Node.systemid
                FirmwareID = $node.Node.firmwareid
                Components = @{}
            })

            # Set system object type
            $System.PSObject.TypeNames.Insert(0,'SPPSystem')
            $System.PSObject.TypeNames.Insert(1,'SPPFilter')

            # Add system
            $Script:Systems.Add($System.Key, $System)
        }

        # Add bundle
        $System.Components.Add($BundleFile, @{})

        # Add components
        foreach ($component in Select-Xml -XPath "product_version" -Xml $node.Node) {
            # Read component key
            $ComponentKey = $component.Node.id.product + "-" + $component.Node.id.version

            # Add component
            if (!$System.Components[$BundleFile].ContainsKey($ComponentKey)) { $System.Components[$BundleFile].Add($ComponentKey, $null) }
        }
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading operating systems" -PercentComplete 20

    # Set operating systems
    if (!$Script:OperatingSystems) {
        $Script:OperatingSystems = @{}
    }

    # Read operating system xml
    $OperatingSystemXML = Select-Xml -XPath "//operating_systems/operating_system" -Path $OperatingSystemFile

    # Add operating systems
    foreach ($node in $OperatingSystemXML) {
        # Read operating system key
        $OperatingSystemKey = $node.Node.id

        # Check operating system key
        if ($Script:OperatingSystems.ContainsKey($OperatingSystemKey)) {
            # Select operating system
            $OperatingSystem = $Script:OperatingSystems[$OperatingSystemKey]
        } else {
            # Create operating system object
            $OperatingSystem = New-Object PSObject -Property ([ordered]@{
                Key = $node.Node.id
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                Platform = $node.Node.supported_operating_system.platform
                Major = $node.Node.supported_operating_system.major
                Minor = $node.Node.supported_operating_system.minor
                Components = @{}
            })

            # Set operating system object type
            $OperatingSystem.PSObject.TypeNames.Insert(0,'SPPOperatingSystem')
            $OperatingSystem.PSObject.TypeNames.Insert(1,'SPPFilter')

            # Add operating system
            $Script:OperatingSystems.Add($OperatingSystem.Key, $OperatingSystem)
        }

        # Add bundle
        $OperatingSystem.Components.Add($BundleFile, @{})

        # Add components
        foreach ($component in Select-Xml -XPath "product_version" -Xml $node.Node) {
            # Read component key
            $ComponentKey = $component.Node.id.product + "-" + $component.Node.id.version

            # Add component
            if (!$OperatingSystem.Components[$BundleFile].ContainsKey($ComponentKey)) { $OperatingSystem.Components[$BundleFile].Add($ComponentKey, $null) }
        }
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading categories" -PercentComplete 30

    # Set categories
    if (!$Script:Categories) {
        $Script:Categories = @{}
    }

    # Read category xml
    $CategoryXML = Select-Xml -XPath "//categories/category" -Path $CategoryFile

    # Add categories
    foreach ($node in $CategoryXML) {
        # Read category key
        $CategoryKey = $node.Node.id

        # Check category key
        if ($Script:Categories.ContainsKey($CategoryKey)) {
            # Select category
            $Category = $Script:Categories[$CategoryKey]
        } else {
            # Create category object
            $Category = New-Object PSObject -Property ([ordered]@{
                Key = $node.Node.id
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                Components = @{}
            })

            # Set category object type
            $Category.PSObject.TypeNames.Insert(0,'SPPCategory')
            $Category.PSObject.TypeNames.Insert(1,'SPPFilter')

            # Add category
            $Script:Categories.Add($Category.Key, $Category)
        }

        # Add bundle
        $Category.Components.Add($BundleFile, @{})

        # Add components
        foreach ($component in Select-Xml -XPath "product_version" -Xml $node.Node) {
            # Read component key
            $ComponentKey = $component.Node.id.product + "-" + $component.Node.id.version

            # Add component
            if (!$Category.Components[$BundleFile].ContainsKey($ComponentKey)) { $Category.Components[$BundleFile].Add($ComponentKey, $null) }
        }
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading devices" -PercentComplete 40

    # Set devices
    if (!$Script:Devices) {
        $Script:Devices = @{}
    }

    # Read device xml
    $DeviceXML = Select-Xml -XPath "//devices/device" -Path $DeviceFile

    # Add devices
    foreach ($node in $DeviceXML) {
        # Read device key
        $DeviceKey = $node.Node.id

        # Check device key
        if ($Script:Devices.ContainsKey($DeviceKey)) {
            # Select device
            $Device = $Script:Devices[$DeviceKey]
        } else {
            # Create device object
            $Device = New-Object PSObject -Property ([ordered]@{
                Key = $node.Node.id
                Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText
                DeviceType = $node.Node.device_type
                DeviceIDString = $node.Node.deviceidstring
                Components = @{}
            })

            # Set device object type
            $Device.PSObject.TypeNames.Insert(0,'SPPDevice')
            $Device.PSObject.TypeNames.Insert(1,'SPPFilter')

            # Add device
            $Script:Devices.Add($Device.Key, $Device)
        }

        # Add bundle
        $Device.Components.Add($BundleFile, @{})

        # Add components
        foreach ($component in Select-Xml -XPath "product_version" -Xml $node.Node) {
            # Read component key
            $ComponentKey = $component.Node.id.product + "-" + $component.Node.id.version

            # Add component
            if (!$Device.Components[$BundleFile].ContainsKey($ComponentKey)) { $Device.Components[$BundleFile].Add($ComponentKey, $null) }
        }
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading types" -PercentComplete 50

    # Set types
    if (!$Script:Types) {
        $Script:Types = @{}
    }

    # Read type xml
    $TypeXML = Select-Xml -XPath "//types/type" -Path $TypeFile

    # Add types
    foreach ($node in $TypeXML) {
        # Read type key
        $TypeKey = $node.Node.id

        # Check type key
        if ($Script:Types.ContainsKey($TypeKey)) {
            # Select type
            $Type = $Script:Types[$TypeKey]
        } else {
            # Create type object
            $Type = New-Object PSObject -Property ([ordered]@{
                Key = $node.Node.id
                Name = $node.Node.id
                Components = @{}
            })

            # Set type object type
            $Type.PSObject.TypeNames.Insert(0,'SPPType')
            $Type.PSObject.TypeNames.Insert(1,'SPPFilter')

            # Add type
            $Script:Types.Add($Type.Key, $Type)
        }

        # Add bundle
        $Type.Components.Add($BundleFile, @{})

        # Add components
        foreach ($component in Select-Xml -XPath "product_version" -Xml $node.Node) {
            # Read component key
            $ComponentKey = $component.Node.id.product + "-" + $component.Node.id.version

            # Add component
            if (!$Type.Components[$BundleFile].ContainsKey($ComponentKey)) { $Type.Components[$BundleFile].Add($ComponentKey, $null) }
        }
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading revision history" -PercentComplete 60

    # Set revision history
    $RevisionHistory = @{}

    # Read revision history xml
    $RevisionHistoryXML = Select-Xml -XPath "//revision_history/product_version" -Path $RevisionHistoryFile

    # Add components
    foreach ($component in $RevisionHistoryXML) {
        # Read component key
        $ComponentKey = $component.Node.id.product + "-" + $component.Node.id.version

        # Set revisions
        $Revisions = @()

        # Add revisions
        foreach ($node in Select-Xml -XPath "revision_history/revision/version[@value]/.." -Xml $component.Node) {
            # Create revision
            $Revision = New-Object PSObject -Property ([ordered]@{
                Version = $node.Node.version.value
                Revision = $node.Node.version.revision
                FullVersion = $node.Node.version.value + $node.Node.version.revision
                TypeOfChange = @{'0'='Optional';'1'='Recommended';'2'='Critical'}[$node.Node.version.type_of_change]
                Enhancements = @()
                Fixes = @()
            })

            # Add enhancements
            foreach ($enhancement in Select-Xml -XPath "revision_enhancements_xlate[@lang='en']/revision_enhancements_xlate_part" -Xml $node.Node) {
                if ($enhancement.Node.InnerText) {
                    $Revision.Enhancements += $enhancement.Node.InnerText.Replace("&nbsp;", " ")
                }
            }

            # Add fixes
            foreach ($fix in Select-Xml -XPath "revision_fixes_xlate[@lang='en']/revision_fixes_xlate_part" -Xml $node.Node) {
                if ($fix.Node.InnerText) {
                    $Revision.fixes += $fix.Node.InnerText.Replace("&nbsp;", " ")
                }
            }

            # Add revision
            $Revisions += $Revision
        }

        # Add revision history
        if (!$RevisionHistory.ContainsKey($ComponentKey)) { $RevisionHistory.Add($ComponentKey, $Revisions) }
    }

    # Display progress
    Write-Progress -Activity "Adding bundle $($BundleFile) ..." -Status "Reading components" -PercentComplete 80

    # Read component xml
    $ComponentXML = Select-Xml -XPath "//product/product_version" -Path $ComponentFile

    # Add components
    foreach ($node in $ComponentXML) {
        # Read component key
        $ComponentKey = $node.Node.id.product + "-" + $node.Node.id.version

        # Create component object
        $Component = New-Object PSObject -Property ([ordered]@{
            Key = $ComponentKey
            Bundle = $Bundle.Version
            BundleFile = $BundleFile
            Name = $node.Node.SelectSingleNode("name/name_xlate[@lang='en']").InnerText.Trim()
            ProductID = $node.Node.id.product
            VersionID = $node.Node.id.version
            Description = $node.Node.SelectSingleNode("description/description_xlate[@lang='en']").InnerText
            AltName = $node.Node.SelectSingleNode("alt_name/alt_name_xlate[@lang='en']").InnerText
            FileName = $node.Node.filename
            Version = $node.Node.version.value
            Revision = $node.Node.version.revision
            FullVersion = $node.Node.version.value + $node.Node.version.revision
            UpgradeRequirement = $node.Node.upgrade_requirement.value
            Different = $node.Node.different
            BuildNumber = $node.Node.build.number
            Type = $node.Node.type.id
            Category = $node.Node.SelectSingleNode("category/category_xlate[@lang='en']").InnerText
            Manufacturer = $node.Node.SelectSingleNode("manufacturer_name/manufacturer_name_xlate[@lang='en']").InnerText
            TypeOfChange = @{'0'='Optional';'1'='Recommended';'2'='Critical'}[$node.Node.version.type_of_change]
            ReleaseYear = $node.Node.release_date.year
            ReleaseMonth = $node.Node.release_date.month
            ReleaseDay = $node.Node.release_date.day
            ReleaseHour = $node.Node.release_date.hour
            ReleaseMinute = $node.Node.release_date.minute
            ReleaseSecond = $node.Node.release_date.second
            RequiredDiskSpaceKB = $node.Node.prerequisites.required_diskspace.size_kb
            OperatingSystems = @()
            Divisions = @()
            Files = @()
            PrerequisiteNotes = @()
            InstallationNotes = @()
            AvailabilityNotes = @()
            DocumentationNotes = @()
            RevisionHistory = @()
        })

        # Set component object type
        $Component.PSObject.TypeNames.Insert(0,'SPPComponent')

        # Add operating systems
        foreach ($operatingsystem in Select-Xml -XPath "operating_systems/operating_system/operating_system_xlate" -Xml $node.Node) {
            $Component.OperatingSystems += $operatingsystem.Node.InnerText
        }

        # Add divisions
        foreach ($division in Select-Xml -XPath "divisions/division/division_xlate[@lang='en']" -Xml $node.Node) {
            $Component.Divisions += $division.Node.InnerText
        }

        # Add files
        foreach ($item in Select-Xml -XPath "files/file" -Xml $node.Node) {
            $Component.Files += New-Object PSObject -Property ([ordered]@{
                Name = $item.Node.name
                FileUrl = Join-Path $PackagesDirectory $item.Node.name
                Size = $item.Node.size
                DateModified = $item.Node.date_modified
                Md5Sum = $item.Node.md5sum
                Sha1Sum = $item.Node.sha1sum
            })
        }

        # Add prerequisite notes
        foreach ($item in Select-Xml -XPath "prerequisite_notes/prerequisite_notes_xlate[@lang='en']/prerequisite_notes_xlate_part" -Xml $node.Node) {
            $Component.PrerequisiteNotes += $item.Node.InnerText.Replace("&nbsp;", " ")
        }

        # Add installation notes
        foreach ($item in Select-Xml -XPath "installation_notes/installation_notes_xlate[@lang='en']/installation_notes_xlate_part" -Xml $node.Node) {
            $Component.InstallationNotes += $item.Node.InnerText.Replace("&nbsp;", " ")
        }

        # Add availability notes
        foreach ($item in Select-Xml -XPath "availability_notes/availability_notes_xlate[@lang='en']/availability_notes_xlate_part" -Xml $node.Node) {
            $Component.AvailabilityNotes += $item.Node.InnerText.Replace("&nbsp;", " ")
        }

        # Add documentation notes
        foreach ($item in Select-Xml -XPath "documentation_notes/documentation_notes_xlate[@lang='en']/documentation_notes_xlate_part" -Xml $node.Node) {
            $Component.DocumentationNotes += $item.Node.InnerText.Replace("&nbsp;", " ")
        }

        # Add revision history
        if ($RevisionHistory.ContainsKey($ComponentKey)) {
            $Component.RevisionHistory = $RevisionHistory[$ComponentKey]
        }

        # Add component
        $Bundle.Components += $Component
    }
}

Function Get-SPPBundle {
<#
    .SYNOPSIS
        Get SPP bundles.

    .DESCRIPTION
        The Get-SPPBundle command gets Service Pack for ProLiant bundles.

    .PARAMETER File
        Bundle files. Wildcards are accepted (e.g. *bp*). If omitted the command will get all available bundles.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml
        Get SPP bundle identified by the specified file.

    .EXAMPLE
        Get-SPPBundle
        Get all available SPP bundles.

    .INPUTS
        None

    .OUTPUTS
        SPPBundle[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Bundle files", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $File
    )

    # Set bundles by file
    $BundlesByFile = @{}

    # Check files
    if ($File) {
        # Add bundles by file
        foreach ($value in $File) {
            foreach ($bundle in $Script:Bundles.Values) {
                if (!$BundlesByFile.ContainsKey($bundle.File)) {
                    if ($bundle.File -like $value) {
                        $BundlesByFile.Add($bundle.File, $bundle)
                    }
                }
            }
        }
    } else {
        # Add all bundles
        $BundlesByFile = $Script:Bundles
    }

    # Output bundles
    Write-Output $BundlesByFile.Values | Sort-Object -Property FullVersion -Descending
}

Function Remove-SPPBundle {
<#
    .SYNOPSIS
        Remove SPP bundles.

    .DESCRIPTION
        The Remove-SPPBundle command removes Service Pack for ProLiant bundles.

    .PARAMETER File
        Bundle files. Wildcards are accepted (e.g. *bp*).

    .EXAMPLE
        Remove-SPPBundle E:\packages\bp003135.xml
        Remove SPP bundle identified by the specified file.

    .INPUTS
        None

    .OUTPUTS
        None
#>

    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='Low')]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Bundle file", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $File
    )

    # Confirm removal
    if (!$PSCmdlet.ShouldProcess($File)) {
        return
    }

    # Set bundles by file
    $BundlesByFile = @{}

    # Add bundles by file
    foreach ($value in $File) {
        foreach ($bundle in $Script:Bundles.Values) {
            if (!$BundlesByFile.ContainsKey($bundle.File)) {
                if ($bundle.File -like $value) {
                    $BundlesByFile.Add($bundle.File, $bundle)
                }
            }
        }
    }

    # Remove selected bundles
    foreach ($key in $BundlesByFile.Keys) {
        # Remove bundle
        $Script:Bundles.Remove($key)

        # Set systems to remove
        $SystemsToRemove = @()

        # Remove system components
        foreach ($system in $Script:Systems.Values) {
            # Check bundle
            if ($system.Components.ContainsKey($key)) {
                # Remove bundle
                $system.Components.Remove($key)

                # Check system
                if ($system.Components.Count -eq 0) {
                    $SystemsToRemove += $system
                }
            }
        }

        # Remove systems
        foreach ($system in $SystemsToRemove) {
            $Script:Systems.Remove($system.Key)
        }

        # Set operating systems to remove
        $OperatingSystemsToRemove = @()

        # Remove operating system components
        foreach ($operatingsystem in $Script:OperatingSystems.Values) {
            # Check bundle
            if ($operatingsystem.Components.ContainsKey($key)) {
                # Remove bundle
                $operatingsystem.Components.Remove($key)

                # Check operating system
                if ($operatingsystem.Components.Count -eq 0) {
                    $OperatingSystemsToRemove += $operatingsystem
                }
            }
        }

        # Remove operating systems
        foreach ($operatingsystem in $OperatingSystemsToRemove) {
            $Script:OperatingSystems.Remove($operatingsystem.Key)
        }

        # Set categories to remove
        $CategoriesToRemove = @()

        # Remove category components
        foreach ($category in $Script:Categories.Values) {
            # Check bundle
            if ($category.Components.ContainsKey($key)) {
                # Remove bundle
                $category.Components.Remove($key)

                # Check category
                if ($category.Components.Count -eq 0) {
                    $CategoriesToRemove += $category
                }
            }
        }

        # Remove categories
        foreach ($category in $CategoriesToRemove) {
            $Script:Categories.Remove($category.Key)
        }

        # Set devices to remove
        $DevicesToRemove = @()

        # Remove device components
        foreach ($device in $Script:Devices.Values) {
            # Check bundle
            if ($device.Components.ContainsKey($key)) {
                # Remove bundle
                $device.Components.Remove($key)

                # Check device
                if ($device.Components.Count -eq 0) {
                    $DevicesToRemove += $device
                }
            }
        }

        # Remove devices
        foreach ($device in $DevicesToRemove) {
            $Script:Devices.Remove($device.Key)
        }

        # Set types to remove
        $TypesToRemove = @()

        # Remove type components
        foreach ($type in $Script:Types.Values) {
            # Check bundle
            if ($type.Components.ContainsKey($key)) {
                # Remove bundle
                $type.Components.Remove($key)

                # Check type
                if ($type.Components.Count -eq 0) {
                    $TypesToRemove += $type
                }
            }
        }

        # Remove types
        foreach ($type in $TypesToRemove) {
            $Script:Types.Remove($type.Key)
        }
    }
}

Function ConvertTo-SPPBundleHtml {
<#
    .SYNOPSIS
        Convert SPP bundles to html.

    .DESCRIPTION
        The ConvertTo-SPPBundleHtml command converts Support Pack for ProLiant bundle objects to html format.

    .PARAMETER Bundle
        Bundle objects.

    .PARAMETER File
        Output html file. If omitted the output will be sent to console.

    .PARAMETER Details
        Include bundle details. This adds bundle notes, revision history, and contents to the output.

    .EXAMPLE
        Get-SPPBundle | ConvertTo-SPPBundleHtml 'bundle.html'
        Convert bundle objects to html format and save the output to the file 'bundle.html'.

    .EXAMPLE
        Get-SPPBundle | ConvertTo-SPPBundleHtml -Details | Set-Content 'bundle.html'
        Convert bundle objects, including details, to html format and send the output to console.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        String[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Bundle objects", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Output html file", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $File,

    [Parameter(Mandatory=$false, HelpMessage="Include bundle details")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $Details
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    END {
        # Check html file
        if ($File) {
            # Check file exists
            $FileName = Resolve-Path $File -ErrorAction SilentlyContinue -ErrorVariable ResolvePath
            if ($FileName) {
                if($Host.UI.PromptForChoice("Overwrite File (ConvertTo-SPPBundleHtml)" , "File '$FileName' exists. Do you want to overwriet it?" , @("&Yes", "&No"), 1) -eq 1) {
                    return
                }
                $HtmlFile = $FileName.Path
            } else {
                $HtmlFile = $ResolvePath[0].TargetObject
            }

            # Create stream writer
            $StreamWriter = New-Object System.IO.StreamWriter $HtmlFile
        } else {
            # Create memory stream
            $HtmlStream = New-Object System.IO.MemoryStream

            # Create stream writer
            $StreamWriter =  New-Object System.IO.StreamWriter $HtmlStream
        }

        # Check bundles
        if ($Bundles) {
            # Write output
            $StreamWriter.WriteLine("<html>")
            $StreamWriter.WriteLine("  <head>")
            $StreamWriter.WriteLine("    <style>")
            $StreamWriter.WriteLine("      a {color: teal; text-decoration: none;}")
            $StreamWriter.WriteLine("      a:hover {text-decoration: underline;}")
            $StreamWriter.WriteLine("      .dimmed {color: gray;}")
            $StreamWriter.WriteLine("      .spaced:first-child {display: block; margin-top: 0; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .spaced {display: block; margin-top: 1.5em; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .caps {text-transform: capitalize;}")
            $StreamWriter.WriteLine("      .title {width: 98%; margin: 1em auto; border-bottom: 2px solid lightseagreen; padding-bottom: 0.5em; font-size: 95%; font-family: verdana;color: chocolate;}")
            $StreamWriter.WriteLine("      .titlevalue {width: 100%;}")
            $StreamWriter.WriteLine("      .properties {width: 98%; margin: 1em auto 2em; padding-bottom: 0.5em; font-size: 75%; font-family: verdana;}")
            $StreamWriter.WriteLine("      .propname {width: 15%; vertical-align: top; padding-bottom: 1.5em;}")
            $StreamWriter.WriteLine("      .propvalue {width: 85%; vertical-align: top; padding-bottom: 1.5em;}")
            $StreamWriter.WriteLine("      .note p, .note ul, .note ol {margin: 0 !important; padding: 0 !important;}")
            $StreamWriter.WriteLine("      .note li {list-style-position: inside !important;}")
            $StreamWriter.WriteLine("      .note li ul, .note li ol {margin-left: 1em !important;}")
            $StreamWriter.WriteLine("      .note li p:first-child {display: inline !important;}")
            $StreamWriter.WriteLine("      .note blockquote {margin: 0 !important; padding: 0 !important;}")
            $StreamWriter.WriteLine("      .note table {margin: 1em 0 !important; padding: 0 !important; font-style: normal !important; font-size: 100% !important; font-family: verdana !important; border-collapse: collapse !important;}")
            $StreamWriter.WriteLine("      .note strong, .note b {font-weight: normal !important;}")
            $StreamWriter.WriteLine("      .note u {text-decoration: none !important;}")
            $StreamWriter.WriteLine("      .note br {display: inline !important; line-height: 0 !important;}")
            $StreamWriter.WriteLine("      .note em, .note i {font-style: normal !important;}")
            $StreamWriter.WriteLine("    </style>")
            $StreamWriter.WriteLine("  </head>")
            $StreamWriter.WriteLine("  <body>")

            foreach ($bundle in $Bundles) {
            # Write name
            $StreamWriter.WriteLine("  <table class=`"title`">")
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"titlevalue`">")
            $StreamWriter.WriteLine("        $($bundle.Name)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            $StreamWriter.WriteLine("  </table>")

            # Write version
            $StreamWriter.WriteLine("  <table class=`"properties`">")
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Version")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($bundle.Version)</span>")
            if ($bundle.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($bundle.Revision)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write file
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        File")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        $($bundle.File)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write tag
            if ($bundle.Tag) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Tag")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <span class=`"caps`">$($bundle.Tag)</span>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write category
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Category")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        $($bundle.Category)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write description
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Description")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        $($bundle.Description)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write release date
            if ($bundle.ReleaseYear -and $bundle.ReleaseMonth -and $bundle.ReleaseDay) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Release Date")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        $($bundle.ReleaseYear)-$($bundle.ReleaseMonth)-$($bundle.ReleaseDay)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write divisions
            if ($bundle.Divisions) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Divisions")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <p>")
            foreach ($division in $bundle.Divisions) {
            $StreamWriter.WriteLine("          $division<br/>")
            }
            $StreamWriter.WriteLine("        </p>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write operating systems
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Operating Systems")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <p>")
            foreach ($operatingsystem in $bundle.OperatingSystems | Sort-Object) {
            $StreamWriter.WriteLine("          $operatingsystem<br/>")
            }
            $StreamWriter.WriteLine("        </p>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            if ($Details) {
            # Write notes
            if ($bundle.PrerequisiteNotes -or $bundle.InstallationNotes -or $bundle.AvailabilityNotes -or $bundle.DocumentationNotes) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Notes")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")

            # Write prerequisite notes
            if ($bundle.PrerequisiteNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Prerequisites:</span>")
            foreach ($note in $bundle.PrerequisiteNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write installation notes
            if ($bundle.InstallationNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Installation:</span>")
            foreach ($note in $bundle.InstallationNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write availability notes
            if ($bundle.AvailabilityNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Availability:</span>")
            foreach ($note in $bundle.AvailabilityNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write documentation notes
            if ($bundle.DocumentationNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Documentation:</span>")
            foreach ($note in $bundle.DocumentationNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write revision history
            if ($bundle.RevisionHistory) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Revision History")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")

            # Write revisions
            foreach ($revision in $bundle.RevisionHistory | Sort-Object -Property FullVersion -Descending) {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($revision.Version)</span>")
            if ($revision.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($revision.Revision)</span>")
            }

            # Write enhancements
            if ($revision.Enhancements) {
            $StreamWriter.WriteLine("        <span class=`"spacedin`">Enhancements:</span>")
            foreach ($enhancement in $revision.Enhancements) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $enhancement")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write fixes
            if ($revision.Fixes) {
            $StreamWriter.WriteLine("        <span class=`"spacedin`">Fixes:</span>")
            foreach ($fix in $revision.Fixes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $fix")
            $StreamWriter.WriteLine("        </div>")
            }
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write contents
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Contents")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <p>")
            foreach ($package in $bundle.Contents | Sort-Object) {
            $StreamWriter.WriteLine("          $package<br/>")
            }
            $StreamWriter.WriteLine("        </p>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            $StreamWriter.WriteLine("  </table>")
            }

            # Write html closing
            $StreamWriter.WriteLine("  </body>")
            $StreamWriter.WriteLine("</html>")

            # Flush stream writer
            $StreamWriter.Flush()

            if (!$File) {
                # Create stream reader
                $StreamReader = New-Object System.IO.StreamReader $HtmlStream

                # Write output
                $HtmlStream.Position = 0
                while ($null -ne ($line = $StreamReader.ReadLine())) {
                    Write-Output $line
                }

                # Close stream reader
                $StreamReader.Close()
            }

            # Close stream writer
            $StreamWriter.Close()
        }
    }
}

Function ConvertTo-SPPBundleCsv {
<#
    .SYNOPSIS
        Convert SPP bundles to csv.

    .DESCRIPTION
        The ConvertTo-SPPBundleCsv command converts Support Pack for ProLiant bundle objects to csv format.

    .PARAMETER Bundle
        Bundle objects.

    .PARAMETER File
        Output csv file. If omitted the output will be sent to console.

    .EXAMPLE
        Get-SPPBundle | ConvertTo-SPPBundleCsv -File 'bundle.csv'
        Convert bundle objects to csv format and save the output to the file 'bundle.csv'.

    .EXAMPLE
        Get-SPPBundle | ConvertTo-SPPBundleCsv | Set-Content 'bundle.csv'
        Convert bundle objects to csv format and send the output to console.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        String[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Bundle objects", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Output csv file", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $File
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    END {
        # Check csv file
        if ($File) {
            # Check file exists
            $FileName = Resolve-Path $File -ErrorAction SilentlyContinue -ErrorVariable ResolvePath
            if ($FileName) {
                if($Host.UI.PromptForChoice("Overwrite File (ConvertTo-SPPBundleCsv)" , "File '$FileName' exists. Do you want to overwriet it?" , @("&Yes", "&No"), 1) -eq 1) {
                    return
                }
                $CsvFile = $FileName.Path
            } else {
                $CsvFile = $ResolvePath[0].TargetObject
            }

            # Create stream writer
            $StreamWriter = New-Object System.IO.StreamWriter $CsvFile
        } else {
            # Create memory stream
            $CsvStream = New-Object System.IO.MemoryStream

            # Create stream writer
            $StreamWriter =  New-Object System.IO.StreamWriter $CsvStream
        }

        # Check bundles
        if ($Bundles) {
            # Write header
            $StreamWriter.WriteLine('Name, Version, Revision, File, Tag, Category, Description, Release Date, Division, Operating System')

            # Add lines
            foreach ($bundle in $Bundles) {
                # Write name
                $line = '"' + $bundle.Name + '"'

                # Write version
                $line += ',"' + $bundle.Version + '"'

                # Write revision
                $line += ',"' + $bundle.Revision + '"'

                # Write file
                $line += ',"' + $bundle.File + '"'

                # Write tag
                $line += ',"' + $bundle.Tag + '"'

                # Write category
                $line += ',"' + $bundle.Category + '"'

                # Write description
                $line += ',"' + $bundle.Description + '"'

                # Write release date
                if ($bundle.ReleaseYear -and $bundle.ReleaseMonth -and $bundle.ReleaseDay) {
                    $line += ',"' + $bundle.ReleaseYear + '-' + $bundle.ReleaseMonth + '-' + $bundle.ReleaseDay + '"'
                } else {
                    $line += ',""'
                }

                # Set division lines
                $divisionlines = @()

                # Write divisions
                foreach ($division in $bundle.Divisions) {
                    $divisionlines += $line + ',"' + $division + '"'
                }

                # Set operating system lines
                $operatingsystemlines = @()

                # Write operating systems
                foreach ($line in $divisionlines) {
                    foreach ($operatingsystem in $bundle.OperatingSystems) {
                        $operatingsystemlines += $line + ',"' + $operatingsystem + '"'
                    }
                }

                # Write operating system lines
                foreach ($line in $operatingsystemlines) {
                    $StreamWriter.WriteLine($line)
                }
            }

            # Flush stream writer
            $StreamWriter.Flush()

            if (!$File) {
                # Create stream reader
                $StreamReader = New-Object System.IO.StreamReader $CsvStream

                # Write output
                $CsvStream.Position = 0
                while ($null -ne ($line = $StreamReader.ReadLine())) {
                    Write-Output $line
                }

                # Close stream reader
                $StreamReader.Close()
            }

            # Close stream writer
            $StreamWriter.Close()
        }
    }
}

Function Get-SPPSystem {
<#
    .SYNOPSIS
        Get SPP systems.

    .DESCRIPTION
        The Get-SPPSystem command gets Service Pack for ProLiant systems.

    .PARAMETER Name
        System names. Wildcards are accepted (e.g. *DL360*). If omitted the command will get all systems.

    .PARAMETER Bundle
        System bundles. If omitted the command will get systems from all bundles.

    .PARAMETER PassThru
        Pass through filter objects. This outputs filter objects received from the input pipeline.

    .EXAMPLE
        Get-SPPSystem *DL380*
        Get all systems that contain 'DL360' in their names.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml | Get-SPPSystem -Name *G5*,*G6*
        Get all systems, from the specified bundle, that contain 'G5' or 'G6' in their names.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        SPPFilter[], SPPSystem[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="System names", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name,

    [Parameter(Mandatory=$false, HelpMessage="System bundles", Position=1, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Pass through filter objects")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $PassThru
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Set filters
        $Filters = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            } elseif ($PassThru -and ($object.PSObject.TypeNames -contains "SPPFilter")) {
                $Filters += $object
            }
        }
    }

    END {
        # Set systems by bundle
        $SystemsByBundle = @{}

        # Check bundles
        if ($Bundles.Count -ne 0) {
            # Add systems by bundle
            foreach ($bundle in $Bundles) {
                foreach ($system in $Script:Systems.Values) {
                    if ($system.Components.ContainsKey($bundle.File)) {
                        $SystemsByBundle.Add($system.Key, $system)
                    }
                }
            }
        } else {
            # Add all systems
            $SystemsByBundle = $Script:Systems
        }

        # Set systems by name
        $SystemsByName = @{}

        # Check names
        if ($Name) {
            # Add systems by name
            foreach ($value in $Name) {
                foreach ($system in $SystemsByBundle.Values) {
                    if (!$SystemsByName.ContainsKey($system.Key)) {
                        if ($system.Name -like $value) {
                            $SystemsByName.Add($system.Key, $system)
                        }
                    }
                }
            }
        } else {
            # Add all systems by bundle
            $SystemsByName = $SystemsByBundle
        }

        # Output filters
        Write-Output $Filters

        # Output systems
        Write-Output $SystemsByName.Values | Sort-Object -Property Name
    }
}

Function Get-SPPOperatingSystem {
<#
    .SYNOPSIS
        Get SPP operating systems.

    .DESCRIPTION
        The Get-SPPOperatingSystem command gets Service Pack for ProLiant operating systems.

    .PARAMETER Name
        Operating system names. Wildcards are accepted (e.g. *VMware*). If omitted the command will get all operating systems.

    .PARAMETER Bundle
        Operating system bundles. If omitted the command will get operating systems from all bundles.

    .PARAMETER PassThru
        Pass through filter objects. This outputs filter objects received from the input pipeline.

    .EXAMPLE
        Get-SPPOperatingSystem *VMware*
        Get all operating systems that contain 'VMware' in their names.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml | Get-SPPOperatingSystem -Name *Windows*,*Linux*
        Get all operating systems, from the specified bundle, that contain 'Windows' or 'Linux' in their names.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        SPPFilter[], SPPOperatingSystem[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Operating system names", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name,

    [Parameter(Mandatory=$false, HelpMessage="Operating system bundles", Position=1, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Pass through filter objects")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $PassThru
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Set filters
        $Filters = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            } elseif ($PassThru -and ($object.PSObject.TypeNames -contains "SPPFilter")) {
                $Filters += $object
            }
        }
    }

    END {
        # Set operating systems by bundle
        $OperatingSystemsByBundle = @{}

        # Check bundles
        if ($Bundles.Count -ne 0) {
            # Add operating systems by bundle
            foreach ($bundle in $Bundles) {
                foreach ($operatingsystem in $Script:OperatingSystems.Values) {
                    if ($operatingsystem.Components.ContainsKey($bundle.File)) {
                        $OperatingSystemsByBundle.Add($operatingsystem.Key, $operatingsystem)
                    }
                }
            }
        } else {
            # Add all oeprating systems
            $OperatingSystemsByBundle = $Script:OperatingSystems
        }

        # Set operating systems by name
        $OperatingSystemsByName = @{}

        # Check names
        if ($Name) {
            # Add opearating systems by name
            foreach ($value in $Name) {
                foreach ($operatingsystem in $OperatingSystemsByBundle.Values) {
                    if (!$OperatingSystemsByName.ContainsKey($operatingsystem.Key)) {
                        if ($operatingsystem.Name -like $value) {
                            $OperatingSystemsByName.Add($operatingsystem.Key, $operatingsystem)
                        }
                    }
                }
            }
        } else {
            # Add all operating systems by bundle
            $OperatingSystemsByName = $OperatingSystemsByBundle
        }

        # Output filters
        Write-Output $Filters

        # Output operating systems
        Write-Output $OperatingSystemsByName.Values | Sort-Object -Property Name
    }
}

Function Get-SPPCategory {
<#
    .SYNOPSIS
        Get SPP categories.

    .DESCRIPTION
        The Get-SPPCategory command gets Service Pack for ProLiant categories.

    .PARAMETER Name
        Category names. Wildcards are accepted (e.g. *Firmware*). If omitted the command will get all categories.

    .PARAMETER Bundle
        Category bundles. If omitted the command will get categories from all bundles.

    .PARAMETER PassThru
        Pass through filter objects. This outputs filter objects received from the input pipeline.

    .EXAMPLE
        Get-SPPCategory *Firmware*
        Get all categories that contain 'Firmware' in their names.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml | Get-SPPCategory -Name *Firmware*,*Driver*
        Get all categories, from the specified bundle, that contain 'Firmware' or 'Driver' in their names.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        SPPFilter, SPPCategory[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Category names", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name,

    [Parameter(Mandatory=$false, HelpMessage="Category bundles", Position=1, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Pass through filter objects")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $PassThru
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Set filters
        $Filters = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            } elseif ($PassThru -and ($object.PSObject.TypeNames -contains "SPPFilter")) {
                $Filters += $object
            }
        }
    }

    END {
        # Set categories by bundle
        $CategoriesByBundle = @{}

        # Check bundles
        if ($Bundles.Count -ne 0) {
            # Add categories by bundle
            foreach ($bundle in $Bundles) {
                foreach ($category in $Script:Categories.Values) {
                    if ($category.Components.ContainsKey($bundle.File)) {
                        $CategoriesByBundle.Add($category.Key, $category)
                    }
                }
            }
        } else {
            # Add all categories
            $CategoriesByBundle = $Script:Categories
        }

        # Set categories by name
        $CategoriesByName = @{}

        # Check names
        if ($Name) {
            # Add categories by name
            foreach ($value in $Name) {
                foreach ($category in $CategoriesByBundle.Values) {
                    if (!$CategoriesByName.ContainsKey($category.Key)) {
                        if ($category.Name -like $value) {
                            $CategoriesByName.Add($category.Key, $category)
                        }
                    }
                }
            }
        } else {
            # Add all categories by bundle
            $CategoriesByName = $CategoriesByBundle
        }

        # Output filters
        Write-Output $Filters

        # Output categories
        Write-Output $CategoriesByName.Values | Sort-Object -Property Name
    }
}

Function Get-SPPDevice {
<#
    .SYNOPSIS
        Get SPP devices.

    .DESCRIPTION
        The Get-SPPDevice command gets Service Pack for ProLiant devices.

    .PARAMETER Name
        Device names. Wildcards are accepted (e.g. *VMware*). If omitted the command will get all devices.

    .PARAMETER Bundle
        Device bundles. If omitted the command will get devices from all bundles.

    .PARAMETER PassThru
        Pass through filter objects. This outputs filter objects received from the input pipeline.

    .EXAMPLE
        Get-SPPDevice *HBA*
        Get all devices that contain 'HBA' in their names.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml | Get-SPPDevice -Name *HBA*,*NIC*
        Get all devices, from the specified bundle, that contain 'HBA' or 'NIC' in their names.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        SPPFilter[], SPPDevice[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Device names", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name,

    [Parameter(Mandatory=$false, HelpMessage="Device bundles", Position=1, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Pass through filter objects")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $PassThru
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Set filters
        $Filters = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            } elseif ($PassThru -and ($object.PSObject.TypeNames -contains "SPPFilter")) {
                $Filters += $object
            }
        }
    }

    END {
        # Set devices by bundle
        $DevicesByBundle = @{}

        # Check bundles
        if ($Bundles.Count -ne 0) {
            # Add devices by bundle
            foreach ($bundle in $Bundles) {
                foreach ($device in $Script:Devices.Values) {
                    if ($device.Components.ContainsKey($bundle.File)) {
                        $DevicesByBundle.Add($device.Key, $device)
                    }
                }
            }
        } else {
            # Add all devices
            $DevicesByBundle = $Script:Devices
        }

        # Set devices by name
        $DevicesByName = @{}

        # Check names
        if ($Name) {
            # Add devices by name
            foreach ($value in $Name) {
                foreach ($device in $DevicesByBundle.Values) {
                    if (!$DevicesByName.ContainsKey($device.Key)) {
                        if ($device.Name -like $value) {
                            $DevicesByName.Add($device.Key, $device)
                        }
                    }
                }
            }
        } else {
            # Add all devices by bundle
            $DevicesByName = $DevicesByBundle
        }

        # Output filters
        Write-Output $Filters

        # Output devices
        Write-Output $DevicesByName.Values | Sort-Object -Property Name
    }
}

Function Get-SPPType {
<#
    .SYNOPSIS
        Get SPP types.

    .DESCRIPTION
        The Get-SPPType command gets Service Pack for ProLiant types.

    .PARAMETER Name
        Type names. Wildcards are accepted (e.g. *VMware*). If omitted the command will get all types.

    .PARAMETER Bundle
        Type bundles. If omitted the command will get types from all bundles.

    .PARAMETER PassThru
        Pass through filter objects. This outputs filter objects received from the input pipeline.

    .EXAMPLE
        Get-SPPType *component*
        Get all types that contain 'component' in their names.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml | Get-SPPType -Name *component*,*software*
        Get all types, from the specified bundle, that contain 'component' or 'software' in their names.

    .INPUTS
        SPPBundle[]

    .OUTPUTS
        SPPFilter[], SPPType[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Type names", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name,

    [Parameter(Mandatory=$false, HelpMessage="Type bundles", Position=1, ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Pass through filter objects")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $PassThru
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Set filters
        $Filters = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            } elseif ($PassThru -and ($object.PSObject.TypeNames -contains "SPPFilter")) {
                $Filters += $object
            }
        }
    }

    END {
        # Set types by bundle
        $TypesByBundle = @{}

        # Check bundles
        if ($Bundles.Count -ne 0) {
            # Add types by bundle
            foreach ($bundle in $Bundles) {
                foreach ($type in $Script:Types.Values) {
                    if ($type.Components.ContainsKey($bundle.File)) {
                        $TypesByBundle.Add($type.Key, $type)
                    }
                }
            }
        } else {
            # Add all types
            $TypesByBundle = $Script:Types
        }

        # Set types by name
        $TypesByName = @{}

        # Check names
        if ($Name) {
            # Add types by name
            foreach ($value in $Name) {
                foreach ($type in $TypesByBundle.Values) {
                    if (!$TypesByName.ContainsKey($type.Key)) {
                        if ($type.Name -like $value) {
                            $TypesByName.Add($type.Key, $type)
                        }
                    }
                }
            }
        } else {
            # Add all types by bundle
            $TypesByName = $TypesByBundle
        }

        # Output filters
        Write-Output $Filters

        # Output types
        Write-Output $TypesByName.Values | Sort-Object -Property Name
    }
}

Function Get-SPPComponent {
<#
    .SYNOPSIS
        Get SPP components.

    .DESCRIPTION
        The Get-SPPComponent command gets Service Pack for ProLiant components.

    .PARAMETER Name
        Component names. Wildcards are accepted (e.g. *iLO*). If omitted the command will get all components.

    .PARAMETER Bundle
        Component bundles. If omitted the command will get components from all bundles.

    .PARAMETER Filter
        Component filter objects. Can be any of System, OperatingSystem, Category, Device, or Type objects used to narrow down, i.e. filter, components selection.

    .PARAMETER Versions
        Component versions (All, Changed, or Unchanged). 'Changed' will only get components that have different versions across the selected bundles. 'Unchanged' will only get components that have the same version across the selected bundles. All, the default, will get all components from the selected bundles.

    .EXAMPLE
        Get-SPPComponent *iLO*
        Get all components that contain 'iLO' in their names.

    .EXAMPLE
        Get-SPPBundle E:\packages\bp003135.xml | Get-SPPComponent -Name *iLO*,*SmartArray*
        Get all components, from the specified bundle, that contain 'iLO' or 'SmartArray' in their names.

    .EXAMPLE
        Get-SPPSystem *BL460c* | Get-SPPComponent
        Get all components for the specified systems.

    .EXAMPLE
        Get-SPPOperatingSystem *VMware* | Get-SPPComponent -Versions Changed
        Get only components that have different versions across bundles, for the specified operating systems.

    .INPUTS
        SPPBundle[], SPPSystem[], SPPOperatingSystem[], SPPCategory[], SPPDevice[], SPPType[]

    .OUTPUTS
        SPPCompoent[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$false, HelpMessage="Component names", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String[]]
    $Name,

    [Parameter(Mandatory=$false, HelpMessage="Component bundles", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Bundle,

    [Parameter(Mandatory=$false, HelpMessage="Component filter objects", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Filter,

    [Parameter(Mandatory=$false, HelpMessage="Component versions")]
    [ValidateSet("All", "Changed", "Unchanged", "Unique")]
    [ValidateNotNullOrEmpty()]
    [String]
    $Versions="All"
    )

    BEGIN {
        # Set bundles
        $Bundles = @()

        # Set systems
        $Systems = @()

        # Set operating systems
        $OperatingSystems = @()

        # Set categories
        $Categories = @()

        # Set devices
        $Devices = @()

        # Set types
        $Types = @()

        # Process bundles from variable
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }

        # Process filters from variable
        foreach ($object in $Filter) {
            if ($object.PSObject.TypeNames -contains "SPPSystem") {
                $Systems += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPOperatingSystem") {
                $OperatingSystems += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPCategory") {
                $Categories += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPDevice") {
                $Devices += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPType") {
                $Types += $object
            }
        }
    }

    PROCESS {
        # Process bundles from pipeline
        foreach ($object in $Bundle) {
            if ($object.PSObject.TypeNames -contains "SPPBundle") {
                $Bundles += $object
            }
        }

        # Process systems from pipeline
        foreach ($object in $Filter) {
            if ($object.PSObject.TypeNames -contains "SPPSystem") {
                $Systems += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPOperatingSystem") {
                $OperatingSystems += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPCategory") {
                $Categories += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPDevice") {
                $Devices += $object
            } elseif ($object.PSObject.TypeNames -contains "SPPType") {
                $Types += $object
            }
        }
    }

    END {
        # Set components
        $Components = @()

        # Set component groups
        $Groups = @{}

        # Check bundles
        if ($Bundles.Count -eq 0) {
            $Bundles = $Script:Bundles.Values
        }

        # Check components
        foreach ($bundle in $Bundles) {
            foreach ($component in $bundle.Components) {
                # Check systems
                if ($Systems.Count -ne 0 ) {
                    # Reset keep flag
                    $keep = $false

                    # Check component
                    foreach ($filter in $Systems) {
                        if ($filter.Components.ContainsKey($bundle.File)) {
                            $keep = $filter.Components[$bundle.File].ContainsKey($component.Key)
                            if ($keep) { break }
                        }
                    }

                    # Check keep flag
                    if (!$keep) { continue }
                }

                # Check operating systems
                if ($OperatingSystems.Count -ne 0 ) {
                    # Reset keep flag
                    $keep = $false

                    # Check component
                    foreach ($filter in $OperatingSystems) {
                        if ($filter.Components.ContainsKey($bundle.File)) {
                            $keep = $filter.Components[$bundle.File].ContainsKey($component.Key)
                            if ($keep) { break }
                        }
                    }

                    # Check keep flag
                    if (!$keep) { continue }
                }

                # Check categories
                if ($Categories.Count -ne 0 ) {
                    # Reset keep flag
                    $keep = $false

                    # Check component
                    foreach ($filter in $Categories) {
                        if ($filter.Components.ContainsKey($bundle.File)) {
                            $keep = $filter.Components[$bundle.File].ContainsKey($component.Key)
                            if ($keep) { break }
                        }
                    }

                    # Check keep flag
                    if (!$keep) { continue }
                }

                # Check devices
                if ($Devices.Count -ne 0 ) {
                    # Reset keep flag
                    $keep = $false

                    # Check component
                    foreach ($filter in $Devices) {
                        if ($filter.Components.ContainsKey($bundle.File)) {
                            $keep = $filter.Components[$bundle.File].ContainsKey($component.Key)
                            if ($keep) { break }
                        }
                    }

                    # Check keep flag
                    if (!$keep) { continue }
                }

                # Check types
                if ($Types.Count -ne 0 ) {
                    # Reset keep flag
                    $keep = $false

                    # Check component
                    foreach ($filter in $Types) {
                        if ($filter.Components.ContainsKey($bundle.File)) {
                            $keep = $filter.Components[$bundle.File].ContainsKey($component.Key)
                            if ($keep) { break }
                        }
                    }

                    # Check keep flag
                    if (!$keep) { continue }
                }

                # Check names
                if ($Name) {
                    # Reset keep flag
                    $keep = $false

                    # Check component
                    foreach ($value in $Name) {
                        $keep = ($component.Name -like $value)
                        if ($keep) { break }
                    }

                    # Check keep flag
                    if (!$keep) { continue }
                }

                # Check component group
                if (!$Groups.ContainsKey($component.ProductID)) {
                    # Add component group
                    $Groups.Add($component.ProductID, @())
                }

                # Add component
                $Groups[$component.ProductID] += $component
            }
        }

        # Check versions
        if ($Versions -eq "Unique") {
            # Check component groups
            foreach ($group in $Groups.Values) {
                # Check component count
                if ($group.count -eq 1) {
                    # Add unique components
                    $Components += $group
                }
            }
        } elseif ($Versions -eq "Changed") {
            # Check component groups
            foreach ($group in $Groups.Values) {
                # Check component count
                if ($group.count -eq 1) {
                    continue
                }

                # Set version keys
                $keys = @{}

                # Reset common flag
                $common = $false

                # Check components
                foreach ($component in $group) {
                    if ($keys.Count -eq 0) {
                        $keys.Add($component.FullVersion, $null)
                    } else {
                        $common = $keys.ContainsKey($component.FullVersion)
                        if ($common) {
                            break
                        } else {
                            $keys.Add($component.FullVersion, $null)
                        }
                    }
                }

                # Check common flag
                if (!$common) {
                    # Add changed components
                    $Components += $group
                } else {
                    continue
                }
            }
        } elseif ($Versions -eq "Unchanged") {
            # Check component groups
            foreach ($group in $Groups.Values) {
                # Check component count
                if ($group.count -eq 1) {
                    continue
                }

                # Set version keys
                $keys = @{}

                # Reset common flag
                $common = $false

                # Check components
                foreach ($component in $group) {
                    if ($keys.Count -eq 0) {
                        $keys.Add($component.FullVersion, $null)
                    } else {
                        $common = $keys.ContainsKey($component.FullVersion)
                        if (!$common) {
                            break
                        }
                    }
                }

                # Check common flag
                if ($common) {
                    # Add unchanged components
                    $Components += $group
                } else {
                    continue
                }
            }
        } else {
            # Check component groups
            foreach ($group in $Groups.Values) {
                # Add all components
                $Components += $group
            }
        }

        # Output selected components
        Write-Output $Components | Sort-Object -Property Name, FullVersion -Descending
    }
}

Function Copy-SPPComponent {
<#
    .SYNOPSIS
        Copy SPP component files.

    .DESCRIPTION
        The Copy-SPPComponent command copies Service Pack for ProLiant component files to a directory.

    .PARAMETER Component
        Component objects.

    .PARAMETER Directory
        Destination directory. The directory must exist. If a file to be copied already exists in the destination directory it will be overwritten.

    .EXAMPLE
        Get-SPPComponent | Copy-SPPComponent C:\Components
        Copy component files to destination directory 'C:\Components'.

    .INPUTS
        SPPComponent[]

    .OUTPUTS
        None
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Component objects", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Component,

    [Parameter(Mandatory=$true, HelpMessage="Destination directory", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $Directory
    )

    BEGIN {
        # Set components
        $Components = @()

        # Process components from variable
        foreach ($object in $Component) {
            if ($object.PSObject.TypeNames -contains "SPPComponent") {
                $Components += $object
            }
        }
    }

    PROCESS {
        # Process components from pipeline
        foreach ($object in $Component) {
            if ($object.PSObject.TypeNames -contains "SPPComponent") {
                $Components += $object
            }
        }
    }

    END {
        # Check destination directory
        if (Test-Path -PathType Container -Path $Directory) {
            # Set destination directory full path
            $DestinationDirectory = (Resolve-Path -Path $Directory).Path
        } else {
            Throw "Destination directory '$Directory' not found"
        }

        # Set processed count
        $processed = 0

        # Copy components
        foreach ($component in $Components) {
            # Display progress
            Write-Progress -Activity "Copying components to '$DestinationDirectory' ..." -Status "Copying '$($component.FileName)'" -PercentComplete ([int]($processed++ / $Components.Count * 100))

            # Copy component files
            foreach ($file in $component.Files) {
                # Set file item
                $item = Get-Item $file.FileUrl

                # Copy files
                Copy-Item -Path "$(Join-Path $item.Directory $item.BaseName)*" -Destination $DestinationDirectory -Force
            }
        }
    }
}

Function ConvertTo-SPPComponentHtml {
<#
    .SYNOPSIS
        Convert SPP components to html.

    .DESCRIPTION
        The ConvertTo-SPPComponentHtml command converts Support Pack for ProLiant component objects to html format.

    .PARAMETER Component
        Component objects.

    .PARAMETER File
        Output html file. If omitted the output will be sent to console.

    .PARAMETER Details
        Include component details. This adds component notes, revision history, and contents to the output.

    .PARAMETER Combine
        Combine component versions. This combines all versions of each component under one section.

    .EXAMPLE
        Get-SPPComponent | ConvertTo-SPPComponentHtml -File 'components.html'
        Convert component objects to html format and save the output to the file 'components.html'.

    .EXAMPLE
        Get-SPPComponent | ConvertTo-SPPComponentHtml -Details | Set-Content 'components.html'
        Convert component objects, including details, to html format and send the output to console.

    .INPUTS
        SPPComponent[]

    .OUTPUTS
        String[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Component objects", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Component,

    [Parameter(Mandatory=$false, HelpMessage="Output html file", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $File,

    [Parameter(Mandatory=$false, HelpMessage="Include component details")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $Details,

    [Parameter(Mandatory=$false, HelpMessage="Combine component versions")]
    [ValidateNotNullOrEmpty()]
    [Switch]
    $Combine
    )

    BEGIN {
        # Set components
        $Components = @()

        # Process components from variable
        foreach ($object in $Component) {
            if ($object.PSObject.TypeNames -contains "SPPComponent") {
                $Components += $object
            }
        }
    }

    PROCESS {
        # Process components from pipeline
        foreach ($object in $Component) {
            if ($object.PSObject.TypeNames -contains "SPPComponent") {
                $Components += $object
            }
        }
    }

    END {
        # Check html file
        if ($File) {
            # Check file exists
            $FileName = Resolve-Path $File -ErrorAction SilentlyContinue -ErrorVariable ResolvePath
            if ($FileName) {
                if($Host.UI.PromptForChoice("Overwrite File (ConvertTo-SPPComponentHtml)" , "File '$FileName' exists. Do you want to overwriet it?" , @("&Yes", "&No"), 1) -eq 1) {
                    return
                }
                $HtmlFile = $FileName.Path
            } else {
                $HtmlFile = $ResolvePath[0].TargetObject
            }

            # Create stream writer
            $StreamWriter = New-Object System.IO.StreamWriter $HtmlFile
        } else {
            # Create memory stream
            $HtmlStream = New-Object System.IO.MemoryStream

            # Create stream writer
            $StreamWriter =  New-Object System.IO.StreamWriter $HtmlStream
        }

        # Check components
        if ($Components -and $Combine) {
            # Set component groups
            $Groups = @{}

            # Set component revisions
            $RevisionHistory = @{}

            # Add components
            foreach ($component in $Components) {
                # Check component group
                if (!$Groups.ContainsKey($component.ProductID)) {
                    # Add component group
                    $Groups.Add($component.ProductID, @())

                    # Add revision history
                    $RevisionHistory.Add($component.ProductID, @())
                }

                # Add component
                $Groups[$component.ProductID] += $component

                # Add revisions
                $RevisionHistory[$component.ProductID] += $component.RevisionHistory
            }

            # Write output
            $StreamWriter.WriteLine("<html>")
            $StreamWriter.WriteLine("  <head>")
            $StreamWriter.WriteLine("    <style>")
            $StreamWriter.WriteLine("      a {color: teal; text-decoration: none;}")
            $StreamWriter.WriteLine("      a:hover {text-decoration: underline;}")
            $StreamWriter.WriteLine("      .dimmed {color: gray;}")
            $StreamWriter.WriteLine("      .spaced:first-child {display: block; margin-top: 0; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .spaced {display: block; margin-top: 1.5em; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .spacedin {display: block; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .caps {text-transform: capitalize;}")
            $StreamWriter.WriteLine("      .title {width: 98%; margin: 1em auto; border-bottom: 2px solid lightseagreen; padding-bottom: 0.5em; font-size: 95%; font-family: verdana;color: chocolate;}")
            $StreamWriter.WriteLine("      .titlevalue {width: 100%;}")
            $StreamWriter.WriteLine("      .properties {width: 98%; margin: 1em auto 2em; padding-bottom: 0.5em; font-size: 75%; font-family: verdana;}")
            $StreamWriter.WriteLine("      .propname {width: 15%; vertical-align: top; padding-bottom: 2em;}")
            $StreamWriter.WriteLine("      .propvalue {width: 85%; vertical-align: top; padding-bottom: 2em;}")
            $StreamWriter.WriteLine("      .note p, .note ul, .note ol {margin: 0 !important; padding: 0 !important;}")
            $StreamWriter.WriteLine("      .note li {list-style-position: inside !important;}")
            $StreamWriter.WriteLine("      .note li ul, .note li ol {margin-left: 1em !important;}")
            $StreamWriter.WriteLine("      .note li p:first-child {display: inline !important;}")
            $StreamWriter.WriteLine("      .note blockquote {margin: 0 !important; padding: 0 !important;}")
            $StreamWriter.WriteLine("      .note table {margin: 1em 0 !important; padding: 0 !important; font-style: normal !important; font-size: 100% !important; font-family: verdana !important; border-collapse: collapse !important;}")
            $StreamWriter.WriteLine("      .note strong, .note b {font-weight: normal !important;}")
            $StreamWriter.WriteLine("      .note u {text-decoration: none !important;}")
            $StreamWriter.WriteLine("      .note br {display: inline !important; line-height: 0 !important;}")
            $StreamWriter.WriteLine("      .note em, .note i {font-style: normal !important;}")
            $StreamWriter.WriteLine("    </style>")
            $StreamWriter.WriteLine("  </head>")
            $StreamWriter.WriteLine("  <body>")

            foreach ($key in $Groups.Keys) {
            # Set group
            $group = $Groups[$key]

            # Set latest component
            $componentx = $group | Sort-Object -Property FullVersion, Bundle -Descending | Select-Object -First 1

            # Set revisions
            $revisions = $RevisionHistory[$key]

            # Write name
            $StreamWriter.WriteLine("  <table class=`"title`">")
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"titlevalue`">")
            $StreamWriter.WriteLine("        $($componentx.Name)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            $StreamWriter.WriteLine("  </table>")

            # Write versions
            $StreamWriter.WriteLine("  <table class=`"properties`">")
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Versions")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            foreach ($component in $group | Sort-Object -Property FullVersion, Bundle -Descending) {
            if ($component.TypeOfChange) {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Version) ($($component.TypeOfChange))</span>")
            } else {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Version)</span>")
            }
            if ($component.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($component.Revision)</span><br/>")
            }
            if ($component.ReleaseYear -and $component.ReleaseMonth -and $component.ReleaseDay) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Release Date: $($component.ReleaseYear)-$($component.ReleaseMonth)-$($component.ReleaseDay)</span><br/>")
            }
            if ($component.BuildNumber) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Build Number: $($component.BuildNumber)</span>")
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write bundle information
            foreach ($component in $group | Sort-Object -Property FullVersion, Bundle -Descending) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">Bundle</span>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Version: $($component.Version)</span><br/>")
            if ($component.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($component.Revision)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Bundle)</span>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">File: $($component.BundleFile)</span>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write category
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Category")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($componentx.Category)</span>")
            if ($component.Type) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Type: $($componentx.Type)</span></br>")
            }
            if ($component.AltName) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Alternative Name: $($componentx.AltName)</span></br>")
            }
            if ($component.Manufacturer) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Manufacturer: $($componentx.Manufacturer)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write description
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Description")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        $($componentx.Description)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write operating systems
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Operating Systems")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <p>")
            foreach ($operatingsystem in $componentx.OperatingSystems | Sort-Object) {
            $StreamWriter.WriteLine("          $operatingsystem<br/>")
            }
            $StreamWriter.WriteLine("        </p>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write files
            foreach ($component in $group | Sort-Object -Property FullVersion, Bundle -Descending) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">Files</span>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Version: $($component.Version)</span><br/>")
            if ($component.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($component.Revision)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            foreach ($item in $component.Files) {
            $StreamWriter.WriteLine("        <a class=`"spaced`" href=`"$($item.FileUrl)`">$($item.Name)</a>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Size: $($item.Size)</span><br/>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Date: $($item.DateModified)</span><br/>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Md5sum: $($item.Md5Sum)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            if ($Details) {
            # Write notes
            foreach ($component in $group | Sort-Object -Property FullVersion, Bundle -Descending) {
            if ($component.PrerequisiteNotes -or $component.InstallationNotes -or $component.AvailabilityNotes -or $component.DocumentationNotes) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">Notes</span>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Version: $($component.Version)</span><br/>")
            if ($component.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($component.Revision)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")

            # Write prerequisite notes
            if ($component.PrerequisiteNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Prerequisites:</span>")
            foreach ($note in $component.PrerequisiteNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write installation notes
            if ($component.InstallationNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Installation:</span>")
            foreach ($note in $component.InstallationNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write availability notes
            if ($component.AvailabilityNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Availability:</span>")
            foreach ($note in $component.AvailabilityNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write documentation notes
            if ($component.DocumentationNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Documentation:</span>")
            foreach ($note in $component.DocumentationNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }
            }

            # Write revision history
            if ($revisions) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Revision History")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")

            # Write revisions
            foreach ($revision in $revisions | Sort-Object -Property FullVersion -Descending -Unique) {
            if ($revision.TypeOfChange) {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($revision.Version) ($($revision.TypeOfChange))</span>")
            } else {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($revision.Version)</span>")
            }
            if ($revision.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($revision.Revision)</span>")
            }

            # Write enhancements
            if ($revision.Enhancements) {
            $StreamWriter.WriteLine("        <span class=`"spacedin`">Enhancements:</span>")
            foreach ($enhancement in $revision.Enhancements) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $enhancement")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write fixes
            if ($revision.Fixes) {
            $StreamWriter.WriteLine("        <span class=`"spacedin`">Fixes:</span>")
            foreach ($fix in $revision.Fixes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $fix")
            $StreamWriter.WriteLine("        </div>")
            }
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }
            }

            $StreamWriter.WriteLine("  </table>")
            }

            # Write html closing
            $StreamWriter.WriteLine("  </body>")
            $StreamWriter.WriteLine("</html>")

            # Flush stream writer
            $StreamWriter.Flush()

            if (!$File) {
                # Create stream reader
                $StreamReader = New-Object System.IO.StreamReader $HtmlStream

                # Write output
                $HtmlStream.Position = 0
                while ($null -ne ($line = $StreamReader.ReadLine())) {
                    Write-Output $line
                }

                # Close stream reader
                $StreamReader.Close()
            }

            # Close stream writer
            $StreamWriter.Close()
        }

        # Check components
        if ($Components -and !$Combine) {
            # Write output
            $StreamWriter.WriteLine("<html>")
            $StreamWriter.WriteLine("  <head>")
            $StreamWriter.WriteLine("    <style>")
            $StreamWriter.WriteLine("      a {color: teal; text-decoration: none;}")
            $StreamWriter.WriteLine("      a:hover {text-decoration: underline;}")
            $StreamWriter.WriteLine("      .dimmed {color: gray;}")
            $StreamWriter.WriteLine("      .spaced:first-child {display: block; margin-top: 0; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .spaced {display: block; margin-top: 1.5em; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .spacedin {display: block; margin-bottom: 0.5em;}")
            $StreamWriter.WriteLine("      .caps {text-transform: capitalize;}")
            $StreamWriter.WriteLine("      .title {width: 98%; margin: 1em auto; border-bottom: 2px solid lightseagreen; padding-bottom: 0.5em; font-size: 95%; font-family: verdana;color: chocolate;}")
            $StreamWriter.WriteLine("      .titlevalue {width: 100%;}")
            $StreamWriter.WriteLine("      .properties {width: 98%; margin: 1em auto 2em; padding-bottom: 0.5em; font-size: 75%; font-family: verdana;}")
            $StreamWriter.WriteLine("      .propname {width: 15%; vertical-align: top; padding-bottom: 1.5em;}")
            $StreamWriter.WriteLine("      .propvalue {width: 85%; vertical-align: top; padding-bottom: 1.5em;}")
            $StreamWriter.WriteLine("      .note p, .note ul, .note ol {margin: 0 !important; padding: 0 !important;}")
            $StreamWriter.WriteLine("      .note li {list-style-position: inside !important;}")
            $StreamWriter.WriteLine("      .note li ul, .note li ol {margin-left: 1em !important;}")
            $StreamWriter.WriteLine("      .note li p:first-child {display: inline !important;}")
            $StreamWriter.WriteLine("      .note blockquote {margin: 0 !important; padding: 0 !important;}")
            $StreamWriter.WriteLine("      .note table {margin: 1em 0 !important; padding: 0 !important; font-style: normal !important; font-size: 100% !important; font-family: verdana !important; border-collapse: collapse !important;}")
            $StreamWriter.WriteLine("      .note strong, .note b {font-weight: normal !important;}")
            $StreamWriter.WriteLine("      .note u {text-decoration: none !important;}")
            $StreamWriter.WriteLine("      .note br {display: inline !important; line-height: 0 !important;}")
            $StreamWriter.WriteLine("      .note em, .note i {font-style: normal !important;}")
            $StreamWriter.WriteLine("    </style>")
            $StreamWriter.WriteLine("  </head>")
            $StreamWriter.WriteLine("  <body>")

            foreach ($component in $Components) {
            # Write name
            $StreamWriter.WriteLine("  <table class=`"title`">")
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"titlevalue`">")
            $StreamWriter.WriteLine("        $($component.Name)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            $StreamWriter.WriteLine("  </table>")

            # Write version
            $StreamWriter.WriteLine("  <table class=`"properties`">")
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Version")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            if ($component.TypeOfChange) {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Version) ($($component.TypeOfChange))</span>")
            } else {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Version)</span>")
            }
            if ($component.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($component.Revision)</span><br/>")
            }
            if ($component.ReleaseYear -and $component.ReleaseMonth -and $component.ReleaseDay) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Release Date: $($component.ReleaseYear)-$($component.ReleaseMonth)-$($component.ReleaseDay)</span><br/>")
            }
            if ($component.BuildNumber) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Build Number: $($component.BuildNumber)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write bundle information
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Bundle")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Bundle)</span>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">File: $($component.BundleFile)</span>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write category
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Category")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($component.Category)</span>")
            if ($component.Type) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Type: $($component.Type)</span></br>")
            }
            if ($component.AltName) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Alternative Name: $($component.AltName)</span></br>")
            }
            if ($component.Manufacturer) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Manufacturer: $($component.Manufacturer)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write description
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Description")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        $($component.Description)")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write operating systems
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Operating Systems")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            $StreamWriter.WriteLine("        <p>")
            foreach ($operatingsystem in $component.OperatingSystems | Sort-Object) {
            $StreamWriter.WriteLine("          $operatingsystem<br/>")
            }
            $StreamWriter.WriteLine("        </p>")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            # Write files
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Files")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")
            foreach ($item in $component.Files) {
            $StreamWriter.WriteLine("        <a class=`"spaced`" href=`"$($item.FileUrl)`">$($item.Name)</a>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Size: $($item.Size)</span><br/>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Date: $($item.DateModified)</span><br/>")
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Md5sum: $($item.Md5Sum)</span>")
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")

            if ($Details) {
            # Write notes
            if ($component.PrerequisiteNotes -or $component.InstallationNotes -or $component.AvailabilityNotes -or $component.DocumentationNotes) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Notes")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")

            # Write prerequisite notes
            if ($component.PrerequisiteNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Prerequisites:</span>")
            foreach ($note in $component.PrerequisiteNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write installation notes
            if ($component.InstallationNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Installation:</span>")
            foreach ($note in $component.InstallationNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write availability notes
            if ($component.AvailabilityNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Availability:</span>")
            foreach ($note in $component.AvailabilityNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write documentation notes
            if ($component.DocumentationNotes) {
            $StreamWriter.WriteLine("      <span class=`"spaced`">Documentation:</span>")
            foreach ($note in $component.DocumentationNotes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $note")
            $StreamWriter.WriteLine("        </div>")
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }

            # Write revision history
            if ($component.RevisionHistory) {
            $StreamWriter.WriteLine("    <tr>")
            $StreamWriter.WriteLine("      <td class=`"propname`">")
            $StreamWriter.WriteLine("        Revision History")
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("      <td class=`"propvalue`">")

            # Write revisions
            foreach ($revision in $component.RevisionHistory | Sort-Object -Property FullVersion -Descending) {
            if ($revision.TypeOfChange) {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($revision.Version) ($($revision.TypeOfChange))</span>")
            } else {
            $StreamWriter.WriteLine("        <span class=`"spaced`">$($revision.Version)</span>")
            }
            if ($revision.Revision) {
            $StreamWriter.WriteLine("        <span class=`"dimmed`">Revision: $($revision.Revision)</span>")
            }

            # Write enhancements
            if ($revision.Enhancements) {
            $StreamWriter.WriteLine("        <span class=`"spacedin`">Enhancements:</span>")
            foreach ($enhancement in $revision.Enhancements) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $enhancement")
            $StreamWriter.WriteLine("        </div>")
            }
            }

            # Write fixes
            if ($revision.Fixes) {
            $StreamWriter.WriteLine("        <span class=`"spacedin`">Fixes:</span>")
            foreach ($fix in $revision.Fixes) {
            $StreamWriter.WriteLine("        <div class=`"note dimmed spacedin`">")
            $StreamWriter.WriteLine("          $fix")
            $StreamWriter.WriteLine("        </div>")
            }
            }
            }
            $StreamWriter.WriteLine("      </td>")
            $StreamWriter.WriteLine("    </tr>")
            }
            }

            $StreamWriter.WriteLine("  </table>")
            }

            # Write html closing
            $StreamWriter.WriteLine("  </body>")
            $StreamWriter.WriteLine("</html>")

            # Flush stream writer
            $StreamWriter.Flush()

            if (!$File) {
                # Create stream reader
                $StreamReader = New-Object System.IO.StreamReader $HtmlStream

                # Write output
                $HtmlStream.Position = 0
                while ($null -ne ($line = $StreamReader.ReadLine())) {
                    Write-Output $line
                }

                # Close stream reader
                $StreamReader.Close()
            }

            # Close stream writer
            $StreamWriter.Close()
        }
    }
}

Function ConvertTo-SPPComponentCsv {
<#
    .SYNOPSIS
        Convert SPP components to csv.

    .DESCRIPTION
        The ConvertTo-SPPComponentCsv command converts Support Pack for ProLiant component objects to csv format.

    .PARAMETER Component
        Component objects.

    .PARAMETER File
        Output csv file. If omitted the output will be sent to console.

    .EXAMPLE
        Get-SPPComponent | ConvertTo-SPPComponentCsv -File 'components.csv'
        Convert component objects to csv format and save the output to the file 'components.csv'.

    .EXAMPLE
        Get-SPPComponent | ConvertTo-SPPComponentCsv | Set-Content 'components.csv'
        Convert component objects to csv format and send the output to console.

    .INPUTS
        SPPComponent[]

    .OUTPUTS
        String[]
#>

    [CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true, HelpMessage="Component objects", ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]
    $Component,

    [Parameter(Mandatory=$false, HelpMessage="Output csv file", Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]
    $File
    )

    BEGIN {
        # Set components
        $Components = @()

        # Process components from variable
        foreach ($object in $Component) {
            if ($object.PSObject.TypeNames -contains "SPPComponent") {
                $Components += $object
            }
        }
    }

    PROCESS {
        # Process components from pipeline
        foreach ($object in $Component) {
            if ($object.PSObject.TypeNames -contains "SPPComponent") {
                $Components += $object
            }
        }
    }

    END {
        # Check csv file
        if ($File) {
            # Check file exists
            $FileName = Resolve-Path $File -ErrorAction SilentlyContinue -ErrorVariable ResolvePath
            if ($FileName) {
                if($Host.UI.PromptForChoice("Overwrite File (ConvertTo-SPPComponentCsv)" , "File '$FileName' exists. Do you want to overwriet it?" , @("&Yes", "&No"), 1) -eq 1) {
                    return
                }
                $CsvFile = $FileName.Path
            } else {
                $CsvFile = $ResolvePath[0].TargetObject
            }

            # Create stream writer
            $StreamWriter = New-Object System.IO.StreamWriter $CsvFile
        } else {
            # Create memory stream
            $CsvStream = New-Object System.IO.MemoryStream

            # Create stream writer
            $StreamWriter =  New-Object System.IO.StreamWriter $CsvStream
        }

        # Check components
        if ($Components) {
            # Write header
            $StreamWriter.WriteLine('Name, Version, Revision, Build Number, Update, Bundle, Bundle File, Category, Description, Release Date, Type, Alternative Name, Manufacturer, File Name, File Size, File Date, File Md5sum, Operating System')

            # Add lines
            foreach ($component in $Components) {
                # Write name
                $line = '"' + $component.Name + '"'

                # Write version
                $line += ',"' + $component.Version + '"'

                # Write revision
                $line += ',"' + $component.Revision + '"'

                # Write build number
                $line += ',"' + $component.BuildNumber + '"'

                # Write update recommendation
                $line += ',"' + $component.TypeOfChange + '"'

                # Write bundle version
                $line += ',"' + $component.Bundle + '"'

                # Write bundle file
                $line += ',"' + $component.BundleFile + '"'

                # Write category
                $line += ',"' + $component.Category + '"'

                # Write description
                $line += ',"' + $component.Description + '"'

                # Write release date
                if ($component.ReleaseYear -and $component.ReleaseMonth -and $component.ReleaseDay) {
                    $line += ',"' + $component.ReleaseYear + '-' + $component.ReleaseMonth + '-' + $component.ReleaseDay + '"'
                } else {
                    $line += ',""'
                }

                # Write type
                $line += ',"' + $component.Type + '"'

                # Write alternative name
                $line += ',"' + $component.AltName + '"'

                # Write manufacturer
                $line += ',"' + $component.Manufacturer + '"'

                # Set file lines
                $filelines = @()

                # Write files
                foreach ($item in $component.Files) {
                    $filelines += $line + ',"' + $item.Name + '","' + $item.Size + '","' + $item.DateModified + '","' + $item.Md5Sum + '"'
                }

                # Set operating system lines
                $operatingsystemlines = @()

                # Write operating systems
                foreach ($line in $filelines) {
                    foreach ($operatingsystem in $component.OperatingSystems) {
                        $operatingsystemlines += $line + ',"' + $operatingsystem + '"'
                    }
                }

                # Write operating system lines
                foreach ($line in $operatingsystemlines) {
                    $StreamWriter.WriteLine($line)
                }
            }

            # Flush stream writer
            $StreamWriter.Flush()

            if (!$File) {
                # Create stream reader
                $StreamReader = New-Object System.IO.StreamReader $CsvStream

                # Write output
                $CsvStream.Position = 0
                while ($null -ne ($line = $StreamReader.ReadLine())) {
                    Write-Output $line
                }

                # Close stream reader
                $StreamReader.Close()
            }

            # Close stream writer
            $StreamWriter.Close()
        }
    }
}
