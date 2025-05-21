#Requires -Modules ActiveDirectory,ImportExcel,@{ModuleName="OZO";ModuleVersion="1.6.0"},OZOAD -Version 5.1

<#PSScriptInfo
    .VERSION 1.0.0
    .GUID 212d85f0-e58c-4578-8fce-5cd4c33f1c75
    .AUTHOR Andy Lievertz <alievertz@onezeroone.dev>
    .COMPANYNAME One Zero One
    .COPYRIGHT This script is released under the terms of the GNU General Public License ("GPL") version 2.0.
    .TAGS 
    .LICENSEURI https://github.com/onezeroone-dev/OZO-AD-Active-Users-Report/blob/main/LICENSE
    .PROJECTURI https://github.com/onezeroone-dev/OZO-AD-Active-Users-Report
    .ICONURI 
    .EXTERNALMODULEDEPENDENCIES ActiveDirectory,ImportExcel
    .REQUIREDSCRIPTS 
    .EXTERNALSCRIPTDEPENDENCIES 
    .RELEASENOTES https://github.com/onezeroone-dev/OZO-AD-Active-Users-Report/blob/main/CHANGELOG.md
#>

<# 
    .SYNOPSIS
    See description.
    .DESCRIPTION 
    Produces an Excel report of active domain users including display name, title, office, department, and password status.
    .PARAMETER OutDir
    Directory for the Excel report. Defaults to the current directory.
    .LINK
    https://github.com/onezeroone-dev/OZO-AD-Active-Users-Report/blob/main/README.md
    .NOTES
    Empty cells in the password status field indicates the password has never been set.
#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess = $true)] Param (
    [Parameter(Mandatory=$false,HelpMessage="Directory for the Excel report")][String]$OutDir = (Get-Location)
)

# CLASSES
Class OZOADActiveUsers {
    # PROPERTIES: Booleans, DateTimes, Strings
    [Boolean]  $Validates    = $true
    [DateTime] $todayDate    = (Get-Date)
    [String]   $excelPath    = $null
    [String]   $outDir       = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $adDomain  = $null
    [PSCustomObject] $ozoLogger = $null
    # PROPERTIES: PSCustomObject Lists
    [System.Collections.Generic.List[PSCustomObject]] $activeADUsers = @()
    # METHODS: Constructor method
    OZOADActiveUsers($OutDir) {
        # Set properties
        $this.outDir = $OutDir
        # Create a logger
        $this.ozoLogger = (New-OZOLogger)
        # Log a process start message
        $this.ozoLogger.Write("Process starting.","Information")
        # Determine if the environment validates
        If ($this.ValidateEnvironment() -eq $true) {
            # Environment validates; get the active users
            $this.GetActiveUsers()
        } Else {
            # Environment does not validate; log error
            $this.ozoLogger.Write("Environment does not validate.","Error")
            $this.Validates -eq $false
        }
        # Report
        $this.Report()
        # Log a process end message
        $this.ozoLogger.Write("Process complete.","Information")
    }
    # METHODS: Validate environment method
    Hidden [Boolean] ValidateEnvironment() {
        # Control variable
        [Boolean] $Return = $true
        # Try to get domain information
        Try {
            $this.adDomain = (Get-ADDomain -ErrorAction Stop)
            # Success
        } Catch {
            # Failure; create a dummy object to support the OutDir test
            $this.adDomain = [PSCustomObject]@{Name = "NULL"}
            # Log
            $this.ozoLogger.Write(("Unable to obtain AD domain information. Error message is: " + $_),"Error")
            $Return = $false
        }
        # Determine if OutDir is writable
        If ((Test-OZOPath -Path $this.outDir -Writable)) {
            # OutDir is writable
            $this.excelPath = (Join-Path -Path $this.outDir -ChildPath ((Get-OZO8601Date -Time) + "-" + (Get-ADDomain).Name + "-OZO-AD-Active-Users-Report.xlsx"))
        } Else {
            $this.ozoLogger.Log(($this.outDir + " does not exist or is not writable."),"Error")
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: Get active users method
    Hidden [Void] GetActiveUsers() {
        Try {
            $this.activeADUsers = (Get-OZOADUsers -Enabled -UserProperties Department,GivenName,"msDS-UserPasswordExpiryTimeComputed",Office,PasswordLastSet,SAMAccountName,Surname,Title -ErrorAction Stop)
            # Success
        } Catch {
            # Failure
            $this.ozoLogger.Write(("Failed to get AD users. Error message is: " + $_),"Error")
        }
    }
    # METHODS: Report
    Hidden [Void] Report() {
        # Determine if we found at least one user
        If ($this.activeADUsers.Count -gt 0) {
            # We found at least one user; export desired properties to Excel
            $this.activeADUsers | Select-Object -Property Surname,
            GivenName,
            SamAccountName,
            PasswordLastSet,
            @{Name = "Days Since Password Last Set"; Expression = {(New-TimeSpan -Start $_.PasswordLastSet -End $this.todayDate).Days}},
            @{Name = "Password Expiration Date"; Expression = {If ($_."msDS-UserPasswordExpiryTimeComputed" -gt 0) { [DateTime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed") } Else { $null }}},
            Title,
            Office,
            Department | Export-Excel -Path $this.excelPath
            # Log a help message
            $this.ozoLogger.Write(("Export Complete. Please see " + $this.excelPath + "."),"Information")
        }
    }
}

# Create an instance of the OZOADActiveUsers object
[OZOADActiveUsers]::new($OutDir) | Out-Null
