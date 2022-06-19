<#
.Synopsis
   Create an DomainJoinFile to offline join a domain
.DESCRIPTION
   This is a PowerShell frontend to the DJOIN.exe command which provides better discoverability and usability.
   Use -Verbose to see what DJOIN command gets executed, use -Verbose.
.EXAMPLE
   PS> New-DomainJoinFile -Domain dev.contoso.com -ComputerName server1 -Path c:\temp\oldj
   PS> New-PSDrive -Name S1 -PSProvider FileSystem -Root \\Server1\c$
   PS> mkdir S1:\Temp
   PS> Copy-Item c:\temp\oldj S1:\temp
   PS> Invoke-Command -computer server1 { Join-DomainUsingFile -Path c:\temp\oldb -WindowsPath C:\Windows -JoinLocalOS }
 
#>
function New-DomainJoinFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param
    (
        # Name of the domain to join e.g. dev.Contoso.com
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$Domain,

        # Name of the computer to join the domain
        [Parameter(Mandatory = $true, Position = 1)]
        [String]$ComputerName,

        #
        [Parameter(Mandatory = $true, Position = 2)]
        [String]$Path,

        #Organization Unit (OU) where the account is created
        [Parameter()]
        [String]$MachineOU,

        # Reuse any existing account (password will be reset)
        [Parameter()]
        [Switch]$Reuse,

        # Skip account conflict detection, requires DCNAME (faster)
        [Parameter()]
        [Switch]$NoSearch,

        # Support using a Windows Server 2008 DC or earlier
        [Parameter()]
        [Switch]$DownLevel,

        # Return base64 encoded metadata blob for an answer file
        [Parameter()]
        [Switch]$Printable,

        # Include root Certificate Authority certificates.
        [Parameter()]
        [Switch]$RootCACerts,

        # Machine certificate template.
        # Includes root Certificate Authority certificates.
        [Parameter()]
        [String]$CertTemplate,

        # Semicolon-separated list of policy names.
        # Each name is the displayName of the GPO in AD.
        [Parameter()]
        [String]$PolicyNames,

        # Semicolon-separated list of policy paths.
        # Each path is a path to a registry policy file.
        [Parameter()]
        [String]$PolicyPaths,

        # Netbios Name of the computer joining the domain
        [Parameter()]
        [String]$NetBIOS,

        # Name of persistent site to put the computer joining the domain in.
        [Parameter()]
        [String]$PersistentSite,

        # Name of dynamic site to initially put the computer joining the domain in.
        [Parameter()]
        [String]$DynamicSite,

        # Name of primary DNS domain of the computer joining the domain.
        [Parameter()]
        [String]$PrimaryDNS
    )

    # These statements and the use of single quotes are SECURITY CRITICAL.
    # Without these someone could do an injection attack (e.g. provider a parameter with a ";" to terminate the statement
    # and then start a new, evil, command.
    $Domain = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($Domain)
    $ComputerName = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($ComputerName)
    $Path = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($Path)
    $MachineOU = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($MachineOU)
    $CertTemplate = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($CertTemplate)
    $PolicyNames = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($PolicyNames)
    $PolicyPaths = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($PolicyPaths)
    $NetBIOS = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($NetBIOS)
    $PersistentSite = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($PersistentSite)
    $DynamicSite = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($DynamicSite)
    $PrimaryDNS = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($PrimaryDNS)


    $cmd = "djoin.exe /provision /domain '$Domain' /machine '$ComputerName' /savefile '$Path' "

    if ($PSBoundParameters.ContainsKey('MachineOU')) {
        $cmd += "/MACHINEOU '$MachineOU' "
    }

    if ($Reuse.IsPresent) {
        $cmd += "/Reuse "
    }

    if ($NoSearch.IsPresent) {
        $cmd += "/NoSearch "
    }

    if ($DOWNLEVEL.IsPresent) {
        $cmd += "/DOWNLEVEL "
    }

    if ($Printable.IsPresent) {
        $cmd += "/PRINTBLOB "
    }

    if ($RootCACerts.IsPresent) {
        $cmd += "/RootCACerts "
    }

    if ($PSBoundParameters.ContainsKey('MachineOU')) {
        $cmd += "/MACHINEOU '$MachineOU' "
    }

    if ($PSBoundParameters.ContainsKey('CertTemplate')) {
        $cmd += "/CertTemplate '$CertTemplate' "
    }

    if ($PSBoundParameters.ContainsKey('PolicyNames')) {
        $cmd += "/POLICYNAMES '$PolicyNames' "
    }

    if ($PSBoundParameters.ContainsKey('PolicyPaths')) {
        $cmd += "/POLICYPaths '$PolicyPaths' "
    }

    if ($PSBoundParameters.ContainsKey('NetBIOS')) {
        $cmd += "/NetBIOS '$NetBIOS' "
    }

    if ($PSBoundParameters.ContainsKey('PersistentSite')) {
        $cmd += "/PSITE '$PersistentSite' "
    }

    if ($PSBoundParameters.ContainsKey('DynamicSite')) {
        $cmd += "/DSITE '$DynamicSite' "
    }

    if ($PSBoundParameters.ContainsKey('PrimaryDNS')) {
        $cmd += "/PRIMARYDNS '$PrimaryDNS' "
    }

    if ($PSBoundParameters.ContainsKey("Verbose")) {
        Write-Verbose $cmd
    }

    if ($PSCmdlet.ShouldProcess($domain, "Domain join computer [$computerName]")) {
        Invoke-Expression $cmd
    }
}

try {
    do {
        $computerName = Read-Host "コンピュータ名を入力してください"
    } while ($computerName -notmatch "^COMPANY[0-9][0-9][0-9|R][0-9][0-9][D|N|T]*[A|C|L|R]*")

    $path = Join-Path $PSScriptRoot ($computerName + ".txt")

    $OU = "*"

    New-DomainJoinFile -Domain "tenant.local" -ComputerName $computerName -Path $path -MachineOU "OU=$($OU),DC=tenant,DC=local" -Reuse -Confirm
} catch {
    $_.Exception.Message
}