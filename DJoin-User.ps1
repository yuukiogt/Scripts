$ConfirmPreference = "None"
$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

<#
.Synopsis
   Join a domain using a DomainJoinFile
.DESCRIPTION
   This is a PowerShell frontend to the DJOIN.exe command which provides better discoverability and usability.
.EXAMPLE
   PS> New-DomainJoinFile -Domain dev.contoso.com -ComputerName server1 -Path c:\temp\oldj
   PS> New-PSDrive -Name S1 -PSProvider FileSystem -Root \\Server1\c$
   PS> mkdir S1:\Temp
   PS> Copy-Item c:\temp\oldj S1:\temp
   PS> Invoke-Command -computer server1 { Join-DomainUsingFile -Path c:\temp\oldb -WindowsPath C:\Windows -JoinLocalOS }
#>
function Join-DomainUsingFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param
    (
        # Path to the DomainJoinFile
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript( { test-path $_ })]
        [String]$Path,

        # Path to Windows directory in an offline image
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateScript( { test-path $_ })]
        [String]$WindowsPath,

        # WindowsPath specifies the locally running OS
        [Parameter()]
        [Switch]$JoinLocalOS
    )

    # These statements and the use of single quotes are SECURITY CRITICAL.
    # Without these someone could do an injection attack (e.g. provider a parameter with a ";" to terminate the statement
    # and then start a new, evil, command.
    $Path = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($Path)
    $WindowsPath = [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($WindowsPath)

    if ($JoinLocalOS.IsPresent) {
        $TargetString = "Local computer [$(hostname)]"
    }
    else {
        $TargetString = "Offine computer"
    }

    if ($PSCmdlet.ShouldProcess($TargetString, "Use [$Path] to domain join")) {
        if ($JoinLocalOS.IsPresent) {
            djoin.exe /requestodj /LoadFile '$Path' /Windowspath '$WindowsPath' /LocalOS
        }
        else {
            djoin.exe /requestodj /LoadFile '$Path' /Windowspath '$WindowsPath' 
        }
    }
}

try {
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $argsToAdminProcess = ""
        $Args.ForEach{ $argsToAdminProcess += " `"$PSItem`"" }
        Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess"  -Verb RunAs
        exit
    }

    $localUser = Get-LocalUser -Name root
    if($Null -eq $localUser) {
        $UserFile = Join-Path $CurrentDir ".\DJoin_AddLocalUser.txt"
        $UserString = Get-Content $UserFile | ConvertTo-SecureString -Key (1..16)
        New-LocalUser -Name root -Password $UserString -PasswordNeverExpires
    }

    $computerName = Get-ChildItem -Path . "*.txt"
    Rename-Computer -NewName $computerName.BaseName -Force

    $path = Join-Path $PSScriptRoot ($computerName.BaseName + ".txt")

    Join-DomainUsingFile -Path $path -WindowsPath "C:\Windows" -JoinLocalOS -Confirm
}
catch {
    $_.Exception.Message
}
finally {
    Read-Host "Press any key..."
}