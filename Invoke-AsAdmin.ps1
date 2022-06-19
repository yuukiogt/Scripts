if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    $argsToAdminProcess = ""
    $Args.ForEach{ $argsToAdminProcess += "`"$PSItem`"" }
    Start-Process powershell.exe "-File `"$PSCommandPath`" $argsToAdminProcess" -Verb RunAs
    exit
}