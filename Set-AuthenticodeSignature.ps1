$cert = @(Get-ChildItem -Path cert:\CurrentUser\my -CodeSigningCert)

$cert

Set-AuthenticodeSignature -Filepath "" -Cert $cert[2]