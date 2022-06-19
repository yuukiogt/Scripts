
$csvPath = ".\Users.csv"
$top = ""

Import-Csv $csvPath |
Foreach-Object {
    if ($_.UserPrincipalName -eq $top) {
        $manager = ""
    }
    else {
        $manager = (Get-ADUser -Identity $_.Manager).DistinguishedName
    }

    if(![string]::IsNullOrEmpty($_.Mobile) -and !$_.Mobile.StartsWith('0')) {
        $_.Mobile = '0' + $_.Mobile
    }

    $setuser_args = @{
        Identity    = $_.UserPrincipalName   
        Department  = $_.Department
        Title       = $_.Title
        Manager     = $manager            
        Office      = $_.Office
        MobilePhone = $_.Mobile        
        Company     = $_.Company
    };

    $keys = @($setuser_args.Keys | Where-Object { [string]::IsNullOrEmpty($setuser_args[$_]) })
    $keys | ForEach-Object { $setuser_args.Remove($_) }
        
    Set-ADUser @setuser_args;
    
    Set-ADUser $_.UserPrincipalName -add @{
        'msDS-PhoneticDisplayName' = $_.PhoneticDisplayName
    }
}