function get-FolderPermissions {

    

[CmdletBinding()]
param (
    [string]$SharePath
    
)
 $OutputCSV = "C:\users\Public\$($env:COMPUTERNAME).csv"
Get-ChildItem -Path $SharePath -Directory -Recurse -Depth 2 | 
ForEach-Object {
    $Path = $_
    $acl = Get-Acl -Path $_.FullName
    $acl.Access | ForEach-Object {
        $str = "";
        if ( $_ ) { 
            $str += $_.IdentityReference.ToString();
            $str += " ";
            $str += $_.AccessControlType.ToString();
            $str += "  ";
            $str += $_.FileSystemRights.ToString();
        }
        [pscustomobject]@{
            Access = $str
            Group  = $acl.Group
            Owner  = $acl.Owner
            Sddl   = $acl.Sddl
            Path   = $Path
        } | Select-Object -Property Path, Access, Group, Owner, Sddl
    }
} | Export-Csv -Path $OutputCSV -Encoding utf8 -Delimiter ','
}