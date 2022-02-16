param($name)
$customerExclude = $env:customerExclude -split ','
$Authheaders = Get-GraphToken -tenantid $env:Tenantid
[System.Collections.Generic.List[PSCustomObject]]$Customers = (New-GraphGetRequest -uri "https://graph.microsoft.com/beta/contracts?`$top=999" -tenantid $env:Tenantid -Headers $Authheaders) | Select-Object CustomerID, DefaultdomainName, DisplayName | Where-Object -Property DisplayName -NotIn $customerExclude
if ([System.Convert]::ToBoolean($env:DocumentPartnerTenant) -eq $true) {
    $null = $Customers.add([PSCustomObject]@{
        customerId        = $env:TenantID
        defaultDomainName = $env:PartnerDefaultDomain
        displayName       = $env:PartnerDisplayName
    })
}
return $Customers
