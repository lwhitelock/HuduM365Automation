param($name)

$customerExclude = $env:customerExclude -split ','

try {
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

} catch {
    Write-Host "Error fetching customers: $_"

    $WebhookURL = $env:WebhookURL

    $AdaptiveCard = [pscustomobject]@{
        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
        type      = 'AdaptiveCard'
        version   = '1.2'
        body      = @(@{
                type  = 'TextBlock'
                text  = "Error Occured Fetching Companies - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                style = 'heading'
            },
            @{
                type = "TextBlock"
                text = "Error: $_"
                wrap = $true
            }
        )
    }


    $TeamsMessageBody = [PSCustomObject]@{
        type        = 'message'
        attachments = @(@{
                contentType = 'application/vnd.microsoft.card.adaptive'
                contentURL  = $null
                content     = $AdaptiveCard
                width       = 'stretch'
            })

    }


    $parameters = @{
        "URI"         = $WebhookURL
        "Method"      = 'POST'
        "Body"        = $TeamsMessageBody | convertto-json -depth 100
        "ContentType" = 'application/json'
    }

    Invoke-RestMethod @parameters
    return $Null

}