param($Context)

$Customers = Invoke-DurableActivity -FunctionName 'Get-M365Tenants'

$ProcessingCompanies = foreach ($Customer in $Customers) {
    Invoke-DurableActivity -FunctionName 'Exec-HuduM365ProcessTenant' -Input $Customer -NoWait
}

$Outputs = Wait-ActivityFunction -Task $ProcessingCompanies

Write-Host "Complete: $($Outputs | Convertto-json | out-string)"

$WebhookURL = $env:WebhookURL

$ErrorOutputs = $Outputs | where-object { $_.Errors }
$ErrorOutputs | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'ErrorCount' -Value "$(($_.errors | measure-object).count)" }

$AdaptiveCard = [pscustomobject]@{
    '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
    type      = 'AdaptiveCard'
    version   = '1.2'
    body      = @(@{
            type  = 'TextBlock'
            text  = "M365 to Hudu Sync Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            style = 'heading'
        },
        @{
            type  = 'FactSet'
            width = 'stretch'
            facts = @(
                @{
                    title = 'Total Customers'
                    value = ($Outputs | measure-object).count
                },
                @{
                    title = 'Total Users'
                    value = ($Outputs.users | measure-object -sum).sum
                },
                @{
                    title = 'Total Devices'
                    value = ($Outputs.devices | measure-object -sum).sum
                },
                @{
                    title = 'Total Errors'
                    value = ($ErrorOutputs.ErrorCount | measure-object -sum).sum
                },
                @{
                    title = 'Customers with Errors'
                    value = ($ErrorOutputs | measure-object).count
                }
            )
        }
    )
}



$TeamsMessageBody = [PSCustomObject]@{
    type        = 'message'
    attachments = @(@{
            contentType = 'application/vnd.microsoft.card.adaptive'
            contentURL  = $null
            content     = $AdaptiveCard
            width = 'stretch'
        })

}


$parameters = @{
    "URI"         = $WebhookURL
    "Method"      = 'POST'
    "Body"        = $TeamsMessageBody | convertto-json -depth 100
    "ContentType" = 'application/json'
}

Invoke-RestMethod @parameters




for ($i = 0; $i -lt $ErrorOutputs.count; $i += 20) {                                                                                                                                              
    [System.Collections.Generic.List[PSCustomObject]]$AdaptiveBody = @()

    $CustomerHeading = [pscustomobject]@{
        type                = 'TextBlock'
        weight              = 'Bolder'
        text                = "M365 to Hudu Sync Errors - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        size                = 'small'
        wrap                = $true
        seperator           = $true
        horizontalAlignment = 'center'
    }
    $AdaptiveBody.add($CustomerHeading)

    [System.Collections.Generic.List[PSCustomObject]]$CustomerDetailsColumns = @()

    $CustomerDetailsColumns.add((Get-AdaptiveColumn -Strings $ErrorOutputs[$i..($i + 19)].name -Title 'Customer'))
    $CustomerDetailsColumns.add((Get-AdaptiveColumn -Strings $ErrorOutputs[$i..($i + 19)].ErrorCount -Title 'Errors'))

    $AdaptiveBody.add([pscustomobject]@{
            type    = 'ColumnSet'
            columns = $CustomerDetailsColumns
            width   = 'strech'
        })


    $CustomersWithErrors = $ErrorOutputs[$i..($i + 19)] | Where-Object { $_.errors } | Select-Object name, @{n = 'Errors'; e = { $_.errors -join ', ' } }

    foreach ($Customer in $CustomersWithErrors) {
        $ErrorParsed = $Customer.Errors | foreach-object { "- $_`n" }
        $Message = [pscustomobject]@{
            type      = 'TextBlock'
            text      = "**$($Customer.name)**`n$ErrorParsed"
            wrap      = $true
            seperator = $true
        }
        $AdaptiveBody.add($Message)
    }


    $AdaptiveCard = [pscustomobject]@{
        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
        type      = 'AdaptiveCard'
        version   = '1.2'
        body      = $AdaptiveBody
    }

    $TeamsMessageBody = [PSCustomObject]@{
        type        = 'message'
        attachments = @(@{
                contentType = 'application/vnd.microsoft.card.adaptive'
                contentURL  = $null
                content     = $AdaptiveCard
            })

    }

    $parameters = @{
        "URI"         = $WebhookURL
        "Method"      = 'POST'
        "Body"        = $TeamsMessageBody | convertto-json -depth 100
        "ContentType" = 'application/json'
    }

    Invoke-RestMethod @parameters
}

