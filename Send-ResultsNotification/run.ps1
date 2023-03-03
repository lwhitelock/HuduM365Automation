Param($Outputs)

try {
    $WebhookURL = $env:WebhookURL
    if ($WebhookURL) {

        $ErrorOutputs = $Outputs | Where-Object { $_.Errors }
        $ErrorOutputs | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'ErrorCount' -Value $(($_.Errors | Measure-Object).count) }

        #Write-Output "Send notification: $($ErrorOutputs | ConvertTo-Json)"
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
                            value = ($Outputs | Measure-Object).count
                        },
                        @{
                            title = 'Total Users'
                            value = ($Outputs.users | Measure-Object -Sum).sum
                        },
                        @{
                            title = 'Total Devices'
                            value = ($Outputs.devices | Measure-Object -Sum).sum
                        },
                        @{
                            title = 'Total Errors'
                            value = ($ErrorOutputs.ErrorCount | Measure-Object -Sum).sum
                        },
                        @{
                            title = 'Customers with Errors'
                            value = ($ErrorOutputs | Measure-Object).count
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
                    width       = 'stretch'
                })

        }


        $parameters = @{
            'URI'         = $WebhookURL
            'Method'      = 'POST'
            'Body'        = $TeamsMessageBody | ConvertTo-Json -Depth 100
            'ContentType' = 'application/json'
        }

        Invoke-RestMethod @parameters
        $ErrorCount = ($ErrorOutputs | Measure-Object).Count
        for ($i = 0; $i -lt $ErrorCount; $i += 20) {
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

            $CustomerDetailsColumns.add((Get-AdaptiveColumn -Strings $ErrorOutputs[$i..($i + 19)].Name -Title 'Customer'))
            $CustomerDetailsColumns.add((Get-AdaptiveColumn -Strings $ErrorOutputs[$i..($i + 19)].ErrorCount -Title 'Errors'))

            $AdaptiveBody.add([pscustomobject]@{
                    type    = 'ColumnSet'
                    columns = $CustomerDetailsColumns
                    width   = 'strech'
                })

            #Write-Output "ErrorOutputs: $($ErrorOutputs | ConvertTo-Json -Depth 10)"
            $CustomersWithErrors = $ErrorOutputs | Select-Object -First 20 -Skip $i Name, Errors

            foreach ($Customer in $CustomersWithErrors) {
                Write-Output "Customer: $($Customer | ConvertTo-Json -Depth 10)"
                $ErrorParsed = ($Customer.Errors | ForEach-Object { "- $_" }) -join "`n"
                $Message = [pscustomobject]@{
                    type      = 'TextBlock'
                    text      = "**$($Customer.Name)**`n$ErrorParsed"
                    wrap      = $true
                    seperator = $true
                }
                $AdaptiveBody.add($Message)

                #Write-Output ($Message | ConvertTo-Json -Depth 10)
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
                'URI'         = $WebhookURL
                'Method'      = 'POST'
                'Body'        = $TeamsMessageBody | ConvertTo-Json -Depth 100
                'ContentType' = 'application/json'
            }

            Invoke-RestMethod @parameters
        }

    }
} catch {
    Write-Host "Error sending webhook $($_.Exception.Message)"
}
