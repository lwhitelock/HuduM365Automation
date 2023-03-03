param($Context)

$DurableRetryOptions = @{
    FirstRetryInterval  = (New-TimeSpan -Minutes 2)
    MaxNumberOfAttempts = 2
    BackoffCoefficient  = 2
}
$RetryOptions = New-DurableRetryOptions @DurableRetryOptions

try {

    $Customers = Invoke-ActivityFunction -FunctionName 'Get-M365Tenants'

    $ProcessingCompanies = foreach ($Customer in $Customers) {
        Invoke-DurableActivity -FunctionName 'Exec-HuduM365ProcessTenant' -Input $Customer -NoWait -RetryOptions $RetryOptions
    }
    $Outputs = Wait-ActivityFunction -Task $ProcessingCompanies
    Write-Host "Complete: $($Outputs | ConvertTo-Json | Out-String)"

    Invoke-ActivityFunction -FunctionName 'Send-ResultsNotification' -Input ($Outputs | ConvertTo-Json -Compress -Depth 10)

    Write-Host 'HuduM365Orchestrator completed.'

} catch {
    Write-Host $_.Exception.Message
}