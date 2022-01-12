param($Context)

$Customers = Invoke-DurableActivity -FunctionName 'Get-M365Tenants'

$ProcessingCompanies = foreach ($Customer in $Customers){
    Invoke-DurableActivity -FunctionName 'Exec-HuduM365ProcessTenant' -Input $Customer -NoWait
}

$Outputs = Wait-ActivityFunction -Task $ProcessingCompanies

Write-Host "Complete: $($Outputs | Where-Object {$_.Errors.count -gt 0}|convertto-json -depth 100 | out-string)"

