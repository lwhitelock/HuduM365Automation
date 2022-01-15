param($Context)

$Customers = Invoke-DurableActivity -FunctionName 'Get-M365Tenants'

$ProcessingCompanies = foreach ($Customer in $Customers){
    Invoke-DurableActivity -FunctionName 'Exec-HuduM365ProcessTenant' -Input $Customer -NoWait
}

$Outputs = Wait-ActivityFunction -Task $ProcessingCompanies

$OutJSON = $Outputs | convertto-json -depth 100 | out-string

$null = Invoke-DurableActivity -FunctionName 'Exec-HuduM365ProcessResults' -Input $OutJSON

