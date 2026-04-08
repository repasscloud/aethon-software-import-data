$jobsToRun = @(
    [PSCustomObject]@{
        SearchUrl  = "https://careers-inc.nttdata.com/search/?q=&locationsearch=PH&location=PH&sortColumn=referencedate&sortDirection=desc"
        OutputPath = "${PSScriptRoot}/nttdata-ph.json"
    }
    [PSCustomObject]@{
        SearchUrl  = "https://careers-inc.nttdata.com/search/?q=&locationsearch=US&location=US&sortColumn=referencedate&sortDirection=desc"
        OutputPath = "${PSScriptRoot}/nttdata-us.json"
    }
)

foreach ($job in $jobsToRun) {
    $PSScriptRoot/Get-NttDataJobs.ps1 `
        -SearchUrl $job.SearchUrl `
        -OutputPath $job.OutputPath
}
