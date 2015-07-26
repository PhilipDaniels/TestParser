###$username = "systeam@landmarkinfo.co.uk"
###$password =  "…"

###$projectId = ${bamboo.TESTRAIL_PROJECT_ID}

###$getcases_url = "https://landmark.testrail.com/index.php?/api/v2/get_cases/"
###$addrun_url = "https://landmark.testrail.com/index.php?/api/v2/add_run/"
###$addresultforcase_url = "https://landmark.testrail.com/index.php?/api/v2/add_result_for_case/"

###$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username, $password)))

#Get Cases For project and get ID of "Behaviour Tests"
###$cases = Invoke-RestMethod -Headers @{Authorization =("Basic {0}" -f $base64AuthInfo)} -Method Get -ContentType "application/json" -Uri $getcases_url$projectid 
###$case = $cases | where { $_.title -eq "Feature Behaviour Tests"} | Foreach-object {echo $_.id}

# add New Test run for project
$datetime = Get-Date
$hash = @{ "name" = "Behaviour Tests: " + $datetime}
$json = $hash | ConvertTo-Json

###$newrun = Invoke-RestMethod -Headers @{Authorization =("Basic {0}" -f $base64AuthInfo)} -Method Post -ContentType "application/json" -Body $json -Uri $addrun_url$projectid

#Add Results to new test run
[xml]$results = Get-Content "C:\temp\QDW.trx"

echo $results.TestRun.Results | % { 
$status = $_.UnitTestResult | select outcome
$duration = $_.UnitTestResult | select duration
$text = $_.UnitTestResult.Output.StdOut }


$text = [uri]::EscapeDataString($text)


if ( $status.outcome -eq "passed") {
$status_id = 1
}else{
$status_id = 5
}

#$formatduration = $duration.duration -replace ",","."
#$seconds = ([TimeSpan]::Parse($formatduration)).TotalSeconds
#$seconds = "{0:N0}" -f $seconds
#$seconds = $seconds+"s"
#$seconds

$hash = @{ "status_id" = $status_id; "comment" = $text; "elapsed" = 0}
$json = $hash | ConvertTo-Json
###$newrunid = $newrun.id
###$addresult = Invoke-RestMethod -Headers @{Authorization =("Basic {0}" -f $base64AuthInfo)} -Method Post -ContentType "application/json" -Body $json -Uri $addresultforcase_url$newrunid"/"$case


Write-Host $json
