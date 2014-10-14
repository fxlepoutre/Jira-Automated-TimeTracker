#############################
##### GLOBAL PARAMETERS #####
#############################

$jiraURL = "http://jira.sophieparis.com"
$secondsPerDay = 8*3600


#############################
#####     FUNCTIONS     #####
#############################

# Calls REST procedure to read details of selected issue.
Function GetDetailsOfJiraIssue ($issueCode) {
    $uri = $Script:jiraURL+"/rest/api/2/issue/"+$issueCode
    return Invoke-RestMethod -ContentType "application/json" -Method Get -Uri $uri -TimeoutSec 5  -WebSession $Script:session
}

# Calls REST procedure to add worklog to specified issue.
Function AddWorklogToJiraIssue ($issueCode, $timeSpentSeconds, $dateWork) {
    $dateWorkJiraFormat = ($dateWork.ToString("s")) +".000+0700"
    $uri = $Script:jiraURL+"/rest/api/2/issue/"+$issueCode+"/worklog?adjustEstimate=AUTO"
    $contents = @{
        timeSpentSeconds = $timeSpentSeconds;
        started = $dateWorkJiraFormat
    } | ConvertTo-Json
    Invoke-RestMethod -ContentType "application/json" -Method Post -Uri $uri -Body $contents -TimeoutSec 5  -WebSession $Script:session
}

# Calls REST procedure to remove specified worklog.
Function RemoveWorklogFromJiraIssue ($issueCode, $worklogCode) {
    $uri = $Script:jiraURL+"/rest/api/2/issue/"+$issueCode+"/worklog/"+$worklogCode+"?adjustEstimate=new&newEstimate=0"
    Invoke-RestMethod -ContentType "application/json" -Method Delete -Uri $uri -TimeoutSec 5 -WebSession $Script:session
}

#############################
#####    MAIN SCRIPT    #####
#############################

# Read credentials.
$username = Read-Host 'Jira username'
$password = Read-Host 'Jira password' -AsSecureString
$passwordDecoded = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

# Create authentication header.
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username,$passwordDecoded)))

# Read input date and set timestamp based on this.
$worklogDateInput = Read-Host 'Input date (dd/mm/yy) or leave blank for today'
if(!$worklogDateInput) {
    $worklogDate = Get-Date
    $worklogDateStr = ""+$worklogDate.Day+"/"+$worklogDate.Month+"/"+$worklogDate.Year
    $worklogDate = Get-Date $worklogDateStr
} else {
    $worklogDate = Get-Date $worklogDateInput
}
$dateStart = Get-Date $worklogDate -UFormat %s
$dateEnd = (Get-Date $worklogDate.AddDays(1) -UFormat %s) - 1

# Try login and read XML feed of activities..
try {
    # Read XML feed of activities.
    $uri = $Script:jiraURL+"/activity?streams=user+IS+fx&streams=update-date+BETWEEN+"+$dateStart+"000+"+$dateEnd+"999&maxResults=500"
    [xml]$activityStream = Invoke-WebRequest -Uri $uri -Method Get -SessionVariable session -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -TimeoutSec 10
} catch {
    write-Host "Login error, check if login on Jira is possible."
    Break
}

# Parse XML feed of activities and get interresting details.
$activityEntries = @()
ForEach ($entry in $activityStream.feed.entry) {
    
    # Find issue related to activity entry
    $issueKey = $entry.target.title.InnerText
    if((!$issueKey) -or ($issueKey.Trim() -eq "")) {
        $issueKey = $entry.object.title.innerText
    }
    
    # Find action that generated activity entry
    $action = $entry.category.term
    if(!$action -or ($action.Trim() -eq "")) {
        $action = $entry.verb.SubString($entry.verb.Length - 6, 6)
    }
    
    # Assign coefficient to defined action
    $coef = 0
    switch($action) {
        "closed" {$coef = 3; break}
        "created" {$coef = 5; break}
        "update" {$coef = 1; break}
        "comment" {$coef = 2; break}
        "reopened" {$coef = 4; break}
        "0/post" {$coef = 3; break}
        default {$coef = 0}
    }

    # Get date of the activity
    $datePublished = $entry.published

    # Load all results into the datastore of activity entries.
    $activityEntries += [pscustomobject]@{issueKey=$issueKey;coef=$coef;date=$datePublished}
}

# Always add entry to this worklog.
$activityEntries += [pscustomobject]@{issueKey="ITHD-8454";coef=1;date=$worklogDate}

# Create a new datastore which groups the results by issue, in order to have only one worklog by issue. 
$issuesWithWork = $activityEntries | Group-Object issueKey | %{
    New-Object psobject -Property @{
        Item = $_.Name
        Sum = ($_.Group | Measure-Object coef -Sum).Sum
        Date = $worklogDate
    }
}

# Get current status of items
$issuesWithWork | Add-Member -NotePropertyName Status -NotePropertyValue "Unknown"
ForEach ($issueWithWork in $($issuesWithWork)) {
    $issueDetails = GetDetailsOfJiraIssue $issueWithWork.Item
    $issueWithWork.Status = $issueDetails.fields.status.name
}
# Remove entries that are closed
$issuesWithWork = $issuesWithWork | where {$_.Status -ne "Closed"}

# Record the worklog depending on the 
$totalCoeff = ($issuesWithWork | Measure-Object Sum -Sum).Sum
ForEach ($issueWithWork in $issuesWithWork) {
    $timeSpentInSec = $issueWithWork.Sum / $totalCoeff * $secondsPerDay
    
    Write-Host "Adding worklog on" $issueWithWork.Item "-" $timeSpentInSec "seconds" "-" $issueWithWork.Date
    $worklog = AddWorklogToJiraIssue $issueWithWork.Item $timeSpentInSec $issueWithWork.Date
    #RemoveWorklogFromJiraIssue $issueWithWork.Item $worklog.id
}


