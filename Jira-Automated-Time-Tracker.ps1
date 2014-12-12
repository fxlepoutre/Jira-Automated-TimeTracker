Param(
  [string]$cred,
  [string]$date,
  [string]$post,
  [string]$debug
)

#USAGE: powershell -ExecutionPolicy ByPass -File Jira-Automated-Time-Tracker.ps1 -cred E:\Path-to-Credentials-File\pwd_username.txt -date 24/10/14 -post true

#############################
##### GLOBAL PARAMETERS #####
#############################

$jiraURL = "http://jira.sophieparis.com"
$secondsPerDay = 8*3600
$modulo = 360
$cocu = "ITHD-8454"
$coefcocumax = 12

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

# Writes debug data to main screen
Function Debug($object, $title) {
    if ($debug -eq "true") {
        Write-Host "===================================================================="
        Write-Host "=== Debug data ===" $title
        Write-Host "===================================================================="
        $object | Format-Table -AutoSize
        Write-Host "===================================================================="
        Write-Host " "
    }
}

#############################
#####    MAIN SCRIPT    #####
#############################

# Read credentials.
if (!$cred) {
	$username = Read-Host 'Jira username'
	$password = Read-Host 'Jira password' -AsSecureString
	$password | ConvertFrom-SecureString | Out-File ".\pwd_$username.txt"
} else {
	$username = $cred | select-string -pattern ".*pwd_(.*).txt" | %{ $_.Matches[0].Groups[1].Value }
	if (test-path $cred) {
		# Read the password from file.
		$password = Get-Content $cred | ConvertTo-SecureString
	} else {
        Write-Host "File """ $cred """ not found."
        Break
    }
}

# Create authentication header.
$passwordDecoded = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username,$passwordDecoded)))

# Read input date and set timestamp based on this.
if ($date) {
	$worklogDate = Get-Date $date
} else {
	if (!$cred) {
		$worklogDateInput = Read-Host 'Input date (dd/mm/yy) or leave blank for today'
		if ($worklogDateInput) {
			$worklogDate = Get-Date $worklogDateInput
		}
        else
                {
          # If blank, default assign it to today.
          $worklogDate = Get-Date
        }
	}
}

# If Date not set, default assign it to yesterday.
if(!$worklogDate) {
  $worklogDate = Get-Date
  $worklogDate = $worklogDate.AddDays(-1)
}

$worklogDateStr = ""+$worklogDate.Day+"/"+$worklogDate.Month+"/"+$worklogDate.Year
$worklogDate = Get-Date $worklogDateStr

$dateStart = Get-Date $worklogDate -UFormat %s
$dateEnd = (Get-Date $worklogDate.AddDays(1) -UFormat %s) - 1

# Try login and read XML feed of activities..
try {
    # Read XML feed of activities.
    $uri = $Script:jiraURL+"/activity?streams=user+IS+"+$username+"&streams=update-date+BETWEEN+"+$dateStart+"000+"+$dateEnd+"999&maxResults=500"
    Write-Host $uri
    [xml]$activityStream = Invoke-WebRequest -Uri $uri -Method Get -SessionVariable session -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -TimeoutSec 10
} catch [Exception] {
    write-Host "Login error, check if login on Jira is possible."
    write-Host $_.Exception|format-list -force
    Break
}

# Parse XML feed of activities and get interresting details.
$activityEntries = @()
ForEach ($entry in $activityStream.feed.entry) { 
    # Initialize number of attachments
    $nbfiles = 0
    $issueKey = ""

    # Find issue related to activity entry in target first
    ForEach ($target in $entry.target) {
    	if($target.'object-type' -eq "http://streams.atlassian.com/syndication/types/issue") {
    		$issueKey = $entry.target.title.InnerText
    		break
    	}
    }
    
    # Looks at the activity object for issues or files
	  ForEach ($object in $entry.object) { 	
	    # Find issue in object if no issue found yet
    	if((!$issueKey) -or ($issueKey.Trim() -eq "")) {
	    	if(($object.'object-type' -eq "http://streams.atlassian.com/syndication/types/issue")) {
	    		# Check if entry is only a worklog entry. If so, don't include it.
                $excludedStringInAction = "Logged"
                $concat = " " + $entry.title.innerText + $entry.content.innerText
                if(-not($concat.ToLower() -match $excludedStringInAction)) {
                    $issueKey = $entry.object.title.innerText
                }
	    	}
	    }
    	# Count the number of attachments
    	if($object.'object-type' -eq "http://activitystrea.ms/schema/1.0/file") {
    		$nbfiles = $nbfiles + 1
    	}
    }

    # Ignore issue if no key found
    if((!$issueKey) -or ($issueKey.Trim() -eq "")) {
    	continue
    }
    
    # Find action that generated activity entry
    $action = $entry.category.term
    if(!$action -or ($action.Trim() -eq "")) {
        $action = $entry.verb.SubString($entry.verb.Length - 6, 6)
    }
    
    # Assign coefficient to defined action
    $coef = 0
    switch($action) {
        "closed" {$coef = 2; break}
        "created" {$coef = 5; break}
        "update" {$coef = 1; break}
        "comment" {$coef = 2; break}
        "started" {$coef = 5; break}
        "reopened" {$coef = 4; break}
        "resolved" {$coef = 5; break}
        "0/post" {$coef = 2; break}
        default {$coef = 0}
    }
    
    # Add weight according to nb of attachments
    if ($nbfiles > 0) {
    	$coef = $coef + 1 * $nbfiles
    }
	
    # Get date of the activity
    $datePublished = $entry.published

    # Load all results into the datastore of activity entries.
    $activityEntries += [pscustomobject]@{issueKey=$issueKey;action=$action;coef=$coef;date=$datePublished}
}

# Always add entry to this "cocu" worklog.
$coefcocu = [math]::Max(($coefcocumax - ($activityEntries | Measure-Object coef -Sum).Sum), 1)
$activityEntries += [pscustomobject]@{issueKey=$cocu;action="misc. work";coef=$coefcocu;date=$worklogDate}

Debug $activityEntries "All activity entries"

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
$issuesWithWork | Add-Member -NotePropertyName Assignee -NotePropertyValue $username
ForEach ($issueWithWork in $($issuesWithWork)) {
    $issueDetails = GetDetailsOfJiraIssue $issueWithWork.Item
    $issueWithWork.Status = $issueDetails.fields.status.name
    $issueWithWork.Assignee = $issueDetails.fields.assignee.name
    if ($issueWithWork.Assignee -eq $username) {
    	$issueWithWork.Sum = $issueWithWork.Sum + 10
    }
}

Debug $issuesWithWork "All issues with work"

# Remove entries that are closed
$issuesWithWork = $issuesWithWork | where {$_.Status -ne "Closed"}

# Allocate time according to coef
$totalCoeff = ($issuesWithWork | Measure-Object Sum -Sum).Sum
$timeRemaining = $secondsPerDay
$issuesWithWork | Add-Member -NotePropertyName TimeSpent -NotePropertyValue 0
ForEach ($issueWithWork in $($issuesWithWork)) {
    $timeSpentInSec = $issueWithWork.Sum / $totalCoeff * $secondsPerDay
    $timeSpentInSec = $timeSpentInSec - ($timeSpentInSec % $modulo)
    $timeRemaining = $timeRemaining - $timeSpentInSec
    $issueWithWork.TimeSpent = $timeSpentInSec
}

Debug $issuesWithWork "All issues to be updated"

# Record the worklog
ForEach ($issueWithWork in $issuesWithWork) {
    if ($issueWithWork.Item -eq $cocu) {
    	$issueWithWork.TimeSpent = $issueWithWork.TimeSpent + $timeRemaining
    }
    Write-Host "Adding worklog on" $issueWithWork.Item "-" $issueWithWork.Assignee "-" $issueWithWork.TimeSpent "seconds" "-" $issueWithWork.Date
    if (($post -eq "true") -and ($debug -ne "true")) {
    	if ($issueWithWork.Sum -ne 0) {
    		$worklog = AddWorklogToJiraIssue $issueWithWork.Item $issueWithWork.TimeSpent $issueWithWork.Date
    	}
    }
}

# Confirm saving.
if (($post -eq "true") -and ($debug -ne "true")) {
    Write-Host "Done, changes saved to Jira."
} else {
    Write-Host "Changes not saved. Launch again with parameter ""-post true"" to save (or remove debug switch)."
}
