param(
    [switch]$Authenticate, [switch]$DryRun, [switch]$FunctionsOnly, [switch]$Background,
    [string]$ExpectedAccount = $env:TODO_SYNC_ACCOUNT,
    [string]$SourceTimeZone = 'UTC',
    [string]$CalendarTimeZone = 'UTC',
    [string]$CalendarName = 'To Do Sync',
    [string]$TenantId = 'consumers',
    [string]$StateDirectory = (Join-Path $HOME '.todo-calendar-sync')
)
$ErrorActionPreference='Stop'

$stateFile=Join-Path $StateDirectory "sync.json"
$tagId='String {7CAEA90D-2B90-4AE1-9305-F6376CC89C61} Name LocalTodoSync'
function Save-State {
    $state | ConvertTo-Json -Depth 20 | Set-Content "$stateFile.tmp" -Encoding utf8
    Move-Item "$stateFile.tmp" $stateFile -Force
}
function Graph($method,$uri,$body=$null) {
    if (!$uri.StartsWith('https://graph.microsoft.com/v1.0/')) { throw 'Unexpected Graph URL' }
    if ($null -eq $body) { return Invoke-MgGraphRequest -Method $method -Uri $uri -OutputType Hashtable }
    Invoke-MgGraphRequest -Method $method -Uri $uri -Body ($body|ConvertTo-Json -Depth 20 -Compress) -ContentType 'application/json' -OutputType Hashtable
}
function All($uri) {
    $items=[Collections.Generic.List[object]]::new()
    do {
        $page=Graph GET $uri
        if (!$page.ContainsKey('value')) { throw 'Incomplete collection response' }
        foreach ($item in $page.value) {$items.Add($item)}
        $uri=$page['@odata.nextLink']
    } while ($uri)
    return ,$items.ToArray()
}
function Task-Key($list,$task) {
    [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes("$list`n$task"))).ToLower()
}
function Date-Text($value) {
    if ($value -is [DateTime] -or $value -is [DateTimeOffset]) {return $value.ToString('yyyy-MM-dd')}
    return ([string]$value).Substring(0,10)
}
function Due-Day($due) {
    # Interpret normalized due dates in the configured source time zone.
    $raw=$due.dateTime
    $wall=if($raw -is [DateTime]){[DateTime]::SpecifyKind($raw,[DateTimeKind]::Unspecified)}else{[DateTime]::SpecifyKind([DateTime]::Parse([string]$raw,[Globalization.CultureInfo]::InvariantCulture),[DateTimeKind]::Unspecified)}
    $from=[TimeZoneInfo]::FindSystemTimeZoneById([string]$due.timeZone)
    $display=[TimeZoneInfo]::FindSystemTimeZoneById($SourceTimeZone)
    return [TimeZoneInfo]::ConvertTime($wall,$from,$display).ToString('yyyy-MM-dd')
}
function Payload($task) {
    $day=[DateTime]::ParseExact((Due-Day $task.dueDateTime),'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
    return @{
        subject=$task.title; isAllDay=$true; showAs='free'; isReminderOn=$false; sensitivity='private'
        start=@{dateTime=$day.ToString('yyyy-MM-ddT00:00:00');timeZone=$CalendarTimeZone}
        end=@{dateTime=$day.AddDays(1).ToString('yyyy-MM-ddT00:00:00');timeZone=$CalendarTimeZone}
    }
}
function Matches($event,$payload) {
    return ($event.subject -ceq $payload.subject -and $event.isAllDay -eq $true -and $event.showAs -eq 'free' -and $event.isReminderOn -eq $false -and $event.sensitivity -eq 'private' -and (Date-Text $event.start.dateTime) -eq (Date-Text $payload.start.dateTime) -and (Date-Text $event.end.dateTime) -eq (Date-Text $payload.end.dateTime))
}
function Sync-Once {
    $desired=@{}
    $lists=All 'https://graph.microsoft.com/v1.0/me/todo/lists?$top=100'
    foreach ($list in $lists) {
        $lid=[Uri]::EscapeDataString($list.id)
        $tasks=All "https://graph.microsoft.com/v1.0/me/todo/lists/$lid/tasks?`$top=100"
        foreach ($task in $tasks) {
            if ($task.status -ne 'completed' -and $task.dueDateTime -and $task.dueDateTime.dateTime) {
                $key=Task-Key $list.id $task.id
                $desired[$key]=Payload $task
            }
        }
    }
    # All source pages must succeed before any calendar mutation.
    $calendars=All 'https://graph.microsoft.com/v1.0/me/calendars?$top=100'
    if (!$state.calendarId) {
        if ($DryRun) { Write-Output "Dry run: $($desired.Count) dated unfinished tasks; calendar not created"; return }
        $existing=@($calendars|Where-Object {$_.name -eq $CalendarName})
        if ($existing.Count -gt 0) {throw 'Calendar with automatic-sync name already exists; explicit recovery required'}
        $calendar=Graph POST 'https://graph.microsoft.com/v1.0/me/calendars' @{name=$CalendarName}
        $state.calendarId=$calendar.id; Save-State
    }
    if (@($calendars|Where-Object {$_.id -eq $state.calendarId}).Count -eq 0 -and !$calendar) {throw 'Sync calendar missing; refusing to recreate automatically'}
    $cid=[Uri]::EscapeDataString($state.calendarId)
    $base="https://graph.microsoft.com/v1.0/me/calendars/$cid/events"
    $expand=[Uri]::EscapeDataString("singleValueExtendedProperties(`$filter=id eq '$tagId')")
    $events=All "$base`?`$top=100&`$expand=$expand"
    $owned=@{}
    foreach ($event in $events) {
        $tags=@($event.singleValueExtendedProperties | Where-Object {$_.id -eq $tagId})
        if ($tags.Count -eq 1 -and $tags[0].value -match '^[a-f0-9]{64}$') {
            $key=$tags[0].value
            if ($owned.ContainsKey($key)) {throw 'Duplicate owned events found; refusing mutations'}
            $owned[$key]=$event
        }
    }
    $created=0;$updated=0;$removed=0
    foreach ($key in $desired.Keys) {
        $payload=$desired[$key]
        if ($owned.ContainsKey($key)) {
            if (!(Matches $owned[$key] $payload)) {
                if (!$DryRun) {Graph PATCH "$base/$([Uri]::EscapeDataString($owned[$key].id))" $payload | Out-Null};$updated++
            }
            $state.pending.Remove($key)
        } else {
            if (!$DryRun) {
                if (!$state.pending.ContainsKey($key)) {$state.pending[$key]=[Guid]::NewGuid().ToString();Save-State}
                $payload.transactionId=$state.pending[$key]
                $payload.singleValueExtendedProperties=@(@{id=$tagId;value=$key})
                Graph POST $base $payload | Out-Null
                $state.pending.Remove($key);Save-State
            };$created++
        }
    }
    foreach ($key in $owned.Keys) {
        if (!$desired.ContainsKey($key)) {
            if (!$DryRun) {Graph DELETE "$base/$([Uri]::EscapeDataString($owned[$key].id))" | Out-Null};$removed++
        }
    }
    if (!$DryRun) {$state.lastSuccess=[DateTimeOffset]::UtcNow.ToString('o');Save-State}
    Write-Output "Sync: desired=$($desired.Count) created=$created updated=$updated removed=$removed dryRun=$DryRun"
}
if ($FunctionsOnly) { return }
if (-not $ExpectedAccount) { throw 'Set TODO_SYNC_ACCOUNT or pass -ExpectedAccount before syncing.' }
Import-Module Microsoft.Graph.Authentication
New-Item -ItemType Directory -Path $StateDirectory -Force | Out-Null
$state=if(Test-Path $stateFile){Get-Content $stateFile -Raw|ConvertFrom-Json -AsHashtable}else{@{calendarId=$null;pending=@{};lastSuccess=$null}}
$lock=[IO.File]::Open((Join-Path $StateDirectory "sync.lock"),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
try {
    if ($Authenticate) {
        Connect-MgGraph -TenantId $TenantId -Scopes Tasks.Read,Calendars.ReadWrite -ContextScope CurrentUser -UseDeviceAuthentication -ClientTimeout 900 -NoWelcome
    } elseif (!(Get-MgContext)) {
        # Device mode reuses the encrypted cache without opening a browser. If
        # access has been revoked it times out, leaving the calendar untouched.
        if ($Background) {
            Connect-MgGraph -TenantId $TenantId -Scopes Tasks.Read,Calendars.ReadWrite -ContextScope CurrentUser -UseDeviceAuthentication -NoWelcome
        } else {
            Connect-MgGraph -TenantId $TenantId -Scopes Tasks.Read,Calendars.ReadWrite -ContextScope CurrentUser -NoWelcome
        }
    }
    # Verify the mailbox belongs to the already verified personal account, not the school.
    $mailboxCalendars=All 'https://graph.microsoft.com/v1.0/me/calendars?$top=100'
    if (@($mailboxCalendars|Where-Object {$_.owner.address -eq $ExpectedAccount}).Count -eq 0) {throw 'Unexpected personal mailbox; refusing sync'}
    Sync-Once
} finally {$lock.Dispose()}
