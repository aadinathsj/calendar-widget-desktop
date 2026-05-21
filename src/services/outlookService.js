const { spawn } = require('child_process');
const path = require('path');

// PowerShell script to interact with Outlook COM
function executePowerShell(script) {
  return new Promise((resolve, reject) => {
    const ps = spawn('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-Command', script
    ]);

    let stdout = '';
    let stderr = '';

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(stderr || `PowerShell exited with code ${code}`));
      } else {
        resolve(stdout);
      }
    });
  });
}

async function checkOutlookAvailable() {
  const script = `
    try {
      $outlook = New-Object -ComObject Outlook.Application
      $namespace = $outlook.GetNamespace("MAPI")
      $folder = $namespace.GetDefaultFolder(9) # 9 = Calendar folder
      Write-Output "available"
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    } catch {
      Write-Error $_.Exception.Message
      exit 1
    }
  `;

  try {
    await executePowerShell(script);
    return true;
  } catch (error) {
    console.error('Outlook check failed:', error);
    return false;
  }
}

async function getCalendarEvents(startDate, endDate) {
  const startDateStr = startDate.toISOString();
  const endDateStr = endDate.toISOString();

  const script = `
    try {
      $outlook = New-Object -ComObject Outlook.Application
      $namespace = $outlook.GetNamespace("MAPI")

      $startDate = [DateTime]::Parse("${startDateStr}")
      $endDate = [DateTime]::Parse("${endDateStr}")

      $accountColors = @('#6366f1', '#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#06b6d4')
      $allEvents = [System.Collections.Generic.List[object]]::new()
      $accountIndex = 0

      foreach ($store in $namespace.Stores) {
        try {
          $calFolder = $store.GetDefaultFolder(9)
          $accountName = $store.DisplayName
          $accountColor = $accountColors[$accountIndex % $accountColors.Count]

          $appointments = $calFolder.Items

          # CRITICAL: IncludeRecurrences MUST be set BEFORE Sort() and Restrict()
          $appointments.IncludeRecurrences = $true
          $appointments.Sort("[Start]")

          # Get ALL items - don't filter by response status or meeting type
          $filter = "[Start] >= '$($startDate.ToString("g"))' AND [Start] <= '$($endDate.ToString("g"))'"
          $items = $appointments.Restrict($filter)

          Write-Host "Store: $accountName - Calendar: $($calFolder.Name) - Total items before filter: $($appointments.Count) - After filter: $($items.Count)"

          foreach ($item in $items) {
            # Include all meeting types: olNonMeeting(0), olMeeting(1), olMeetingReceived(3)
            # Skip only canceled meetings: olMeetingCanceled(5)
            $meetingStatus = 0
            try { $meetingStatus = $item.MeetingStatus } catch {}

            if ($meetingStatus -eq 5) {
              # Skip canceled meetings
              continue
            }

            $teamsLink = $null
            $location = $item.Location

            if ($item.Body -match 'https://teams\\.microsoft\\.com/(?:l/meetup-join|meet)/[^\\s<>"]+') {
              $teamsLink = $matches[0]
            }
            if (-not $teamsLink -and $item.HTMLBody) {
              if ($item.HTMLBody -match 'https://teams\\.microsoft\\.com/(?:l/meetup-join|meet)/[^"\\s]+') {
                $teamsLink = $matches[0]
              }
            }
            if (-not $teamsLink -and $item.IsOnlineMeeting) {
              try {
                if ($item.OnlineMeetingConfLink) { $teamsLink = $item.OnlineMeetingConfLink }
              } catch {}
            }

            $organizer = "Unknown"
            try { if ($item.Organizer) { $organizer = $item.Organizer } } catch {}

            # Get meeting status and response status for debugging
            $meetingStatus = 0
            $responseStatus = 0
            try { $meetingStatus = $item.MeetingStatus } catch {}
            try { $responseStatus = $item.ResponseStatus } catch {}

            $eventObj = [PSCustomObject]@{
              id             = $item.EntryID
              subject        = $item.Subject
              start          = $item.Start.ToString("o")
              end            = $item.End.ToString("o")
              location       = $location
              teamsLink      = $teamsLink
              organizer      = $organizer
              body           = $item.Body
              isAllDay       = $item.AllDayEvent
              accountName    = $accountName
              accountIndex   = $accountIndex
              accountColor   = $accountColor
              meetingStatus  = $meetingStatus
              responseStatus = $responseStatus
            }

            Write-Host "  Event: $($item.Subject) | Start: $($item.Start) | MeetingStatus: $meetingStatus | Response: $responseStatus"

            $allEvents.Add($eventObj)
          }
        } catch {
          # Store has no accessible calendar folder, skip it
        }
        $accountIndex++
      }

      # Sort by start time; deduplicate by EntryID (shared calendars can duplicate)
      $seen = [System.Collections.Generic.HashSet[string]]::new()
      $sorted = $allEvents |
        Sort-Object { [DateTime]::Parse($_.start) } |
        Where-Object { $seen.Add($_.id) }

      $json = $sorted | ConvertTo-Json -Depth 3 -Compress
      Write-Output $json

      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

    } catch {
      Write-Error $_.Exception.Message
      exit 1
    }
  `;

  try {
    const result = await executePowerShell(script);

    // Parse the JSON output
    const lines = result.trim().split('\n');
    const jsonLine = lines[lines.length - 1];

    if (!jsonLine || jsonLine.trim() === '') {
      return [];
    }

    let events = JSON.parse(jsonLine);

    // Handle case where single event is returned as object instead of array
    if (!Array.isArray(events)) {
      events = [events];
    }

    return events;
  } catch (error) {
    console.error('Error fetching calendar events:', error);
    throw error;
  }
}

// Combined check + fetch in a single PowerShell process — saves one full PS startup (~2 s).
// Always resolves (never rejects): returns { available: false, events: [] } when Outlook
// is not running so the caller can handle it gracefully.
async function checkAndGetEvents(startDate, endDate) {
  const startDateStr = startDate.toISOString();
  const endDateStr   = endDate.toISOString();

  const script = `
    try {
      $outlook    = New-Object -ComObject Outlook.Application
      $namespace  = $outlook.GetNamespace("MAPI")
      $startDate  = [DateTime]::Parse("${startDateStr}")
      $endDate    = [DateTime]::Parse("${endDateStr}")

      $accountColors = @('#6366f1','#3b82f6','#10b981','#f59e0b','#ef4444','#06b6d4')
      $allEvents     = [System.Collections.Generic.List[object]]::new()
      $accountIndex  = 0

      foreach ($store in $namespace.Stores) {
        try {
          $calFolder   = $store.GetDefaultFolder(9)
          $accountName = $store.DisplayName
          $accountColor = $accountColors[$accountIndex % $accountColors.Count]

          $appointments = $calFolder.Items

          # CRITICAL: IncludeRecurrences MUST be set BEFORE Sort() and Restrict()
          $appointments.IncludeRecurrences = $true
          $appointments.Sort("[Start]")

          $filter = "[Start] >= '$($startDate.ToString("g"))' AND [Start] <= '$($endDate.ToString("g"))'"
          $items  = $appointments.Restrict($filter)

          foreach ($item in $items) {
            # Skip only canceled meetings (MeetingStatus = 5)
            $meetingStatus = 0
            try { $meetingStatus = $item.MeetingStatus } catch {}
            if ($meetingStatus -eq 5) { continue }

            $teamsLink = $null
            $location  = $item.Location

            if ($item.Body -match 'https://teams\\.microsoft\\.com/(?:l/meetup-join|meet)/[^\\s<>"]+') {
              $teamsLink = $matches[0]
            }
            if (-not $teamsLink -and $item.HTMLBody) {
              if ($item.HTMLBody -match 'https://teams\\.microsoft\\.com/(?:l/meetup-join|meet)/[^"\\s]+') {
                $teamsLink = $matches[0]
              }
            }
            if (-not $teamsLink -and $item.IsOnlineMeeting) {
              try { if ($item.OnlineMeetingConfLink) { $teamsLink = $item.OnlineMeetingConfLink } } catch {}
            }

            $organizer = "Unknown"
            try { if ($item.Organizer) { $organizer = $item.Organizer } } catch {}

            $allEvents.Add([PSCustomObject]@{
              id           = $item.EntryID
              subject      = $item.Subject
              start        = $item.Start.ToString("o")
              end          = $item.End.ToString("o")
              location     = $location
              teamsLink    = $teamsLink
              organizer    = $organizer
              isAllDay     = $item.AllDayEvent
              accountName  = $accountName
              accountIndex = $accountIndex
              accountColor = $accountColor
            })
          }
        } catch {}
        $accountIndex++
      }

      $seen   = [System.Collections.Generic.HashSet[string]]::new()
      $sorted = @($allEvents | Sort-Object { [DateTime]::Parse($_.start) } | Where-Object { $seen.Add($_.id) })

      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

      $eventsJson = if ($sorted.Count -eq 0)  { "[]" }
                    elseif ($sorted.Count -eq 1) { "[" + ($sorted | ConvertTo-Json -Depth 4 -Compress) + "]" }
                    else                         { $sorted | ConvertTo-Json -Depth 4 -Compress }

      Write-Output ('{"available":true,"events":' + $eventsJson + '}')
    } catch {
      Write-Output '{"available":false,"events":[]}'
    }
  `;

  try {
    const output = await executePowerShell(script);
    const lines  = output.trim().split('\n');
    const parsed = JSON.parse(lines[lines.length - 1].trim());

    if (!Array.isArray(parsed.events)) {
      parsed.events = parsed.events ? [parsed.events] : [];
    }
    return { available: parsed.available === true, events: parsed.events };
  } catch (error) {
    console.error('checkAndGetEvents failed:', error);
    return { available: false, events: [] };
  }
}

module.exports = {
  checkOutlookAvailable,
  getCalendarEvents,
  checkAndGetEvents
};
