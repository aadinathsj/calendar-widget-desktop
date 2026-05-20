const { spawn } = require('child_process');
const path = require('path');

// PowerShell script to interact with Outlook COM
function executePowerShell(script) {
  return new Promise((resolve, reject) => {
    const ps = spawn('powershell.exe', [
      '-NoProfile',
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
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
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
      $calendar = $namespace.GetDefaultFolder(9)

      $startDate = [DateTime]::Parse("${startDateStr}")
      $endDate = [DateTime]::Parse("${endDateStr}")

      $appointments = $calendar.Items
      $appointments.Sort("[Start]")
      $appointments.IncludeRecurrences = $true

      $filter = "[Start] >= '$($startDate.ToString("g"))' AND [Start] <= '$($endDate.ToString("g"))'"
      $items = $appointments.Restrict($filter)

      $events = @()

      foreach ($item in $items) {
        $teamsLink = $null
        $location = $item.Location

        # Extract Teams join link - supports both URL formats:
        #   Old: https://teams.microsoft.com/l/meetup-join/...
        #   New: https://teams.microsoft.com/meet/...
        if ($item.Body -match 'https://teams\\.microsoft\\.com/(?:l/meetup-join|meet)/[^\\s<>"]+') {
          $teamsLink = $matches[0]
        }

        # Fallback: check HTMLBody - some received invites only have the link in HTML
        if (-not $teamsLink -and $item.HTMLBody) {
          if ($item.HTMLBody -match 'https://teams\\.microsoft\\.com/(?:l/meetup-join|meet)/[^"\\s]+') {
            $teamsLink = $matches[0]
          }
        }

        # Fallback: OnlineMeetingConfLink (only populated for meetings you organised)
        if (-not $teamsLink -and $item.IsOnlineMeeting) {
          try {
            if ($item.OnlineMeetingConfLink) {
              $teamsLink = $item.OnlineMeetingConfLink
            }
          } catch {}
        }

        $organizer = "Unknown"
        try {
          if ($item.Organizer) {
            $organizer = $item.Organizer
          }
        } catch {}

        $eventObj = @{
          id = $item.EntryID
          subject = $item.Subject
          start = $item.Start.ToString("o")
          end = $item.End.ToString("o")
          location = $location
          teamsLink = $teamsLink
          organizer = $organizer
          body = $item.Body
          isAllDay = $item.AllDayEvent
        }

        $events += $eventObj
      }

      # Convert to JSON
      $json = $events | ConvertTo-Json -Depth 3 -Compress
      Write-Output $json

      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()

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

module.exports = {
  checkOutlookAvailable,
  getCalendarEvents
};
