# Prerequisites
- [PowerShell](https://github.com/PowerShell/PowerShell/releases) - `winget install Microsoft.PowerShell`
- [GnuWin32 Make](https://gnuwin32.sourceforge.net/packages/make.htm) - `winget install GnuWin32.Make`

All other dependencies (Python, uv, fnm, Node.js, Appium, WinAppDriver) are installed
automatically by the Makefile.

# Setup
Enable Developer Mode in Windows Settings.

Open PowerShell and set the current user's execution
policy to RemoteSigned:

    $ Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

Then bootstrap the project:

    $ & 'C:\Program Files (x86)\GnuWin32\bin\make.exe' bootstrap

This runs `install-os-deps` (installs Python 3.13, uv, fnm, Node.js 22, and WinAppDriver
via winget/msiexec) followed by `install-deps` (installs Python packages, npm packages,
and the Appium Windows driver).

Create a `.env` file in the project root with the required credentials:

    NIELSEN_USERNAME=your_username
    NIELSEN_PASSWORD=your_password

For the `gdrive-weekly` command you also need:

    GDRIVE_SERVICE_ACCOUNT_CREDENTIALS=path/to/service-account.json
    GOOGLE_DRIVE_IN_FOLDER_ID=<folder-id>
    GOOGLE_DRIVE_OUT_FOLDER_ID=<folder-id>
    GOOGLE_DRIVE_DRIVE_ID=<drive-id>          # optional, for shared drives

# Run

## `files` — Refresh local Excel files

    $ uv run python main.py files path\to\file1.xlsx path\to\file2.xlsx

Opens each Excel file via Appium, logs into NielsenIQ Discover, refreshes all
data ranges, saves the file, and renames it to include the Nielsen period
dates (e.g. `Report - 030126 to 030826.xlsx`).

## `gdrive-weekly` — Refresh files from Google Drive

    $ uv run python main.py gdrive-weekly

Downloads every Excel file from the Google Drive input folder, refreshes them
the same way as the `files` command, and uploads the results to a dated
subfolder in the output folder on Google Drive.

# How it works

## Automation stack

The tool drives Excel through the Windows desktop UI using
[Appium](https://appium.io/) with the
[Windows Application Driver (WinAppDriver)](https://github.com/microsoft/WinAppDriver).
At startup, `main.py` launches both WinAppDriver (port 4724) and the Appium
server (port 10000), then creates a remote WebDriver session for each Excel
file.

## NielsenIQ login and data refresh

After opening a file, the automation clicks the **NielsenIQ Discover** add-in
button on the Insert ribbon, selects the **US** region, and either logs in
(username + password) or skips login if the session is still active. It then
clicks **Refresh All** to update every data range and waits up to 10 minutes
for the refresh to finish.

## Period date validation and file renaming

Once a file is refreshed, the tool reads the `Periods` column from the
workbook to determine the minimum and maximum week-ending dates. It verifies
that the latest period matches the expected Nielsen week ending (the Saturday
one week before the most recent Saturday). The output file is renamed to
include these dates — for example, `Report.xlsx` becomes
`Report - 030126 to 030826.xlsx`. If the `Periods` column is missing, the
expected week ending date is used as a fallback.

## Error handling

- **Data plan limit reached** — If NielsenIQ reports that the data plan limit
  has been reached, the run stops immediately with a `DataPlanLimitReachedError`.
- **Failure limit** — If 2 or more files fail to refresh, the run stops with a
  `FailureLimitExceededError` rather than continuing to fail.
- **Week ending date mismatch** — If the period dates in the refreshed file do
  not match the expected week ending, a `WeekEndingDateError` is raised.
- On any error, a screenshot is taken and (for `gdrive-weekly`) uploaded to a
  `logs` subfolder on Google Drive alongside the main log file.

# Development

## Inspecting UI elements

The automation locates buttons, inputs, and other controls by their UI
Automation properties (Name, AutomationId, ClassName, etc.). When adding or
updating interactions you need to inspect the live UI to find the right
locators. Useful tools:

- **[Accessibility Insights for Windows](https://accessibilityinsights.io/downloads/)** —
  Microsoft's free inspection tool. Use the **Live Inspect** mode to hover over
  any UI element and see its full set of UI Automation properties (Name,
  AutomationId, ClassName, ControlType, etc.). The **UI Automation tree** view
  shows the element hierarchy, which is helpful for building XPath selectors.
- **[Inspect.exe](https://learn.microsoft.com/en-us/windows/win32/winauto/inspect-objects)** —
  Ships with the Windows SDK. Similar to Accessibility Insights but more
  lightweight. Select an element and the properties pane shows its
  UIA properties. Install via `winget install Microsoft.WindowsSDK` (the tool
  is found under `C:\Program Files (x86)\Windows Kits\10\bin\<version>\x64\inspect.exe`).
- **[FlaUI Inspect](https://github.com/FlaUI/FlaUInspect)** — An open-source
  alternative with a tree view of the UI Automation element hierarchy. Download
  from the GitHub releases page.

## Locator strategies

The code uses several Appium locator strategies. In order of preference:

1. **`AppiumBy.NAME`** — Matches the element's `Name` property
   (e.g. `"Got it"`, `"Refresh All"`). Simplest and most readable.
2. **`AppiumBy.ACCESSIBILITY_ID`** — Matches the `AutomationId` property
   (e.g. `"input27"`, `"FileSave"`). More stable than Name when the display
   text might change due to localisation or UI updates.
3. **XPath** — Use when Name or AutomationId are not unique enough, e.g.
   `'//Button[@ClassName="NetUIRibbonButton"][@Name="NielsenIQ Discover"]'`.
   XPath selectors are slower and more brittle, so prefer the simpler
   strategies when possible.

Use an inspection tool to identify which property is most stable for the
element you need to target.

## Tips for making changes

- **Run the `files` command on a single test file first** to verify your
  changes before running against a full batch or Google Drive.
- **Use `wait_until_visible`** with an appropriate `WebDriverWait` timeout
  rather than fixed `time.sleep` calls. The UI can take variable amounts of
  time to respond.
- **Check for competing UI states** — Some steps need to handle multiple
  possible outcomes (e.g. after clicking NielsenIQ Discover, the add-in may
  show a login form, a "Got it" button, or a "Data plan limit reached"
  message). See `_update_sales` for an example of polling for multiple elements.
- **Screenshots on failure** — `_take_screenshot` saves a screenshot to the
  logs directory. When debugging a new interaction, you can add temporary
  screenshot calls to capture the UI state at specific points.
- **Appium server logs** — The Appium server output is logged under the
  `appium_server` logger. Check these logs when session creation or element
  lookup fails unexpectedly.
