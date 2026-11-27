<#
.SYNOPSIS
  Find a ribbon button in a running Office window by its label, through UI
  Automation, and optionally click it. Used by the live-reload tests to prove a
  button is really on the ribbon Office built, not only in the file's XML.

.PARAMETER ProcessId  Office process whose main window holds the ribbon.
.PARAMETER Tab        Label of the ribbon tab to select first.
.PARAMETER Button     Label of the button to look for.
.PARAMETER Click      Invoke the button once found.

Prints FOUND / CLICKED, or MISSING: <what>; exit code 0 only when found
(and clicked, with -Click).
#>
param(
    [Parameter(Mandatory = $true)][int]$ProcessId,
    [Parameter(Mandatory = $true)][string]$Tab,
    [Parameter(Mandatory = $true)][string]$Button,
    [switch]$Click,
    [int]$TimeoutSeconds = 20
)

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$ae = [System.Windows.Automation.AutomationElement]
$tree = [System.Windows.Automation.TreeScope]

function Find-Named($root, [string]$name, $controlType) {
    $byName = New-Object System.Windows.Automation.PropertyCondition($ae::NameProperty, $name)
    $byType = New-Object System.Windows.Automation.PropertyCondition($ae::ControlTypeProperty, $controlType)
    $both = New-Object System.Windows.Automation.AndCondition($byName, $byType)
    return $root.FindFirst($tree::Descendants, $both)
}

function Wait-For([scriptblock]$probe) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $found = & $probe
        if ($found) { return $found }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    return $null
}

$byPid = New-Object System.Windows.Automation.PropertyCondition($ae::ProcessIdProperty, $ProcessId)

# Office can have several top-level windows (a dialog, a second document);
# look in each for the tab.
function Find-InWindows([string]$name, $controlType) {
    foreach ($w in $ae::RootElement.FindAll($tree::Children, $byPid)) {
        $hit = Find-Named $w $name $controlType
        if ($hit) { return @($w, $hit) }
    }
    return $null
}

function Describe-Windows {
    foreach ($w in $ae::RootElement.FindAll($tree::Children, $byPid)) {
        Write-Output ("  window [" + $w.Current.Name + "] class=" + $w.Current.ClassName)
        $isTab = New-Object System.Windows.Automation.PropertyCondition(
            $ae::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)
        $names = @($w.FindAll($tree::Descendants, $isTab) | ForEach-Object { $_.Current.Name })
        Write-Output ("    tabs: " + ($names -join " | "))
    }
}

$found = Wait-For { Find-InWindows $Tab ([System.Windows.Automation.ControlType]::TabItem) }
if (-not $found) { Write-Output "MISSING: tab $Tab"; Describe-Windows; exit 1 }
$window = $found[0]
$found[1].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()

$target = Wait-For { Find-Named $window $Button ([System.Windows.Automation.ControlType]::Button) }
if (-not $target) { Write-Output "MISSING: button $Button"; Describe-Windows; exit 1 }
Write-Output "FOUND"
if ($Click) {
    $target.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
    Write-Output "CLICKED"
}
exit 0
