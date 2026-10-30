# Runs the real Excel what-if tools on output/PhanB_Dap_an_tham_khao.xlsx via COM:
#   Goal Seek, Data Table, Scenario Manager (+ summaries), Solver (+ Answer/Sensitivity reports).
# Configuration comes from output/excel_tools.json, written by model/build_model.py.
#
# Usage (from the project root, after build_model.py):
#   powershell -NoProfile -ExecutionPolicy Bypass -File model/apply_excel_tools.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xlsxPath = Join-Path $root 'output\PhanB_Dap_an_tham_khao.xlsx'
$tools = Get-Content (Join-Path $root 'output\excel_tools.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$missing = [Type]::Missing

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

# Leading comma stops PowerShell from unrolling a multi-cell Range into single cells
function Nm([string]$name) { , $script:wb.Names.Item($name).RefersToRange }
function Addr([string]$name) { (Nm $name).Address($true, $true) }

# Multi-area range from named ranges on one sheet ($xl.Union does not bind from PowerShell)
function Union-Names([string[]]$names) {
    $sheet = (Nm $names[0]).Worksheet
    $address = ($names | ForEach-Object { Addr $_ }) -join ','
    , $sheet.Range($address)
}

# JSON numbers arrive as Decimal, which COM cannot marshal into a Variant array
function To-Doubles($values) { , [object[]]@($values | ForEach-Object { [double]$_ }) }

function Rename-NewSheet([string]$pattern, [string]$newName) {
    foreach ($sh in $script:wb.Worksheets) {
        if ($sh.Name -like $pattern) {
            $sh.Name = $newName
            $sh.Move($missing, $script:wb.Worksheets.Item($script:wb.Worksheets.Count))
            return
        }
    }
    throw "No sheet matching '$pattern' to rename to $newName"
}

function Invoke-Solver($model, [object[]]$reports) {
    $script:wb.Worksheets.Item($model.sheet).Activate()
    $xl.Run('Solver.xlam!SolverReset') | Out-Null
    $xl.Run('Solver.xlam!SolverOk', (Addr $model.objective), [int]$model.sense, 0, (Addr $model.vars), [int]$model.engine) | Out-Null
    foreach ($c in $model.constraints) {
        $xl.Run('Solver.xlam!SolverAdd', (Addr $c[0]), [int]$c[1], ('=' + (Addr $c[2]))) | Out-Null
    }
    $code = $xl.Run('Solver.xlam!SolverSolve', $true)
    if ($code -notin 0, 1, 2) { throw "Solver failed on $($model.sheet), code $code" }
    if ($reports) { $xl.Run('Solver.xlam!SolverFinish', 1, $reports) | Out-Null }
    else { $xl.Run('Solver.xlam!SolverFinish', 1) | Out-Null }
    $code
}

try {
    $script:wb = $xl.Workbooks.Open($xlsxPath)
    $xl.Calculation = -4105  # xlCalculationAutomatic
    $xl.CalculateFull()

    # --- Solver add-in
    $solver = $xl.AddIns | Where-Object { $_.Name -like 'SOLVER*' } | Select-Object -First 1
    if (-not $solver) { throw 'Solver add-in not found' }
    $solver.Installed = $true
    $null = $xl.Workbooks.Open($solver.FullName)
    $xl.Run('Solver.xlam!Auto_open') | Out-Null
    $script:wb.Activate()

    # --- Goal Seek: record the answer, then restore the input cell
    foreach ($g in $tools.goal_seek) {
        $change = Nm $g.change
        $original = $change.Value2
        $ok = (Nm $g.set).GoalSeek([double]$g.target, $change)
        if (-not $ok) { throw "Goal Seek did not converge: $($g.result)" }
        (Nm $g.result).Value2 = $change.Value2
        $change.Value2 = $original
        "Goal Seek  {0,-22} = {1:N4}" -f $g.result, (Nm $g.result).Value2
    }

    # --- Data Tables
    foreach ($d in $tools.data_tables) {
        $row = if ($d.row_input) { Nm $d.row_input } else { $missing }
        $col = if ($d.col_input) { Nm $d.col_input } else { $missing }
        $null = (Nm $d.range).Table($row, $col)
        "Data Table $($d.range)"
    }

    # --- Input-only scenarios (Bai 1)
    foreach ($s in $tools.scenarios) {
        $ws = $script:wb.Worksheets.Item($s.sheet)
        $changing = Union-Names $s.changing
        foreach ($item in $s.items) {
            $null = $ws.Scenarios().Add($item.name, $changing, (To-Doubles $item.values), $item.comment)
        }
        $null = $ws.Scenarios().CreateSummary(1, (Union-Names $s.results))
        Rename-NewSheet 'Scenario Summary*' $s.summary_name
        $null = $ws.Scenarios().Item($s.base).Show()
        "Scenarios $($s.summary_name)"
    }

    # --- Solver models: solve each scenario, store it, then solve the base case with reports
    foreach ($m in $tools.solver) {
        $ws = $script:wb.Worksheets.Item($m.sheet)
        $sc = $m.scenario
        $paramNames = @(if ($sc.params) { $sc.params } else { $sc.param })
        $changing = Union-Names ($paramNames + $m.vars)
        foreach ($item in $sc.items) {
            $paramValues = @(if ($item.values) { $item.values } else { $item.value })
            for ($i = 0; $i -lt $paramNames.Count; $i++) { (Nm $paramNames[$i]).Value2 = [double]$paramValues[$i] }
            $null = Invoke-Solver $m $null
            $varValues = @((Nm $m.vars).Value2 | ForEach-Object { $_ })
            $null = $ws.Scenarios().Add($item.name, $changing, (To-Doubles ($paramValues + $varValues)), $sc.comment)
            "Solver $($m.sheet) / $($item.name): objective = {0:N2}" -f (Nm $m.objective).Value2
        }
        # Show the base case first so the summary's "Current Values" column is the base case
        $null = $ws.Scenarios().Item($sc.base).Show()
        $null = $ws.Scenarios().CreateSummary(1, (Union-Names $sc.results))
        Rename-NewSheet 'Scenario Summary*' $sc.summary_name

        $code = Invoke-Solver $m ([object[]](1, 2))
        Rename-NewSheet 'Answer Report*' ('AR_' + $m.report_suffix)
        Rename-NewSheet 'Sensitivity Report*' ('SR_' + $m.report_suffix)
        "Solver $($m.sheet) base: code $code, objective = {0:N2}" -f (Nm $m.objective).Value2
    }

    $xl.CalculateFull()
    $script:wb.Worksheets.Item('Huong_dan').Activate()
    $script:wb.Save()
    "Saved $xlsxPath"
}
finally {
    foreach ($b in @($xl.Workbooks)) { $b.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
