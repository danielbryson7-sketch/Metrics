Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

$csvPath = "C:\tab3\Metrics\people.csv"

if (-not (Test-Path $csvPath)) {
    [System.Windows.Forms.MessageBox]::Show("people.csv not found at $csvPath")
    exit
}

# =========================
# Form
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Employee Matrix Editor"
$form.Size = New-Object System.Drawing.Size(800,520)
$form.StartPosition = "CenterScreen"

# =========================
# Buttons
# =========================
$saveBtn = New-Object System.Windows.Forms.Button
$saveBtn.Text = "Save"
$saveBtn.Location = New-Object System.Drawing.Point(10,10)

$addBtn = New-Object System.Windows.Forms.Button
$addBtn.Text = "Add Employee"
$addBtn.Location = New-Object System.Drawing.Point(100,10)

$removeBtn = New-Object System.Windows.Forms.Button
$removeBtn.Text = "Remove Selected"
$removeBtn.Location = New-Object System.Drawing.Point(230,10)

# =========================
# Grid
# =========================
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10,50)
$grid.Size = New-Object System.Drawing.Size(760,420)
$grid.AllowUserToAddRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoGenerateColumns = $false

# =========================
# Columns
# =========================
$nameCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$nameCol.HeaderText = "Name"
$nameCol.DataPropertyName = "Name"
$nameCol.Width = 260

$shiftCol = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$shiftCol.HeaderText = "Shift"
$shiftCol.DataPropertyName = "Shift"
$shiftCol.Items.AddRange("1st","2nd","3rd")
$shiftCol.Width = 100

$teamCol = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$teamCol.HeaderText = "Team"
$teamCol.DataPropertyName = "Team"
$teamCol.Items.AddRange("Team A","Team B")
$teamCol.Width = 120

$activeCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$activeCol.HeaderText = "Active"
$activeCol.DataPropertyName = "Active"
$activeCol.Width = 80

$grid.Columns.Add($nameCol)
$grid.Columns.Add($shiftCol)
$grid.Columns.Add($teamCol)
$grid.Columns.Add($activeCol)

# =========================
# DataTable (source of truth)
# =========================
$table = New-Object System.Data.DataTable
$table.Columns.Add("Name",[string])   | Out-Null
$table.Columns.Add("Shift",[string])  | Out-Null
$table.Columns.Add("Team",[string])   | Out-Null
$table.Columns.Add("Active",[bool])   | Out-Null

# =========================
# Load CSV immediately
# =========================
Import-Csv $csvPath | ForEach-Object {
    $row = $table.NewRow()
    $row.Name   = $_.Name
    $row.Shift  = $_.Shift
    $row.Team   = $_.Team
    $row.Active = ($_.Active -eq "TRUE")
    $table.Rows.Add($row)
}

$grid.DataSource = $table

# =========================
# Save
# =========================
$saveBtn.Add_Click({
    $table |
        Select-Object Name,Shift,Team,
            @{Name="Active";Expression={ if ($_.Active) { "TRUE" } else { "FALSE" } }} |
        Export-Csv $csvPath -NoTypeInformation

    [System.Windows.Forms.MessageBox]::Show("Saved.")
})

# =========================
# Add Employee
# =========================
$addBtn.Add_Click({
    $row = $table.NewRow()
    $row.Name   = ""
    $row.Shift  = "1st"
    $row.Team   = ""
    $row.Active = $true
    $table.Rows.Add($row)
})

# =========================
# Remove Selected
# =========================
$removeBtn.Add_Click({
    foreach ($row in $grid.SelectedRows) {
        $grid.Rows.RemoveAt($row.Index)
    }
})

# =========================
# Controls
# =========================
$form.Controls.Add($saveBtn)
$form.Controls.Add($addBtn)
$form.Controls.Add($removeBtn)
$form.Controls.Add($grid)

$form.ShowDialog()
