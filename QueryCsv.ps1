
function Get-CsvFilePath() 
{ 
    param (
        [string]$FilePath
    );

    $csvFile = [System.IO.Path]::GetFullPath($FilePath);

    return $csvFile;
}

# file / provider 
$csv = ""; 

if (($provider = (New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object {$_.SOURCES_NAME -like "Microsoft.ACE.OLEDB.*"}) -is [system.array]) 
{ 
   # Write-Host "Going with: $($provider[0].SOURCES_NAME)"
    $provider = $provider[0].SOURCES_NAME; 
} 

else
{ 
    $provider = $provider.SOURCES_NAME;
}

# Strings
[string]$ConnectionString = "Provider=$provider;Data Source=$(Split-Path -Path $csv);Extended Properties='text;HDR=Yes;'"
[string]$tableName = (Split-Path -Path $csv -Leaf).Replace(".","#"); 
[string]$Filters = ""; 
[string]$query = "SELECT * FROM [$tableName] $Filters"; 

# Objects
$Connection = New-Object System.Data.OleDb.OleDbConnection; 
$Command = New-Object System.Data.OleDb.OleDbCommand;
$DataTable = New-Object System.Data.DataTable;

# Connection 
$Connection.ConnectionString = $ConnectionString;
$Connection.Open();

# Command
$Command.Connection = $Connection;
$Command.CommandText = $query; 

# DataTable
try
{
    $DataTable.Load($Command.ExecuteReader("CloseConnection")); 
    $rows = $DataTable.Select(); 

    for ($i = 0; $i -lt $rows.length; $i ++ ) 
    { 
        Write-Host $rows[$i][""]
    }
}
catch
{
    Write-Error "Can't Complete the Read";
}

$Command.Dispose(); 
$Connection.Dispose(); 

