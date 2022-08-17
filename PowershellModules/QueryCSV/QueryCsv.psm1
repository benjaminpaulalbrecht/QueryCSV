
function DisposeOf 
{ 
    param ( 
        [object]$Command, 
        [object]$Connection
    ); 

    $Command.Dispose(); 
    $Connection.Dispose(); 
}



function Get-CSVDataByQuery { 
    param ( 
        [string]$DataFilePath, 
        [string]$SQLQuery
    )

    # provider sources name string may change to a var in the params...
    if (($provider = (New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object {$_.SOURCES_NAME -like "Microsoft.ACE.OLEDB.*"}) -is [system.array]) 
    { 

        $provider = $provider[0].SOURCES_NAME; 
    } 
    else
    { 
        $provider = $provider.SOURCES_NAME;
    }

    # Strings
    [string]$ConnectionString = "Provider=$provider;Data Source=$(Split-Path -Path $DataFilePath);Extended Properties='text;HDR=Yes;'"
    [string]$tableName = (Split-Path -Path $DataFilePath -Leaf).Replace(".","#"); 
    [string]$Filters = ""; 

    if ([string]::IsNullOrEmpty($SQLQuery)) { 

        # default
        [string]$query = "SELECT * FROM [$tableName] $Filters"; 
    }
    else 
    { 
        $query = $SQLQuery;
    }

    # Objects
    $Connection = New-Object System.Data.OleDb.OleDbConnection; 
    $Command = New-Object System.Data.OleDb.OleDbCommand;
    $DataTable = New-Object System.Data.DataTable;

    # Connection 
    $Connection.ConnectionString = $ConnectionString;
    
    try
    {
        $Connection.Open();
    }
    catch [OleDbException]
    {
        Write-Host $Error[0].Exception.GetType().FullName;
        exit;
    }

    # Command
    $Command.Connection = $Connection;
    $Command.CommandText = $query; 

    try {
        $table =  $DataTable.Load($Command.ExecuteReader("CloseConnection")); 
    }
    catch  {
        Write-Host $Error[0].Exception.GetType().FullName;
        Exit;
    }
    
    DisposeOf -Command $Command -Connection $Connection;

    return $table;
} 

