Import-Module WebAdministration; 

# AppPools
$AppPoolUserInformation = @(
    Get-ChildItem -Path 'IIS:\AppPools\' | 
        Select-Object -Property @{e={$_.Name};l='AppPoolName'}, @{e={$_.processModel.username};l='Username'}`
        ,@{e={$_.processModel.password};l='Password'}
);  

$objAppPoolInfo = $AppPoolUserInformation | ForEach-Object { 
    [PSCustomObject]@{
        AppPoolName = $_.AppPoolName
        UserName = $_.Username 
        Password = $_.Password
    }
}

# Services
$ServiceAccountInfomation = @( 
    Get-WmiObject -Class win32_service | Select-Object -Property @{e={$_.Name};l='ServiceName'}, StartName;
); 

$objSvcAcctInfo = $ServiceAccountInfomation | ForEach-Object { 
    [PSCustomObject]@{
        ServiceName = $_.ServiceName
        ServiceAccount = $_.StartName
    }
}

# Create an array to hold our objects
$report = @() 

# AppPools
foreach ($appPool in $objAppPoolInfo) { 
    $report += [PSCustomObject]@{
    AppPoolName = $appPool.AppPoolName 
    AppPoolUserID = $appPool.AppPoolName
    AppPoolUserPassword = $appPool.Password
    ServiceName = '' 
    ServiceAccount = ''
    }
}

# Services
foreach ($service in $objSvcAcctInfo) { 
        $report += [PSCustomObject]@{
        AppPoolName = '' 
        AppPoolUserID = '' 
        AppPoolUserPassword = ''
        ServiceName = $service.ServiceName 
        ServiceAccount =  $service.ServiceAccount
    }
}

$report

