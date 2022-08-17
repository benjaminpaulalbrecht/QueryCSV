Import-Module WebAdministration; 

$AppPoolUserInformation = @(
    Get-ChildItem -Path 'IIS:\AppPools\' | 
        Select-Object -Property Name, @{e={$_.processModel.username};l='Username'}`
        ,@{e={$_.processModel.password};l='Password'}
);  

$objAppPoolInfo = $AppPoolUserInformation | ForEach-Object { 
    [PSCustomObject]@{
        AppPoolName = $_.Name
        UserName = $_.Username 
        Password = $_.Password
        Type = 'AppPool'
    }
}


$ServiceAccountInfomation = @( 
    Get-WmiObject -Class win32_service | Select-Object -Property Name, StartName;
); 

$objSvcAcctInfo = $ServiceAccountInfomation | ForEach-Object { 
    [PSCustomObject]@{
        ServiceName = $_.Name
        SeviceAccount = $_.StartName
        Type = 'Service'
    }
}

$report = @() 

foreach ($appPool in $objAppPoolInfo) { 
    $report += [pscustomobject]@{
    AppPoolName = $appPool.AppPoolName 
    AppPoolUserID = $appPool.AppPoolName
    AppPoolUserPassword = $appPool.Password
    ServiceName = '' 
    ServiceAccount = '' 
    Type = $appPool.Type
    }
}

foreach ($service in $objSvcAcctInfo) { 
    $report += [PSCustomObject]@{
        AppPoolName = '' 
        AppPoolUserID = '' 
        AppPoolUserPassword = ''
        ServiceName = $service.Name 
        ServiceAccount =  $service.ServiceAccount
        Type = $service.Type

    }
}


$report

