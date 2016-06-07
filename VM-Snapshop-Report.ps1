
Set-Location (Get-Item ($MyInvocation.MyCommand.Definition)).DirectoryName

Add-PSSnapin *vmware* -ErrorAction:SilentlyContinue

$vCenters = 'vc001.company.com', 'vc002.company.com', 'vc003.company.com'

$Date = Get-Date
$strDate = $Date.ToString('yyyy-MM-dd')

$ReportTable = @{
    Windows = @{
        Enabled = $true;
        Filter = 'windows';
        EmailSubject = "[$strDate] Monthly Windows snapshot report.";
        EmailFrom = "$($env:COMPUTERNAME)@company.com";
        EmailTo = 'larrysong@company.com';
        EmailSMTP = 'smtp.company.com';
    };
    Linux = @{
        Enabled = $true;
        Filter = 'linux';
        EmailSubject = "[$strDate] Monthly Linux snapshot report.";
        EmailFrom = "$($env:COMPUTERNAME)@company.com";
        EmailTo = 'larrysong@company.com';
        EmailSMTP = 'smtp.company.com';
    };
}

$LogFolder = $strDate

try {
    New-Item -Name $LogFolder -ItemType Directory -Force | Out-Null
} catch {
    $LogFolder = '.'
}

$strLogFile = "$LogFolder\${strDate}.log"

function Add-Log{
    PARAM(
        [String]$Path,
        [String]$Value,
        [String]$Type = 'NULL'
    )
    $Type = $Type.ToUpper()
    Write-Host "$((Get-Date).ToString('[HH:mm:ss] '))[$Type] $Value"
    if($Path){
        Add-Content -Path $Path -Value "$((Get-Date).ToString('[HH:mm:ss] '))[$Type] $Value"
    }
}

$objCSVBase = New-Object PSObject
'VM Name', 'Snapshot name', 'Created time' | %{
    Add-Member -InputObject $objCSVBase -Name $_ -Value $null -MemberType NoteProperty -Force
}

Add-Log -Path $strLogFile -Value '====== Script initialized.' -Type Info
foreach($vCenter in $vCenters){
    Add-Log -Path $strLogFile -Value "Trying connect to ${vCenter}" -Type Info
    try{
        Connect-VIServer -Server $vCenter -Force
    } catch {
        Add-Log -Path $strLogFile -Value "Failed connect to ${vCenter}, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        continue
    }
    Add-Log -Path $strLogFile -Value "Successfully connect to ${vCenter}" -Type Info
    $VM_ALL = $null
    try {
        $VM_ALL = @(Get-VM)
    } catch {
        Add-Log -Path $strLogFile -Value "Failed to run Get-VM cmdlet, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        continue
    }
    Add-Log -Path $strLogFile -Value "All VM count: $($VM_ALL.Count)" -Type Info
    $ReportTable.Keys | %{
        $Type = $_
        if(!$ReportTable.$Type.Enabled){return}
        Add-Log -Path $strLogFile -Value "Collecting data for $Type on $vCenter" -Type Info
        $ReportTable.$Type.Add($vCenter, @())
        $VM_FILTERED = $null
        $VM_FILTERED = @($VM_ALL | ?{$_.ExtensionData.Config.GuestFullName -imatch $ReportTable.$Type.Filter})
        Add-Log -Path $strLogFile -Value "VM filtered: $($VM_FILTERED.Count)" -Type Info
        if($VM_FILTERED.Count -eq 0){
            Add-Log -Path $strLogFile -Value '0 VM captured, unlikely' -Type Warning
            return
        }
        $VM_SNAPSHOTS = $null
        try {
            $VM_SNAPSHOTS = @(Get-Snapshot -VM $VM_FILTERED)
        } catch {
            Add-Log -Path $strLogFile -Value "Failed to run Get-Snapshot cmdlet, cause:" -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            return
        }
        Add-Log -Path $strLogFile -Value "Snapshot captured: $($VM_SNAPSHOTS.Count)" -Type Info
        if($VM_SNAPSHOTS.Count -eq 0){
            Add-Log -Path $strLogFile -Value 'No snapshot captured' -Type Info
            return
        }
        foreach($snapshot in $VM_SNAPSHOTS){
            $ReportTable.$Type.$vCenter += $objCSVBase.PSObject.Copy()
            $ReportTable.$Type.$vCenter[-1].'VM Name' = $snapshot.VM.Name
            $ReportTable.$Type.$vCenter[-1].'Snapshot name' = $snapshot.Name
            $ReportTable.$Type.$vCenter[-1].'Created time' = $snapshot.Created.ToString('yyyy-MM-dd HH:mm:ss')
        }
        if($ReportTable.$Type.$vCenter){
            $ReportTable.$Type.$vCenter | Export-Csv -NoTypeInformation -Delimiter ',' -Encoding Unicode -Path "$LogFolder\${vCenter}.${Type}.CSV"
        }
    }
    Disconnect-VIServer -Server $vCenter -Force -Confirm:$false
    Add-Log -Path $strLogFile -Value "Disconnected from $vCenter" -Type Info
}

$ReportTable.Keys | %{
    if(!$ReportTable.$Type.Enabled){return}
    Add-Log -Path $strLogFile -Value 'Start sending report' -Type Info
    $CSVFiles = $null
    $CSVFiles = ls -Filter "$LogFolder\*.${_}.CSV" | %{$_.FullName}
    if($CSVFiles){
        try {
            Send-MailMessage -From $ReportTable.$_.EmailFrom -To $ReportTable.$_.EmailTo -Subject $ReportTable.$_.EmailSubject -Attachments $CSVFiles -SmtpServer $ReportTable.$_.EmailSMTP
        } catch {
            Add-Log -Path $strLogFile -Value 'Failed to send email, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        }
    }else{
        try {
            Send-MailMessage -From $ReportTable.$_.EmailFrom -To $ReportTable.$_.EmailTo -Subject "$($ReportTable.$_.EmailSubject) - [No snapshot found]" -SmtpServer $ReportTable.$_.EmailSMTP
        } catch {
            Add-Log -Path $strLogFile -Value 'Failed to send email, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        }
        Add-Log -Path $strLogFile "No CSV files folder under $LogFolder for $_, no attachment to send." -Type Info
    }
}
