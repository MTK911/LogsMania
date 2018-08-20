Param(
    [parameter(Mandatory=$false)]
    [Bool]$AutoOpen,
	[parameter(Mandatory=$false)]
    [String]$OutFile,
	[parameter(Mandatory=$false)]
    [Bool]$IgnoreInParams,
	[parameter(Mandatory=$false)]
    [Bool]$IgnoreOutParams)

$Error.Clear()
$DefaultFolder=[Environment]::CurrentDirectory
$Destination = "8 Emails recived from outside.CSV"

$Dire = "\..\Logs Mania\LOGS_DATA\"
$Destination = $DefaultFolder + $Dire + "\" + $Destination

if($OutFile -ne [String]::Empty)
{
	$OutFileType = [System.IO.Path]::GetExtension($OutFile.ToUpper())
	$OriginalFileType = [System.IO.Path]::GetExtension($Destination.ToUpper())
	if($OutFileType -ne $OriginalFileType)
	{
		Write-Host "You have chosen" $OutFileType "as the output, but this script was originally generated as" $OriginalFileType -ForegroundColor Red
		Write-Host "Either change -OutFile to" $OriginalFileType "or generate the script again with the output as" $OutFileType  -ForegroundColor Red
		Write-Host "You can also modify the OutputFormat variable in this script to match the correct Log Parser 2.2 COM output format." -ForegroundColor Red
		[System.Environment]::NewLine
		return
	}
	else
	{
		if($true -ne $OutFile.Contains("\"))
		{
		  $Destination = $DefaultFolder + "\" + $OutFile
		}
		else
		{
		  $Destination = $OutFile
		}
	}
}

$LogQuery = New-Object -ComObject "MSUtil.LogQuery"
$InputFormat = New-Object -ComObject "MSUtil.LogQuery.CSVInputFormat"
$OutputFormat = New-Object -ComObject "MSUtil.LogQuery.CSVOutputFormat"

if($IgnoreInParams-eq $false){
     $InputFormat.headerRow=1
     #$InputFormat.dQuotes=0  #deprecated try 'useDoubleQuotes' instead
     $InputFormat.nFields=-1
     $InputFormat.nSkipLines=4
     $InputFormat.iTsFormat="yyyy-MM-dd hh:mm:ss"
     $InputFormat.iCodePage=0
     $InputFormat.fixedFields=1
     $InputFormat.dtLines=10
}

if($IgnoreOutParams -eq $false){
     $OutputFormat.Headers="AUTO"
     $OutputFormat.oDQuotes="AUTO"
     $OutputFormat.tabs=0
     $OutputFormat.oTsFormat="yyyy-MM-dd hh:mm:ss"
     $OutputFormat.oCodepage=0
     $OutputFormat.fileMode=1
}
Write-Progress -Activity "Executing query, please wait..."  -Status " "

$SQLQuery = "SELECT TO_TIMESTAMP( EXTRACT_PREFIX(TO_STRING([#Fields: date-time]),0,'.') , 'yyyy-MM-ddThh:mm:ss') AS DATE-TIME,[client-ip],[server-ip],[event-id],[directionality],[recipient-address],[sender-address],[message-subject],[total-bytes],[message-id],[server-hostname] INTO '" + $Destination + "' FROM '\Logs Mania\LOGS\*.LOG' WHERE ([sender-address] not LIKE '%org.com%' and [recipient-address] LIKE '%org.com%' AND [client-ip] IS NOT NULL)"
$rtnVal = $LogQuery.ExecuteBatch($SQLQuery, $InputFormat, $OutputFormat);
$OutputFormat = $null;
$InputFormat = $null;
$LogQuery = $null;

if($AutoOpen)
{
	try
	{
		Start-Process($Destination)
	}
	catch
	{
		Write-Host $_.Exception.Message  -ForegroundColor Red
		Write-Host $_.Exception.GetType().FullName  -ForegroundColor Red
		Write-Host "NOTE: No output file will be created if the query returned zero records!"  -ForegroundColor Gray
	}
}
