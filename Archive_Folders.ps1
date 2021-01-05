<#
        .SYNOPSIS
        To compress source folder and remove them.

        .DESCRIPTION
        This script is used to compress the source folder and remove it after compression
        Date - 26/12/2020
        Developer - K.Janarthanan
        Version - 1

        .PARAMETER ConfigFile
        The Config File location.

        .OUTPUTS
        Log file will be created in the same location of script.

        .EXAMPLE
        C:\PS> ./Archive_Folders.ps1 -ConfigFile "C:\PS\Config.json"
#>

Param(
    [Parameter(Mandatory)]
    [string]$ConfigFile
)

$Global:LogFile = "{0}\{1}.Log" -f $PSScriptRoot,(Get-Date -Format "MM_dd_yyyy_HH-mm-ss")

#Function for logging
function Write-Log
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Validateset("INFO","ERROR","WARN")]
        [string]$Type="INFO"
    )

    $DateTime = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
    $FinalMessage = "[{0}]::[{1}]::[{2}]" -f $DateTime,$Type,$Message

    $FinalMessage | Out-File -FilePath $LogFile -Append
    Write-Host $FinalMessage
}

#Main
try 
{
    Write-Log "Script Execution started"

    if(Test-Path -Path $ConfigFile -PathType Leaf)
    {
        $Config = Get-Content -Path $ConfigFile -ErrorAction Stop | ConvertFrom-Json

        foreach ($Folders in $Config.Archive)
        {
            try
            {
                if(Test-Path -Path $Folders.SourceFolder)
                {
                    $ArchiveFolders = (Get-ChildItem -Path $Folders.SourceFolder -Directory -EA Stop)
                    
                    foreach($SingleFolder in $ArchiveFolders)
                    {
                        try 
                        {
                            
                            Write-Log "Working on folder $($SingleFolder.Name)"

                            if($SingleFolder -match '\d{2}-\d{2}-\d{4}') #To check pattern
                            {
                                $DateObject = [Datetime]::ParseExact($SingleFolder.Name, 'MM-dd-yyyy', $null)

                                if($DateObject -lt (Get-Date).AddDays(-$Folders.NumberOfDays))
                                {
                                    Write-Log "Going to archive the folder $($SingleFolder.Name)"
                                    
                                    if(-not(Test-Path -Path $Folders.DestinationFolder))
                                    {
                                        Write-Log "Destination Folder $($Folders.DestinationFolder) is not found. Therefore will create it"
                                        New-Item -ItemType Directory -Path $Folders.DestinationFolder -Force -EA Stop
                                    }
        
                                    Compress-Archive -Path "$($SingleFolder.FullName)\*" -DestinationPath ("{0}\{1}.zip" -f $Folders.DestinationFolder,$SingleFolder.Name) -Force -EA Stop
                                    Write-Log "Successfully archived the folder $($SingleFolder.Name)"

                                    Write-Log "Going to remove the folder $($SingleFolder.Name)"
                                    Remove-Item -Path $SingleFolder.FullName -Force -Recurse -EA Stop
                                    Write-Log "Successfully removed the folder $($SingleFolder.Name)"
                                }
                                else 
                                {
                                    Write-Log "Folder $($SingleFolder.Name) is not archived, as it is not less than $($Folders.NumberOfDays) days"    
                                }          
                            }
                            else 
                            {
                                Write-Log "Skipping the folder $($SingleFolder.Name) as it is not matching name pattern"    
                            }
                        }  
                        catch
                        {
                            Write-Log "$_" -Type Error
                        }
                    }
                }
                else 
                {
                    Write-Log "Source Folder $($Folders.SourceFolder) is not available" -Type ERROR    
                }
            }
            catch
            {
                Write-Log "$_" -Type Error
            }
        } 
    }
    else 
    {
        throw "Configuration JSON file not found"    
    }
    
    Write-Log "Script Execution Completed" 
}
catch 
{
    Write-Log "$_" -Type ERROR 
    Write-Log "Script Execution Completed" 
}
