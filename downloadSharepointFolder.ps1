#Based on:
#https://learn.microsoft.com/en-us/answers/questions/788342/powershell-script-to-download-specific-folders-fro
#Set Parameters
$SiteURL = "https://yourtenantname.sharepoint.com/sites/sitename/subsite"
#$ListName="Información General"
$FolderURL="/LibraryName/Folder/SubFolder"

$DownloadPath ="\\?\F:\SharepointBackup"
$LogDir = "F:\Logs"

#Function to Download All Files from a SharePoint Online Folder - Recursively 
Function Download-SPOFolder([Microsoft.SharePoint.Client.Folder]$Folder, $DestinationFolder, $logFile)
{
    #Get the Folder's Site Relative URL
    $FolderURL = $Folder.ServerRelativeUrl.Substring($Folder.Context.Web.ServerRelativeUrl.Length)
    $LocalFolder = $DestinationFolder + ($FolderURL -replace "/","\")
    #Create Local Folder, if it doesn't exist
    $ErrorToCreateLocalFolder = 0
    $ErrorMessage = ''
    If (!(Test-Path -Path $LocalFolder)) {
            try
            {
                New-Item -ItemType Directory -Path $LocalFolder|Out-Null
                #Write-host -f Yellow "Created a New Folder '$LocalFolder'"
            }
            catch
            {
                $ErrorMessage = $_.CategoryInfo.Activity + ": " + $_.Exception.Message
                Write-Error "Failed to create local folder: $LocalFolder, $ErrorMessage"
                $ErrorToCreateLocalFolder = 1
                Add-Content $OutFile "Local,'$LocalFolder',ERROR,'$ErrorMessage'"
            }
    }

    if ($ErrorToCreateLocalFolder -eq 0)
    {
        #Get all Files from the folder
        $FilesColl = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType File
        #Iterate through each file and download
        $counter = 0
        $arraySize = $FilesColl.count
        Foreach($File in $FilesColl)
        {
            [int]$PercentComplete = ($counter / $arraySize) * 100
            $counter++
            $FileName = $File.Name
            Write-Progress -Activity "Downloading files from: $FolderURL" -CurrentOperation "File Name: '$FileName' ($counter/$arraySize)" -PercentComplete $PercentComplete
            $ErrorMessage = $File.Length
            try
            {
                #There will be errors if file has % in the named
                Get-PnPFile -ServerRelativeUrl $File.ServerRelativeUrl -Path $LocalFolder -FileName $File.Name -AsFile -force
                Add-Content $OutFile "Remote,'$FolderURL/$FileName',OK,'$ErrorMessage'"
                #Write-host -f Green "`tDownloaded File from '$($File.ServerRelativeUrl)'"
            }
            catch
            {
                $ErrorMessage = $_.CategoryInfo.Activity + ": " + $_.Exception.Message
                if ( $ErrorMessage -like 'Get-PnPFile: El archivo * no existe.' -and $File.Name -like '*%*' )
                {
                    $ErrorMessage = "Get-PnPFile: File '" + $File.ServerRelativeUrl + "' Has % in the name. Remove % then try again"      
                }
                Write-Error "Failed to dowload file: $FolderURL/$FileName, $ErrorMessage"
                Add-Content $OutFile "Remote,'$FolderURL/$FileName',ERROR,'$ErrorMessage'"
            }
        }
        #Get Subfolders of the Folder and call the function recursively
        $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType Folder | Where {$_.Name -ne "Forms"}    
        
        $subdirectoryCount = 0
        $subdirectoryTotal = $SubFolders.count
        $parentFolderURL = $FolderURL
        Foreach ($Folder in $SubFolders)
        { 
            $subdirectoryCount++
            $FolderURL = $Folder.ServerRelativeUrl.Substring($Folder.Context.Web.ServerRelativeUrl.Length+$parentFolderURL.Length)
            try
            {
                [int]$PercentCompleteDirectory = ($subdirectoryCount / $subdirectoryTotal) * 100
            }
            catch
            {
                [int]$PercentCompleteDirectory = 100
                $subdirectoryTotal = 1
            }
            Write-Progress -Activity "Working on: $parentFolderURL" -CurrentOperation "Subfolder: $FolderURL ($subdirectoryCount/$subdirectoryTotal)" -PercentComplete $PercentCompleteDirectory
            Download-SPOFolder $Folder $DestinationFolder
        }
    }
} 

$CurrentDate = Get-Date -UFormat "%Y-%m-%d_%H-%M-%S"
$OutFile = $LogDir + "\" + "downloadSharepointFolder.${CurrentDate}.log.csv"

# https://www.sharepointdiary.com/2021/04/connect-pnponline-command-was-found-in-module-pnp-powershell-but-module-could-not-be-loaded.html
# Install-Module PnP.PowerShell -RequiredVersion 1.12 -Force -Scope CurrentUser
#Connect to PnP Online
Import-Module PnP.PowerShell
#Connect-PnPOnline -Url $SiteURL -UseWebLogin 
Connect-PnPOnline -Url $SiteURL -Interactive

#Get The Root folder of the Library
#$ListID = ((Get-PnPList -ThrowExceptionIfListNotFound | Where-Object { $_.Title -eq $ListName }).Id).Guid
#$Folder = Get-PnPFolder -List $ListID|Where-Object { $_.Name -eq $FolderName }
$Folder = Get-PnPFolder -Url $FolderURL
if ( $Folder -eq $null)
{
    throw "Not found any forlder with name: '$FolderName'"
    exit 1
}
if ( $Folder.Count -gt 1)
{
    throw "There are more than one folder with name: '$FolderName'"
    exit 1
}

$StartTime = Get-Date -UFormat "%Y-%m-%d_%H-%M-%S"
Add-Content $OutFile "#downloadSharepointFolder, URL: $SiteURL$FolderURL, Start Time: '$StartTime'"
Add-Content $OutFile "Souce,Path,State,FileLength"
#Call the function to download the document library
Download-SPOFolder $Folder $DownloadPath $OutFile
$EndTime = Get-Date -UFormat "%Y-%m-%d_%H-%M-%S"
Add-Content $OutFile "#downloadSharepointFolder, URL: $SiteURL$FolderURL, End Time: '$EndTime'"
