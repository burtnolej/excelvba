 param (
    [string]$outputfile = "C:\Users\burtn\Deploy\.folders.csv"
 )

#$utils_file = Join-Path -Path $working_dir -ChildPath "Tools-Utils.ps1"
#. $utils_file

function Log-Output {
    [CmdletBinding()]
    param (
        [ref]$result,
        [String]$status,
        [String]$action,
        [String]$object,
        [String]$message,
        [String]$errormsg
    )

    $LOGTIME=Get-Date -Format "MMddyyyy_HHmmss"
    
    $sb = New-Object -TypeName System.Text.StringBuilder
    
    $null = $sb.Append($LOGTIME.PadRight(18," "))
    $null = $sb.Append($status.PadRight(7," "))
    $null = $sb.Append($action.PadRight(25," "))
    $null = $sb.Append($message.PadRight(50," "))
    $null = $sb.Append($object.PadRight(100," "))
    

    if ($PSBoundParameters.ContainsKey('errormsg')) {
        $null = $sb.Append($errormsg)
    }
    $result.value = $sb.ToString()
}

function Get-OneDriveSubFolders {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$AdminName,
        [Parameter(Mandatory)]
        [String]$SiteUrl,
        [Parameter(Mandatory)]
        [String]$FileFolder
    )

    $output=$null

    $CLIENTDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    $CLIENTRUNTIMEDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    if (Test-Path -Path $CLIENTDLL) {
        Log-Output -result ([ref]$output) -status "OK" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Found!"
        #Write-Information  $output
    }
    else {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
        Write-Error  $output
    }

    Add-Type -Path $CLIENTDLL
    Add-Type -Path $CLIENTRUNTIMEDLL

    if ($AdminName -eq "jon.butler@veloxfintech.com") {
        $AdminPassword ="4o5yWohgxOB8"
        $SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
    }

    
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)

    try {
        if (-not ([string]::IsNullOrEmpty($AdminPassword)))
        {
            $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$SecurePassword)
        }
        else {
            $Credential =Get-Credential -Credential $AdminName
            $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password) 
        }
    }
    catch {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Get Credential" -object $AdminName -message "Failed" -errormsg  $_
        Write-Error $output
        exit
    }

    Try {

        #Get the Folder and Files
        $Folder=$Context.Web.GetFolderByServerRelativeUrl($FileFolder)
        $Context.Load($Folder)
        #$Context.Load($Folder.Files)
        $Context.Load($Folder.Folders)
        #$Context.Load($Folder.Author)
        #$Context.Load($Folder.ModifiedBy)        
        $Context.ExecuteQuery()
 
        $sb = New-Object -TypeName System.Text.StringBuilder
        $sb2 = New-Object -TypeName System.Text.StringBuilder

        #Iterate through each File in the folder
        #Foreach($File in $Folder.Files)
        Foreach($Subfolder in $Folder.Folders)
        {

            $date1 = Get-Date "2023/10/01"
            if ($Subfolder.TimeLastModified -gt $date1) {
                $Context.Load($Subfolder.Files)
                $Context.ExecuteQuery()
                
                $null = $sb2.Clear()
                Foreach($file in $Subfolder.Files)
                {
                    $null = $sb2.Append($file.Name)
                    $null = $sb2.Append(";;")
                }

                $null = $sb.Append($Subfolder.Name)
                $null = $sb.Append(",")
                $null = $sb.Append($Subfolder.TimeCreated)
                $null = $sb.Append(",")
                $null = $sb.Append($Subfolder.TimeLastModified)
                $null = $sb.Append(",")
                $null = $sb.Append($Subfolder.ServerRelativeURL)
                $null = $sb.Append(",")
                $null = $sb.Append("") #Size
                $null = $sb.Append(",")
                $null = $sb.Append($Subfolder.ItemCount)
                $null = $sb.Append(",")
                $null = $sb.Append($sb2.ToString())
                $null = $sb.Append("`n")
                #Write-Host  $Subfolder.Name
                #Write-Host  $sb2.ToString()
            }

        }
    }
    Catch{
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "List Folder" `
                -object $FileFolder `
                -message "Failed" `
                -errormsg  $_
        Write-Error $output
        exit
    }

    return $sb.ToString()
}

$output=$null

Read-Host "Installing ..... Press RETURN to continue"

Log-Output -result ([ref]$output) `
                -status "OK" `
                -action "Argument passed" `
                -object $outputfile `
                -message "output file"

Write-Host $output

#$MYHOME=Get-Content -Path Env:\HOMEPATH

#$folderfile=Join-Path -Path $MYHOME -ChildPath "Deploy\.folders.csv"

Get-OneDriveSubFolders "jon.butler@veloxfintech.com" `
    "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
    "/sites/VeloxSharedDrive/Shared%20Documents/General/Monday" `
    | Out-File -FilePath $outputfile -encoding utf7

Read-Host "Installing ..... Press RETURN to continue"