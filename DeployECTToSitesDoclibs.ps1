<#
        This script applies the OnePlaceMail Email Columns to an existing site collection, creates Site Content Types, adds them to Document Libraries and creates a default view.
        Please check the README.md on Github before using this script.
#>
$ErrorActionPreference = 'Stop'

#Columns to add to the Email View if we are creating one. Edit as required based on Internal Naming
[string[]]$script:emailViewColumns = @("EmHasAttachments","EmSubject","EmTo","EmDate","EmFromName")

$script:logFile = "OPSScriptLog.txt"
$script:logPath = "$env:userprofile\Documents\$script:logFile"

$script:extractedTenant = ""

function Write-Log { 
    <#
        .NOTES 
            Created by: Jason Wasser @wasserja 
            Modified by: Ashley Gregory
        .LINK (original)
            https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
        #>
    [CmdletBinding()] 
    Param ( 
        [Parameter(Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
    
        [Parameter(Mandatory = $false)] 
        [Alias('LogPath')] 
        [string]$Path = $script:logPath, 
            
        [Parameter(Mandatory = $false)] 
        [ValidateSet("Error", "Warn", "Info")] 
        [string]$Level = "Info", 
            
        [Parameter(Mandatory = $false)] 
        [switch]$NoClobber 
    ) 
    
    Begin {
        $VerbosePreference = 'SilentlyContinue' 
        $ErrorActionPreference = 'Continue'
    } 
    Process {
        # If the file already exists and NoClobber was specified, do not write to the log. 
        If ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
        } 
    
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        ElseIf (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
        } 
    
        Else { 
            # Nothing to see here yet. 
        } 
    
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss K" 
    
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        Switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
            } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
            } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
            } 
        } 
            
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    }
    End {
        $ErrorActionPreference = 'Stop'
    } 
}

Add-Type -AssemblyName System.Windows.Forms
Try {
    Set-ExecutionPolicy Bypass -Scope Process

    Clear-Host
    Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
    Write-Host 'Welcome to the OnePlace Solutions Content Type Deployment Script' -ForegroundColor Green
    Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red

    Write-Host "Beginning script. `nLogging script actions to $script:logPath" -ForegroundColor Cyan
    Write-Host "Performing Pre-Requisite checks, please wait..." -ForeGroundColor Yellow
    Start-Sleep -Seconds 3

    #Check for module versions of PnP / SPOMS
    Try {
        Write-Host "Checking if PnP / SPOMS installed via Module..." -ForegroundColor Cyan
        $pnpModule = Get-InstalledModule SharePointPnPPowerShell* | Select-Object Name, Version
        $spomsModule = Get-InstalledModule Microsoft.Online.SharePoint.PowerShell | Select-Object Name, Version
    }
    Catch {
        #Couldn't check PNP or SPOMS Module versions, Package Manager may be absent
    }
    Finally {
        Write-Log -Level Info -Message "PnP Module Installed: $pnpModule"
        Write-Log -Level Info -Message "SPOMS Module Installed: $spomsModule"
    }

    #Check for MSI versions of PnP / SPOMS
    Try {
        Write-Host "Checking if PnP / SPOMS installed via MSI..." -ForegroundColor Cyan
        $pnpMSI = Get-WmiObject Win32_Product | Where-Object {$_.Name -match "PnP PowerShell*"} | Select-Object Name, Version
        $spomsMSI = Get-WmiObject Win32_Product | Where-Object {$_.Name -match "SharePoint Online Management Shell"} | Select-Object Name, Version
    }
    Catch {
        #Couldn't check PNP or SPOMS MSI versions
    }
    Finally {
        Write-Log -Level Info -Message "PnP MSI Installed: $pnpMSI"
        Write-Log -Level Info -Message "SPOMS MSI Installed: $spomsMSI"
    }

    If (($null -ne $pnpModule.Count) -or ((0 -ne $pnpMSI.Count) -and (1 -ne $pnpMSI.Count))) {
        Write-Log -Level Warn -Message "Multiple versions of PnP may be installed. This is not supported by PnP and will likely cause issues when running this script.`nPlease uninstall the versions not applicable to your SharePoint version and re-run this script."
        Pause
    }

    $preReqMissing = $false
    If (($null -eq $pnpModule) -and ($null -eq $pnpMSI)) {
        Write-Log -Level Warn -Message "No SharePoint PnP Cmdlets installation detected! This is required."
        $preReqMissing = $true
    }
    If (($null -eq $spomsModule) -and ($null -eq $spomsMSI)) {
        Write-Log -Level Warn -Message "No SharePoint Online Management Shell installation detected! This is required."
        $preReqMissing = $true
    }
    If ($preReqMissing) {
        Write-Host "`nPlease ensure you have checked and installed the pre-requisites listed in the GitHub documentation prior to running this script."
        Write-Host "!!! If pre-requisites for the Content Type Deployment have not been completed this script/process may fail !!!" -ForegroundColor Yellow
        Pause
    }
    Start-Sleep -Seconds 2

    #Contains all our Site Collections as siteCol objects
    $script:siteColsHT = @{ }

    #Flag for whether we are working in SharePoint Online or on-premises.
    [boolean]$script:isSPOnline = $true

    #Flag for whether we create default email views or not, and if so what name to use
    [boolean]$script:createDefaultViews = $false
    $script:emailViewName = $null

    #Flag for whether we automatically create the OnePlaceMail Email Columns
    [boolean]$script:createEmailColumns = $false

    #Name of Column group containing the Email Columns, and an object to contain the Email Columns
    $script:groupName = $null
    $script:emailColumns = $null

    #Credentials object to hold On-Premises credentials, so we can iterate across site collections with it
    $script:onPremisesCred

    #Tells us if we are using token auth
    [boolean]$script:usingTokenAuth = $false

    #Holds our OAuth 2.0 token if using SharePoint Online
    $script:token = $null

    #Contains all the data we need relating to the Site Collection we are working with, including the Document Libraries and the Site Content Type names
    class siteCol {
        [String]$name
        [String]$url
        [String]$web
        [Hashtable]$documentLibraries = @{ }
        [Array]$contentTypes
        [Boolean]$isSubSite

        siteCol([string]$name, $url) {
            If ($name -eq "") {
                $this.name = $url
            }
            Else {
                $this.name = $name
            }
            $filler = $this.name
            $filler = "Creating siteCol object with name '$filler'"
            Write-Host $filler -ForegroundColor Yellow
            Write-Log -Level Info -Message $filler

            $this.contentTypes = @()

            $urlArray = $url.Split('/')
            $rootUrl = $urlArray[0] + '//' + $urlArray[2] + '/'

            If ($urlArray[3] -eq "") {
                #This is the root site collection
                $this.isSubSite = $false
            }
            ElseIf (($urlArray[3] -ne "sites") -and ($urlArray[3] -ne "teams")) {
                #This is a subsite in the root site collection
                For ($i = 3; $i -lt $urlArray.Length; $i++) {
                    If ($urlArray[$i] -ne "") {
                        $this.web += '/' + $urlArray[$i]
                    }
                }
                $this.isSubSite = $true
            }
            Else {
                #This is a site collection with a possible subweb
                $rootUrl += $urlArray[3] + '/' + $urlArray[4] + '/'
                For ($i = 3; $i -lt $urlArray.Length; $i++) {
                    If ($urlArray[$i] -ne "") {
                        $this.web += '/' + $urlArray[$i]
                    }
                }
                If ($urlArray[5].Count -ne 0) {
                    $this.isSubSite = $true
                }
                Else {
                    $this.isSubSite = $false
                }
            }
            [boolean]$temp = $this.isSubSite
            $filler = "Is this a subsite? $temp"
            Write-Host $filler -ForegroundColor Yellow
            Write-Log -Level Info -Message $filler
            $this.url = $rootUrl
        }

        [void]addContentTypeToDocumentLibrary($contentTypeName, $docLibName) {
            #Check we aren't working without a Document Library name, otherwise assume that we just want to add a Site Content Type
            If (($null -ne $docLibName) -and ($docLibName -ne "")) {
                If ($this.documentLibraries.ContainsKey($docLibName)) {
                    $this.documentLibraries.$docLibName
                    Write-Log -Level Info -Message "Document Library '$docLibName' already listed for this Site Collection."
                }
                Else {
                    $tempDocLib = [docLib]::new("$docLibName")
                    $this.documentLibraries.Add($docLibName, $tempDocLib)
                    Write-Log -Level Info -Message "Document Library '$docLibName' not listed for this Site Collection, added."
                }
                
                $this.documentLibraries.$docLibName.addContentType($contentTypeName)
                Write-Log -Level Info -Message "Listing Content Type '$contentTypeName' for Document Library '$docLibName'."
            }
            
            #If the named Content Type is not already listed in Site Content Types, add it to the Site Content Types
            If (-not $this.contentTypes.Contains($contentTypeName)) {
                $this.contentTypes += $contentTypeName
                Write-Log -Level Info -Message "Content Type '$contentTypeName' not listed in Site Content Types, adding."
            }
            Else {
                Write-Log -Level Info -Message "Content Type '$contentTypeName' listed in Site Content Types."
            }
        }
    }

    #Contains all the data we need relating to the Document Library we are working with, including the Site Content Type names we are adding to it
    class docLib {
        [String]$name
        [Array]$contentTypes

        docLib([String]$name) {
            $filler = "Creating docLib object with name $name"
            Write-Host $filler -ForegroundColor Yellow
            Write-Log -Level Info -Message $filler
            $this.name = $name
            $this.contentTypes = @()
        }

        [void]addContentType([string]$contentTypeName) {
            If (-not $this.contentTypes.Contains($contentTypeName)) {
                $filler = $this.name
                $filler = "Adding Content Type '$contentTypeName' to '$filler' Document Library Content Types"
                Write-Host $filler -ForegroundColor Yellow
                Write-Log -Level Info -Message $filler
                $this.contentTypes += $contentTypeName
            }
            Else {
                $temp = $this.name
                $filler = "Content Type '$contentTypeName' already listed in Document Library $temp"
                Write-Host $filler -ForegroundColor Red
                Write-Log -Level Info -Message $filler
            }
        }
    }

    #Grabs the CSV file and enumerate it into siteColHT as siteCol and docLib objects to work with later
    function EnumerateSitesDocLibs([string]$csvFile) {
        If ($csvFile -eq "") {
            Write-Host "Please select your customized CSV containing the Site Collections and Document Libraries to create the Content Types in"
            Start-Sleep -seconds 1
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
                InitialDirectory = [Environment]::GetFolderPath('Desktop') 
                Filter           = 'Comma Separates Values (*.csv)|*.csv'
                Title            = 'Select your CSV file'
            }
            $null = $FileBrowser.ShowDialog()
            
            $csvFile = $FileBrowser.FileName
            Write-Log -Level Info -Message "Using CSV at path '$csvFile'"
        }

        $script:siteColsHT = [hashtable]::new
        $script:siteColsHT = @{ }

        Try {
            $csv = Import-Csv -Path $csvFile -ErrorAction Continue

            Write-Host "Enumerating Site Collections and Document Libraries from CSV file." -ForegroundColor Yellow
            foreach ($element in $csv) {
                $csv_siteName = $element.SiteName
                $csv_siteUrl = $element.SiteUrl -replace '\s', '' #remove any whitespace from URL
                $csv_docLib = $element.DocLib
                $csv_contentType = $element.CTName

                If("" -eq $script:extractedTenant) {
                    $script:extractedTenant = $csv_siteUrl  -match 'https://(?<Tenant>.+)\.sharepoint.com'
                    $script:extractedTenant = $Matches.Tenant
                    Write-Log -Level Info -Message "Extracted Tenant name '$script:extractedTenant'"
                }

                #Don't create siteCol objects that do not have a URL, this also accounts for empty lines at EOF
                If ($csv_siteUrl -ne "") {
                    #If a name is not defined, use the URL
                    If ($csv_siteName -eq "") { $csv_siteName = $element.SiteUrl }

                    If ($script:siteColsHT.ContainsKey($csv_siteUrl)) {
                        $script:siteColsHT.$csv_siteUrl.addContentTypeToDocumentLibrary($csv_contentType, $csv_docLib)
                    }
                    Else {
                        $newSiteCollection = [siteCol]::new($csv_siteName, $csv_siteUrl)
                        $newSiteCollection.addContentTypeToDocumentLibrary($csv_contentType, $csv_docLib)
                        $script:siteColsHT.Add($csv_siteUrl, $newSiteCollection)
                    }
                }
            }
            $filler = "Completed Enumerating Site Collections and Document Libraries from CSV file!"
            Write-Host $filler -ForegroundColor Green
            Write-Log -Level Info -Message $filler
        }
        Catch {
            Write-Host "Error parsing CSV file. Is this filepath for a a valid CSV file?" -ForegroundColor Red
            $csvFile
            Throw
        }
    }

    #Facilitates connection to the SharePoint Online site collections through the SharePoint Online Management Shell
    #This automatically takes place on 'Connect-PnPOnline -SPOManagementShell' calls, but we can explicitly connect first here
    function ConnectToSharePointOnlineAdmin([string]$tenant) {
        #Prompt for SharePoint Management Site Url     
        Try {
            If ($tenant -eq "") {
                $tenant = Read-Host -Prompt "Please enter the name of your SharePoint Online tenant, eg for 'https://contoso.sharepoint.com' just enter 'contoso'."
                $tenant = $tenant.Trim()
                If (($tenant.Contains("sharepoint")) -and (-not $tenant.Contains('-admin'))) {
                    $tenant = $tenant.trim("https://")
                    $charArray = $tenant.Split(".")
                    $tenant = ($charArray[$charArray.IndexOf('sharepoint') - 1])
                    $adminSharePointUrl = "https://$tenant-admin.sharepoint.com"
                }
                ElseIf ($tenant.Contains('-admin.sharepoint.com')) {
                    $adminSharePointUrl = $tenant
                }
                Else {
                    $adminSharePointUrl = "https://$tenant-admin.sharepoint.com"
                }
            }
            Else {
                $adminSharePointUrl = "https://$tenant-admin.sharepoint.com"
            }
            #Connect to site collection
        
            Write-Host "Enter SharePoint credentials(your email address for SharePoint Online):" -ForegroundColor Green
            #Connect-SPOService -Url $adminSharePointUrl
            Connect-PnPOnline -Url $adminSharePointUrl -SPOManagementShell -ClearTokenCache
            #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
            Start-Sleep -Seconds 3
            Get-PnPWeb
            Pause
            $filler = "Testing connection with 'Get-PnPWeb'..."
            Write-Log -Level Info -Message $filler
            Write-Host $filler
        }
        Catch [System.Management.Automation.ParameterBindingException]{
            If ($($_.Exception.Message) -like "A parameter cannot be found that matches parameter name 'SPOManagementShell'.") {
                Write-Log -Level Warn "Cannot authenticate using PnP and SharePoint Online Management Shell. Please check which version of SharePoint PnP PowerShell is installed, and that the SharePoint Online Management Shell is installed."
            }
            Throw
        }

        Catch [System.Net.WebException] {
            If ($($_.Exception.Message) -like "*(401) Unauthorized*") {
                Write-Log -Level Warn "Cannot authenticate with SharePoint Admin Site. Please check if an authentication prompt appeared on your machine prior to the last interaction with this script."
            }
            Throw
        }
    }

    #Facilitates connection to the SharePoint Online site collections through an OAUTH 2.0 token
    function ConnectToSharePointOnlineOAuth([string]$rootSharePointUrl) {
        #Prompt for SharePoint Root Site Url
        If ($rootSharePointUrl -eq "") {
            $rootSharePointUrl = Read-Host -Prompt "Please enter the URL of your SharePoint Online Root Site Collection, eg 'https://contoso.sharepoint.com'."
            $rootSharePointUrl = $rootSharePointUrl.Trim()

            If (-not $rootSharePointUrl.Contains('sharepoint')) {
                $rootSharePointUrl = "https://" + $rootSharePointUrl + ".sharepoint.com"
            }
        } 

        Write-Host "Please authenticate against your Office 365 tenant by pasting the code copied to your clipboard and signing in. App access must be granted to the Office 365 PnP Management Shell to continue." -ForegroundColor Green  
        Try {
            Connect-PnPOnline -url $rootSharePointUrl -PnPO365ManagementShell -LaunchBrowser
        }
        Catch {
            $exMessage = $($_.Exception.Message)
            #These messages can be ignored. If we have an empty token we will throw an exception further down
            If (($exMessage -notmatch 'The handle is invalid') -and ($exMessage -notmatch 'Object reference not set to an instance of an object')) {
                Throw
            }
        }

        #workaround for PnP handling the token. Login to the root site normally and retrieve the token from there, then we will test it.
        Write-Host "`nEnter SharePoint credentials(your email address for SharePoint Online):`n" -ForegroundColor Green
        Connect-PnPOnline -url $rootSharePointUrl -UseWebLogin
        $script:token = Get-PnPAccessToken
        Disconnect-PnPOnline

        Write-Host "Testing OAuth token to ensure we don't have an issue later...`n"
        Connect-PnPOnline -url $rootSharePointUrl -AccessToken $script:token
        $web = Get-PnPWeb
        If ($null -ne $web) {
            Write-Host "Success!`n" -ForegroundColor Green
        }
    }

    #Facilitates connection to the on premises site collections through the root site collection
    function ConnectToSharePointOnPremises([string]$rootsite) {
        $script:isSPOnline = $false
        #Prompt for SharePoint Root Site Url     
        If ($rootsite -eq "") {
            $rootsite = Read-Host -Prompt "Please enter the URL of your on premises SharePoint root site collection"
        }
        
        Write-Host "Enter SharePoint credentials(your domain\username login for Sharepoint):" -ForegroundColor Green
        $tempCred = Get-Credential -Credential $null
        $script:onPremisesCred = $tempCred
        Connect-PnPOnline -url $rootsite -Credentials $script:onPremisesCred | Out-Null
    }

    #Creates the Email Columns in the given Site Collection. Taken from the existing OnePlaceSolutions Email Column deployment script
    function CreateEmailColumns([string]$siteCollection) {
        If ($siteCollection -eq "") {
            $siteCollection = Read-Host -Prompt "Please enter the Site Collection URL to add the OnePlace Solutions Email Columns to"
        }
        
        Try {
            $tempColumns = Get-PnPField -Group $script:groupName
            $emailColumnCount = 0
            ForEach ($col in $tempColumns) {
                If (($col.InternalName -match 'Em') -or ($col.InternalName -match 'Doc')) {
                    $emailColumnCount++
                }
            }
        }
        Catch {
            #This is fine, we will just try to add the columns anyway
            Write-Host "Couldn't check email columns, will attempt to add them anyway..." -ForegroundColor Yellow
        }

        #Check if we have 35 columns in our Column Group
        If ($emailColumnCount -eq 35) {
            Write-Host "All Email columns already present in group '$script:groupName', skipping adding."
        }
        #Create the Columns if we didn't find 35
        Else {
            $script:columnsXMLPath = "$env:temp\email-columns.xml"
            If (-not (Test-Path $script:columnsXMLPath)) {
                #From 'https://github.com/OnePlaceSolutions/EmailColumnsPnP/blob/master/installEmailColumns.ps1'
                #Download xml provisioning template
                $WebClient = New-Object System.Net.WebClient
                $Url = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
                
                Write-Host "Downloading provisioning xml template:" $script:columnsXMLPath -ForegroundColor Green 
                $WebClient.DownloadFile( $Url, $script:columnsXMLPath )
            }

            #Apply xml provisioning template to SharePoint
            Write-Host "Applying email columns template to SharePoint:" $siteCollection -ForegroundColor Green 
        
            $rawXml = Get-Content $script:columnsXMLPath
        
            #To fix certain compatibility issues between site template types, we will just pull the Field XML from the template
            ForEach ($line in $rawXml) {
                Try {
                    If (($line.ToString() -match 'Name="Em') -or ($line.ToString() -match 'Name="Doc')) {
                        $fieldAdded = Add-PnPFieldFromXml -fieldxml $line -ErrorAction Stop | Out-Null
                    }
                }
                Catch {
                    Write-Host $_.Exception.Message
                }
            }
        }
    }

    function CreateEmailView([string]$library, [string]$web) {
        Try {
            If ($script:createDefaultViews) {
                Try {
                    $view = Get-PnPView -List $libName -Identity $script:emailViewName -Web $web -ErrorAction Stop
                    Write-Host "View '$script:emailViewName' in Document Library '$libName' already exists, skipping." -ForegroundColor Green
                }
                Catch [System.NullReferenceException]{
                    #View does not exist, this is good
                    Write-Host "Adding Default View '$script:emailViewName' to Document Library '$libName'." -Foregroundcolor Yellow
                    $view = Add-PnPView -List $libName -Title $script:emailViewName -Fields $script:emailViewColumns -SetAsDefault -RowLimit 100 -Web $web -ErrorAction Continue
                    #Let SharePoint catch up for a moment
                    Start-Sleep -Seconds 2
                    $view = Get-PnPView -List $libName -Identity $script:emailViewName -Web $web -ErrorAction Stop
                    Write-Host "Success" -ForegroundColor Green 
                }
                Catch{
                    Throw
                }
            }
        }
        Catch {
            Write-Host "Error checking/creating Default View '$script:emailViewName' to Document Library '$libName'. Details below." -ForegroundColor Red
            $_
            Write-Host "`nContinuing Script...`n"
        }
    }
    #Starting menu for selection between SharePoint Online or SharePoint On-Premises, or exiting the script
    function showEnvMenu { 
        Clear-Host 
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
        Write-Host 'Welcome to the OnePlace Solutions Content Type Deployment Script' -ForegroundColor Green
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
        Write-Host 'Please make a selection:' -ForegroundColor Yellow
        Write-Host "1: SharePoint Online (365)" 
        Write-Host "2: SharePoint On-Premises (2016/2019)"
        Write-Host "Q: Press 'Q' to quit." 
    }

    #Menu to check if the user wants us to create a default Email View in the Document Libraries
    function emailViewMenu {
        $script:emailViewName = $null
        do { 
            Write-Host "Would you like to create an Email View in your Document Libraries?"
            Write-Host "N: No" 
            Write-Host "Y: Yes"
            Write-Host "Q: Press 'Q' to quit."  
            $input = Read-Host "Please select an option" 
            switch ($input) { 
                'N' {
                    $script:createDefaultViews = $false
                }
                'Y' {
                    $script:createDefaultViews = $true
                    $script:emailViewName = Read-Host -Prompt "Please enter the name for the Email View to be created (leave blank for default 'Emails')"

                    If ($script:emailViewName.Length -eq 0) { $script:emailViewName = "Emails" }
                    Write-Host "View will be created with name '$script:emailViewName' in listed Document Libraries in the CSV"
                }
                'q' { Exit }
            }
        } 
        until(($input -eq 'q') -or ($script:createDefaultViews -ne $null))

        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
    }

    function emailColumnsMenu {
        $script:groupName = $null
        do { 
            Write-Host "Would you like to automatically add the OnePlaceMail Email Columns to the listed Site Collections?"
            Write-Host "N: No" 
            Write-Host "Y: Yes"
            Write-Host "Q: Press 'Q' to quit."  
            $input = Read-Host "Please select an option" 
            $input = $input[0]

            switch ($input) { 
                'N' {
                    $script:createEmailColumns = $false
                    #Get the Group name containing the OnePlaceMail Email Columns for use later per site, default is 'OnePlaceMail Solutions'
                    $script:groupName = Read-Host -Prompt "Please enter the Group name containing the OnePlaceMail Email Columns in your SharePoint Site Collections (leave blank for default 'OnePlace Solutions')"
                    If ($script:groupName.Length -eq 0) { $script:groupName = "OnePlace Solutions" }
                    Write-Host "Will check for columns under group '$script:groupName'"
                }
                'Y' {
                    $script:createEmailColumns = $true
                    #Get the Group name we will create the OnePlaceMail Email Columns in for use later per site, default is 'OnePlaceMail Solutions'
                    $script:groupName = Read-Host -Prompt "Please enter the Group name to create the OnePlaceMail Email Columns in, in your SharePoint Site Collections (leave blank for default 'OnePlace Solutions')"
                    If ($script:groupName.Length -eq 0) { $script:groupName = "OnePlace Solutions" }
                    Write-Host "Will create and check for columns under group '$script:groupName'"
                }
                'q' { Exit }
            }
        } 
        until(($input -eq 'q') -or ($null -ne $script:createEmailColumns))

        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
    }

    function Deploy {
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red

        emailColumnsMenu
        emailViewMenu

        #Go through our siteCol objects in siteColsHT
        ForEach ($site in $script:siteColsHT.Values) {
            $siteName = $site.name
            $siteWeb = $site.web
            $filler = "Working with Site Collection: $siteName"
            Write-Host $filler -ForegroundColor Yellow
            Write-Log -Level Info -Message $filler
            $filler = "Working with Web: $siteWeb"
            Write-Host $filler -ForegroundColor Yellow
            Write-Log -Level Info -Message $filler
            #Authenticate against the Site Collection we are currently working with

            Try {
                If ($script:isSPOnline -and (-not $script:usingTokenAuth)) {
                    Write-Log -Level Info -Message "Connecting using SPOMS Auth"
                    Connect-pnpOnline -url $site.url -SPOManagementShell
                    Start-Sleep -Seconds 2
                }
                ElseIf ($script:isSPOnline -and $script:usingTokenAuth) {
                    Write-Log -Level Info -Message "Connecting using Token Auth"
                    Connect-PnPOnline -url $site.url -AccessToken $script:token
                }
                Else {
                    Write-Log -Level Info -Message "Connecting using On Prem Auth"
                    Connect-pnpOnline -url $site.url -Credentials $script:onPremisesCred
                }
                #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
                Start-Sleep -seconds 3
                Get-PnPWeb -ErrorAction Continue
                Write-Log -Level Info -Message "Authenticated"
            }
            Catch {
                Write-Host "Error connecting to SharePoint Site Collection '$siteName'. Is this URL correct?" -ForegroundColor Red
                $site.url
                Write-Host "Other Details below. Halting script." -ForegroundColor Red
                Throw
            }

            #Check if we are creating email columns, if so, do so now
            If ($script:createEmailColumns) {
                Write-Log -Level Info -Message "User has opted to create email columns"
                CreateEmailColumns -siteCollection $site.url
            }

            #Retrieve all the columns/fields for the group specified in this Site Collection, we will add these to the named Content Types shortly. If we do not get the site columns, skip this Site Collection
            $script:emailColumns = Get-PnPField -Group $script:groupName
            If (($null -eq $script:emailColumns) -or ($null -eq $script:groupName)) {
                $filler = "Email Columns not found in Site Columns group '$script:groupName' for Site Collection '$siteName'. Skipping."
                Write-Log -Level Warn -Message $filler
                Pause
                Continue
            }
            Write-Host "Columns found for group '$script:groupName':"
            $script:emailColumns | Format-Table
            Write-Host "These Columns will be added to the Site Content Types listed in your CSV file."
            
            #Get the Content Type Object for 'Document' from SP, we will use this as the parent Content Type for our email Content Type
            $DocCT = Get-PnPContentType -Identity "Document"
            If ($null -eq $DocCT) {
                $filler = "Couldn't get 'Document' Site Content Type in $siteName. Skipping Site Collection: $siteName"
                Write-Log -Level Warn -Message $filler
                Pause
                Continue
            }
            #For each Site Content Type listed for this siteCol/Site Collection, try and create it and add the email columns to it
            ForEach ($ct in $site.contentTypes) {
                Try {
                    $filler = "Checking if Content Type '$ct' already exists"
                    Write-Host $filler -ForegroundColor Yellow
                    Write-Log -Level Info -Message $filler
                    $foundContentType = Get-PnPContentType -Identity $ct
                
                    #If Content Type object returned is null, assume Content Type does not exist, create it. 
                    #If it does exist and we just failed to find it, this will throw exceptions for 'Duplicate Content Type found', and then continue.
                    If ($null -eq $foundContentType) {
                        $filler = "Couldn't find Content Type '$ct', might not exist"
                        Write-Host $filler -ForegroundColor Red
                        Write-Log -Level Info -Message $filler
                        #Creating content type
                        Try {
                            $filler = "Creating Content Type '$ct' with parent of 'Document'"
                            Write-Host $filler -ForegroundColor Yellow
                            Write-Log -Level Info -Message $filler
                            Add-PnPContentType -name $ct -Group "Custom Content Types" -ParentContentType $DocCT -Description "Email Content Type"
                        }
                        Catch {
                            Write-Host "Error creating Content Type '$ct' with parent of Document. Details below. Halting script." -ForegroundColor Red
                            Throw
                        } 
                    }
                }
                Catch {
                    Write-Host "Error checking for existence of Content Type '$ct'. Details below. Halting script." -ForegroundColor Red
                    Throw
                }

                #Try adding columns to the Content Type
                Try {
                    $filler = "Adding email columns to Site Content Type '$ct'"
                    Write-Host $filler -ForegroundColor Yellow
                    Write-Log -Level Info -Message $filler
                    Start-Sleep -Seconds 2
                    $numColumns = $script:emailColumns.Count
                    $i = 0
                    $emSubjectFound = $false
                    ForEach ($column in $script:emailColumns) {
                        $column = $column.InternalName
                        If ($column -eq 'EmSubject') {
                            $emSubjectFound = $true
                        }
                        Add-PnPFieldToContentType -Field $column -ContentType $ct
                        Write-Progress -Activity "Adding column: $column" -Status "To Site Content Type: $ct in Site Collection: $siteName. Progress:" -PercentComplete ($i / $numColumns * 100)
                        $i++
                    }
                    If (($false -eq $emSubjectFound) -or ($numColumns -ne 35)) {
                        Throw "Not all Email Columns present. Please check you have added the columns to the Site Collection or elected to do so when prompted with this script."
                    }
                    Write-Progress -Activity "Done adding Columns" -Completed
                }
                Catch {
                    Write-Host "Error adding email columns to Site Content Type '$ct'. Details below. Halting script." -ForegroundColor Red
                    Throw
                }
            }

            #For each docLib/Document Library in our siteCol/Site Collection, get it's list of Content Types we want to add
            ForEach ($library in $site.documentLibraries.Values) {
                $libName = $library.name
                $filler = "`nWorking with Document Library: $libName" 
                Write-Host $filler -ForegroundColor Yellow
                Write-Log -Level Info -Message $filler
                Write-Host "Which has Content Types:" -ForegroundColor Yellow
                $library.contentTypes | Format-Table

                Write-Host "`nEnabling Content Type Management in Document Library '$libName'." -ForegroundColor Yellow
                Set-PnPList -Identity $libName -EnableContentTypes $true -Web $site.web

                #For each Site Content Type listed for this docLib/Document Library, try to add it to said Document Library
                Try {
                    ForEach ($ct in $library.contentTypes) {
                        $filler = "Adding Site Content Type '$ct' to Document Library '$libName'..."
                        Write-Host $filler -ForegroundColor Yellow
                        Write-Log -Level Info -Message $filler
                        Add-PnPContentTypeToList -List $libName -ContentType $ct -Web $site.web
                    }
                }
                Catch {
                    Write-Host "Error adding Site Content Type '$ct' to Document Library '$libName'. Details below. Continuing script." -ForegroundColor Red
                    $_
                }

                CreateEmailView -library $libName -web $site.web
            }
            Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red

            Try {
                If ($script:usingTokenAuth) {
                    #refresh our token if we are using one
                    $script:token = Get-PnPAccessToken
                    Write-Log -Level Info -Message "Refreshing Access Token"
                }
            }
            Catch {
                Write-Log -level Warn -Message "Failed to refresh PnP Auth Token, will attempt to continue:`n"
                Write-Log -level Info -Message $_
            }
        
        }   
        Write-Host "Deployment complete! Please check your SharePoint Environment to verify completion. If you would like to copy the output above, do so now before pressing 'Enter'." 
        Write-Log -Level Info -Message "Deployment complete." 
    }

    #Start of Script
    #----------------------------------------------------------------

    do {
        showEnvMenu 
        $input = Read-Host "Please select an option" 
        switch ($input) { 
            '1' {
                #Online
                Write-Log -Level Info -Message "User has selected Option 1 for SPO."
                Clear-Host

                If((($pnpModule -notlike "*Online*") -or ($null -eq $pnpModule)) -and (($pnpMSI -notlike "*Online*") -or ($null -eq $pnpMSI))){
                    Write-Log -Level Warn -Message "SharePoint Online selected for deployment, but SharePoint Online PnP CmdLets not installed. Please check installed version before continuing."
                    Pause
                }

                #Start with getting the CSV file of Site Collections, Document Libraries and Content Types
                EnumerateSitesDocLibs
                #Connect to SharePoint Online, using SharePoint Management Shell against the Admin site
                If("" -ne $script:extractedTenant) {
                    Write-Host "`nExtracted Tenant Name '$script:extractedTenant' from CSV, is this correct?"
                    Write-Host "N: No" 
                    Write-Host "Y: Yes"
                    $otherInput = Read-Host "Please select an option" 
                    If($otherInput[0] -eq 'Y') {
                        Write-Log -Level Info -Message "User has confirmed extracted Tenant name."
                        ConnectToSharePointOnlineAdmin -Tenant $script:extractedTenant
                    }
                    Else {
                        ConnectToSharePointOnlineAdmin
                    }
                }
                Else {
                    ConnectToSharePointOnlineAdmin
                }
                
                Deploy
            }
            's' {
                #Online
                Write-Log -Level Info -Message "User has selected Option S for skipping Admin Site pre-auth."
                Clear-Host
                #Start with getting the CSV file of Site Collections, Document Libraries and Content Types
                EnumerateSitesDocLibs
                #Skip connecting to the Admin Site, we will automatically connect using SPO management shell when required.
                #This will possibly not prompt for credentials
                Deploy
            }
            't' {
                #Online
                Write-Log -Level Info -Message "User has selected Option T for token auth."
                Clear-Host
                #Start with getting the CSV file of Site Collections, Document Libraries and Content Types
                EnumerateSitesDocLibs
                #Connect to SharePoint Online, using token based login to iterate over the site collections
                $script:usingTokenAuth = $true
                ConnectToSharePointOnlineOAuth
                Deploy
            }
            '2' {
                #On-Premises
                Write-Log -Level Info -Message "User has selected Option 2 for SP On Prem."
                Clear-Host

                If(($pnpModule -like "*Online*") -or ($pnpMSI -like "*Online*")){
                    Write-Log -Level Warn -Message "SharePoint On-Premises selected for deployment, but SharePoint Online PnP CmdLets installed. Please check installed version before continuing."
                    Pause
                }

                #Start with getting the CSV file of Site Collections, Document Libraries and Content Types
                EnumerateSitesDocLibs
                #Connect to SharePoint On-Premises, specifically the root site so we can iterate over the site collections
                ConnectToSharePointOnPremises
                Deploy
            }
            'c' {
                #clear logins
                Write-Log -Level Info -Message "User has selected Option C for clear logins."
                Clear-Host
                Try {
                    Disconnect-PnPOnline
                    Write-Host "Cleared PnP Connection!"
                }
                Catch{}
                Try {
                    Disconnect-SPOService
                    Write-Host "Cleared SPO Management Shell Connection!"
                }
                Catch{}
                Start-Sleep -Seconds 3
                showEnvMenu
            }
            'q' { return }
        } 
        pause 
    } 
    until($input -eq 'q') {
    }
}
Catch {
    $exType = $($_.Exception.GetType().FullName)
    $exMessage = $($_.Exception.Message)
    Write-Host "`nCaught an exception, further debugging information below:" -ForegroundColor Red
    Write-Log -Level Error -Message "Caught an exception. Exception Type: $exType. $exMessage"
    Write-Host $exMessage -ForegroundColor Red
    Write-Host "`nPlease send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance." -ForegroundColor Yellow
    Pause
}
Finally {
    Try {
        Disconnect-SPOService
    }
    Catch {
        #Sink everything, this is just trying to tidy up
    }
    Try {
        Disconnect-PnPOnline
    }
    Catch {
        #Sink everything, this is just trying to tidy up
    }
}
