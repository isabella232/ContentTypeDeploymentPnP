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

    #Check for module versions of PnP
    Try {
        Write-Host "Checking if PnP installed via Module..." -ForegroundColor Cyan
        $pnpModule = Get-InstalledModule "PnP.PowerShell" | Select-Object Name, Version
    }
    Catch {
        #Couldn't check PNP or SPOMS Module versions, Package Manager may be absent
    }
    Finally {
        Write-Log -Level Info -Message "PnP Module Installed: $pnpModule"
    }


    #Contains all our Site Collections as siteCol objects
    $script:siteColsHT = @{ }

    #Flag for whether we create  email views or not, and if so what name to use
    [boolean]$script:createEmailViews = $false
    [string]$script:emailViewName = "Emails"
    [boolean]$script:emailViewDefault = $true

    #Flag for whether we automatically create the OnePlaceMail Email Columns
    [boolean]$script:createEmailColumns = $true

    #Name of Column group containing the Email Columns, and an object to contain the Email Columns
    [string]$script:groupName = "OnePlace Solutions"
    $script:emailColumns = $null

    $script:csvFilePath = ""

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
            $script:csvFilePath = $csvFile
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

    #Facilitates connection to the SharePoint Online site collections through the PnP Management Shell
    function ConnectToSharePointOnlineAdmin([string]$tenant) {
        #Prompt for SharePoint Root Site Url     
        Try {
            If ($tenant -eq "") {
                $rootSharePointUrl = Read-Host -Prompt "Please enter your SharePoint Online Root Site Collection URL, eg (without quotes) 'https://contoso.sharepoint.com'"
                Write-Log -Level Info -Message "Root SharePoint: $rootSharePointUrl"
                $rootSharePointUrl = $rootSharePointUrl.Trim("'")
                $rootSharePointUrl = $rootSharePointUrl.Trim("/")
                Write-Log -Level Info -Message "Sanitized: $rootSharePoint"
            }
            Else {
                $rootSharePointUrl = "https://$tenant.sharepoint.com"
            }
            #Connect to site collection
        
            Write-Host "Prompting for PnP Management Shell Authentication. Please copy the code displayed into the browser as directed and log in." -ForegroundColor Green
            Connect-PnPOnline -Url $rootSharePointUrl -PnPManagementShell -LaunchBrowser
            #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
            Start-Sleep -Seconds 3
            Pause
            
            $filler = "Testing connection with 'Get-PnPWeb'..."
            Write-Log -Level Info -Message $filler
            Write-Host $filler
            Get-PnPWeb
        }
        Catch [System.Net.WebException] {
            If ($($_.Exception.Message) -like "*(401) Unauthorized*") {
                Write-Log -Level Warn "Cannot authenticate with SharePoint Admin Site. Please check if an authentication prompt appeared on your machine prior to the last interaction with this script."
            }
            ElseIf ($($_.Exception.Message) -like "*(403) Forbidden*") {
                Write-Log -Level Warn "Cannot login to SharePoint Admin Site, Access Denied. Please check the permissions of the credentials/account you are using to authenticate with."
            }
            Throw
        }
        Catch {
            Throw
        }
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
            If ($script:createEmailViews) {
                Try {
                    $view = Get-PnPView -List $libName -Identity $script:emailViewName -Web $web -ErrorAction Stop
                    Write-Host "View '$script:emailViewName' in Document Library '$libName' already exists, skipping." -ForegroundColor Green
                }
                Catch [System.NullReferenceException]{
                    #View does not exist, this is good
                    Write-Host "Adding Email View '$script:emailViewName' to Document Library '$libName'." -Foregroundcolor Yellow
                    If($script:emailViewDefault) {
                        $view = Add-PnPView -List $libName -Title $script:emailViewName -Fields $script:emailViewColumns -SetAsDefault -RowLimit 100 -Web $web -ErrorAction Continue
                    }
                    Else {
                        Write-Host "Email View will be created as default view"
                        $view = Add-PnPView -List $libName -Title $script:emailViewName -Fields $script:emailViewColumns -RowLimit 100 -Web $web -ErrorAction Continue
                    }
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
        Write-Host 'Please make a selection to set, toggle, change or execute:' -ForegroundColor Yellow
        Write-Host "1: Select CSV file. Path: $($script:csvFilePath)"
        Write-Host "2: Enable Email Column Creation: $($script:createEmailColumns)"
        Write-Host "3: Email Column Group: $($script:groupName)"
        Write-Host "4: Enable Email View Creation: $($script:createEmailViews)"
        Write-Host "5: Email View Name: $($script:emailViewname)"
        Write-Host "6: Set View '$($script:emailViewName)' as default: $($script:emailViewDefault)"
        Write-Host "7: Deploy"
        Write-Host "Q: Press 'Q' to quit."
    }

    #Menu to check if the user wants us to create a default Email View in the Document Libraries

    function Deploy {
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red

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
                Write-Log -Level Info -Message "Connecting to $siteName using PnP Management Shell"
                Connect-pnpOnline -url $site.url -PnPManagementShell
                Start-Sleep -Seconds 5
                #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
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
                Write-Log "User has selected Option $($input)"
                EnumerateSitesDocLibs
            }
            '2' {
                Write-Log "User has selected Option $($input)"
                $script:createEmailColumns = -not $script:createEmailColumns
            }
            '3' {
                Write-Log "User has selected Option $($input)"
                $script:groupName = Read-Host -Prompt "`nPlease enter a Group name to check for or create Email columns in. Current value: $($script:groupName)"
                If("" -eq $script:groupName) {
                    $script:groupName = "OnePlace Solutions"
                }
            }
            '4' {
                Write-Log "User has selected Option $($input)"
                $script:createEmailViews = -not $script:createEmailViews
            }
            '5' {
                Write-Log "User has selected Option $($input)"
                $script:emailViewname = Read-Host -Prompt "`nPlease enter a View name for the email view (if being created). Current value: $($script:emailViewname)"
                If("" -eq $script:emailViewname) {
                    $script:emailViewname = "Emails"
                }
            }
            '6' {
                Write-Log "User has selected Option $($input)"
                $script:emailViewDefault = -not $script:emailViewDefault
            }
            '7' {
                Write-Log "User has selected Option $($input)"
                Clear-Host
                If("" -ne $script:csvFilePath) {
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
                Else {
                    Write-Log -Level Warn -Message "No CSV file has been defined. Please select Option 1 and select your CSV file."
                    Pause
                }
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
                Start-Sleep -Seconds 3
                showEnvMenu
            }
            'q' { return }
        }
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
        Disconnect-PnPOnline
    }
    Catch {
        #Sink everything, this is just trying to tidy up any open PnP connections
    }
}
