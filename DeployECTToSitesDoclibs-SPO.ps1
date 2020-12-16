<#
        This script applies the OnePlaceMail Email Columns to an existing site collection, creates Site Content Types, adds them to Document Libraries and creates a default view.
        Please check the README.md on Github before using this script.
#>
$ErrorActionPreference = 'Stop'

#Columns to add to the Email View if we are creating one. Edit as required based on Internal Naming
[string[]]$script:emailViewColumns = @("EmHasAttachments","EmSubject","EmTo","EmDate","EmFromName")

[string]$script:logFile = "OPSScriptLog.txt"
[string]$script:logPath = "$env:userprofile\Documents\$script:logFile"

[string]$script:extractedTenant = ""

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
        [switch]$NoOutput = $false,

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
                If($NoOutput) {
                    Write-Verbose $Message
                }
                Else {
                    Write-Verbose $Message
                    Write-Host $Message -ForegroundColor Yellow
                }
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

    #Check for module versions of PnP
    Try {
        Write-Log "Checking if PnP installed via Module..."
        $pnpModule = Get-InstalledModule "PnP.PowerShell" | Select-Object Name, Version
    }
    Catch {
        #Couldn't check PNP Module versions, Package Manager may be absent
    }
    Finally {
        Write-Log "PnP Module Installed: $pnpModule"
    }


    #Contains all our Site Collections as siteCol objects
    $script:siteColsHT = @{ }

    #Flag for whether we create  email views or not, and if so what name to use
    [boolean]$script:createEmailViews = $false
    [string]$script:emailViewName = "Emails"
    [boolean]$script:emailViewDefault = $false

    #Flag for whether we automatically create the OnePlaceMail Email Columns
    [boolean]$script:createEmailColumns = $true

    #Name of Column group containing the Email Columns, and an object to contain the Email Columns
    [string]$script:columnGroupName = "OnePlace Solutions"
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
            Write-Log "Creating siteCol object with name '$($this.name)'"

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
            Write-Log "Is this a subsite? $($this.isSubSite)"
            $this.url = $rootUrl
        }

        [boolean]connect() {
            Try {
                #Connect to site collection
                Connect-PnPOnline -Url $this.url -PnPManagementShell
                #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
                Start-Sleep -Seconds 3
                Try {
                    Write-Log "Testing connection with 'Get-PnPWeb'..."
                    $currentWeb = Get-PnPWeb -ErrorAction Stop
                    Write-Log "Connected to $($currentWeb.Title)"
                }
                Catch {
                    Write-Log -Level Warn -Message "Error testing connection with 'Get-PnPWeb'. Logging but continuing"
                }
                Return $true
            }
            Catch [System.Net.WebException] {
                If ($($_.Exception.Message) -like "*(401)*") {
                    Write-Log -Level Warn "Cannot authenticate with SharePoint Site at $($this.url). Please check if an authentication prompt appeared on your machine prior to the last interaction with this script."
                }
                ElseIf ($($_.Exception.Message) -like "*(403)*") {
                    Write-Log -Level Warn "Cannot login to SharePoint Site at $($this.url), Access Denied. Please check the permissions of the credentials/account you are using to authenticate with, and that the Site Collection is not still being provisioned by Microsoft 365."
                }
                Return $false
            }
            Catch {
                Write-Log -Level Error -Message "Unhandled exception encountered: $($_.Exception.Message)"
                Return $false
            }
        }

        [void]addContentTypeToDocumentLibraryObj($contentTypeName, [string]$docLibName) {
            #Check we aren't working without a Document Library name, otherwise assume that we just want to add a Site Content Type
            If (($null -ne $docLibName) -and ($docLibName -ne "")) {
                If ($this.documentLibraries.ContainsKey($docLibName)) {
                    $this.documentLibraries.$docLibName
                    Write-Log "Document Library '$docLibName' already listed for this Site Collection."
                }
                Else {
                    $tempDocLib = [docLib]::new($docLibName,$($this.web))
                    $this.documentLibraries.Add($docLibName, $tempDocLib)
                    Write-Log "Document Library '$docLibName' not listed for this Site Collection, added."
                }
                
                $this.documentLibraries.$docLibName.addContentTypeToObj($contentTypeName)
                Write-Log "Listing Content Type '$contentTypeName' for Document Library '$docLibName'."
            }
            
            #If the named Content Type is not already listed in Site Content Types, add it to the Site Content Types
            If (-not $this.contentTypes.Contains($contentTypeName)) {
                $this.contentTypes += $contentTypeName
                Write-Log "Content Type '$contentTypeName' not listed in Site Content Types, adding."
            }
            Else {
                Write-Log "Content Type '$contentTypeName' listed in Site Content Types."
            }
        }

        [void]addContentTypesToDocumentLibariesSPO() {
            Write-Log "Adding the Content Types to the Document Libraries in SPO..."
            ForEach($lib in $this.documentLibraries.Values) {
                $lib.addContentTypeInSPO()
            }
        }

        #Creates the Email Columns in the given Site Collection. Taken from the existing OnePlaceSolutions Email Column deployment script
        [void]createEmailColumns() {
            Try {
                $tempColumns = Get-PnPField -Group $script:columnGroupName
                $script:emailColumnCount = 0
                ForEach ($col in $tempColumns) {
                    If (($col.InternalName -match 'Em') -or ($col.InternalName -match 'Doc')) {
                        $script:emailColumnCount++
                    }
                }
            }
            Catch {
                #This is fine, we will just try to add the columns anyway
                Write-Log -Level Warn -Message "Couldn't check email columns, will attempt to add them anyway..."
            }

            #Check if we have 35 columns in our Column Group
            If ($script:emailColumnCount -eq 35) {
                Write-Log "All Email columns already present in group '$script:columnGroupName', skipping adding."
            }
            #Create the Columns if we didn't find 35
            Else {
                $script:columnsXMLPath = "$env:temp\email-columns.xml"
                If (-not (Test-Path $script:columnsXMLPath)) {
                    #From 'https://github.com/OnePlaceSolutions/EmailColumnsPnP/blob/master/installEmailColumns.ps1'
                    #Download xml provisioning template
                    $WebClient = New-Object System.Net.WebClient
                    $downloadUrl = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
                
                    Write-Log "Downloading provisioning xml template:" $script:columnsXMLPath
                    $WebClient.DownloadFile( $downloadUrl, $script:columnsXMLPath )
                }

                #Apply xml provisioning template to SharePoint
                Write-Log "Applying email columns template to SharePoint: $($this.url)"
        
                $rawXml = Get-Content $script:columnsXMLPath
        
                #To fix certain compatibility issues between site template types, we will just pull the Field XML entries from the template
                ForEach ($line in $rawXml) {
                    Try {
                        If (($line.ToString() -match 'Name="Em') -or ($line.ToString() -match 'Name="Doc')) {
                            $fieldAdded = Add-PnPFieldFromXml -fieldxml $line -ErrorAction Stop | Out-Null
                        }
                    }
                    Catch {
                        If($($_.Exception.Message) -match 'duplicate') {
                            Write-Log -Level Warn -Message "Duplicate fields detected. $($_.Exception.Message). Continuing script."
                        }
                        Else {
                            Write-Log -Level Error -Message "Error creating Email Column. Error: $($_.Exception.Message)"
                        }
                    }
                }
            }
        }
        
        #Retrieve all the columns/fields for the group specified in this Site Collection, we will add these to the named Content Types shortly. If we do not get the site columns, skip this Site Collection
        [void]createContentTypes() {
            $script:emailColumns = Get-PnPField -Group $script:columnGroupName
            If (($null -eq $script:emailColumns) -or ($null -eq $script:columnGroupName)) {
                Write-Log -Level Warn -Message "Email Columns not found in Site Columns group '$script:columnGroupName' for Site Collection '$($this.name)'. Skipping."
            }
            Else {
                Write-Log "Columns found for group '$script:columnGroupName':"
                $script:emailColumns | Format-Table
                Write-Host "These Columns will be added to the Site Content Types extracted from your CSV file:"
                Write-Log "$($this.contentTypes)"
                $this.contentTypes | Format-Table
            
                #Get the Content Type Object for 'Document' from SP, we will use this as the parent Content Type for our email Content Type
                $DocCT = Get-PnPContentType -Identity "Document"
                If ($null -eq $DocCT) {
                    Write-Log -Level Warn -Message "Couldn't get 'Document' Parent Site Content Type in $($this.name). Skipping this Site Collection."
                }
                #For each Site Content Type listed for this siteCol/Site Collection, try and create it and add the email columns to it
                Else {
                    ForEach ($ct in $this.contentTypes) {
                        Try {
                            Write-Log "Checking if Content Type '$ct' already exists"
                            $foundContentType = Get-PnPContentType -Identity $ct
                
                            #If Content Type object returned is null, assume Content Type does not exist, create it. 
                            #If it does exist and we just failed to find it, this will throw exceptions for 'Duplicate Content Type found', and then continue.
                            If ($null -eq $foundContentType) {
                                Write-Log "Couldn't find Content Type '$ct', might not exist"
                                #Creating content type
                                Try {
                                    Write-Log "Creating Content Type '$ct' with parent of 'Document'"
                                    Add-PnPContentType -name $ct -Group "Custom Content Types" -ParentContentType $DocCT -Description "Email Content Type"
                                }
                                Catch {
                                    Write-Log -Level Error -Message "Error creating Content Type '$ct' with parent of Document. Details below."
                                } 
                            }
                        }
                        Catch {
                            Write-Log -Level Error -Message "Error checking for existence of Content Type '$ct'."
                        }

                        #Try adding columns to the Content Type
                        Try {
                            Write-Log "Adding email columns to Site Content Type '$ct'"
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
                                Write-Progress -Activity "Adding column: $column" -Status "To Site Content Type: $ct in Site Collection: $($this.name). Progress:" -PercentComplete ($i / $numColumns * 100)
                                $i++
                            }
                            If (($false -eq $emSubjectFound) -or ($numColumns -ne 35)) {
                                Write-Log -Level Warn -Message "Not all Email Columns present. Please check you have added the columns to the Site Collection or elected to do so when prompted with this script."
                            }
                        }
                        Catch {
                            Write-Log -Level Error -Message "Error adding email columns to Site Content Type '$ct': $($_.Exception.Message)"
                        }
                        Finally {
                            Write-Progress -Activity "Done adding Columns" -Completed
                        }
                    }
                }
            }
        }
    }

    #Contains all the data we need relating to the Document Library we are working with, including the Site Content Type names we are adding to it
    class docLib {
        [String]$name
        [Array]$contentTypes
        [string]$web

        docLib([String]$name,[string]$web) {
            Write-Log "Creating docLib object with name $name"
            $this.name = $name
            $this.contentTypes = @()
            $this.web = $web
        }

        [void]addContentTypeToObj([string]$contentTypeName) {
            If (-not $this.contentTypes.Contains($contentTypeName)) {
                Write-Log "Adding Content Type '$contentTypeName' to '$($this.name)' Document Library Content Types"
                $this.contentTypes += $contentTypeName
            }
            Else {
                $temp = $this.name
                $filler = "Content Type '$contentTypeName' already listed in Document Library $temp"
                Write-Host $filler -ForegroundColor Red
                Write-Log -Level Info -Message $filler
            }
        }

        [void]addContentTypeInSPO() {
            Write-Log "`nWorking with Document Library: $($this.name)"
            Write-Host "Which has Content Types:" -ForegroundColor Yellow
            $this.contentTypes | Format-Table

            Write-Host "`nEnabling Content Type Management in Document Library '$($this.name)'." -ForegroundColor Yellow
            Set-PnPList -Identity $($this.name) -EnableContentTypes $true -Web $this.web

            #For each Site Content Type listed for this docLib/Document Library, try to add it to said Document Library
            $this.contentTypes | ForEach-Object {
                Try{
                    Write-Log "Adding Site Content Type '$($_)' to Document Library '$($this.name)'..."
                    Add-PnPContentTypeToList -List $($this.name) -ContentType $($_) -Web $this.web
                }
                Catch {
                    Write-Log -Level Error -Message "Error adding Site Content Type '$($_)' to Document Library '$($this.name)': $($_.Exception.Message)"
                }
            }

            If($script:createEmailViews) {
                $this.createEmailView($script:emailViewName)
            }
        }

        [void]createEmailView([string]$viewName) {
            Try {
                Try {
                    $view = Get-PnPView -List $this.name -Identity $viewName -Web $this.web -ErrorAction Stop
                    Write-Log "View '$viewName' in Document Library '$($this.name)' already exists, will set as Default View if required but otherwise skipping."
                    If($script:emailViewDefault) {
                        Set-PnPView -List $this.name -Identity $viewName -Values @{DefaultView =$True}
                    }
                }
                Catch [System.NullReferenceException]{
                    #View does not exist, this is good
                    Write-Log "Adding Email View '$viewName' to Document Library '$($this.name)'."
                    If($script:emailViewDefault) {
                        Write-Log "Email View will be created as default view..."
                        $view = Add-PnPView -List $this.name -Title $viewName -Fields $script:emailViewColumns -SetAsDefault -RowLimit 100 -Web $this.web -ErrorAction Continue
                    }
                    Else {
                        Write-Log "Email View will not be created as default view..."
                        $view = Add-PnPView -List $this.name -Title $viewName -Fields $script:emailViewColumns -RowLimit 100 -Web $this.web -ErrorAction Continue
                    }
                    #Let SharePoint catch up for a moment
                    Start-Sleep -Seconds 2
                    $view = Get-PnPView -List $this.name -Identity $viewName -Web $this.web -ErrorAction Stop
                    Write-Log "Email View $($viewName) created successfully. As Default? $($script:emailViewDefault)"
                }
                Catch{
                    Throw
                }
                
            }
            Catch {
                Write-Log -Level Error -Message "Error checking/creating View '$viewName' in Document Library '$($this.name)': $($_.Exception.Message)"
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
                    Write-Log "Extracted Tenant name '$script:extractedTenant'"
                }

                #Don't create siteCol objects that do not have a URL, this also accounts for empty lines at EOF
                If ($csv_siteUrl -ne "") {
                    #If a name is not defined, use the URL
                    If ($csv_siteName -eq "") { $csv_siteName = $element.SiteUrl }

                    If ($script:siteColsHT.ContainsKey($csv_siteUrl)) {
                        $script:siteColsHT.$csv_siteUrl.addContentTypeToDocumentLibraryObj($csv_contentType, $csv_docLib)
                    }
                    Else {
                        $newSiteCollection = [siteCol]::new($csv_siteName, $csv_siteUrl)
                        $newSiteCollection.addContentTypeToDocumentLibraryObj($csv_contentType, $csv_docLib)
                        $script:siteColsHT.Add($csv_siteUrl, $newSiteCollection)
                    }
                }
            }
            Write-Log "Completed Enumerating Site Collections and Document Libraries from CSV file!"
        }
        Catch {
            Write-Log -Level Error -Message "Error parsing CSV file. Is this filepath for a a valid CSV file?"
            $csvFile
            Throw
        }
        Pause
    }

    #Facilitates connection to the SharePoint Online site collections through the PnP Management Shell
    function ConnectToSharePointOnline([string]$tenant) {
        #Prompt for SharePoint Root Site Url     
        Try {
            If ($tenant -eq "") {
                $rootSharePointUrl = Read-Host -Prompt "Please enter your SharePoint Online Root Site Collection URL, eg (without quotes) 'https://contoso.sharepoint.com'"
                Write-Log "Root SharePoint URL entered: $rootSharePointUrl"
                $rootSharePointUrl = $rootSharePointUrl.Trim("'")
                $rootSharePointUrl = $rootSharePointUrl.Trim("/")
                Write-Log "Sanitized: $rootSharePointUrl"
            }
            Else {
                $rootSharePointUrl = "https://$tenant.sharepoint.com"
            }
            
            #Check we aren't already logged in to this SharePoint Tenant
            Try {
                $currentWeb = Get-PnPWeb -ErrorAction Continue
            }
            Catch{
                #No issue if we throw an exception here, we will just login later
            }

            If($currentWeb.url -ne $rootSharePointUrl) {
                #Connect to site collection
                Write-Host "Prompting for PnP Management Shell Authentication. Please copy the code displayed into the browser as directed and log in." -ForegroundColor Green
                $conn = Connect-PnPOnline -Url $rootSharePointUrl -PnPManagementShell -LaunchBrowser
                #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
                Write-Log "Testing connection with 'Get-PnPWeb'..."
                Start-Sleep -Seconds 3
                Get-PnPWeb -Connection $conn
            }
            Else {
                Write-Log "Already connected to SharePoint with Root URL '$rootSharePointUrl'. Skipping login"
            }
        }
        Catch [System.Net.WebException] {
            If ($($_.Exception.Message) -like "*(401) Unauthorized*") {
                Write-Log -Level Warn "Cannot authenticate with SharePoint Root Site. Please check if an authentication prompt appeared on your machine prior to the last interaction with this script."
            }
            ElseIf ($($_.Exception.Message) -like "*(403) Forbidden*") {
                Write-Log -Level Warn "Cannot login to SharePoint Root Site, Access Denied. Please check the permissions of the credentials/account you are using to authenticate with."
            }
            Throw
        }
        Catch {
            Throw
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
        Write-Host "3: Email Column Group: $($script:columnGroupName)"
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
        $i = 1
        $j = $script:siteColsHT.Values.Count
        ForEach ($site in $script:siteColsHT.Values) {
            Write-Log "Site Collection $i / $j"
            Write-Log "Working with Site Collection: $($site.name)"
            Write-Log "Working with Web: $($site.web)"
            
            #Authenticate against the Site Collection we are currently working with
            $connected = $site.connect()

            If($connected) {
                #Check if we are creating email columns, if so, do so now
                If ($script:createEmailColumns) {
                    Write-Log "User has opted to create email columns"
                    $site.createEmailColumns()
                }
                $site.createContentTypes()
                $site.addContentTypesToDocumentLibariesSPO()
            }
            Else {
                Write-Log "Issue connecting to SharePoint Online when iterating over Site Collection $($site.name) at URL $($site.url). Skipping"
            }
            $i++
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
                Write-Log "User has selected Option $($input)" -NoOutput
                EnumerateSitesDocLibs
            }
            '2' {
                Write-Log "User has selected Option $($input)" -NoOutput
                $script:createEmailColumns = -not $script:createEmailColumns
            }
            '3' {
                Write-Log "User has selected Option $($input)" -NoOutput
                $script:columnGroupName = Read-Host -Prompt "`nPlease enter a Group name to check for or create Email columns in. Current value: $($script:columnGroupName)"
                If([string]::IsNullOrWhiteSpace($script:columnGroupName)) {
                    $script:columnGroupName = "OnePlace Solutions"
                }
            }
            '4' {
                Write-Log "User has selected Option $($input)" -NoOutput
                $script:createEmailViews = -not $script:createEmailViews
            }
            '5' {
                Write-Log "User has selected Option $($input)" -NoOutput
                $script:emailViewname = Read-Host -Prompt "`nPlease enter a View name for the email view (if being created). Current value: $($script:emailViewname)"
                If([string]::IsNullOrWhiteSpace($script:emailViewname)) {
                    $script:emailViewname = "Emails"
                }
            }
            '6' {
                Write-Log "User has selected Option $($input)" -NoOutput
                $script:emailViewDefault = -not $script:emailViewDefault
            }
            '7' {
                Write-Log "User has selected Option $($input)" -NoOutput
                Clear-Host

                If([string]::IsNullOrWhiteSpace($script:csvFilePath)) {
                    Write-Log -Level Warn -Message "No CSV file has been defined. Please select Option 1 and select your CSV file."
                    Pause
                }
                Else {
                    #Connect to SharePoint Online, using PnP Management Shell
                    If("" -ne $script:extractedTenant) {
                        Write-Host "`nExtracted Tenant Name '$script:extractedTenant' from CSV, is this correct?"
                        Write-Host "N: No" 
                        Write-Host "Y: Yes"
                        $otherInput = Read-Host "Please select an option" 
                        If($otherInput[0] -eq 'Y') {
                            Write-Log -Level Info -Message "User has confirmed extracted Tenant name '$($script:extractedTenant)'."
                            ConnectToSharePointOnline -Tenant $script:extractedTenant
                        }
                        Else {
                            ConnectToSharePointOnline
                        }
                    }
                    Else {
                        ConnectToSharePointOnline
                    }
                    Write-Log "Create Email Views: $script:createEmailViews"
                    Write-Log "Email View Name:$script:emailViewName"
                    Write-Log "Email View set as Default: $script:emailViewDefault"
                    Write-Log "Create View with columns: $script:emailViewColumns"
                    Write-Log "Create Email Columns: $script:createEmailColumns"
                    Write-Log "Email Columns to create/find under Group: $script:columnGroupName"

                    #Deploy
                }
            }
            'c' {
                #clear logins
                Write-Log "User has selected Option C for clear logins."
                Clear-Host
                Try {
                    Disconnect-PnPOnline
                    Write-Log "Cleared PnP Connection!"
                }
                Catch{
                    Write-Log "No PnP Connection to clear!"
                }
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
    Write-Log -Level Error -Message "Caught an exception at the bottom level. `nException Type: $($_.Exception.GetType().FullName) `nException Message: $($_.Exception.Message)"
    Write-Host "`n!!! Please send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance !!!" -ForegroundColor Yellow
    Write-Host "`n!!! Please send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance !!!" -ForegroundColor Red
    Write-Host "`n!!! Please send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance !!!" -ForegroundColor Cyan
    Pause
}
Finally {
    #If running in ISE let's not cut off the PnP sessions
    If($host.name -notmatch 'ISE') {
        Try {
            Disconnect-PnPOnline
        }
        Catch {
            #Sink everything, this is just trying to tidy up any open PnP connections
        }
    }

}
