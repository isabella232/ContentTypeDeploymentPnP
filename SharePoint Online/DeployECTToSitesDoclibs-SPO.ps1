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
        Write-Log "Couldn't check PNP Module versions, Package Manager may be absent"
    }
    Finally {
        Write-Log "PnP Module Installed: $pnpModule"
        Start-Sleep -Seconds 2
    }
    

    #Contains all our Site Collections as siteCol objects
    $script:siteColsHT = @{ }

    #Name of Column group containing the Email Columns, and an object to contain the Email Columns
    [string]$script:columnGroupName = "OnePlace Solutions"
    $script:emailColumns = $null

    [string]$script:csvFilePath = ""

    #Contains all the data we need relating to the Site Collection we are working with, including the Document Libraries and the Site Content Type names
    class siteCol {
        [string]$name
        [string]$url
        [string]$web
        [hashtable]$documentLibraries = @{ }
        [array]$contentTypes
        [boolean]$isSubSite
        [boolean]$createColumns

        siteCol([string]$name, [string]$url, [boolean]$columns) {
            If ([string]::IsNullOrWhiteSpace($name)) {
                $this.name = $url
            }
            Else {
                $this.name = $name
            }
            Write-Log "Creating siteCol object with name '$($this.name)'"

            $this.createColumns = $columns
            Write-Log "Are we going to create columns in this siteCol: $($this.createColumns)"

            $this.contentTypes = @()

            #URL parsing and checking
            If($url -match "(https://.+/)((sites/)|(teams/))([^/]+)/?$") {
                Write-Log "Valid Site Collection"
                $this.isSubSite = $false
                $this.url = $url
                $this.web = $url
            }
            ElseIf($url -match "(https://.+/)((sites/)|(teams/))([^/]+)(.+)$") {
                Write-Log "Valid Sub-Site / Sub-Web of Site Collection"
                $this.isSubSite = $true
                $this.url = $Matches[1] + $Matches[2] + $Matches[5]
                $this.web = $Matches[2] + $Matches[5] + $Matches[6]
            }
            ElseIf(($url -match "(https://.+/)([^/]+)$") -and ($url -notlike "*sites*") -and ($url -notlike "*team*")) {
                #This is a Sub Site / Sub Web of the Root Site Collection. Not recommended but will allow it
                $this.isSubSite = $true
                $this.url = $Matches[1]
                $this.web = $Matches[2]
            }
            ElseIf($url -match "(https://.+)(\.com)(/)?$") {
                #This is likely the Root Site Collection, not recommended but will allow it
                $this.isSubSite = $false
                $this.url = $url
                $this.web = $url
            }
            Else {
                Write-Log -Level Error -Message "Site Collection URL could not be parsed. Please check this entry in your CSV"
                $this.url = ""
            }

            <#
            $url = 'https://tenant.sharepoint.com/sites/sc/sw'
            $urlArray = $url.Split('/')
            $rootUrl = $urlArray[0] + '//' + $urlArray[2] + '/'

            If ($urlArray[3] -eq "") {
                #This is the root site collection
                $isSubSite = $false
            }
            ElseIf (($urlArray[3] -ne "sites") -and ($urlArray[3] -ne "teams")) {
                #This is a subsite in the root site collection
                For ($i = 3; $i -lt $urlArray.Length; $i++) {
                    If ($urlArray[$i] -ne "") {
                        $web += '/' + $urlArray[$i]
                    }
                }
                $isSubSite = $true
            }
            Else {
                #This is a site collection with a possible subweb
                $rootUrl += $urlArray[3] + '/' + $urlArray[4] + '/'
                For ($i = 3; $i -lt $urlArray.Length; $i++) {
                    If ($urlArray[$i] -ne "") {
                        $web += '/' + $urlArray[$i]
                    }
                }
                If ($urlArray[5].Count -ne 0) {
                    $isSubSite = $true
                }
                Else {
                    $isSubSite = $false
                }
            }
            #>
            Write-Log "What is the Web we are in? $($this.web)"
            Write-Log "What is the Root Site Collection for this Site / Web? $($this.url)"
            Write-Log "Is this Site a Sub-Site? $($this.isSubSite)"
        }

        [boolean]connect() {
            If(-not ([string]::IsNullOrWhiteSpace($this.url))) {
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
            Else {
                Write-Log -Level Warn -Message "No valid URL for Site Collection $($this.name), skipping."
                Return $false
            }
        }

        [void]addContentTypeToDocumentLibrary([string]$contentTypeName,[string]$docLibName,[string]$viewName,[boolean]$viewDefault) {
            #Check we aren't working without a Document Library name, otherwise assume that we just want to add a Site Content Type
            If (-not [string]::IsNullOrWhiteSpace($docLibName)) {
                If ($this.documentLibraries.ContainsKey($docLibName)) {
                    $this.documentLibraries.$docLibName
                    Write-Log "Document Library '$docLibName' already listed for this Site Collection."
                }
                Else {
                    $tempDocLib = [docLib]::new($docLibName,$($this.web),$viewName,$viewDefault)
                    $this.documentLibraries.Add($docLibName, $tempDocLib)
                    Write-Log "Document Library '$docLibName' not listed for this Site Collection, added."
                }
                $this.documentLibraries.$docLibName.addContentType($contentTypeName)
            }
            
            #If the named Content Type is not already listed in Site Content Types, add it to the Site Content Types
            If(-not ([string]::IsNullOrWhiteSpace($contentTypeName))) {
                If (-not $this.contentTypes.Contains($contentTypeName)) {
                    $this.contentTypes += $contentTypeName
                    Write-Log "Content Type '$contentTypeName' not listed in Site Content Types, adding."
                }
                Else {
                    Write-Log "Content Type '$contentTypeName' already listed in Site Content Types."
                }
            }
            Else {
                Write-Log "No Content Type specified, only adding Site Columns."
            }
        }

        [void]processDocLibs() {
            Write-Log "Adding the Content Types to the Document Libraries in SPO..."
            ForEach($lib in $this.documentLibraries.Values) {
                $lib.processDocLib()
            }
        }

        #Creates the Email Columns in the given Site Collection. Taken from the existing OnePlaceSolutions Email Column deployment script
        [void]createEmailColumns() {
            If($this.createColumns) {
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
                    Start-Sleep -Seconds 2
                }
            }
            Else {
                Write-Log "Not creating columns in this siteCol."
            }
        }
        
        #Retrieve all the columns/fields for the group specified in this Site Collection, we will add these to the named Content Types shortly. If we do not get the site columns, skip this Site Collection
        [void]createContentTypes() {
            $script:emailColumns = Get-PnPField -Group $script:columnGroupName
            If (($null -eq $script:emailColumns) -or ($null -eq $script:columnGroupName)) {
                Write-Log -Level Warn -Message "Email Columns not found in Site Columns group '$script:columnGroupName' for Site Collection '$($this.name)'. Skipping."
            }
            Else {
                Write-Log "Email Columns found for group '$script:columnGroupName':"
                $script:emailColumns | Format-Table
                Write-Host "The Email Columns will be added to the Site Content Types extracted from your CSV file:"
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
        [string]$name
        [array]$contentTypes
        [string]$web
        [boolean]$createView
        [string]$viewName
        [boolean]$viewDefault

        docLib([String]$name,[string]$web,[string]$viewName,[boolean]$viewDefault) {
            Write-Log "`tCreating docLib object with name $name."
            $this.name = $name
            $this.contentTypes = @()
            $this.web = $web
            If(-not ([string]::IsNullOrWhiteSpace($viewName))) {
                $this.createView = $true
                $this.viewName = $viewName
                $this.viewDefault = $viewDefault
                Write-Log "`tView information present. Name: $($this.viewName); Default: $($this.viewDefault)"
            }
            Else {
                $this.createView = $false
                $this.viewName = ""
                $this.viewDefault = ""
            }
        }

        docLib([String]$name,[string]$web) {
            Write-Log "`tCreating docLib object with name $name, no view settings passed"
            $this.name = $name
            $this.contentTypes = @()
            $this.web = $web
            $this.createView = $false
            $this.viewName = ""
            $this.viewDefault = ""
        }

        [void]addContentType([string]$contentTypeName) {
            If (-not $this.contentTypes.Contains($contentTypeName)) {
                Write-Log "`tAdding Content Type '$contentTypeName' to '$($this.name)' Document Library Content Types"
                $this.contentTypes += $contentTypeName
            }
            Else {
                Write-Log "`tContent Type '$contentTypeName' already listed in Document Library $($this.name)"
            }
        }

        [void]processDocLib() {
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

        [void]createEmailView() {
            If($this.createView) {
                Try {
                    Try {
                        #assign any output from SPO to a variable to sink any bugs with PnP
                        $view = Get-PnPView -List $this.name -Identity $this.viewName -Web $this.web -ErrorAction Stop
                        Write-Log "View '$($this.viewName)' in Document Library '$($this.name)' already exists, will set as Default View if required but otherwise skipping."
                        If($this.viewDefault) {
                            Write-Log "Setting View '$($this.viewName)' as Default"
                            Set-PnPView -List $this.name -Identity $this.viewName -Values @{DefaultView =$True}
                        }
                        Else {
                            Write-Log "Not Setting View '$($this.viewName)' as Default"
                        }
                    }
                    Catch [System.NullReferenceException]{
                        #View does not exist, this is good
                        Write-Log "Adding Email View '$($this.viewName)' to Document Library '$($this.name)'."
                        #assign any output from SPO to a variable to sink any bugs with PnP
                        $view = Add-PnPView -List $this.name -Title $this.viewName -Fields $script:emailViewColumns -RowLimit 100 -Web $this.web -ErrorAction Continue
                        
                        #Let SharePoint catch up for a moment
                        Start-Sleep -Seconds 2
                        If($this.viewDefault) {
                            Write-Log "Setting View '$($this.viewName)' as Default"
                            Set-PnPView -List $this.name -Identity $this.viewName -Values @{DefaultView =$True}
                        }
                        Else {
                            Write-Log "Not Setting View '$($this.viewName)' as Default"
                        }
                        
                        $view = Get-PnPView -List $this.name -Identity $this.viewName -Web $this.web -ErrorAction Stop
                        Write-Log "Email View '$($this.viewName)' created successfully."
                    }
                    Catch{
                        Throw
                    }
                    
                }
                Catch {
                    Write-Log -Level Error -Message "Error checking/creating View '$($this.viewName)' in Document Library '$($this.name)': $($_.Exception.Message)"
                }
            }
            Else {
                Write-Log "Not creating views here"
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

            Write-Host "Enumerating information from CSV file...`n" -ForegroundColor Yellow
            ForEach ($element in $csv) {
                [string]$csv_siteName = $element.SiteName
                [string]$csv_siteUrl = $element.SiteUrl -replace '\s', '' #remove any whitespace from URL
                [string]$csv_docLib = $element.DocLib
                [string]$csv_contentType = $element.CTName
                Try{
                    [boolean]$csv_viewDefault = [System.Convert]::ToBoolean($element.viewDefault)
                }
                Catch {
                    #null values will be treated as $false
                    [boolean]$csv_viewDefault = $false
                }
                Try{
                    [boolean]$csv_createColumns = [System.Convert]::ToBoolean($element.CreateColumns)
                }
                Catch {
                    #null values will be treated as $false
                    [boolean]$csv_createColumns = $false
                }
                [string]$csv_viewName = $element.viewName


                If([string]::IsNullOrWhiteSpace($script:extractedTenant)) {
                    $script:extractedTenant = $csv_siteUrl  -match 'https://(?<Tenant>.+)\.sharepoint.com'
                    $script:extractedTenant = $Matches.Tenant
                    Write-Log "Extracted Tenant name '$script:extractedTenant'`n"
                }

                #Don't create siteCol objects that do not have a URL, this also accounts for empty lines at EOF
                If (-not ([string]::IsNullOrWhiteSpace($csv_siteUrl))) {
                    #If a name is not defined, use the URL
                    If ([string]::IsNullOrWhiteSpace($csv_siteName)) { 
                        $csv_siteName = $element.SiteUrl 
                    }

                    If ($script:siteColsHT.ContainsKey($csv_siteUrl)) {
                        #Site Collection already listed, just add the Content Type and the Document Library if required
                        $script:siteColsHT.$csv_siteUrl.addContentTypeToDocumentLibrary($csv_contentType,$csv_docLib,$csv_viewName,$csv_viewDefault)
                    }
                    Else {
                        $newSiteCollection = [siteCol]::new($csv_siteName, $csv_siteUrl,$csv_createColumns)
                        $newSiteCollection.addContentTypeToDocumentLibrary($csv_contentType,$csv_docLib,$csv_viewName,$csv_viewDefault)
                        $script:siteColsHT.Add($csv_siteUrl, $newSiteCollection)
                    }
                }
                Write-Host "`n"
            }
            Write-Log "Completed Enumerating Site Collections and Document Libraries from CSV file!"
        }
        Catch {
            Write-Log -Level Error -Message "Error parsing CSV file. Is this filepath for a a valid CSV file?"
            Write-Host "`nPATH: $csvFile`n" -ForegroundColor Red
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
                Write-Host "Prompting for PnP Management Shell Authentication. Please copy the code displayed into the browser as directed and log in.`nIf nothing happens in the PowerShell session after logging in through the browser, please click in this window and press 'Enter'." -ForegroundColor Green
                $conn = Connect-PnPOnline -Url $rootSharePointUrl -PnPManagementShell -LaunchBrowser
                #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
                Write-Log "Testing connection with 'Get-PnPWeb'..."
                Start-Sleep -Seconds 3
                Get-PnPWeb -Connection $conn | Out-Null
            }
            Else {
                Write-Log "Already connected to SharePoint with Root URL '$rootSharePointUrl'. Skipping login"
            }
        }
        Catch [System.Net.WebException] {
            If ($($_.Exception.Message) -like "*(401)*") {
                Write-Log -Level Warn "Cannot authenticate with SharePoint Root Site. Please check if an authentication prompt appeared on your machine prior to the last interaction with this script."
            }
            ElseIf ($($_.Exception.Message) -like "*(403)*") {
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
        Write-Host "2: Deploy"
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
                $site.createEmailColumns()
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
                    Write-Log "Create Email Views: $script:createEmailViews" -NoOutput
                    Write-Log "Email View Name:$script:emailViewName" -NoOutput
                    Write-Log "Email View set as Default: $script:emailViewDefault" -NoOutput
                    Write-Log "Create View with columns: $script:emailViewColumns" -NoOutput
                    Write-Log "Create Email Columns: $script:createEmailColumns" -NoOutput
                    Write-Log "Email Columns to create/find under Group: $script:columnGroupName" -NoOutput

                    Write-Log "Beginning Deployment" -NoOutput
                    Deploy
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
    Write-Log -Level Error -Message "Further exception detail. `n`nException Type: $($_.Exception.GetType().FullName) `n`nException Message: $($_.Exception.Message)"
    Write-Host "`n!!! Please send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance !!!" -ForegroundColor Yellow
    Write-Host "`n!!! Please send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance !!!" -ForegroundColor Red
    Write-Host "`n!!! Please send the log file at '$script:logPath' to 'support@oneplacesolutions.com' for assistance !!!`n" -ForegroundColor Cyan
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
