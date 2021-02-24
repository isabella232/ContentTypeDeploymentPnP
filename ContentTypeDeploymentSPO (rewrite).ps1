Try{
    Set-ExecutionPolicy Bypass -Scope Process
    #Comment out the above line if you allow PowerShell scripts but do not allow execution policy bypass

    Add-Type -AssemblyName System.Windows.Forms

    $ErrorActionPreference = 'Stop'

    #Columns to add to the Email View(s) if/where we are creating them. Edit as required based on Internal Naming
    [string[]]$script:emailViewColumns = @("EmHasAttachments","EmSubject","EmTo","EmDate","EmFromName")

    [string]$script:logFile = "OPSScriptLog$(Get-Date -Format "MMddyyyy").txt"
    [string]$script:logPath = "$env:userprofile\Documents\$script:logFile"
    [string]$script:csvFile = "None"
    [opsTenant]$script:tenant = [opsTenant]::new()

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

    class opsTenant {
        [string]$tenantName
        [string]$rootUrl

        #keys for siteCols should be the URL, since site collections names aren't distinctly unique
        [hashtable]$sites = [ordered]@{}
        [boolean]$authorizedPnP = $false

        opsTenant() {
            $this.tenantName = "No Tenant Name"
            $this.rootUrl = "No Root URL"
        }

        [void]enumerateCSV() {

            Write-Host "Please select your customized CSV containing the Site Collections and Document Libraries to create the Content Types in"
            Start-Sleep -seconds 1
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
                InitialDirectory = [Environment]::GetFolderPath('Desktop') 
                Filter           = 'Comma Separates Values (*.csv)|*.csv'
                Title            = 'Select your CSV file'
            }
            $null = $FileBrowser.ShowDialog()
            
            $script:csvFile = $FileBrowser.FileName
            Write-Log -Level Info -Message "Using CSV at path '$($script:csvFile)'"
            

            $script:siteColsHT = [hashtable]::new
            $script:siteColsHT = @{ }

            Try {
                $csv = Import-Csv -Path $script:csvFile -ErrorAction Continue

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
                    [string]$csv_viewName = $element.viewName

                    $this.processLine($csv_siteName, $csv_siteUrl, $csv_contentType, $csv_docLib, $csv_viewDefault, $csv_viewName, $csv_viewDefault)

                    Write-Host "`n"
                }
                Write-Log "Completed Enumerating Site Collections and Document Libraries from CSV file!"
            }
            Catch {
                Write-Log -Level Error -Message "Error parsing CSV file. Is this filepath for a a valid CSV file?"
                Write-Host "`nPATH: $script:csvFile`n" -ForegroundColor Red
                Throw
            }
            Pause
        }

        <#
        Inputs: (required)name = The name of the Site Collection
                (required)url = The URL of the Site Collection (whole, not relative)
                (required)contentTypeName = The name of the Content Type you wish to create with the OnePlaceMail Email Columns
                (optional)docLibName = The Name of the Document Library you wish to add the Content Type(s) and Views to
                (optional)contentTypeDefault = Whether we set the Content Type as Default in the Document Library. Default $false
                (optional)viewName = The name of the View to create in the Document Library with some of the OnePlaceMail Email Columns visible
                (optional)viewDefault = Whether the View created will be the Default View
        #>
        [void]processLine($name,$url,$contentTypeName,$docLibName,$contentTypeDefault,$viewName,$viewDefault) {
            If([string]::IsNullOrWhiteSpace($name)) {
                $name = $url
            }
            
            $site = $this.addSite($name,$url)
            $site.addContentType($contentTypeName)
            $site.addDocLib($docLibName,$contentTypeName,$contentTypeDefault,$viewName,$viewDefault)

            If("No Tenant Name" -eq $this.tenantName) {
                $url -match 'https://(?<Tenant>.+)\.sharepoint.com' | Out-Null
                $this.tenantName = $Matches.Tenant
                $this.rootUrl = "https://$($this.tenantName).sharepoint.com"
                Write-Log "Extracted Tenant name '$($this.tenantName)`n"
            }
        }

        [void]auth() {
            Try{
                Write-Log "Prompting for PnP Management Shell Authentication"
                Connect-PnPOnline -url $this.rootUrl -PnPManagementShell -LaunchBrowser
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                Write-Log "Connected to '$($currentWeb.Title)' at '$($this.rootUrl)'"
                $this.authorizedPnP = $true
            }
            Catch{
                Write-Log -Level Error -Message "Error authorizing against root Site Collection"
                Write-Log "$_"
            }
        }

        [void]execute() {
            If(0 -eq $this.sites.Count) {
                If($script:csvFile -eq "None") {
                    $this.enumerateCSV()
                }
                $this.auth()
            }
            Else {
                If(-not $($this.authorizedPnP)) {
                    $this.auth()
                }
            }
            $this.sites.Keys | ForEach-Object {
                Write-Log "Executing on Site Object $_"
                $this.sites.Item($_).execute()
            }
        }

        [opsSite]addSite([string]$name,[string]$url) {
            If(-not $this.sites.ContainsKey($url)) {
                $sc = [opsSite]::new($name,$url)
                Write-Log "Adding Site $($name) with URL '$($url)'"
                $this.sites.Add($url,$sc)
                return $sc
            }
            Else {
                Write-Log "Site $($name) already listed with URL '$($url)'"
                return $this.sites.$url
            }
        }
    }

    class opsSite {
        [string]$name
        [string]$url
        [boolean]$connected = $false
        [array]$contentTypes = @()
        [boolean]$contentTypesCreated = $false
        [boolean]$emailColumnsCreated = $false
        [hashtable]$docLibs = [ordered]@{}
        [string]$isSiteCollection = $false
        $currentWeb
        $currentSite

        opsSite([string]$name,[string]$url) {
            If([string]::IsNullOrWhiteSpace($name)) {
                $this.name = $url
            }
            Else {
                $this.name = $name
            }
            $this.url = $url
        }

        [void]connect() {
            Try{
                Connect-PnPOnline -url $this.url -PnPManagementShell
                $this.currentWeb = Get-PnPWeb -ErrorAction Stop
                $this.currentSite = Get-PnPSite -ErrorAction Stop
                Write-Log "Connected to $($this.currentWeb.Title)"

                #If the Site Collection URL contains the whole Web URL then we are in the Site Collection, not a subweb
                $this.isSiteCollection = ($this.currentSite.url -match $this.currentWeb.url)
                $this.connected = $true
            }
            Catch{
                Write-Log -Level Error -Message "Error connecting to Site '$($this.name)' at URL '$($this.url)'."
                $this.connected = $false
            }
        }

        [boolean]addContentType([string]$name) {
            If(-not ($this.contentTypes.Contains($name))) {
                Write-Log "Adding Content Type '$($name)' to Site '$($this.name)' with URL '$($this.url)'"
                $this.contentTypes += $name
                return $true
            }
            Else {
                Write-Log "Content Type '$($name)' already listed in Site '$($this.name)' with URL '$($this.url)'"
                return $false
            }
        }

        [opsDocLib]addDocLib($name,$contentTypeName,$contentTypeDefault,$viewName,$viewDefault) {
            If(-not ($this.docLibs.ContainsKey($name))) {
                $dl = [opsDocLib]::new($name,$this,$contentTypeDefault,$viewName,$viewDefault,$contentTypeName)
                Write-Log "Adding Document Library $($name) with to Site $($this.name) with URL '$($this.url)'"
                Write-Log "Content Type Name: $($contentTypeName)"
                Write-Log "Is Content Type Default: $($contentTypeDefault)"
                Write-Log "View Name: $($viewName)"
                Write-Log "Is View Default: $($viewDefault)"
                $this.docLibs.Add($name,$dl)
                return $dl
            }
            Else {
                return $this.docLibs.$name
            }
        }

        #Creates the Email Columns in the (parent) Site. Taken and modified from the existing OnePlaceSolutions Email Column deployment script
        [void]createEmailColumns() {
            $emailColumns = $null
            
            If(-not $this.isSiteCollection) {
                #Create Columns and Content Types at the Site Collection level for best practice
                Connect-PnPOnline -Url ($this.currentSite.url) -PnPManagementShell
            }
            Try {
                $emailColumns = Get-PnPField -Group 'OnePlace Solutions'
            }
            Catch {
                #This is fine, we will just try to add the columns anyway
                Write-Log -Level Warn -Message "Couldn't check email columns, will attempt to add them anyway..."
            }

            #Check if we have 35 columns in our Column Group
            If ($emailColumns.Count -eq 35) {
                Write-Log "All Email columns already present in group 'OnePlace Solutions', skipping adding."
                $this.emailColumnsCreated = $true
            }
            #Create the Columns if we didn't find 35
            Else {
                $columnsXMLPath = "$env:temp\email-columns.xml"
                If (-not (Test-Path $columnsXMLPath)) {
                    #From 'https://github.com/OnePlaceSolutions/EmailColumnsPnP/blob/master/installEmailColumns.ps1'
                    #Download xml provisioning template
                    $WebClient = New-Object System.Net.WebClient
                    $downloadUrl = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
                
                    Write-Log "Downloading provisioning xml template:" $columnsXMLPath
                    $WebClient.DownloadFile( $downloadUrl, $columnsXMLPath )
                }

                #Apply xml provisioning template to SharePoint
                Write-Log "Applying email columns template to SharePoint Site Collection: $($this.currentSite.Url)"
        
                $rawXml = Get-Content $columnsXMLPath
        
                #To fix certain compatibility issues between site template types, we will just pull the Field XML entries from the template
                ForEach ($line in $rawXml) {
                    Try {
                        If (($line.ToString() -match 'Name="Em') -or ($line.ToString() -match 'Name="Doc')) {
                            Add-PnPFieldFromXml -fieldxml $line -ErrorAction Stop | Out-Null
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
                
                $columnCheckRetry = 5
                Do {
                    $emailColumns = Get-PnPField -Group 'OnePlace Solutions'
                    If($emailColumns.Count -ne 35) {
                        $columnCheckRetry--
                        Start-Sleep -Seconds 1
                    }
                    Else {
                        $columnCheckRetry = 0
                        $this.emailColumnsCreated = $true
                    }
                }
                Until($columnCheckRetry -eq 0)
            }

            If(-not $this.isSiteCollection) {
                #Connect back to the current sub-web
                Connect-PnPOnline -Url ($this.currentWeb.url) -PnPManagementShell
            }
        }

        [void]createContentTypes() {
            If($this.emailColumnsCreated) {
                If(-not $this.isSiteCollection) {
                    #Create Columns and Content Types at the Site Collection level for best practice
                    Connect-PnPOnline -Url ($this.currentSite.url) -PnPManagementShell
                }

                $emailColumns = Get-PnPField -Group "OnePlace Solutions" -InSiteHierarchy
                Start-Sleep -Seconds 2
                If ($null -eq $emailColumns) {
                    Write-Log -Level Warn -Message "Email Columns not found in Site Columns group "OnePlace Solutions" for Site  '$($this.name)'. Skipping."
                }
                Else {
                    Write-Log "Email Columns found for group 'OnePlace Solutions':"
                    $emailColumns | Format-Table
                    Write-Host "The Email Columns will be added to the Site Content Types extracted from your CSV file:"
                    $this.contentTypes
                
                    #Get the Content Type Object for 'Document' from SP, we will use this as the parent Content Type for our email Content Type
                    $DocCT = Get-PnPContentType -Identity 0x0101
                    If ($null -eq $DocCT) {
                        Write-Log -Level Warn -Message "Couldn't get 'Document' Parent Site Content Type in $($this.name). Skipping this Site Collection."
                    }
                    #For each Site Content Type listed for this Site, try and create it and add the email columns to it
                    Else {
                        ForEach ($ct in $this.contentTypes) {
                            Try {
                                $foundContentType = $null
                                Try {
                                    Write-Log "Checking if Content Type '$ct' already exists"
                                    $foundContentType = Get-PnPContentType -Identity $ct
                                }
                                Catch {
                                    #If we encounter any errors on this check we can assume it doesn't exist, but if we failed to find it we should try to create it anyway.
                                    Write-Log "Couldn't find Content Type '$ct', likely does not exist."
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
                                Throw $_
                            }

                            #Try adding columns to the Content Type
                            Try {
                                Write-Log "Adding email columns to Site Content Type '$ct'"
                                Start-Sleep -Seconds 2

                                $numColumns = $emailColumns.Count
                                $i = 0
                                $emSubjectFound = $false
                                ForEach ($column in $emailColumns) {
                                    $column = $column.InternalName
                                    If ($column -eq 'EmSubject') {
                                        $emSubjectFound = $true
                                    }
                                    Write-Progress -Activity "Adding column: $column" -Status "To Site Content Type: $ct in Site: $($this.name). Progress:" -PercentComplete ($i / $numColumns * 100)
                                    Try {
                                        Add-PnPFieldToContentType -Field $column -ContentType $ct
                                    }
                                    Catch {
                                        #Graph is catching up, back off and retry once
                                        Start-Sleep -Seconds 2
                                        Add-PnPFieldToContentType -Field $column -ContentType $ct
                                    }
                                    
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
                        $this.contentTypesCreated = $true
                    }
                }

                If(-not $this.isSiteCollection) {
                    #Connect back to the current sub-web
                    Connect-PnPOnline -Url ($this.currentWeb.url) -PnPManagementShell
                }
            }
            Else {
                $this.createEmailColumns()
                $this.createContentTypes()
            }
        }

        [void]execute() {
            $this.connect()
            If($this.connected) {
                $this.createContentTypes()

                $this.docLibs.Values | ForEach-Object {
                    $_.execute()
                }
            }
            Else {
                Write-Log -Level Warn -Message "Couldn't connect to Site, skipping"
            }
        }
    }

    class opsDocLib {
        [string]$name
        [opsSite]$site
        [boolean]$contentTypeDefault = $false
        [string]$viewName = ""
        [boolean]$viewDefault = $false
        [array]$contentTypes = @()

        opsDocLib([string]$name,[opsSite]$site,[boolean]$contentTypeDefault,[string]$viewName,[boolean]$viewDefault,[string]$contentType) {
            $this.name = $name
            $this.site = $site
            $this.contentTypeDefault = $contentTypeDefault
            $this.viewName = $viewName
            $this.viewDefault = $viewDefault
            $this.addContentType($contentType)
        }

        [void]addContentType($name) {
            If((-not ($this.contentTypes.Contains($name))) -and (-not [string]::IsNullOrWhiteSpace($name))) {
                $this.contentTypes += $name
            }
        }
        [void]execute() {
            $this.contentTypes | ForEach-Object {
                Write-Log "Adding Content Type '$($_)' to Document Library '$($this.name)' in Site '$($this.site.name)'"
                Try {
                    Add-PnPContentTypeToList -List $this.name -ContentType $_
                    If($contentTypeDefault -and (1 -eq $this.contentTypes.Count)) {
                        Write-Log "Default Content Type flag has been set and there is only one Content Type listed for creation. Setting $($_) as Default Content Type"
                        Set-PnPDefaultContentTypeToList -List $this.name -ContentType $_
                    }
                }
                Catch {
                    Write-Log -Level Error -Message "Error adding Content Type '$($_)' to Document Library '$($this.name)' in Site '$($this.site.name)'. Does this Library exist? `nSkipping."
                }
            }

            If( -not [string]::IsNullOrWhiteSpace($this.viewName)) {
                Write-Log "Adding View $($this.viewName) to Document Library $($this.name) in Web $($this.web)"
                Try {
                    Try {
                        #Check if view exists
                        Get-PnPView -List $this.name -Identity $this.viewName
                    }
                    Catch {
                        #View does not exist
                        Add-PnPView -List $this.name -Title $this.viewName -Fields $script:emailViewColumns
                        
                    }
                    If($this.viewDefault) {
                        Write-Log "Default View flag has been set, modifying view to be Default"
                        Set-PnPView -List $this.name -Identity $this.viewName -Values @{DefaultView =$True}
                    }
                }
                Catch {
                    Write-Log -Level Error -Message "Error adding View '$($this.viewName)' to Document Library '$($this.name)' in Site '$($this.site.name)'. Does this Library exist? `nSkipping."
                }
            }
            Else {
                Write-Log "No view name defined, skipping view operations for Document Library $($this.name)"
            }
        }
    }


    #Start of Script
    #----------------------------------------------------------------

    function showEnvMenu { 
        Clear-Host 
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
        Write-Host 'Welcome to the OnePlace Solutions Content Type Deployment Script' -ForegroundColor Green
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
        Write-Host 'Please make a selection to set, toggle, change or execute:' -ForegroundColor Yellow
        Write-Host "1: Select CSV File (currently: '$script:csvFile')"
        Write-Host "2: Deploy"
        Write-Host "`nAdditional Configuration Options:" -ForegroundColor Yellow
        Write-Host "L: Change Log file path (currently: '$script:logPath')"
        Write-Host "`nQ: Press 'Q' to quit."
    }

    do {
        showEnvMenu 
        $userInput = Read-Host "Please select an option" 
        switch ($userInput) { 
            '1' {
                $script:tenant = [opsTenant]::new()
                $script:tenant.enumerateCSV()
            }
            '2' {
                $script:tenant.execute()
            }
            'c' {
                #clear logins
                Write-Log "User has selected Option $($userInput)" -NoOutput
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
            'l' {
                Write-Log "User has selected Option $($userInput)" -NoOutput
                $newLogPath = (Read-Host "Please enter a new path including 'OPSScriptLog.txt' and quotes for the new log file. Eg, 'C:\Users\John\Documents\OPSScriptLog.txt'.")
                $newLogPath = $newLogPath.Replace('\\','\')
                If ([string]::IsNullOrWhiteSpace($newLogPath)) {
                    Write-Host "No path entered, keeping default '$script:logPath'"
                }
                Else {
                    If(-not (Test-Path $newLogPath)) {
                        Move-Item -Path $script:logPath -Destination $newLogPath
                    }
                    Else {
                        Write-Log "Log file exists at $newLogPath, changing path and appending log data."
                        Add-Content -Path $newLogPath -Value (Get-Content -Path $script:logPath)
                        Remove-Item -Path $script:logPath
                    }
                    $script:logPath = $newLogPath
                }
                Pause
            }
            'q' { return }
        }
    } 
    Until($userInput -eq 'q') {
    }

    Try {
        Disconnect-PnPOnline
    }
    Catch{
        #just cleaning up, no issue if we can't disconnect
    }
}
Catch{
    Write-Output "Uncaught error:"
    $_
    Pause
}