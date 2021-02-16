
{
    Add-Type -AssemblyName System.Windows.Forms

    $ErrorActionPreference = 'Stop'

    #Columns to add to the Email View(s) if we are creating them. Edit as required based on Internal Naming
    [string[]]$script:emailViewColumns = @("EmHasAttachments","EmSubject","EmTo","EmDate","EmFromName")

    [string]$script:logFile = "OPSScriptLog$(Get-Date -Format "MMddyyyy").txt"
    [string]$script:logPath = "$env:userprofile\Documents\$script:logFile"

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
        [hashtable]$siteCols = [ordered]@{}
        [boolean]$authorizedPnP = $false

        [void]enumerateCSV() {
            #TODO
            #And extract $tenantName
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
            $sc = $this.addSiteCol($name,$url)
            $sc.addContentType($contentTypeName)
            
            #TODO add parsing for the $url passed so we can create a proper opsWeb object
            ($sc.addWeb($name,$url,$sc)).addDocLib($docLibName,$contentTypeName,$contentTypeDefault,$viewName,$viewDefault)

        }

        [boolean]auth() {
            Try{
                Connect-PnPOnline -url $this.rootUrl -PnPManagementShell -LaunchBrowser
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                Write-Log "Connected to $($currentWeb.Title)"
                $this.authorizedPnP = $true
                return $true
            }
            Catch{
                return $false
            }
        }

        [void]execute() {
            If(0 -eq $this.siteCols.Count) {
                $this.enumerateCSV()
                $this.auth()
            }
            Else {
                If(-not $this.authorizedPnP) {
                    $this.auth()
                }
            }
            $this.siteCols | ForEach-Object {
                $_.execute()
            }
        }

        [opsSiteCol]addSiteCol([string]$name,[string]$url) {
            If(-not $this.siteCols.ContainsKey($url)) {
                $sc = [opsSiteCol]::new($name,$url)
                $this.siteCols.Add($url,$sc)
                return $sc
            }
            Else {
                return $this.siteCols.$url
            }
        }
    }

    class opsSiteCol {
        [string]$name
        [string]$url
        [array]$contentTypes = @()

        #keys for webs should be the URL, since web names arent distinctly unique
        [hashtable]$webs = [ordered]@{}

        opsSiteCol([string]$name,[string]$url) {
            #We might not know the name of the Site Collection because we only care about the Web in this context
            If([string]::IsNullOrWhiteSpace($name)) {
                $this.name = $url
            }
            Else {
                $this.name = $name
            }
            $this.url = $url
        }

        [boolean]addContentType([string]$name) {
            If(-not ($this.contentTypes.Contains($name))) {
                $this.contentTypes += $name
                return $true
            }
            Else {
                return $false
            }
        }

        [opsWeb]addWeb([string]$name, [string]$url) {
            If(-not ($this.webs.ContainsKey($url))) {
                $sw = [opsWeb]::new($name,$url,$this)
                $this.webs.Add($url,$sw)
                return $sw
            }
            Else {
                return $this.webs.$url
            }
        }

        [void]execute() {
            $this.webs | ForEach-Object {
                $_.execute()
            }
        }
    }

    class opsWeb {
        [string]$name
        [opsSiteCol]$siteCol
        [string]$relativeUrl

        #keys for docLibs should be the name, since their names can be distinct and unique
        [hashtable]$docLibs = [ordered]@{}

        opsWeb([string]$name,[string]$url,[opsSiteCol]$siteCol) {
            $this.name = $name

            #TODO Add parsing to check we actually have a relative Url
            $this.relativeUrl = $url

            $this.siteCol = $siteCol
        }

        [opsDocLib]addDocLib($name,$contentTypeName,$contentTypeDefault,$viewName,$viewDefault) {
            If(-not ($this.docLibs.ContainsKey($name))) {
                $dl = [opsDocLib]::new($name,$this,$contentTypeDefault,$viewName,$viewDefault,$contentTypeName)
                $this.docLibs.Add($name,$dl)
                return $dl
            }
            Else {
                return $this.docLibs.$name
            }
        }

        [void]execute() {
            $this.docLibs | ForEach-Object {
                $_.execute
            }
        }
    }

    class opsDocLib {
        [string]$name
        [opsWeb]$web
        [boolean]$contentTypeDefault = $false
        [string]$viewName = ""
        [boolean]$viewDefault = $false
        [array]$contentTypes = @()

        opsDocLib([string]$name,[opsWeb]$web,[boolean]$contentTypeDefault,[string]$viewName,[boolean]$viewDefault,[string]$contentType) {
            $this.name = $name
            $this.web = $web
            $this.contentTypeDefault = $contentTypeDefault
            $this.viewName = $viewName
            $this.viewDefault = $viewDefault
            $this.contentTypes += $contentType
        }

        [boolean]addContentType($name) {
            If(-not ($this.contentTypes.Contains($name))) {
                $this.contentTypes += $name
                return $true
            }
            Else {
                return $false
            }
        }
        [void]execute() {
            $this.contentTypes | ForEach-Object {
                Write-Log "Adding Content Type $($_) to Document Library $($this.name) in Web $($this.web)"
                Add-PnPContentTypeToList -List $this.name -ContentType $_ -Web $this.web
                If($contentTypeDefault -and (1 -eq $this.contentTypes.Count)) {
                    Write-Log "Default Content Type flag has been set and there is only one Content Type listed. Setting $($_) as Default Content Type"
                    Set-PnPDefaultContentTypeToList -List $this.name -ContentType $_ -Web $this.web
                }
            }

            Write-Log "Adding View $($this.viewName) to Document Library $($this.name) in Web $($this.web)"
            Add-PnPView -List $this.name -Title $this.viewName -Fields $script:emailViewColumns
            If($this.viewDefault) {
                Write-Log "Default View flag has been set, modifying view to be Default"
                Set-PnPView -List $this.name -Identity $this.viewName -Values @{DefaultView =$True}
            }
        }
    }
}