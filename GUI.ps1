Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:connectedToSPO = $False

Function expandWeb($node){
    $subWebsNode = $node.Nodes.Add("Subwebs")
    $listsNode = $node.Nodes.Add("Lists")
    $subWebs = Get-PnPSubWebs
    $lists = Get-PnPList

    ForEach($list in $lists){
        If(-not($list.Hidden)){
           $newNode = $listsNode.Nodes.Add($list.Title)
           $newNode.Tag = $list
        }
    }

    ForEach($subWeb in $subWebs){
        $newNode = $subWebsNode.Nodes.Add($subWeb.Title)
        $newNode.Tag = $subWeb
        Connect-PnPOnline -url $subWeb.url -UseWebLogin
        expandWeb -node $newNode
    }
}


<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,600'
$Form.text                       = "Deploy OnePlaceMail Email Content Types"
$Form.TopMost                    = $false

$treeViewNavigation              = New-Object system.Windows.Forms.TreeView
$treeViewNavigation.width        = 250
$treeViewNavigation.height       = 320
$treeViewNavigation.location     = New-Object System.Drawing.Point(25,80)

$treeViewAddedLocations          = New-Object system.Windows.Forms.TreeView
$treeViewAddedLocations.width    = 250
$treeViewAddedLocations.height   = 320
$treeViewAddedLocations.location = New-Object System.Drawing.Point(413,80)

$btn_addLocation                 = New-Object system.Windows.Forms.Button
$btn_addLocation.text            = "Add ->"
$btn_addLocation.width           = 100
$btn_addLocation.height          = 30
$btn_addLocation.location        = New-Object System.Drawing.Point(291,79)
$btn_addLocation.Font            = 'Microsoft Sans Serif,10'
$btn_addLocation.Add_Click({
    $currentNode = $treeViewNavigation.SelectedNode
    $treeViewNavigation.Nodes.Remove($currentNode)
    $treeViewAddedLocations.Nodes.Add($currentNode)
    
})

$btn_removeLocation              = New-Object system.Windows.Forms.Button
$btn_removeLocation.text         = "<- Remove"
$btn_removeLocation.width        = 100
$btn_removeLocation.height       = 30
$btn_removeLocation.location     = New-Object System.Drawing.Point(291,127)
$btn_removeLocation.Font         = 'Microsoft Sans Serif,10'
$btn_removeLocation.Add_Click({
    $currentNode = $treeViewAddedLocations.SelectedNode
    $treeViewAddedLocations.Nodes.Remove($currentNode)
    $currentNode.
    $treeViewNavigation.Nodes.Add($currentNode)
    
})

$txt_columnGroupName             = New-Object system.Windows.Forms.TextBox
$txt_columnGroupName.multiline   = $false
$txt_columnGroupName.text        = "OnePlace Solutions"
$txt_columnGroupName.width       = 178
$txt_columnGroupName.height      = 20
$txt_columnGroupName.location    = New-Object System.Drawing.Point(26,480)
$txt_columnGroupName.Font        = 'Microsoft Sans Serif,10'

$chk_createEmailColumns          = New-Object system.Windows.Forms.CheckBox
$chk_createEmailColumns.text     = "Create Email Columns"
$chk_createEmailColumns.AutoSize  = $false
$chk_createEmailColumns.width    = 178
$chk_createEmailColumns.height   = 20
$chk_createEmailColumns.location  = New-Object System.Drawing.Point(26,421)
$chk_createEmailColumns.Font     = 'Microsoft Sans Serif,10'

$lbl_columnGroupName             = New-Object system.Windows.Forms.Label
$lbl_columnGroupName.text        = "Column Group Name"
$lbl_columnGroupName.AutoSize    = $true
$lbl_columnGroupName.width       = 178
$lbl_columnGroupName.height      = 10
$lbl_columnGroupName.location    = New-Object System.Drawing.Point(26,448)
$lbl_columnGroupName.Font        = 'Microsoft Sans Serif,10'

$lbl_contentTypeName             = New-Object system.Windows.Forms.Label
$lbl_contentTypeName.text        = "Content Type Name"
$lbl_contentTypeName.AutoSize    = $true
$lbl_contentTypeName.width       = 178
$lbl_contentTypeName.height      = 10
$lbl_contentTypeName.location    = New-Object System.Drawing.Point(244,444)
$lbl_contentTypeName.Font        = 'Microsoft Sans Serif,10'

$txt_contentTypeName             = New-Object system.Windows.Forms.TextBox
$txt_contentTypeName.multiline   = $false
$txt_contentTypeName.text        = "OnePlaceMail Email"
$txt_contentTypeName.width       = 178
$txt_contentTypeName.height      = 20
$txt_contentTypeName.location    = New-Object System.Drawing.Point(244,479)
$txt_contentTypeName.Font        = 'Microsoft Sans Serif,10'

$txt_siteCollection              = New-Object system.Windows.Forms.TextBox
$txt_siteCollection.multiline    = $false
$txt_siteCollection.text         = "https://contoso.sharepoint.com/sites/mySiteCollection"
$txt_siteCollection.width        = 343
$txt_siteCollection.height       = 20
$txt_siteCollection.location     = New-Object System.Drawing.Point(26,13)
$txt_siteCollection.Font         = 'Microsoft Sans Serif,10'

$btn_loadSiteCollection          = New-Object system.Windows.Forms.Button
$btn_loadSiteCollection.text     = "Load Site Collection"
$btn_loadSiteCollection.width    = 146
$btn_loadSiteCollection.height   = 30
$btn_loadSiteCollection.location  = New-Object System.Drawing.Point(26,37)
$btn_loadSiteCollection.Font     = 'Microsoft Sans Serif,10'
$btn_loadSiteCollection.Add_Click({
    <#
    If(-not $script:connectedToSPO){
        $urlArray = $txt_siteCollection.Text.Split('/')
        $tenantArray = $urlArray[2].Split('.')
        $adminUrl = $urlArray[0] + '//' + $tenantArray[0] + '-admin' + '.' + $tenantarray[1] + '.' + $tenantarray[2]
        Connect-SPOService -Url $adminUrl
        Start-Sleep -Seconds 3
        $script:connectedToSPO = $True
    }

    Connect-PnPOnline -Url $txt_siteCollection.Text -SPOManagementShell

    $siteCollectionRoot = Get-PnPWeb
    [String]$title = $siteCollectionRoot.Title
    If(-not $treeViewNavigation.Nodes.ContainsKey('$title')){
        $siteNode = $treeViewNavigation.Nodes.Add($title,$title)
        $siteNode.Tag = $siteCollectionRoot

        expandWeb -node $siteNode
        $siteNode.Expand()
    }
    #>

    $siteNode = $treeViewNavigation.Nodes.Add('TestNode','TestNode')
    $siteNode.Nodes.Add('ChildNode','ChildNode')
})

$btn_execute                     = New-Object system.Windows.Forms.Button
$btn_execute.text                = "Execute"
$btn_execute.width               = 60
$btn_execute.height              = 30
$btn_execute.location            = New-Object System.Drawing.Point(448,471)
$btn_execute.Font                = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($treeViewNavigation,$treeViewAddedLocations,$btn_addLocation,$btn_removeLocation,$txt_columnGroupName,$chk_createEmailColumns,$lbl_columnGroupName,$lbl_contentTypeName,$txt_contentTypeName,$txt_siteCollection,$btn_loadSiteCollection,$btn_execute))

$form.ShowDialog()

#$rootNode = New-Object System.Windows.Forms.TreeNode