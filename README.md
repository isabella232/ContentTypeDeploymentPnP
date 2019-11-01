# Deploy a OnePlace Solutions Email Content Type to multiple Site Collections and Document Libraries

A script and sample CSV file to create the OnePlace Solutions Email Columns, add them to Content Types in listed Site Collections, create the named Content Type(s) where necessary, add them to listed Document Libraries, and create a default Email view.

## Getting Started

Please read the entire README (this page) before using the script to ensure you understand it's prerequisites and considerations/limitations.

Download the SitesDocLibs.csv file above and customize it to your requirements. You will be prompted for this file by the script. If you are using Microsoft Edge, you will have to open the CSV file in Github, right click 'Raw' and 'Save Target As'. 

Notes regarding the CSV file:
* You need a new line for each uniquely named Site Content Type, and to define which Site Collection it will be created in, and (optionally) which Document Library it will be added to. 
* When listing a subsite/subweb for the 'SiteUrl', the content type will be created in it's parent Site Collection, eg http://<span>contoso.sharepoint.com/sites/**SiteCollection**/SubSite. You can still list a Document Library within that Subsite to have the Site Content Type added to.
* You may use this script for purely Site Column/Content Type creation by omitting any data for the Document Library column.
* Any Site Content Types listed in the CSV that already exist in your SharePoint Environment will have the Email Columns added to it (and preserve the existing columns).

When you have finished customizing the file, please save and close it to ensure the script can correctly read it.

### Prerequisites

* Administrator rights to your SharePoint Admin Site (for SharePoint Online) and the Site Collections you wish to deploy to.
* [SharePoint PnP CmdLets](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps) - Required for executing the modifications against your Site Collections. Download the appropriate version for your environment.
* [SharePoint Online Management Shell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps) - Required to Authenticate against your Admin Site and access the listed Site Collections through said authentication. (For SharePoint Online only)

### Assumptions and Considerations

* Content Type(s) to be created will have the Site Content Type 'Document' for it's Parent Content Type. 
* Column group name supplied to the script (when prompted) will have all it's columns added to the Content Type(s). If your current Email Columns exist in a group with other columns, please add them to a new Column group to use with this script
* When using this script to add the Email Columns to an existing Content Type, this existing Content Type must be a Site Content Type, and it may be updated to inherit from the 'Document' Site Content Type in the process.
* [OnePlace Solutions Email Columns](https://github.com/OnePlaceSolutions/EmailColumnsPnP) have been installed to the Site Collections you wish to deploy to. This can be done in this script when prompted if not already installed.

### Restrictions

* Only works for SharePoint Online or 2016/2019 environments. SharePoint 2013 is not supported with this script.
* Only works with Site Content Types (for both creation and adding Email Columns to existing) inheriting from the 'Document' Site Content Type. These Site Content Types can however still be added to locations within subsites/subwebs.

### Usage

Please download and modify the CSV before starting. 

1. Start PowerShell on your machine

2. Run the below command to invoke the current(master) version of the script:

```
Invoke-Expression (New-Object Net.WebClient).DownloadString(‘https://github.com/OnePlaceSolutions/ContentTypeDeploymentPnP/raw/master/DeployECTToSitesDoclibs.ps1’)
```

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Colin Wood for his code example on CSV parsing/iterating, and the original Email Columns deployment script.
