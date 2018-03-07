<#
  .Synopsis
  Create the required objects in a Checkpoint R80+ management server to allow Office 365 traffic
  .DESCRIPTION
  This script will connect to https://support.content.office.net/en-us/static/O365IPAddresses.xml
  and download an XML file containing the required objects to allow Office 365 traffic to pass. 
  It will then create the objects (IPv4, IPv6 addresses or URLs) into the Checkpoint management
  server using the R80+ API, according to the selected parameters.
  It depends on the ConvertFrom-O365AddressesXMLFile module 
  (https://github.com/it-praktyk/Convert-Office365NetworksData/tree/master/ConvertFrom-O365AddressesXMLFile).
  .NOTES   
  Name: Create-O365CheckpointObjects
  Version: 1.1.0
  DateCreated: 2017-02-17
  DateUpdated: 2016-03-09
  .PARAMETER Server
  The mandatory Checkpoint management server hostname or IP address
  .PARAMETER Port
  The Checkpoint R80 API port
  By default, 443 will be used
  .PARAMETER Port
  An optional MDS domain name to use
  .PARAMETER Service
  An optional Office 365 to filter on (among "WAC","Sway","Planner","Yammer","OfficeMobile","ProPlus","RCA","OneNote",
  "OfficeiPad","EXO","SPO","Office365Video","LYO","Identity","CRLs","o365","EOP","Teams","EX-Fed")
  If not specified, all Office 365 services objects will be created
  .PARAMETER Type
  A mandatory object type to filter on (among "IPv4","IPv6","URL")
  .PARAMETER Prefix
  A prefix for the Office 365 objects in the Checkpoint management server
  By default, "O365" will be used
  .PARAMETER Category
  The primary category for the Office 365 application objects in the Checkpoint management server
  By default, "Microsoft & Office365 Services" will be used
  .EXAMPLE
  Create-O365CheckpointObjects -Server cpserver -Type IPv4
  Description:
  Will create the IPv4 objects for all the Office 365 apps in a Checkpoint management server
  named "cpserver"
  .EXAMPLE
  Create-O365CheckpointObjects -Server cpserver -Service LYO -Type IPv6 -Verbose
  Description:
  Will create the IPv6 network objects for Skype for Business in a Checkpoint management server 
  named "cpserver"
  .EXAMPLE
  Create-O365CheckpointObjects -Server cpserver -Service EOP -Type URL -Category "Exchange"
  Description:
  Will create an application object for Exchange Online, with the required URLs, and a primary 
  category set to "Exchange"
#>
[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [string]$Server,

  [Parameter()]
  [int]$Port = 443,
  
  [Parameter()]
  [string]$DomainName,  

  [Parameter()]
  [ValidateSet("WAC","Sway","Planner","Yammer","OfficeMobile","ProPlus","RCA","OneNote",
  "OfficeiPad","EXO","SPO","Office365Video","LYO","Identity","CRLs","o365","EOP","Teams","EX-Fed")]
  [string]$Service,

  [Parameter()]
  [string]$Prefix = "O365",

  [Parameter()]
  [string]$Category = "Microsoft & Office365 Services",  

  [Parameter(Mandatory=$true)]
  [ValidateSet("IPv4","IPv6","URL")]
  [string]$Type
)

# Import the required module
If (Get-Module -ListAvailable -Name ConvertFrom-O365AddressesXMLFile) {} Else{
  Write-Host "The O365AddressesXMLFile module is not installed. Exiting" -BackgroundColor Red
  Exit 1
}
If ( ! (Get-module ConvertFrom-O365AddressesXMLFile )) {
  Import-Module ConvertFrom-O365AddressesXMLFile
}

# The URL blacklist
$blacklist  = "facebook|youtube|evernote|google-analytics|wunderlist|flurry|adjust|uservoice|hockeyapp|box.com|webtrends|tific|yahoo|bing|apple"

# Checkpoint API URIs
$loginURI   = "https://${Server}:${Port}/web_api/login"
$logoutURI  = "https://${Server}:${Port}/web_api/logout"
$discardURI = "https://${Server}:${Port}/web_api/discard"
$publishURI = "https://${Server}:${Port}/web_api/publish"
$addNetURI  = "https://${Server}:${Port}/web_api/add-network"
$AddAppURI  = "https://${Server}:${Port}/web_api/add-application-site"
$SetAppURI  = "https://${Server}:${Port}/web_api/set-application-site"
$ShowAppURI = "https://${Server}:${Port}/web_api/show-application-site"
$SetGrpURI  = "https://${Server}:${Port}/web_api/set-group"
$AddGrpURI  = "https://${Server}:${Port}/web_api/add-group"
$ShowGrpURI = "https://${Server}:${Port}/web_api/show-group"


# FUNCTIONS
Function CPAPIRequest {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$true)]
    [string]$uri, 
    
    [Parameter(Mandatory=$true)]
    $body,
    
    [string]$method = "POST",
    $headers,
    [bool]$stoponerror = $False
  ) 
  Process {
    $mybodyjson = $body | convertto-json -compress
    try {
      If ($headers.Length -gt 0) {
        $myresponse = Invoke-WebRequest -uri $uri -ContentType "application/json" -Method $method -headers $headers -body $mybodyjson -ErrorAction Stop
      } Else {
        $myresponse = Invoke-WebRequest -uri $uri -ContentType "application/json" -Method $method -body $mybodyjson -ErrorAction Stop
      }      
    } catch {
        $_.Exception
        
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $reader.ReadToEnd();
        If ($stoponerror) { Exit 1 }
    }
    Return $myresponse
  }
}


# MAIN
Clear-Host
Write-Host "************************************************************"
Write-Host "This script will create the required objects in a Checkpoint" -ForegroundColor Yellow
Write-Host "R80+ management server to allow Office 365 traffic to pass." -ForegroundColor Yellow

Write-Host "************************************************************"
Write-Host "*************** Getting objects from Office 365 ******************" -BackgroundColor Yellow -ForegroundColor Black
Write-Host

$objs = ConvertFrom-O365AddressesXMLFile -RemoveFileAfterParsing
If ($Service) {
  Write-Host "Filtering on service $Service..."
  $objs = $objs | where { $_.Service -eq $Service }
}
Write-Host "Filtering on type $Type..."
$objs = $objs | where { $_.type -eq $Type }

If ($Type -eq "URL") {
  Write-Host "Applying the URL blacklist..."
  $objs = $objs | Where-Object { $_.Url -notmatch $blacklist }
}

Write-Host
Write-Verbose "Objects downloaded from Microsoft :" 
Write-Verbose ($objs | ft | Out-String)

# Count objects
$count =  ($objs | measure).count
If ($count -eq 0) {
  Write-Host "Cannot find Office 365 objects. Exiting" -BackgroundColor Red
  Exit 1
  } Else {
  Write-Host "Found $count objects matching the filters"
  $confirmation = Read-Host "Are you sure you want to proceed (y|n) ?"
  if ($confirmation -ne 'y') {
    Exit 
  }
}

# Prompt for Checkpoint credentials
If ($cred = $host.ui.PromptForCredential('Credentials', 'Please enter the credentials to access the Checkpoint API','', '')){}Else{Exit}
$User = $cred.Username
$Password = $cred.GetNetworkCredential().Password

#create credential json
$myCredentialhash=@{user=$User;password=$Password}

if($DomainName.length -gt 0){$myCredentialhash.add("domain", $DomainName) }

#allow self signed certs
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }

$myresponse = CPAPIRequest -uri $loginURI -body $myCredentialhash -stoponerror $True

#remove objects with password
rv "Password"
if ($credential.password.Length -gt 0) {rv "cred"}

$myresponsecontent=$myresponse.Content | ConvertFrom-Json
$mysid=$myresponsecontent.sid
$myCPHeader=@{"x-chkp-sid"=$mysid}

Write-Host "***************** Creating O365 objects ********************" -BackgroundColor Yellow -ForegroundColor Black
# Looping through Office 365 services
Foreach ($srv in ($objs | Select-Object -Unique Service) ) {
  $grpname = "{0}_{1}_{2}" -f $Prefix, $Type, $srv.Service
  
  If ($Type -eq "URL" ) { # URLs
    $urllist = @()
    Foreach ($node in ( $objs | where {$_.Service -eq $srv.Service } ) ) {
      $URL = $node.Url
      If ($URL -eq "") { Continue } # Skip empty URLs
      
      # Sanitize Url
      $URL = $URL -replace "\.","\." -replace "\*",".*" | ? {$_.trim() -ne "" }
      
      Write-Verbose "Adding object $URL of type $Type"
      $urllist += $URL
    }
    
    $mybody=@{name="$grpname"}
    $myresponse = CPAPIRequest -uri $ShowAppURI -body $mybody -headers $myCPHeader 
    $myresponsecontent=$myresponse.Content | ConvertFrom-Json
    If ($myresponsecontent.name -eq $grpname) {
      Write-Host "$grpname already exists. Updating URL list..."  
      $cmp = Compare-Object -ReferenceObject $myresponsecontent.'url-list' -DifferenceObject $urllist | ft inputobject, @{n="Action";e={ if ($_.SideIndicator -eq '=>') { "ADD" } else { "REMOVE" } }} | out-string
      If ($cmp) { 
        Write-Host $cmp               
        $mybody=@{name=$grpname;"primary-category"=$Category;color="cyan";"urls-defined-as-regular-expression"=$True;"url-list"=$urllist}
        $myresponse = CPAPIRequest -uri $SetAppURI -body $mybody -headers $myCPHeader
        
      } Else { Write-host "No difference in URL list" }      
    } Else {
      Write-Host "Creating application $grpname" -ForegroundColor Green
      $mybody=@{name=$grpname;"primary-category"=$Category;color="cyan";"urls-defined-as-regular-expression"=$True;"url-list"=$urllist}
      $myresponse = CPAPIRequest -uri $AddAppURI -body $mybody -headers $myCPHeader   
    }
  }
  
  Else { # IPv4 or IPv6
    $members = @()
    Foreach ($node in ( $objs | where {$_.Service -eq $srv.Service } ) ) {
      $IPaddress = ($node.IPAddress).IPAddressToString
      $SubNetMaskLength = $node.SubNetMaskLength
      $Name = "{0}_{1}_{2}" -f $Prefix, $Type, $IPaddress
      $members += [pscustomobject]@{ 
            name=$Name
            ipaddress=$IPaddress
            subnetmasklength=$SubNetMaskLength 
      }
    }
    $mybody=@{name=$grpname}
    $myresponse = CPAPIRequest -uri $ShowGrpURI -body $mybody -headers $myCPHeader
    $myresponsecontent=$myresponse.Content | ConvertFrom-Json
    
    If ($myresponsecontent.name -eq $grpname) {
      Write-Host "$grpname already exists. Updating list of members..."  
      $old = $myresponsecontent.members | Foreach { $_.Name}
      $new = $members | Foreach { $_.name}
      $cmp = Compare-Object -ReferenceObject $old -DifferenceObject $new| ft inputobject, @{n="Action";e={ if ($_.SideIndicator -eq '=>') { "ADD" } else { "REMOVE" } }} | out-string
      If ($cmp) { 
        Write-Host $cmp
        
        Foreach ($member in $members) {      
          Write-host "Object $($member.Name) of type $Type"
          $mybody=@{name=$member.Name;color="cyan";subnet=$member.IPAddress;"mask-length"=$member.SubNetMaskLength}
          $myresponse = CPAPIRequest -uri $AddNetURI -body $mybody -headers $myCPHeader   
        }
        
        Write-Host "Updating group $grpname" -ForegroundColor Green
        $members = $members | Foreach { $_.name}
        $mybody=@{name=$grpname;members=$members}
        $myresponse = CPAPIRequest -uri $SetGrpURI -body $mybody -headers $myCPHeader

      } Else { Write-host "No difference in list of members"}
    } 

    Else {
      Foreach ($member in $members) {      
        Write-host "Object $($member.Name) of type $Type"
        $mybody=@{name=$member.Name;color="cyan";subnet=$member.IPAddress;"mask-length"=$member.SubNetMaskLength}
        $myresponse = CPAPIRequest -uri $AddNetURI -body $mybody -headers $myCPHeader   
      }
      Write-Host "Creating group $grpname" -ForegroundColor Green
      $mybody=@{name=$grpname;color="cyan"}
      $myresponse = CPAPIRequest -uri $AddGrpURI -body $mybody -headers $myCPHeader
      
      Write-host "Adding objects to group $grpname" -ForegroundColor Green
      $members = $members | Foreach { $_.name}
      $mybody=@{name=$grpname;members=$members}
      $myresponse = CPAPIRequest -uri $SetGrpURI -body $mybody -headers $myCPHeader       
    }
  }
}

Write-Host
$confirmation = Read-Host "Do you want to publish the objects (y|n) ?"
if ($confirmation -eq 'y') {
  # Publish the objects
  $myresponse = CPAPIRequest -uri $publishURI -body @{} -headers $myCPHeader
  If ($myresponse.statuscode -eq 200){
    Write-Host "Successfully published the objects." -ForegroundColor Green    
  }
  Else {
    Write-Host "Error when publishing the objects." -ForegroundColor Red
  }
}
Else {
  $myresponse = CPAPIRequest -uri $discardURI -body @{} -headers $myCPHeader
  If ($myresponse.statuscode -eq 200){
    Write-Host "Successfully discarded the changes." -ForegroundColor Green    
  }
  Else {
    Write-Host "Error when discarding the changes." -ForegroundColor Red
  }  
}

# logout
$myresponse = CPAPIRequest -uri $logoutURI -body @{} -headers $myCPHeader

Write-Host "********************** End of script ***********************" -BackgroundColor Yellow -ForegroundColor Black

