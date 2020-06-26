cls
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

#region ***Parameters***
[string]$SiteURL = 'https://CONTOSO.sharepoint.com/sites/DEV'
[int]$BatchSize = 500
[int]$RequestTimeOut = 60000
[string]$ReportPath = $PSScriptRoot + '\Reports'
[string]$LibsPath = $PSScriptRoot + '\Libs'
[string]$LogPath = $PSScriptRoot + '\Logs'

[string]$httpUserAgent = 'NONISV|ArtS|SPOPermissionChecker/2.4'
[string]$LogTime = Get-Date -Format yyyy-MM-dd_hh-mm
[string]$LogFile = $LogPath + '\SPOPermissionChecker_' + $LogTime + '.log'
[string]$ReportFile = $ReportPath + '\SitePermissionRpt_' + $LogTime + '.csv'
#endregion

Start-Transcript -Path $LogFile

Function Load-Dependencies
{
#If you are working in the PowerShell and still have error 0x80131515, reopen console after unblocking files
    Unblock-File -Path $LibsPath'\Microsoft.SharePoint.Client.dll'
    Unblock-File -Path $LibsPath'\Microsoft.SharePoint.Client.Runtime.dll'
    Unblock-File -Path $LibsPath'\Microsoft.Online.SharePoint.Client.Tenant.dll'
    Unblock-File -Path $LibsPath'\Microsoft.Graph.Core.dll'
    Unblock-File -Path $LibsPath'\Microsoft.Graph.dll'
    Unblock-File -Path $LibsPath'\Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
    Unblock-File -Path $LibsPath'\Newtonsoft.Json.dll'
    Unblock-File -Path $LibsPath'\SharePointPnP.IdentityModel.Extensions.dll'
    Unblock-File -Path $LibsPath'\System.Web.Http.dll'
    Unblock-File -Path $LibsPath'\System.Net.Http.Formatting.dll'
    Unblock-File -Path $LibsPath'\OfficeDevPnP.Core.dll'

    try
    {
        Add-Type -Path $LibsPath'\Microsoft.SharePoint.Client.dll'
        Add-Type -Path $LibsPath'\Microsoft.SharePoint.Client.Runtime.dll'
        Add-Type -Path $LibsPath'\Microsoft.Online.SharePoint.Client.Tenant.dll'
        Add-Type -Path $LibsPath'\Newtonsoft.Json.dll'
        Add-Type -Path $LibsPath'\Microsoft.Graph.dll'
        Add-Type -Path $LibsPath'\Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
        Add-Type -Path $LibsPath'\SharePointPnP.IdentityModel.Extensions.dll'
        Add-Type -Path $LibsPath'\OfficeDevPnP.Core.dll'
        #Add-Type -Path $LibsPath'\Microsoft.Graph.Core.dll'
        #Add-Type -Path $LibsPath'\System.Web.Http.dll'
        #Add-Type -Path $LibsPath'\System.Net.Http.Formatting.dll'
    }
    catch [System.Reflection.ReflectionTypeLoadException]
    {
        Write-Host -f Red "Message: $($_.Exception.Message)"
        Write-Host "StackTrace: $($_.Exception.StackTrace)"
        Write-Host -f DarkYellow "LoaderExceptions: $($_.Exception.LoaderExceptions)"
    }
}

#Function to call a non-generic method Load
Function Invoke-LoadMethod([Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),[string]$PropertyName) 
{
   $Ctx = $Object.Context
   $Load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
   $Type = $Object.GetType()
   $ClientLoad = $load.MakeGenericMethod($type)
    
   $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
   $Expression = [System.Linq.Expressions.Expression]::Lambda([System.Linq.Expressions.Expression]::Convert([System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),[System.Object] ), $($Parameter))
   $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
   $ExpressionArray.SetValue($Expression, 0)
   $ClientLoad.Invoke($ctx,@($Object,$ExpressionArray))
}
   
#Function to Get Permissions Applied on a particular Object, such as: Web or List
Function Get-Permissions([Microsoft.SharePoint.Client.SecurableObject]$Object)
{
    #Determine the type of the object
    Switch($Object.TypedObject.ToString())
    {
        "Microsoft.SharePoint.Client.Web"  { $ObjectType = "Site" ; $ObjectURL = $Object.URL; $ObjectTitle = $Object.Title }
        "Microsoft.SharePoint.Client.ListItem"
        { 
            If($Object.FileSystemObjectType -eq "Folder")
            {
                $ObjectType = "Folder"
                #Get the URL of the Folder
                Invoke-LoadMethod -Object $Object -PropertyName "Folder"
                Write-Host "'ExecuteQuery' at #Get the URL of the Folder" -ForegroundColor Yellow
                $Ctx.ExecuteQuery()
                $ObjectTitle = $Object.Folder.Name
                $ObjectURL = $("{0}{1}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl,''),$Object.Folder.ServerRelativeUrl)
            }
            Else #File or List Item
            {
                #Get the URL of the Object
                Invoke-LoadMethod -Object $Object -PropertyName "File"
                Write-Host "'ExecuteQuery' at #Get the URL of the Object" -ForegroundColor Yellow
                $Ctx.ExecuteQuery()
                If($Object.File.Name -ne $Null)
                {
                    $ObjectType = "File"
                    $ObjectTitle = $Object.File.Name
                    $ObjectURL = $("{0}{1}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl,''),$Object.File.ServerRelativeUrl)
                }
                else
                {
                    $ObjectType = "List Item"
                    $ObjectTitle = $Object["Title"]
                    #Get the URL of the List Item
                    Invoke-LoadMethod -Object $Object.ParentList -PropertyName "DefaultDisplayFormUrl"
                    Write-Host "'ExecuteQuery' at #Get the URL of the List Item" -ForegroundColor Yellow
                    $Ctx.ExecuteQuery()
                    $DefaultDisplayFormUrl = $Object.ParentList.DefaultDisplayFormUrl
                    $ObjectURL = $("{0}{1}?ID={2}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl,''), $DefaultDisplayFormUrl,$Object.ID)
                }
            }
        }
        Default 
        { 
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            #Get the URL of the List or Library
            $Ctx.Load($Object.RootFolder)
            Write-Host "'ExecuteQuery' at #Get the URL of the List or Library" -ForegroundColor Yellow
            $Ctx.ExecuteQuery()            
            $ObjectURL = $("{0}{1}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl,''), $Object.RootFolder.ServerRelativeUrl)
        }
    }
   
    #Check if Object has unique permissions
    Invoke-LoadMethod -Object $Object -PropertyName "HasUniqueRoleAssignments"
    Write-Host "'ExecuteQuery' at #Check if Object has unique permissions" -ForegroundColor Yellow
    $Ctx.ExecuteQuery()
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
   
    #Get permissions assigned to the object
    $RoleAssignments = $Object.RoleAssignments
    $Ctx.Load($RoleAssignments)
    Write-Host "'ExecuteQuery' at #Get permissions assigned to the object" -ForegroundColor Yellow
    $Ctx.ExecuteQuery()
    
    #Loop through each permission assigned and extract details
    $PermissionCollection = @()
    Foreach($RoleAssignment in $RoleAssignments)
    { 
        $Ctx.Load($RoleAssignment.Member)
        Write-Host "'ExecuteQuery' at #Loop through each permission assigned and extract details" -ForegroundColor Yellow
        Write-Host "'ExecuteQuery' Decorating traffic with User agent" -ForegroundColor Green
        $Ctx.RequestTimeout = $requestTimeOut
        $Ctx.add_ExecutingWebRequest({
        param($Source, $EventArgs)
        $request = $EventArgs.WebRequestExecutor.WebRequest
        $request.UserAgent = $httpUserAgent
        })

        $Ctx.ExecuteQuery()
    
        #Get the Principal Type: User, SP Group, AD Group
        $PermissionType = $RoleAssignment.Member.PrincipalType
    
        #Get the Permission Levels assigned
        $Ctx.Load($RoleAssignment.RoleDefinitionBindings)

        Write-Host "'ExecuteQuery' at #Get the Permission Levels assigned" -ForegroundColor Yellow
        Write-Host "'ExecuteQuery' Decorating traffic with User agent" -ForegroundColor Green
        $Ctx.add_ExecutingWebRequest({
        param($Source, $EventArgs)
        $request = $EventArgs.WebRequestExecutor.WebRequest
        $request.UserAgent = "NONISV|ArtS|SPOPermissionChecker/2.3"
        })

        $Ctx.ExecuteQuery()
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
 
        #Remove Limited Access
        $PermissionLevels = ($PermissionLevels | Where { $_ –ne "Limited Access"}) -join ","
        If($PermissionLevels.Length -eq 0) {Continue}
 
        #Get SharePoint group members
        If($PermissionType -eq "SharePointGroup")
        {
            #Get Group Members
            $Group = $Ctx.web.SiteGroups.GetByName($RoleAssignment.Member.LoginName)
            $Ctx.Load($Group)
            $GroupMembers= $Group.Users
            $Ctx.Load($GroupMembers)

            Write-Host "'ExecuteQuery' at #Get Group Members" -ForegroundColor Yellow
            Write-Host "'ExecuteQuery' Decorating traffic with User agent" -ForegroundColor Green
            $Ctx.add_ExecutingWebRequest({
            param($Source, $EventArgs)
            $request = $EventArgs.WebRequestExecutor.WebRequest
            $request.UserAgent = "NONISV|ArtS|SPOPermissionChecker/2.3"
            })

            $Ctx.ExecuteQuery()
            If($GroupMembers.count -eq 0){Continue}
            $GroupUsersTMP = ($GroupMembers | Select Title, Email)
            
            $b = New-Object System.Collections.ArrayList
            $GroupUsersTMP | ForEach {$b.Add($_.Title + " `"" + $_.Email + "`"")}
            $GroupUsers = $b
 
            #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($GroupUsers)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            $PermissionCollection += $Permissions
        }
        Else
        {

            $c = $RoleAssignment.Member.LoginName
            $d=$c.split("|")
            $LoginName = $d[2]


            #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title + " `"" + $LoginName + "`"")
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
            $PermissionCollection += $Permissions
        }
    }
    #Export Permissions to CSV File
    $PermissionCollection | Export-CSV $ReportFile -NoTypeInformation -Append
}
   
#Function to get sharepoint online site permissions report
Function Generate-SPOSitePermissionRpt()
{
    [cmdletbinding()]     
    Param  
    (    
        [Parameter(Mandatory=$false)] [String] $SiteURL, 
        [Parameter(Mandatory=$false)] [String] $ReportFile,         
        [Parameter(Mandatory=$false)] [switch] $Recursive,
        [Parameter(Mandatory=$false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory=$false)] [switch] $IncludeInheritedPermissions       
    )  
    Try {
        #Get Credentials to connect
        $authManager = New-Object OfficeDevPnP.Core.AuthenticationManager($null)
    
        #Setup the context
        $Ctx = $authManager.GetWebLoginClientContext($SiteURL)

        #Decorating CSOM calls
        $Ctx.add_ExecutingWebRequest({
        param($Source, $EventArgs)
        $request = $EventArgs.WebRequestExecutor.WebRequest
        $request.UserAgent = "NONISV|ArtS|SPOPermissionChecker/2.3"
        })
        Write-Host "'ExecuteQuery' at #Decorating CSOM calls" -ForegroundColor Yellow
        $ctx.ExecuteQuery()
   
        #Get the Web & Root Web
        1..5 | %{
        $Web = $Ctx.Web
        $RootWeb = $Ctx.Site.RootWeb
        $Ctx.Load($Web)
        $Ctx.Load($RootWeb)
        Write-Host "'ExecuteQuery' at #Get the Web & Root Web" -ForegroundColor Yellow
        $Ctx.ExecuteQuery()
        Start-Sleep -Milliseconds 1000
        }
   
        Write-host -f Yellow "Getting Site Collection Administrators..."
        #Get Site Collection Administrators
        $SiteUsers= $RootWeb.SiteUsers 
        $Ctx.Load($SiteUsers)
        Write-Host "'ExecuteQuery' at #Get Site Collection Administrators" -ForegroundColor Yellow
        $Ctx.ExecuteQuery()
        $SiteAdmins = $SiteUsers | Where { $_.IsSiteAdmin -eq $true}

        $SiteCollectionAdminsTMP = ($SiteAdmins | Select Title, Email)
        $a = New-Object System.Collections.ArrayList
        $SiteCollectionAdminsTMP | ForEach {$a.Add($_.Title + " `"" + $_.Email + "`"")}
        $SiteCollectionAdmins = $a

        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($RootWeb.Title)
        $Permissions | Add-Member NoteProperty URL($RootWeb.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
               
        #Export Permissions to CSV File
        $Permissions | Export-CSV $ReportFile -NoTypeInformation
   
        #Function to Get Permissions of All List Items of a given List
        Function Get-SPOListItemsPermission([Microsoft.SharePoint.Client.List]$List)
        {
            Write-host -f Yellow "`t `t Getting Permissions of List Items in the List:"$List.Title
  
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $Query.ViewXml = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query><RowLimit Paged='TRUE'>$BatchSize</RowLimit></View>"
  
            $ItemCounter = 0
            #Batch process list items - to mitigate list threshold issue on larger lists
            Do {  
                #Get items from the list
                $ListItems = $List.GetItems($Query)
                $Ctx.Load($ListItems)
                Write-Host "'ExecuteQuery' at #Get items from the list" -ForegroundColor Yellow
                $Ctx.ExecuteQuery()
            
                $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
   
                #Loop through each List item
                ForEach($ListItem in $ListItems)
                {
                    #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-Permissions -Object $ListItem
                    }
                    Else
                    {
                        Invoke-LoadMethod -Object $ListItem -PropertyName "HasUniqueRoleAssignments"
                        Write-Host "'ExecuteQuery' at #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch" -ForegroundColor Yellow
                        $Ctx.ExecuteQuery()
                        If($ListItem.HasUniqueRoleAssignments -eq $True)
                        {
                            #Call the function to generate Permission report
                            Get-Permissions -Object $ListItem
                        }
                    }
                    $ItemCounter++
                    Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
                }
            } While ($Query.ListItemCollectionPosition -ne $null)
        }
 
        #Function to Get Permissions of all lists from the web
        Function Get-SPOListPermission([Microsoft.SharePoint.Client.Web]$Web)
        {
            #Get All Lists from the web
            $Lists = $Web.Lists
            $Ctx.Load($Lists)
            Write-Host "'ExecuteQuery' at #Get All Lists from the web" -ForegroundColor Yellow
            $Ctx.ExecuteQuery()
   
            #Exclude system lists
            $ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
            "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
            ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library"
            "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")
             
            $Counter = 0
            #Get all lists from the web   
            ForEach($List in $Lists)
            {
                #Exclude System Lists
                If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
                {
                    $Counter++
                    Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Processing Lists $Counter of $($Lists.Count) in $($Web.URL)" -Status "Exporting Permissions from List '$($List.Title)'"
 
                    #Get Item Level Permissions if 'ScanItemLevel' switch present
                    If($ScanItemLevel)
                    {
                        #Get List Items Permissions
                        Get-SPOListItemsPermission -List $List
                    }
 
                    #Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-Permissions -Object $List
                    }
                    Else
                    {
                        #Check if List has unique permissions
                        Invoke-LoadMethod -Object $List -PropertyName "HasUniqueRoleAssignments"
                        Write-Host "'ExecuteQuery' at #Check if List has unique permissions" -ForegroundColor Yellow
                        $Ctx.ExecuteQuery()
                        If($List.HasUniqueRoleAssignments -eq $True)
                        {
                            #Call the function to check permissions
                            Get-Permissions -Object $List
                        }
                    }
                }
            }
        }
   
        #Function to Get Webs's Permissions from given URL
        Function Get-SPOWebPermission([Microsoft.SharePoint.Client.Web]$Web) 
        {
            #Get all immediate subsites of the site
            $Ctx.Load($web.Webs)
            Write-Host "'ExecuteQuery' at #Get all immediate subsites of the site" -ForegroundColor Yellow
            $Ctx.ExecuteQuery()
    
            #Call the function to Get permissions of the web
            Write-host -f Yellow "Getting Permissions of the Web: $($Web.URL)..." 
            Get-Permissions -Object $Web
   
            #Get List Permissions
            Write-host -f Yellow "`t Getting Permissions of Lists and Libraries..."
            Get-SPOListPermission($Web)
 
            #Recursively get permissions from all sub-webs based on the "Recursive" Switch
            If($Recursive)
            {
                #Iterate through each subsite in the current web
                Foreach ($Subweb in $web.Webs)
                {
                    #Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-SPOWebPermission($Subweb)
                    }
                    Else
                    {
                        #Check if the Web has unique permissions
                        Invoke-LoadMethod -Object $Subweb -PropertyName "HasUniqueRoleAssignments"
                        Write-Host "'ExecuteQuery' at #Check if the Web has unique permissions" -ForegroundColor Yellow
                        $Ctx.ExecuteQuery()
   
                        #Get the Web's Permissions
                        If($Subweb.HasUniqueRoleAssignments -eq $true) 
                        { 
                            #Call the function recursively                            
                            Get-SPOWebPermission($Subweb)
                        }
                    }
                }
            }
        }
   
        #Call the function with RootWeb to get site collection permissions
        Get-SPOWebPermission $Web
   
        Write-host -f Green "`n*** Site Permission Report Generated Successfully!***"
     }
    Catch {
        write-host -f Red "Error Generating Site Permission Report!" $_.Exception.Message
   }
}

Load-Dependencies

#Call the function to generate permission report

#Quick check (Root-Site permissions)
#Generate-SPOSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile

#Moderate (List items and inherited permissions are excluded)
#Generate-SPOSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive

#Moderate (List items permissions are excluded)
#Generate-SPOSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive -IncludeInheritedPermissions

#Full check
#Generate-SPOSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive -ScanItemLevel -IncludeInheritedPermissions
Stop-Transcript

#Add Header Agent to all "ExecuteQuery" and decorate them
#Add select scan lvl
#Add global progress bar
#Update loading logic for dependecies