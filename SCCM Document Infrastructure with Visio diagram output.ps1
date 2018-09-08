# SCCM High-Level Infrastructure diagram generator for Visio 2016
# for visually documenting your SCCM Hierachy

# By Jack Rudlin

# 16/07/18

# This is the only variable that needs changing for different SCCM Hierachies. This should be the FQDN of the SCCM Reporting Point server at the top level, so CAS, if it's a hierarchy.
$SCCM_SSRS_FQHostname = "SCCM-RP.domain.local"; # Central Administration Site reporting point
If(-not((Test-NetConnection -ComputerName $SCCM_SSRS_FQHostname -CommonTCPPort HTTP).TcpTestSucceeded)){Write-Error -Message "Could not connect to SSRS on $SCCM_SSRS_FQHostname. Please check in the browser to http://$SCCM_SSRS_FQHostname"}

# Custom SCCM Stencils used for building the infrastructure diagram
$SCCM_Servers_Stencil_Path="\\domain.local\shares\files\visio\it\ConfigMgr_1610_Visio_Stencils_v1.3\ConfigMgr 1610 (Servers).vss"
If(-not(Test-Path -Path $SCCM_Servers_Stencil_Path)){write-error -Message "Could not access or find the stencils @ $SCCM_Servers_Stencil_Path. Please check this location exists"}

# SaveAs/Export path and file name
$SaveAsFileName = "SCCM Infrastructure Diagram"
$SaveAsPath = $(If(test-path ([environment]::getfolderpath("mydocuments"))){
                ([environment]::getfolderpath("mydocuments"))}
              else {
                  $env:TEMP
              })

#region Visio Contstants
# https://msdn.microsoft.com/en-us/vba/visio-vba/articles/viscellindices-enumeration-visio
$visSectionObject = 1
$visRowShapeLayout = 23
$visSLORouteStyle = 10
$visLORouteCenterToCenter = 16
$visSLOLineRouteExt = 19
$visRowTextXForm = 12
$visXFormPinY = 1
$visXFormPinX = 0
$visXFormWidth = 2

# Custom shape data:
$visSectionProp = 243
$visCustPropsLabel = 2
$visCustPropsType = 5
$visCustPropsFormat = 3
$visCustPropsLangID = 14
$visCustPropsCalendar = 15
$visCustPropsPrompt = 1
$visCustPropsValue = 0
$visCustPropsSortKey =  4

# Visio shape data constants
$visRowLast = -2
$visTagDefault = 0

#endregion

# Unlikely this will ever need to change, unless a newer version of SQL Server removes the old ReportExecution service.
$ReportServerUri = "http://$SCCM_SSRS_FQHostname/ReportServer/ReportExecution2005.asmx?wsdl"
$ReportServerServiceUri = "http://$SCCM_SSRS_FQHostname/ReportServer/ReportService2010.asmx";

# Unlikely these will need to be changed. But a quick check after a new version of SCCM is released and installed is all that's needed
$SCCM_SiteRoles_ReportName = 'Site system roles and site system servers for a specific site'
$SCCM_SiteStatus_ReportName = 'Site status for the hierarchy'
$RolesToIgnore = @('Component server','Site system');#Don't include these site components in the Visio text shape as all site systems have these
$SquareSize = 1.4
$InterSitePadRatio = 1.0;# Original was 1.5
$PadRatio = 1.9;# Original was 1.5
$CenterX = 4.0
$CenterY = 7.0

Function Get-VisioContainerTextSize{

    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $SiteSystemsCount,
        [parameter(Mandatory=$true)]
        $SquareSize,
        [parameter(Mandatory=$true)]
        $PadRatio
    )

    if ($SiteSystemsCount -ge 3 )
    {
        $TextSize = [math]::Round(((Get-ContainerSize -items $SiteSystemsCount -SquareSize $SquareSize -PadRatio $PadRatio) * (Get-ContainerSize -items $SiteSystemsCount -SquareSize $SquareSize -PadRatio $PadRatio)) * 2)
    }
    else
    {
        $TextSize = 12
    }
    
    return $TextSize

}

Function New-VisioContainer{

    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $Stencil,
        [parameter(Mandatory=$true)]
        $CenterX,
        [parameter(Mandatory=$true)]
        $CenterY,
        [parameter(Mandatory=$true)]
        $Text,
        [parameter(Mandatory=$false)]
        [string]$TextSize = 100
    )

    $Container=$page.Drop($Stencil,$CenterX,$CenterY)
    $Container.Text=$Text
    $Container.Name=$Text
    $Container.Characters.Text=$Text
    $Container.CellsSRC(3,0,7).FormulaForceU = "$TextSize pt"
    $Container.ContainerProperties.ResizeAsNeeded = 2
    #$ContainerCAS.Resize(1,2,69)

    return $Container

}

Function Get-VisioSCCMStencilShapeName{
    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $SCCMRoles,
        [parameter(Mandatory=$true)]
        [bool]$IsCAS,
        [parameter(Mandatory=$false)]
        $IgnoreRoles

    )

    $RolesFiltered = ($Server.Group.Details_Table0_RoleName | Where-Object {$_ -NotIn $RolesToIgnore})
    # Priorise which image to select where servers contain multiple roles
    If($RolesFiltered -contains "Site server"){
        $SCCMRole = "Site server"
    } elseif($RolesFiltered -contains "Site database server"){
        $SCCMRole = "Site database server"
    } elseif($RolesFiltered -contains "Reporting services point"){
        $SCCMRole = "Reporting services point"
    } else {
        $SCCMRole = $RolesFiltered | Select-Object -Last 1
    }

        switch($SCCMRole){

            "Site server" { If($IsCAS){$return = "CAS"}else{$return = "Primary Site Server"}}
            "Site database server" { If($IsCAS){$return = "SQL Server - CAS"}else{$return = "SQL Server - Primary Site"}}
            "Software update point" {$return = "Software Update Point"}
            "Reporting services point" { If($IsCAS){$return = "Reporting Services Point (CAS)"}else{$return = "Reporting Services Point (Primary)"}}
            "Management point" {$return = "Management Point"}
            "Distribution point" {$return = "Distribution Point Server"}
            "State Migration Point" {$return = "State Migration Point"}
            "Fallback Status Point" {$return = "Fallback Status Point"}
            default {$return = "Management Point"}

        }


        return $return
}

Function Update-VisioShapeText{

    [CmdletBinding(DefaultParameterSetName="Update_Text")]
    param(    
        [parameter(Mandatory=$true)]
        $Shape,
        [parameter(ParameterSetName="Update_Text",Mandatory=$false)]
        [string]$Text,
        [parameter(ParameterSetName="AddIP",Mandatory=$false)]
        [bool]$AddIP
    )

    if($AddIP){
        write-host "`n-ADDIP specified. Resolving IP of $($shape.name)" -ForegroundColor Green
        $address=Resolve-DnsName $shape.name
        $text = $address.IPAddress
    }
  
    Try{
        write-host "Adding text: $text to shape $($shape.name)" -ForegroundColor Green
        $newLabel="{0}`n{1}" -f $Shape.text,$Text
        $Shape.Text=$newLabel
    } catch {
        write-error "Could not set shape $($Shape.Name) to text $Text. Please check Visio is open and shape exists" -ErrorAction SilentlyContinue
    }

}

function Set-VisioServer{

    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $Server,
        [parameter(Mandatory=$true)]
        [double]$x,
        [parameter(Mandatory=$true)]
        [double]$y,
        [parameter(Mandatory=$true)]
        $RolesToIgnore,
        [parameter(Mandatory=$true)]
        $page,
        [parameter(Mandatory=$false)]
        $ConnectToShape,
        [parameter(Mandatory=$false)]
        $SiteContainer,
        [parameter(Mandatory=$false)]
        $IsCAS = $false
    )

        # Pick a visio shape from the custom stencil based on the SCCM roles (but not the ones in $rolestoignore)
        $SCCM_Stencil_Shape_Name = Get-VisioSCCMStencilShapeName -SCCMRoles $server.group.Details_Table0_RoleName -IsCAS $IsCAS -IgnoreRoles $RolesToIgnore

        $Visio_Custom_Stencil=$SCCM_Servers_Stencil.Masters($SCCM_Stencil_Shape_Name)
        
        $VisioServer = $Page.Drop($Visio_Custom_Stencil,$x,$y)
        $VisioServerNetbiosName=($($Server.Name).Split('.'))[0]
        $VisioServer.text=$VisioServerNetbiosName
        $VisioServerFQDN=($($Server.Name).TrimStart($VisioServerNetbiosName).TrimStart('.'))
        Update-VisioShapeText -Shape $VisioServer -Text $VisioServerFQDN
        $VisioServer.Name=$($Server.Name)
        write-host "Drop locations set to X:$x and Y:$y for shape: $($VisioServer.Name)" -ForegroundColor Blue -BackgroundColor white

        Update-VisioShapeText -Shape $VisioServer -AddIP $True
                            
        If($ConnectToShape){
            $ConnectToShape.AutoConnect($VisioServer,0)
            $connector = $page.Shapes | Where-Object {$_.style -eq 'Connector'} | Select-Object -Last 1
            $Connector.Cells('ShapeRouteStyle') = 16
            $Connector.Cells('ConLineRouteExt') = 1
        }
        

        if($SiteContainer)
        {
            write-host "Adding shape $($VisioServer.Name) to container: $($SiteContainer.Name)"
            $SiteContainer.ContainerProperties.AddMember($VisioServer,0)
        }
        
        foreach($role in ($Server.Group | where-object {$_.Details_Table0_RoleName -notin $RolesToIgnore} | select-object -ExpandProperty Details_Table0_RoleName))
        {
            Update-VisioShapeText -Shape $VisioServer -Text $role;
        }

        $TextBoxLineCount = ($VisioServer.text | Measure-Object -Line).Lines
        $TextBoxLineCountSplit = [int[]](($TextBoxLineCount -split '') -ne '')
        If($TextBoxLineCountSplit.count -eq 2){
            $TxtHeight = "Height*-$($TextBoxLineCountSplit[0]).$($TextBoxLineCountSplit[1])"
        } elseif ($TextBoxLineCountSplit.count -eq 1) {
            $TxtHeight = "Height*-0.$($TextBoxLineCountSplit)"
        } else {
            $TxtHeight = "Height*-0.4"
        }
        
        $VisioServer.CellsSRC($visSectionObject, $visRowTextXForm, $visXFormWidth).FormulaU = "Width*4"
        
        If($IsCAS -or ($server.group.Details_Table0_RoleName | Where-Object {$_ -eq "Site server"})){
        } else {
            $VisioServer.CellsSRC($visSectionObject, $visRowTextXForm, $visXFormPinY).FormulaU = $TxtHeight
        }
        
               
}
function Get-Radians {
    [CmdletBinding()]
    param($angle)

    return $angle * ([math]::PI / 180) 

}
  
function Get-ContainerSize
{
    [CmdletBinding()]
    param($items,$SquareSize,$PadRatio)            
    
    $angle = 360.0 / $Items
    
    $toRadians = Get-Radians -angle $angle
    
    $sinVal = [Math]::Sin($toRadians)
    
    $hypLength = ($SquareSize * $PadRatio) / $sinVal 
    
    $MaxSquare = ($hypLength * 2.0 + $SquareSize)
    
    return $MaxSquare
}


function Get-WebServiceConnection
{
    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        [string]$url,
        [parameter(Mandatory=$false)]
        $Creds
    )
  
    $reportServerURI = $url

    Write-Host "Getting Web Proxy Details $url"

    Try{
    $RS = New-WebServiceProxy -Uri $reportServerURI -UseDefaultCredential -ErrorAction Stop
    } Catch {
        Write-Host "Error: $_ when connecting to $reportServerURI" -ForegroundColor Yellow
        write-host "Try providing creds to connect to the report server..." -ForegroundColor Yellow
        $RS = New-WebServiceProxy -Uri $reportServerURI -Credential $Creds
    }

    $RS.Url = $reportServerURI
    return $RS
}

function Get-SQLReport
{
    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $ReportingServicesWebProxy,
        [parameter(Mandatory=$true)]
        [string]$reportPath,
        $parameters
    )

    Write-Host "Getting Report $reportPath"
    
    $ReportingServicesWebProxy.GetType().GetMethod("LoadReport").Invoke($ReportingServicesWebProxy, @($reportPath, $null))
    
    $ReportingServicesWebProxy.SetExecutionParameters($parameters, "en-us") > $null

    $devInfo = "<DeviceInfo></DeviceInfo>"
    $extension = ""
    $mimeType  = ""
    $encoding = ""
    $warnings = $null
    $streamIDs = $null

    $RenderedOutPut = $ReportingServicesWebProxy.Render("XML",$devInfo,[ref]$extension,[ref]$mimeType,[ref]$encoding,[ref]$warnings,[ref]$streamIDs)

    $doc = [System.Xml.XmlDocument]::new()
    $memStream = New-Object System.IO.MemoryStream  @($RenderedOutPut,$false)
    $doc.Load($memStream)
    write-output $doc

    $memStream.Close()
    
}

function New-SSRSParameter
{
    [CmdletBinding()]
    param(    
                    [string]$Name,
        [string]$Value
    )

    $param = New-Object PSObject
    Add-Member -InputObject $param -Name "Name" -Value $Name -MemberType NoteProperty
    Add-Member -InputObject $param -Name "Value" -Value $Value -MemberType NoteProperty
    Write-Output $param 
}

Function Get-SCCMSiteCodesReport{

    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $ReportingServicesWebProxy,
        [parameter(Mandatory=$true)]
        $SiteStatusReportPath
    )

    # Get all the site codes of the SCCM infrastructure/hierarchy from SSRS
    $parameters = @()
    $SiteCodesReport = Get-SQLReport -ReportingServicesWebProxy $ReportingServicesWebProxy -reportPath $SiteStatusReportPath -parameters $parameters

    Write-Output -InputObject $SiteCodesReport
}

#Get the location of the reports we need to list the site code and site systems in the hierarchy
Try{
    $ReportServiceProxy = New-WebServiceProxy -Uri $ReportServerServiceUri -Namespace "SSRS" -UseDefaultCredential -ErrorAction Stop
} Catch {
    Write-Host "Error: $_ when connecting to $reportServerURI" -ForegroundColor Yellow
        write-host "Try providing creds to connect to the report server..." -ForegroundColor Yellow
        $Creds = (Get-Credential)
        $ReportServiceProxy = New-WebServiceProxy -Uri $ReportServerServiceUri -Namespace "SSRS" -Credential $Creds
}
$AllReports=$ReportServiceProxy.ListChildren("/", $true);
#$items | select Type, Path, ID, Name | sort-object Type, Name
$SiteStatusReportPath = ($AllReports | Where-Object -Property Name -like $SCCM_SiteStatus_ReportName).Path
$SiteRolesReportPath = ($AllReports | Where-Object -Property Name -like $SCCM_SiteRoles_ReportName).Path

$webProxy = Get-WebServiceConnection -url $ReportServerUri -Creds $Creds
$SiteCodesReport = Get-SCCMSiteCodesReport -ReportingServicesWebProxy $webProxy -SiteStatusReportPath $SiteStatusReportPath

# List of all the site codes
$SCCM_SiteCodes = $SiteCodesReport.Report.Table0.Detail_Collection.Detail.Details_Table0_SiteCode | ForEach-Object {$_.Trim()}

# Get the table header name of the secondary sites column as for some reason Microsoft have not standardised on the name
$SCCM_SecondarySite_Report_Header_Filter = ($SiteCodesReport.Report.Table0 | ForEach-Object {$_.PSObject.properties} | Where-Object {$_.Value -like "*Secondary Site"}).Name -split "_" | Select-Object -Last 1
If($SCCM_SecondarySite_Report_Header_Filter.count -ne 1){write-error "`nCould not find 'Secondary Site' header from report $SiteStatusReportPath. Please check that Microsoft have not changed the table header values";break}
# Get the the property name for secondary sites
$SCCM_SecondarySite_Report_Property_Name = ($SiteCodesReport.Report.Table0.Detail_Collection.Detail | ForEach-Object {$_.PSObject.properties} | Where-Object {$_.Name -like "*$SCCM_SecondarySite_Report_Header_Filter*"}).Name  | Select-Object -Last 1
# Check if any of the site codes are marked as secondary sites
$AllSecondarySitesCodes = ($SiteCodesReport.Report.Table0.Detail_Collection.Detail | Where-Object {$_.$($SCCM_SecondarySite_Report_Property_Name) -eq "True"}).Details_Table0_SiteCode
If($AllSecondarySitesCodes){
    write-host "`nSCCM Infrastructure has $($AllSecondarySitesCodes.count) Secondary Site/s" -ForegroundColor Green
    $AllSecondarySitesCodes = $AllSecondarySitesCodes | ForEach-Object {$_.Trim()}
} else {
    write-host "`nSCCM Infrastructure doesn't have a Secondary Site" -ForegroundColor Green
}

Function Get-SCCMServers{
    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $SiteCodes,
        [parameter(Mandatory=$true)]
        $WebProxy,
        [parameter(Mandatory=$true)]
        $SiteRolesReportPath
    )
     
    #Loop each of the site codes and get the SCCM system servers/systems in the site    
    $SCCMServers = @()
    ForEach($site in $SiteCodes){
        $parameters = @()
        $parameters += New-SSRSParameter -Name "variable" -Value $site

        write-host "Attempting to load report $ReportPath for site $site "
        $SCCM_SiteRoles_ReportXML = Get-SQLReport -ReportingServicesWebProxy $WebProxy -reportPath $SiteRolesReportPath -parameters $parameters

        $SCCMServers += $SCCM_SiteRoles_ReportXML.Report.Table0.Detail_Collection.Detail
        write-host "$(($SCCM_SiteRoles_ReportXML.Report.Table0.Detail_Collection.Detail).count) SCCM system roles details retrieved from the report"
    }
    write-host "`nTotal roles retrieved: $($SCCMServers.count)" -ForegroundColor Green
    Write-Output -InputObject $SCCMServers

}

$SCCMServers = Get-SCCMServers -SiteCodes $SCCM_SiteCodes -WebProxy $webProxy -SiteRolesReportPath $SiteRolesReportPath

#region Determine top level site code
# Site must be standalone primary site if it only has one site code
write-host "`nDetermine if CAS Hierarchy or Standalone Primary site....." -ForegroundColor Green
If($SCCM_SiteCodes.Count -eq 1 -and (-not($AllSecondarySitesCodes))){
    ######################## Standalone Primary ###############################
    write-host "`nOnly one SCCM site code so infrastructure is an Standalone Primary Site with no secondaries" -ForegroundColor Green
    $SCCM_Standalone_Primary_SiteCode = $SCCM_SiteCodes
    ######################## Standalone Primary ###############################
} elseif (
    ######################## CAS ###############################
    # If there is a site without a management point it must be a CAS site as Primary and Secondary sites both must have at least one MP each
    ($SCCMServers `
     | Where-Object {$_.Details_Table0_RoleName -like "Management Point"} `
     | Select-Object -Property Details_Table0_SiteCode -Unique
    ).Count -ne $SCCM_SiteCodes.Count
){
    write-host "`nSCCM infrastructure is a CAS" -ForegroundColor Green
    $SCCM_CAS_SiteCode = ($SCCMServers `
     | Where-Object {$_.Details_Table0_SiteCode -notin ($SCCMServers | Where-Object {$_.Details_Table0_RoleName -like "Management Point"} | Select-Object -Property Details_Table0_SiteCode -Unique).Details_Table0_SiteCode} `
     | Select-Object -Property Details_Table0_SiteCode -Unique).Details_Table0_SiteCode
    write-host "CAS Site code determined as: $SCCM_CAS_SiteCode" -ForegroundColor Green

    ######################## CAS ###############################
} else {
    ######################## Primary with Secondaries ###############################
    # If there is no CAS but there are multiple sites, then the infrastructure must be a Primary site with secondaries attached
    write-host "`nSCCM Infrastructure is a standalone with secondaries" -ForegroundColor Green
    $SCCM_Primary_with_Secondaries_SiteCode = $SCCM_SiteCodes | Where-Object {$_ -notin $AllSecondarySitesCodes}
    write-host "The Standalone Primary site code (that also has secondaries) is: $SCCM_Primary_with_Secondaries_SiteCode"
    ######################## Primary with Secondaries ###############################
}
#endregion

# Sort SCCM servers into Groups
$SortedSitesWithoutCAS = $SCCMServers | where-object -FilterScript {$_.Details_Table0_SiteCode -ne $SCCM_CAS_SiteCode} | select-object -unique -Property Details_Table0_SiteCode,Details_Table0_ServerName | group-object -property Details_Table0_SiteCode
$AllServers = $SCCMServers | select-object -unique -Property Details_Table0_SiteCode,Details_Table0_ServerName | group-object -property Details_Table0_SiteCode
$MaximumServers = $AllServers | Measure-Object -Property Count -Maximum | Select-Object -ExpandProperty Maximum

# Separate Site Server and Site Systems
$SCCMServers_CAS = $SCCMServers | Where-Object {$_.Details_Table0_SiteCode -eq $SCCM_CAS_SiteCode}
$SCCMServer_CAS_SiteServer = $SCCMServers_CAS | Group-Object -Property Details_Table0_ServerName | where-object {$_.Group.Where({$_.Details_Table0_RoleName -eq "Site Server"}) }
$CASSCCMNonSiteServers = $SCCMServers_CAS | Group-Object -Property Details_Table0_ServerName | where-object {!($_.Group.Where({$_.Details_Table0_RoleName -eq "Site Server"})) }
$SCCM_SiteCodes_WithoutCAS = $SCCM_SiteCodes | Where-Object {$_ -ne $SCCM_CAS_SiteCode}


$MaximumSquare = Get-ContainerSize -items $MaximumServers -SquareSize $SquareSize -PadRatio $PadRatio

$ShapeAngle = 360.0 / $SortedSitesWithoutCAS.Count

if($ShapeAngle -eq 180 -or $ShapeAngle -eq 360)
{
    $ConnectorLength = ($SquareSize * $InterSitePadRatio) * 3.0
}
else {
    $SinAngle = [Math]::Sin((Get-Radians -angle $ShapeAngle))
    $ConnectorLength = (($MaximumSquare * $InterSitePadRatio) / $SinAngle)
}



#region On to the Visio stuff now

# Check if Visio is installed and bind to the COM object if so
Try{
    $Visio=New-Object -ComObject Visio.Application
} Catch {
    write-error "Could not find Visio Com Object on the machine. Please make sure Visio 2016 is installed"; break
}

$Document=$Visio.Documents.Add('')
$Page=$Visio.ActivePage
$Page.AutoSize = $true

#Load Visio Stencils
# Builtin STencils
# https://msdn.microsoft.com/en-us/vba/visio-vba/articles/visbuiltinstenciltypes-enumeration-visio
$BuiltIn_Stencil_Path=$Visio.GetBuiltInStencilFile(2,0)
$BuiltIn_Stencil=$Visio.Documents.OpenEx($BuiltIn_Stencil_Path,64)
#$BuiltIn_Stencil.Masters| Select-Object -ExpandProperty Name
$BuiltIn_Stencil_Classic=$BuiltIn_Stencil.Masters('Classic')
$BuiltIn_Stencil_Translucent=$BuiltIn_Stencil.Masters('Translucent')
$ContainerStencil = $BuiltIn_Stencil_Translucent
# Custom configmgr stencils downloaded from: https://gallery.technet.microsoft.com/System-Center-Configuration-d67b8ac5
$SCCM_Servers_Stencil=$Visio.Documents.OpenEx($SCCM_Servers_Stencil_Path,64)
# $SCCM_Servers_Stencil.Masters | Select-Object -ExpandProperty Name
#endregion

#region loop each start and start to build Visio diagram
# Start with CAS hierarchies
If($SCCM_CAS_SiteCode){
    
    write-host "`nProcessing CAS site: $SCCM_CAS_SiteCode for Visio document creation..." -ForegroundColor Green
    
    # Visio container for CAS
    $TextSize = Get-VisioContainerTextSize -SiteSystemsCount $CASSCCMNonSiteServers.Count -SquareSize $SquareSize -PadRatio $PadRatio
    $ContainerCAS = New-VisioContainer -Stencil $ContainerStencil -CenterX $CenterX -CenterY $CenterY -Text $SCCM_CAS_SiteCode -TextSize $TextSize

    # Create CAS Shape
    Set-VisioServer -Server $SCCMServer_CAS_SiteServer -x $CenterX -y $CenterY -RolesToIgnore $RolesToIgnore -page $Page -IsCAS $True -SiteContainer $ContainerCAS
    
    # Get CAS shape from page so variable can be used as a connector
    $AllShapes = $page.Shapes
    $CASShape = $AllShapes | where-object {$_.Name -eq $($SCCMServer_CAS_SiteServer.Name)} -ErrorAction SilentlyContinue

###################
    $CASSiteServerAngle = 360.0 / ($CASSCCMNonSiteServers.Count)

    if($CASSiteServerAngle -eq 180 -or $CASSiteServerAngle -eq 360)
    {
        $CASSiteServerConnectorLength = ($SquareSize * $PadRatio) * 3.0
    }
    else {
        $SinAngle = [Math]::Sin((Get-Radians -angle $CASSiteServerAngle))
        $CASSiteServerConnectorLength = (($SquareSize * $PadRatio) / $SinAngle)
    }
   
    $CurrentServerCount = 0
    ForEach($SiteSystem in $CASSCCMNonSiteServers){

        $CurentServerAngle = ($CASSiteServerAngle * $CurrentServerCount );
        $currentServerSinAngle = [Math]::Sin((Get-Radians -angle $CurentServerAngle))
        $SiteSeverOpp = $CASSiteServerConnectorLength * $currentServerSinAngle;
        $currentServerCosAngle = [Math]::Cos((Get-Radians -angle $CurentServerAngle));
        $SiteSystemAdj= $CASSiteServerConnectorLength * $currentServerCosAngle ;
        
        Set-VisioServer -Server $SiteSystem -x ($CenterX +$SiteSeverOpp) -y ($CenterY - $SiteSystemAdj) -RolesToIgnore $RolesToIgnore -page $Page -ConnectToShape $CASShape -IsCAS $False -SiteContainer $ContainerCAS
        
        $CurrentServerCount++

    }
    
}
###################
    
$Count = 0
ForEach($Site in $SCCM_SiteCodes_WithoutCAS){

    $SCCMServersCurrent = $SCCMServers | Where-Object {$_.Details_Table0_SiteCode -eq $Site}
    
    $GroupedSCCMSiteServers = $SCCMServersCurrent | Group-Object -Property Details_Table0_ServerName

    $SCCMSiteServers = $GroupedSCCMSiteServers | where-object {$_.Group.Where({$_.Details_Table0_RoleName -eq "Site Server"}) }
    $SCCMNonSiteSystems = $GroupedSCCMSiteServers | where-object {!($_.Group.Where({$_.Details_Table0_RoleName -eq "Site Server"})) }
    
    $currentAngle = ($ShapeAngle * $Count );
    $currentSinAngle = [Math]::Sin((Get-Radians -angle $currentAngle))
    $SiteOpp = $ConnectorLength * $currentSinAngle;
    $currentCosAngle = [Math]::Cos((Get-Radians -angle $currentAngle));
    $SiteAdj= $ConnectorLength * $currentCosAngle ;

    $SiteCenterX = $CenterX + $SiteOpp
    $SiteCenterY = $CenterY - $SiteAdj
    
    $SiteServerAngle = 360.0 / ($SCCMNonSiteSystems.Count)

    if($SiteServerAngle -eq 180 -or $SiteServerAngle -eq 360)
    {
        $SiteServerConnectorLength = ($SquareSize * $PadRatio) * 3.0
    }
    else {
        $SinAngle = [Math]::Sin((Get-Radians -angle $SiteServerAngle))
        $SiteServerConnectorLength = (($SquareSize * $PadRatio) / $SinAngle)
    }


    write-host "`nProcessing servers in site: $Site for Visio document creation..." -ForegroundColor Green
    write-host "Dropping container for site $Site at location $SiteCenterX,$SiteCenterY" -ForegroundColor Red -BackgroundColor White
      
    $TextSize = Get-VisioContainerTextSize -SiteSystemsCount $SCCMNonSiteSystems.Count -SquareSize $SquareSize -PadRatio $PadRatio
    $ContainerSite = New-VisioContainer -Stencil $ContainerStencil -CenterX $SiteCenterX -CenterY $SiteCenterY -Text $Site -TextSize $TextSize

    foreach($SiteServer in $SCCMSiteServers)
    {
        Set-VisioServer -Server $SiteServer -x $SiteCenterX -y $SiteCenterY -RolesToIgnore $RolesToIgnore -page $Page -ConnectToShape $ContainerCAS -SiteContainer $ContainerSite

        $AllShapes = $page.Shapes
        $SiteServerShape = $AllShapes | where-object {$_.Name -eq $($SiteServer.Name)} -ErrorAction SilentlyContinue
        
    }
    
    
   
    $CurrentServerCount = 0
    ForEach($SiteSystem in $SCCMNonSiteSystems){

        $CurentServerAngle = ($SiteServerAngle * $CurrentServerCount );
        $currentServerSinAngle = [Math]::Sin((Get-Radians -angle $CurentServerAngle))
        $SiteSeverOpp = $SiteServerConnectorLength * $currentServerSinAngle;
        $currentServerCosAngle = [Math]::Cos((Get-Radians -angle $CurentServerAngle));
        $SiteSystemAdj= $SiteServerConnectorLength * $currentServerCosAngle ;
        
        Set-VisioServer -Server $SiteSystem -x ($SiteCenterX +$SiteSeverOpp) -y ($SiteCenterY - $SiteSystemAdj) -RolesToIgnore $RolesToIgnore -page $Page -ConnectToShape $SiteServerShape -SiteContainer $ContainerSite
        
        $CurrentServerCount++

    }

    $Count++
}

 
#region Now Standalone Primary site
If($SCCM_Standalone_Primary_SiteCode){

     write-host "`nProcessing Standalone Primary site: $SCCM_Standalone_Primary_SiteCode for Visio document creation..." -ForegroundColor Green

 }
#endregion

Function PopulateVisioRowData {

    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $Shape,
        [parameter(Mandatory=$true)]
        $ShapeDataName,
        [parameter(Mandatory=$true)]
        $ShapeDataValue

    )

    $Row = $Shape.AddRow($visSectionProp,$visRowLast,$visTagDefault)
    $Shape.Section($visSectionProp).Row($Row).NameU = $($ShapeDataName -replace '[\W]', '')
    $Shape.CellsSRC($visSectionProp,$Row,$visCustPropsLabel).FormulaU = """$ShapeDataName"""
    $Shape.CellsSRC($visSectionProp,$Row,$visCustPropsLangID).FormulaU = "2057"
    $Shape.CellsSRC($visSectionProp,$Row,$visCustPropsValue).FormulaU = """$ShapeDataValue"""

}

Function Set-VisioShapeData {

    [CmdletBinding()]
    param(    
        [parameter(Mandatory=$true)]
        $Shape,
        [parameter(Mandatory=$true)]
        $SummaryData,
        [parameter(Mandatory=$true)]
        $AttribsData

    )

    If($AttribsData.Description){PopulateVisioRowData -Shape $Shape -ShapeDataName "Server Description" -ShapeDataValue $AttribsData.Description}
    If($AttribsData.OU){PopulateVisioRowData -Shape $Shape -ShapeDataName "OU" -ShapeDataValue $AttribsData.OU}
    If($SummaryData.BuildType){PopulateVisioRowData -Shape $Shape -ShapeDataName "BuildType" -ShapeDataValue $SummaryData.BuildType}
    If($SummaryData.model){PopulateVisioRowData -Shape $Shape -ShapeDataName "model" -ShapeDataValue $SummaryData.model}
    If($SummaryData.os){PopulateVisioRowData -Shape $Shape -ShapeDataName "os" -ShapeDataValue $SummaryData.os}
    If($AttribsData.VMware){PopulateVisioRowData -Shape $Shape -ShapeDataName "VMware" -ShapeDataValue $AttribsData.VMware}
    
    
}


#region Final shape/page configuratoin
write-host "`nGetting all the shapes on the page..."
$AllShapes = $page.Shapes
If($ContainerCAS){
    write-host "`Bring CAS Container shape to front" -ForegroundColor Green
    $ContainerCAS.BringToFront()
}

If($CASShape){
    write-host "Bring CAS server shape to front" -ForegroundColor Green
    $CASShape.BringToFront()
}

$page.ResizeToFitContents()

#endregion

#region get all the server shapes on the Visio page so that we can pull server info from the web api
$ServersOnPage = ($AllShapes | Where-Object {$_.Style -ne 'Connector' -and $_.NameU -notlike "*Translucent*"} | Select-Object Name).Name
$ServerNetbiosNames = @()
ForEach($Server in $ServersOnPage){
    $server=$server.ToString()
    $ServerNetbiosNames += $server.Split('.')[0]
}
ForEach($Server in $ServerNetbiosNames){
    
    $SummaryData = $null
    write-host "`nTry and get server $Server data from your web API..." -ForegroundColor Green

    Try{
        $SummaryData = Invoke-RestMethoD -URI http://serverdata.farm.domain.local/rest/servers/$server/data -UseDefaultCredentials -ErrorAction SilentlyContinue
        $AttribsData = Invoke-RestMethoD -URI http://serverdata.farm.domain.local/rest/servers/$server/attributes -UseDefaultCredentials -ErrorAction SilentlyContinue
    } Catch{
        write-host "No server name: $Server could be found" -ForegroundColor Yellow
    }

    If($SummaryData){
        write-host "Populating Visio shape data with data retrieved from web API for server: $Server" -ForegroundColor Green
        $VisioServer = $page.Shapes | Where-Object {$_.Name -like "$Server*"}
        Set-VisioShapeData -Shape $VisioServer -SummaryData $SummaryData -AttribsData $AttribsData
 
    }

}
#endregion

# SaveAs the diagram as VSDX and JPG format
Try{
    $VSDX = $SaveAsPath+"\"+$SaveAsFileName+"_$(get-date -Format yyyyMMdd).vsdx"
    $JPG = $SaveAsPath+"\"+$SaveAsFileName+"_$(get-date -Format yyyyMMdd).jpg"
    $Document.SaveAs($VSDX)
    $page.Export($JPG)
} catch {
    write-host "`nCould not save all files. Please manually saveas from Visio."
}
