#########
# au2mator PS Services
# New Service
# AD - Update User Details
# v 1.0 Initial Release
# Init Release: 03.02.2020
# Last Update: 03.02.2020
# Code Template V 1.1
# URL: https://au2mator.com/update-user-details-active-directory-self-service-with-au2mator/
# Github: https://github.com/au2mator/AD-Update-User-Details
#################




#region InputParamaters
##Question in au2mator
param (
    [parameter(Mandatory = $false)] 
    [String]$c_CheckManager,

    [parameter(Mandatory = $false)] 
    [String]$c_Manager,



    [parameter(Mandatory = $false)] 
    [String]$c_CheckPhone,

    [parameter(Mandatory = $false)] 
    [String]$c_MobilePhone,

    [parameter(Mandatory = $false)] 
    [String]$c_HomePhone,

    [parameter(Mandatory = $false)] 
    [String]$c_Fax,

    [parameter(Mandatory = $false)] 
    [String]$c_OfficePhone,
 


    [parameter(Mandatory = $false)] 
    [String]$c_CheckAddress, 

    [parameter(Mandatory = $false)] 
    [String]$c_City,

    [parameter(Mandatory = $false)] 
    [String]$c_StreetAddress,

    [parameter(Mandatory = $false)] 
    [String]$c_PostalCode,

    [parameter(Mandatory = $false)] 
    [String]$c_State,

    [parameter(Mandatory = $false)] 
    [String]$c_POBox,

    [parameter(Mandatory = $false)] 
    [String]$c_Country,




    [parameter(Mandatory = $false)] 
    [String]$c_CheckOrganization,

    [parameter(Mandatory = $false)] 
    [String]$c_Company,

    [parameter(Mandatory = $false)] 
    [String]$c_Department,

    [parameter(Mandatory = $false)] 
    [String]$c_JobTitle,



    [parameter(Mandatory = $false)] 
    [String]$c_Comment, 

    ## au2mator Initialize Data
    [parameter(Mandatory = $false)] 
    [String]$InitiatedBy, 

    [parameter(Mandatory = $false)] 
    [String]$RequestId, 
 
    [parameter(Mandatory = $false)] 
    [String]$Service, 
 
    [parameter(Mandatory = $false)] 
    [String]$TargetUserId
)
#endregion  InputParamaters

#region Variables
##Script Handling
$DoImportPSSession = $false
$ErrorCount = 0


## Environment
[string]$DCServer = 'svdc01'
[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\AD - Update User Details"
[string]$LogfileName = "Update User Details"

## au2mator Settings
[string]$PortalURL = "http://demo01.au2mator.local"
[string]$au2matorDBServer = "demo01"
[string]$au2matorDBName = "au2mator"

## Control Mail
$SendMailToInitiatedByUser = $true #Send a Mail after Service is completed
$SendMailToTargetUser = $true #Send Mail to Target User after Service is completed

## SMTP Settings
$SMTPServer = "smtp.office365.com"
$SMTPUser = "mail@au2mator.com"
$SMTPPassword = "Password1"
$SMPTAuthentication = $true #When True, User and Password needed
$EnableSSLforSMTP = $true
$SMTPSender = "mail@au2mator.com"
#endregion Variables

#region Functions

function ConnectToDB {
    # define parameters
    param(
        [string]
        $servername,
        [string]
        $database
    )
    # create connection and save it as global variable
    $global:Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "server='$servername';database='$database';trusted_connection=false; integrated security='true'"
    $Connection.Open()
    Write-Verbose 'Connection established'
}

function ExecuteSqlQuery {
    # define parameters
    param(
      
        [string]
        $sqlquery
     
    )
     
    Begin {
        If (!$Connection) {
            Throw "No connection to the database detected. Run command ConnectToDB first."
        }
        elseif ($Connection.State -eq 'Closed') {
            Write-Verbose 'Connection to the database is closed. Re-opening connection...'
            try {
                # if connection was closed (by an error in the previous script) then try reopen it for this query
                $Connection.Open()
            }
            catch {
                Write-Verbose "Error re-opening connection. Removing connection variable."
                Remove-Variable -Scope Global -Name Connection
                throw "Unable to re-open connection to the database. Please reconnect using the ConnectToDB commandlet. Error is $($_.exception)."
            }
        }
    }
     
    Process {
        #$Command = New-Object System.Data.SQLClient.SQLCommand
        $command = $Connection.CreateCommand()
        $command.CommandText = $sqlquery
     
        Write-Verbose "Running SQL query '$sqlquery'"
        try {
            $result = $command.ExecuteReader()      
        }
        catch {
            $Connection.Close()
        }
        $Datatable = New-Object "System.Data.Datatable"
        $Datatable.Load($result)
        return $Datatable         
    }
    End {
        Write-Verbose "Finished running SQL query."
    }
}

function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )
       
    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> <{2}> <{3}> {4}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $RequestId, $Service, $Text
    Add-Content -Path $logFile -Value $logEntry
}

function Get-UserInput ($RequestID) {
    [hashtable]$return = @{ }

    ConnectToDB -servername $au2matorDBServer -database $au2matorDBName 

    $Result = ExecuteSqlQuery -sqlquery "SELECT        RPM.Text AS Question, RP.Value
    FROM            dbo.Requests AS R INNER JOIN
                             dbo.RunbookParameterMappings AS RPM ON R.ServiceId = RPM.ServiceId INNER JOIN
                             dbo.RequestParameters AS RP ON RPM.ParameterName = RP.[Key] AND R.RequestId = RP.RequestId
    where RP.RequestId = '$RequestID' order by [Order]"
    
    $html = "<table><tr><td><b>Question</b></td><td><b>Answer</b></td></tr>"
    $html = "<table>"
    foreach ($row in $Result) { 
        $row
        $html += "<tr><td><b>" + $row.Question + "</b></td><td>" + $row.Value + "</td></tr>" 
    }
    $html += "</table>"

    $f_RequestInfo = ExecuteSqlQuery -sqlquery "select InitiatedBy, TargetUserId,[ApprovedBy], [ApprovedTime], Comment from Requests where RequestId =  '$RequestID'"

    $Connection.Close()
    Remove-Variable -Scope Global -Name Connection

    $f_SamInitiatedBy = $f_RequestInfo.InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties Mail

    $f_SamTarget = $f_RequestInfo.TargetUserId.Split("\")[1]
    $f_UserTarget = Get-ADUser -Identity $f_SamTarget -Properties Mail

    $return.InitiatedBy = $f_RequestInfo.InitiatedBy
    $return.MailInitiatedBy = $f_UserInitiatedBy.mail
    $return.MailTarget = $f_UserTarget.mail
    $return.TargetUserId = $f_RequestInfo.TargetUserId
    $return.ApprovedBy = $f_RequestInfo.ApprovedBy
    $return.ApprovedTime = $f_RequestInfo.ApprovedTime
    $return.Comment = $f_RequestInfo.Comment
    $return.HTML = $HTML

    return $return
}

Function Get-MailContent ($RequestID, $RequestTitle, $EndDate, $TargetUserId, $InitiatedBy, $Status, $PortalURL, $RequestedBy, $AdditionalHTML, $InputHTML) {

    $f_RequestID = $RequestID
    $f_InitiatedBy = $InitiatedBy
    
    $f_RequestTitle = $RequestTitle
    $f_EndDate = $EndDate
    $f_RequestStatus = $Status
    $f_RequestLink = "$PortalURL/requeststatus?id=$RequestID"
    $f_RequestedBy = $RequestedBy
    $f_HTMLINFO = $AdditionalHTML
    $f_InputHTML = $InputHTML
    
    $f_SamInitiatedBy = $f_InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties DisplayName
    $f_DisplaynameInitiatedBy = $f_UserInitiatedBy.DisplayName

    
    $HTML = @'    
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 1.5pt; background: #F7F8F3; mso-yfti-tbllook: 1184;" border="0" width="100%" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="padding: .75pt .75pt .75pt .75pt;" valign="top">&nbsp;</td>
    <td style="width: 450.0pt; padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top" width="600">
    <div style="box-sizing: border-box;">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: white; border: solid #E9E9E9 1.0pt; mso-border-alt: solid #E9E9E9 .75pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="1" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="border: none; background: #6ddc36; padding: 15.0pt 0cm 15.0pt 15.0pt;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><img src="https://au2mator.com/wp-content/uploads/2018/02/HPLogoau2mator-1.png" alt="" width="198" height="43" /></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="border: none; padding: 15.0pt 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 55.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="55%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="width: 18.75pt; border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="25">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm; font-color: #0000;"><strong>End Date</strong>: ##EndDate</td>
    </tr>
    <tr style="mso-yfti-irow: 1;">
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Status</strong>: ##Status</td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes;">
    <td style="border: solid #E3E3E3 1.0pt; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border: solid #E3E3E3 1.0pt; border-left: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Requested By</strong>: ##RequestedBy</td>
    </tr>
    </tbody>
    </table>
    </td>
    <td style="width: 5.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="5%">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 9.0pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    <td style="width: 40.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="40%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #FAFAFA; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="width: 100.0%; border: solid #E3E3E3 1.0pt; mso-border-alt: solid #E3E3E3 .75pt; padding: 7.5pt 0cm 1.5pt 3.75pt;" width="100%">
    <p style="text-align: center;" align="center"><span style="font-size: 10.5pt; color: #959595;">Request ID</span></p>
    <p class="MsoNormal" style="text-align: center;" align="center">&nbsp;</p>
    <p style="text-align: center;" align="center"><u><span style="font-size: 12.0pt; color: black;"><a href="##RequestLink"><span style="color: black;">##REQUESTID</span></a></span></u></p>
    <p class="MsoNormal" style="text-align: center;" align="center"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><strong><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">Dear ##UserDisplayname,</span></strong></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">the Request <strong>"##RequestTitle"</strong> has been finished.<br /> <br /> Please see the description for detailed information.<br /><b>##HTMLINFO&nbsp;</b><br /></span></p>
    <div>&nbsp;</div>
    <div>See the Details of the Request</div>
    <div>##InputHTML</div>
    <div>&nbsp;</div>
    <div>&nbsp;</div>
    Kind regards,<br /> au2mator Self Service Team
    <p>&nbsp;</p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';"><a style="border-radius: 3px; -webkit-border-radius: 3px; -moz-border-radius: 3px; display: inline-block;" href="##RequestLink"><strong><span style="color: white; border: solid #50D691 6.0pt; padding: 0cm; background: #50D691; text-decoration: none; text-underline: none;">View your Request</span></strong></a></span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 3; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #333333; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 50.0%; border: none; border-right: solid lightgrey 1.0pt; mso-border-right-alt: solid lightgrey .75pt; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    <td style="width: 50.0%; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </div>
    </td>
    <td style="padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    <p class="MsoNormal"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
'@

    $html = $html.replace('##REQUESTID', $f_RequestID).replace('##UserDisplayname', $f_DisplaynameInitiatedBy).replace('##RequestTitle', $f_RequestTitle).replace('##EndDate', $f_EndDate).replace('##Status', $f_RequestStatus).replace('##RequestedBy', $f_RequestedBy).replace('##HTMLINFO', $f_HTMLINFO).replace('##InputHTML', $f_InputHTML).replace('##RequestLink', $f_RequestLink)

    return $html
}

function Send-ServiceMail ($HTMLBody, $ServiceName, $Recipient, $RequestID, $RequestStatus) {
    $f_Subject = "au2mator - $ServiceName Request [$RequestID] - $RequestStatus"

    if ($SMPTAuthentication) {
        $f_secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
        $f_mycreds = New-Object System.Management.Automation.PSCredential ($SMTPUser, $f_secpasswd)
    
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $f_mycreds -UseSsl
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $f_mycreds
        }
    }
    else {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -UseSsl
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high
        } 
    }   
}

##Custom Functions
Function Update-UserProperty ($SamAccountName, $ADProperty, $ADValue) {
    Write-au2matorLog -Type INFO -Text "Try to update $ADProperty Properties with Value: $ADValue" 

    try {

        switch ($ADProperty) {
            Manager { Set-ADUser -Identity $SamAccountName -Manager $ADValue }

            MobilePhone { Set-ADUser -Identity $SamAccountName -MobilePhone $ADValue } 
            Fax { Set-ADUser -Identity $SamAccountName -Fax $ADValue } 
            HomePhone { Set-ADUser -Identity $SamAccountName -HomePhone $ADValue }         
            OfficePhone { Set-ADUser -Identity $SamAccountName -OfficePhone $ADValue }

            City { Set-ADUser -Identity $SamAccountName -City $ADValue }
            StreetAddress { Set-ADUser -Identity $SamAccountName -StreetAddress $ADValue }
            PostalCode { Set-ADUser -Identity $SamAccountName -PostalCode $ADValue }
            State { Set-ADUser -Identity $SamAccountName -State $ADValue }
            POBox { Set-ADUser -Identity $SamAccountName -POBox $ADValue }
            Country { Set-ADUser -Identity $SamAccountName -Country $ADValue }

            Company { Set-ADUser -Identity $SamAccountName -Company $ADValue }
            Department { Set-ADUser -Identity $SamAccountName -Department $ADValue }
            JobTitle { Set-ADUser -Identity $SamAccountName -Title $ADValue }
        }

        Write-au2matorLog -Type INFO -Text "$ADProperty Property has been updated"
        $f_ErrorCount = 0
    }
    catch {
        $f_ErrorCount = 1
        Write-au2matorLog -Type ERROR -Text "Error to update $ADProperty Property"
        Write-au2matorLog -Type ERROR -Text $Error
    }
    return $f_ErrorCount
}

#endregion Functions

#region Script
Write-au2matorLog -Type INFO -Text "Start Script"

if ($DoImportPSSession) {

    Write-au2matorLog -Type INFO -Text "Import-Pssession"
    $PSSession = New-PSSession -ComputerName $DCServer
    Import-PSSession -Session $PSSession -DisableNameChecking -AllowClobber
}
else {
        
}

Write-au2matorLog -Type INFO -Text "Import AD PS Module"
Import-Module ActiveDirectory

$SamTargetUser = $TargetUserId.Split("\")[1]
Write-au2matorLog -Type INFO -Text "Try to Update User Details for User $SamTargetUser"

If ($c_CheckManager -eq "True") {
    Write-au2matorLog -Type INFO -Text "Manager Checkbox"
    if ($c_Manager) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "Manager" -ADValue $c_Manager
    }
}

if ($c_CheckPhone -eq "True") {
    Write-au2matorLog -Type INFO -Text "Phone Checkbox"
    
    if ($c_MobilePhone) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "MobilePhone" -ADValue $c_MobilePhone
    }
    if ($c_HomePhone) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "HomePhone" -ADValue $c_HomePhone    
    }
    if ($c_Fax) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "Fax" -ADValue $c_Fax
    }
    if ($c_OfficePhone) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "OfficePhone" -ADValue $c_OfficePhone
    }
}

if ($c_CheckAddress -eq "True") {
    Write-au2matorLog -Type INFO -Text "Address Checkbox"
    
    if ($c_City) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "City" -ADValue $c_City
    }
    if ($c_StreetAddress) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "StreetAddress" -ADValue $c_StreetAddress    
    }
    if ($c_PostalCode) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "PostalCode" -ADValue $c_PostalCode
    }
    if ($c_State) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "State" -ADValue $c_State
    }
    if ($c_POBox) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "POBox" -ADValue $c_POBox
    }
    if ($c_Country) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "Country" -ADValue $c_Country
    }
}

if ($c_CheckOrganization -eq "True") {
    Write-au2matorLog -Type INFO -Text "Organization Checkbox"

    if ($c_Company) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "Company" -ADValue $c_Company
    }
    if ($c_Department) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "Department" -ADValue $c_Department    
    }
    if ($c_JobTitle) {
        $ErrorCount += Update-UserProperty -SamAccountName $SamTargetUser -ADProperty "JobTitle" -ADValue $c_JobTitle
    }
}




if ($ErrorCount -eq 0) {
    $au2matorReturn = "Properties for User $c_User updated"
    $AdditionalHTML = "<br>
        User " + (Get-ADUser -identity $c_User).DisplayName + " has nupdated Properties
        <br>
        "
    $Status = "COMPLETED"
}
else {
    $au2matorReturn = "failed to update Properties for $c_User, Error: $Error"
    $Status = "ERROR"
}

#endregion Script


#region Return
## return to au2mator Services

Write-au2matorLog -Type INFO -Text "Service finished"

if ($SendMailToInitiatedByUser) {    
    Write-au2matorLog -Type INFO -Text "Send Mail to Initiated By User"

    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailInitiatedBy)" -RequestStatus $Status -ServiceName $Service
}


if ($SendMailToTargetUser) {    
    Write-au2matorLog -Type INFO -Text "Send Mail to Target User"
    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailTarget)" -RequestStatus $Status -ServiceName $Service
}


return $au2matorReturn    
#endregion Return


