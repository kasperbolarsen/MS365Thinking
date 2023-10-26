
function Add-UserDataToList ($connection, $list, $ct)
{
    #load all users from UPA ( No birthday in graph yet)
    $users = Submit-PnPSearchQuery -Query "*" -SourceId "b09a7990-05ea-4af9-81ef-edfab16c4e31" -SelectProperties "Title,WorkEmail,RefinableDate00, RefinableDate01" -All -Connection $c
    foreach($user in $users.ResultRows)
    {
        # exclude those without an email address that matches this domain
        # add additional logic to exclude users that should not be included
        # like SPS-HideFromAddressLists and so on
        # see https://ms365thinking.blogspot.com/2023/10/disabled-or-inactive-user-accounts-in.html for inspiration
        if($null -eq $User.WorkEmail -or $User.WorkEmail.IndexOf($emaildomain) -eq -1 )
        {
            continue
        }
        if($user.RefinableDate00 -and $user.RefinableDate01)
        {
            $birthdate = [datetime] $user.RefinableDate00
            [datetime]$birthdayThisYear = "$((Get-Date).Year)-$($birthdate.Month)-$($birthdate.Day)"

            $hiredate = [datetime] $user.RefinableDate01
            [int]$nextWorkAnniversaryYear = (Get-Date).Year - $hiredate.Year +1
            [datetime]$hiredateThisYear = "$((Get-Date).Year)-$($hiredate.Month)-$($hiredate.Day)"
        
            Add-PnPListItem -List $list -ContentType $ct -Values @{"Title"=$user.Title;"birthdayhiredate_Account"=$user.WorkEmail;"nextWorkAnniversaryInYears" = $nextWorkAnniversaryYear;"birthdayThisYear"=$birthdayThisYear;"HiredateThisYear"=$hiredateThisYear} -Connection $c
        }
        else
        {
            if($user.RefinableDate00)
            {
                $birthdate = [datetime] $user.RefinableDate00
                [datetime]$birthdayThisYear = "$((Get-Date).Year)-$($birthdate.Month)-$($birthdate.Day)"
                
                Add-PnPListItem -List $list -ContentType $ct -Values @{"Title"=$user.Title;"birthdayhiredate_Account"=$user.WorkEmail;"birthdayThisYear"=$birthdayThisYear} -Connection $c
            }
            elseif ($user.RefinableDate01) 
            {
                $hiredate = [datetime] $user.RefinableDate01
                [datetime]$hiredateThisYear = "$((Get-Date).Year)-$($hiredate.Month)-$($hiredate.Day)"
                
                [int]$nextWorkAnniversaryYear = (Get-Date).Year - $hiredate.Year +1
                Add-PnPListItem -List $list -ContentType $ct -Values @{"Title"=$user.Title;"birthdayhiredate_Account"=$user.WorkEmail;"nextWorkAnniversaryInYears"= $nextWorkAnniversaryYear;"HiredateThisYear"=$hiredateThisYear} -Connection $c
            }
        }
    }    
}
function Create-List($connection, $ListName, $ct) 
{
    try {
        $list = Get-PnPList -Identity $ListName -Connection $connection -ErrorAction Stop    
    }
    catch {
        $list = $null
    }
    
    if($null -eq $list)
    {
        $list = New-PnPList -Title $ListName -Template GenericList -Url $ListName -Connection $connection
    }
    Add-PnPContentTypeToList -List $list -ContentType $ct -DefaultContentType -Connection $connection
    #Create-View -connection $c -list $list  TODO: create view
    return $list
}
function Create-ContentType ($connection, $contentTypeName)
{
    
    $ct = Get-PnPContentType -Identity $contentTypeName  -Connection $connection
    $birthdayhiredate_Account = Get-PnPField -Identity "birthdayhiredate_Account" -Connection $c
    if($birthdayhiredate_Account -eq $null)
    {
        $birthdayhiredate_Account = Add-PnPField -DisplayName "Person" -InternalName "birthdayhiredate_Account" -Type User -Group "Custom date fields" -Connection $c
    }
    
    

    $birthdayThisYear = Get-PnPField -Identity "birthdayThisYear" -Connection $c
    if($birthdayThisYear -eq $null)
    {
        $birthdayThisYear = Add-PnPField -DisplayName "birthdayThisYear" -InternalName "birthdayThisYear" -Type DateTime -Group "Custom date fields" -Connection $c
    }
    
    $HiredateThisYear = Get-PnPField -Identity "HiredateThisYear" -Connection $c #ootb field
    if($null -eq $HiredateThisYear)
    {
        $HiredateThisYear = Add-PnPField -DisplayName "HiredateThisYear" -InternalName "HiredateThisYear" -Type DateTime -Group "Custom date fields" -Connection $c
    }
    
    $nextWorkAnniversaryInYears = Get-PnPField -Identity "nextWorkAnniversaryInYears" -Connection $c
    if ($null -eq $nextWorkAnniversaryInYears) 
    {
        $nextWorkAnniversaryInYears = Add-PnPField -DisplayName "nextWorkAnniversaryInYears" -InternalName "nextWorkAnniversaryInYears" -Type Number -Group "Custom fields" -Connection $c
    }
    
    if ($null -eq $ct) 
    {
        $ct = Add-PnPContentType -Name $contentTypeName -Group "birthday" -Description "BirthdaySyncItem" -Connection $c
        Add-PnPFieldToContentType -Field $birthdayhiredate_Account -ContentType $ct -Connection $c
        Add-PnPFieldToContentType -Field $birthdayThisYear -ContentType $ct -Connection $c
        add-PnPFieldToContentType -Field $HiredateThisYear -ContentType $ct -Connection $c
        add-PnPFieldToContentType -Field $nextWorkAnniversaryInYears -ContentType $ct -Connection $c
    }
    else 
    {
        Add-PnPFieldToContentType -Field $birthdayhiredate_Account -ContentType $ct -Connection $c
        Add-PnPFieldToContentType -Field $birthdayThisYear -ContentType $ct -Connection $c
        add-PnPFieldToContentType -Field $HiredateThisYear -ContentType $ct -Connection $c
        add-PnPFieldToContentType -Field $nextWorkAnniversaryInYears -ContentType $ct -Connection $c
    }
    return $ct
}

#working with birthday and work anniversity in search
$siteUrl = "https://[tenant].sharepoint.com"
$ListName = "BirthdayAndHiredateSyncList"
$emaildomain = "[tenant].onmicrosoft.com"
$contentTypeName = "BirthdayHireDateSyncItem"
if($c -eq $null)
{
    $c = Connect-PnPOnline -url $siteUrl -ReturnConnection -Interactive
}

$ct = Create-ContentType -connection $c -contentTypeName $contentTypeName
$list = Create-List -ListName $ListName -ct $ct -connection $c 
    

#flush the list
$items = Get-PnPListItem -List $list -Connection $c
foreach($item in $items)
{
    Remove-PnPListItem -List $list -Identity $item.Id -Connection $c -Force
}

Add-UserDataToList -connection $c -list $list -ct $ct


    
    
    


