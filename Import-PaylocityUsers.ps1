$CurEmps = Import-Csv '.\Employee Listing 11.20.19.csv'
$adusers = Get-ADUser -Filter *
$names = $adusers | foreach {$_.'GivenName' + $_.'SurName'}
$msolusers = Get-MsolUser -All
$msolnames = $msolusers | foreach {$_.'FirstName' + $_.'LastName'}
$domain = "<Domain>" # ex: "example"

foreach ($emp in $CurEmps) {
    
    $first = (Get-Culture).TextInfo.ToTitleCase($emp.'First Name'.ToLower())
    $last = (Get-Culture).TextInfo.ToTitleCase($emp.'Last Name'.toLower())
    $name =  $first + " " + $last
    $GenSam = $first + $last
    $dep = ($emp.'Department - Current' -split " - ")[1].ToString().trim()
    $depid =($emp.'Department - Current' -split " - ")[0].ToString().trim()
    $empId = $emp.'Employee ID'
    $NHD= $emp.'Hire Date - Current'
    $supFirst = $emp.'Supervisor - Current'.Split(",")[1].trim() 
    $supLast = $emp.'Supervisor - Current'.Split(",")[0].trim()

    if($names -contains $GenSam -and $msolnames -contains $gemsam){
        #Write-Host $emp.'First Name' $emp.'Last Name' -ForegroundColor Green
        if(get-aduser -Filter {givenName -eq $first -AND surname -eq $last} ){
            $aduser = get-aduser -Filter {givenName -eq $first -AND surname -eq $last}
            $aduser | Set-ADUser -Department $dep –replace @{DepartmentNumber=$depid}
            $aduser | Set-ADUser –replace @{NewHireDate=$NHD}
            $aduser | Set-ADUser –replace @{EmployeeID=$empid}
            if(get-aduser -Filter {givenName -eq $supfirst -AND surname -eq $suplast}){
                $aduser | Set-ADUser -Manager $(get-aduser -Filter {givenName -eq $supfirst -AND surname -eq $suplast})
            }
            else{
                write-host $emp.'First Name' $emp.'Last Name' "Has no sup" -ForegroundColor Yellow
                Write-Host $supFirst $supLast -ForegroundColor Yellow
            }
        }
        else{
            write-host "AD ACCOUNT NOT FOUND" -ForegroundColor Red
        }
    }
    else{
        # 
        if($msolnames -contains $gemsam){
            #No ad , Yes Email
            #Write-Host "Email found for $GenSam"
            Write-Host $emp.'First Name' $emp.'Last Name' -ForegroundColor red
            if(get-aduser -Filter {givenName -eq $supfirst -AND surname -eq $suplast}){
                $manager = get-aduser -Filter {givenName -eq $supfirst -AND surname -eq $suplast}
                $name = $name.replace("-","")
                $name = $name.ToString()
                $GenSam = $GenSam.Replace(" ","").replace(".","")
                if($GenSam.Length -gt 20){
                    $gensam = $GenSam.SubString(1,20)
                }
                write-host $GenSam ":" $name
                New-ADUser -SamAccountName $GenSam -Name $name -GivenName $first -Surname $last -path "OU=To Be Moved,DC=$($domain),DC=org" -Enabled $true -Manager $manager -AccountPassword ((New-Guid).Guid  | ConvertTo-SecureString -AsPlainText -Force) -UserPrincipalName ($GenSam + "@$($domain).org")
            }
            else{
                write-host $emp.'First Name' $emp.'Last Name' "Has no sup" -ForegroundColor Yellow
                Write-Host $supFirst $supLast -ForegroundColor Yellow
                $GenSam = $GenSam.Replace(" ","").replace(".","")
                if($GenSam.Length -gt 20){
                    $gensam = $GenSam.SubString(1,20)
                }
                New-ADUser -SamAccountName $GenSam -Name $name -GivenName $first -Surname $last -path "OU=To Be Moved,DC=$($domain),DC=org" -Enabled $true -AccountPassword ((New-Guid).Guid  | ConvertTo-SecureString -AsPlainText -Force) -UserPrincipalName ($GenSam + "@$($domain).org")
            }
       }
       elseif($names -contains $GenSam -and $msolnames -notcontains $gemsam){
       Write-Host write-host $GenSam
       }
       else{
            #No Email No AD
            Write-Host $emp.'First Name' $emp.'Last Name' -ForegroundColor red -BackgroundColor Black
            if(get-aduser -Filter {givenName -eq $supfirst -AND surname -eq $suplast}){
                $manager = get-aduser -Filter {givenName -eq $supfirst -AND surname -eq $suplast}
                $name = $name.replace("-","")
                $name = $name.ToString()
                Write-Host $manager
                $GenSam = $GenSam.Replace(" ","").replace(".","")
                if($GenSam.Length -gt 20){
                    $gensam = $GenSam.SubString(1,20)
                }
                New-ADUser -SamAccountName $GenSam -Name $name -GivenName $first -Surname "$last" -path "OU=To Be Moved,DC=$($domain),DC=org" -Enabled $true -Manager $manager -AccountPassword ((New-Guid).Guid  | ConvertTo-SecureString -AsPlainText -Force) -UserPrincipalName ($GenSam + "@$($domain).org")
            }
      }
    }

}
