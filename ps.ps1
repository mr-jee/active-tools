Import-Module ActiveDirectory
function find_user_by_employeeId{
    param($employee_id)
    $user = Get-ADUser -Filter "EmployeeID -eq $employee_id" -Properties SamAccountName,EmployeeId,DisplayName,mail,Title,Manager,"msDS-UserPasswordExpiryTimeComputed"
    if ($user){
        $expire_time = $user."msDS-UserPasswordExpiryTimeComputed"
        $expire_time_in_days = [math]::Ceiling(($expire_time -(Get-Date).ToFileTime())/864000000000)
        Write-Host "Display Name: " $user.DisplayName
        Write-Host "Username: " $user.samAccountName
        Write-Host "Personnel Code: " $user.EmployeeId
        Write-Host "Email Address: " $user.Mail
        Write-Host "Job Title: " $user.Title
        Write-Host "Report To: " $user.Manager
        Write-Host "Password Validity: " $expire_time_in_days "Days"
        }
    else{
        Write-Host "User Not Found"
    }
}

function find_user_by_samaccountname{
    param($samaccountname)
    $user = Get-ADUser -Filter {sAMAccountName -eq $samaccountname} -Properties sAMAccountName,EmployeeId,DisplayName,mail,Title,Manager,"msDS-UserPasswordExpiryTimeComputed"
    if ($user){
        $expire_time = $user."msDS-UserPasswordExpiryTimeComputed"
        $expire_time_in_days = [math]::Ceiling(($expire_time -(Get-Date).ToFileTime())/864000000000)
        Write-Host "Display Name: " $user.DisplayName
        Write-Host "Username: " $user.samAccountName
        Write-Host "Personnel Code: " $user.EmployeeId
        Write-Host "Email Address: " $user.Mail
        Write-Host "Job Title: " $user.Title
        Write-Host "Report To: " $user.Manager
        Write-Host "Password Validity: " $expire_time_in_days "Days"
        }
    else{
    Write-Host "User Not Found"
    }
}

function is_numeric{
    # This strange function check if user input is number or string
    param ($value)
    return $value -match "^[\d\.]+$"
}

$user_input = Read-Host -Prompt "Please insert Personnel Code or User Name:`n"
if(is_numeric $user_input){
    find_user_by_employeeId -employee_id $user_input
}
else{
    find_user_by_samaccountname -samaccountname $user_input
}
