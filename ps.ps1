Import-Module ActiveDirectory
function find_user(){
    $user_input = Read-Host -Prompt "Please insert Personnel Code or User Name:`n"
    if(is_numeric $user_input){
        $user = find_user_by_employeeId -employee_id $user_input
        return $user
}
    else{
        $user = find_user_by_samaccountname -samaccountname $user_input
        return $user
    }
}

function find_user_by_employeeId{
    param($employee_id)
    $user = Get-ADUser -Filter "EmployeeID -eq $employee_id" -Properties SamAccountName,EmployeeId,DisplayName,mail,Title,Manager,"msDS-UserPasswordExpiryTimeComputed"
    if ($user -ne $null){
        return $user
        }
    else{
        Write-Host "User Not Found"
        exit
    }
}

function find_user_by_samaccountname{
    param($samaccountname)
    $user = Get-ADUser -Filter {sAMAccountName -eq $samaccountname} -Properties sAMAccountName,EmployeeId,DisplayName,mail,Title,Manager,"msDS-UserPasswordExpiryTimeComputed"
    if ($user -ne $null){
        return $user
        }
    else{
    Write-Host "User Not Found"
    exit
    }
}

function is_numeric{
    # This regex function check if user input is number or string
    param ($value)
    return $value -match "^[\d\.]+$"
}

function change_password($user){
    $change_password = Read-Host -Prompt " Do you want to change the password? yes/no"
    $change_password = $change_password.ToLower()

    if ($change_password.Equals( "yes")){
        $new_pass = Read-Host -Prompt "Enter the new password: " -AsSecureString
        Set-ADAccountPassword -Identity $user -NewPassword $new_pass -Reset
        Write-Host "Password of $($user.SamAccountName) changed."
}
}

$user = $null

$user = find_user

change_password -user $user





 <#     $expire_time = $user."msDS-UserPasswordExpiryTimeComputed"
        $expire_time_in_days = [math]::Ceiling(($expire_time -(Get-Date).ToFileTime())/864000000000)
        Write-Host "Display Name: " $user.DisplayName
        Write-Host "Username: " $user.samAccountName
        Write-Host "Personnel Code: " $user.EmployeeId
        Write-Host "Email Address: " $user.Mail
        Write-Host "Job Title: " $user.Title
        Write-Host "Report To: " $user.Manager
        Write-Host "Password Validity: " $expire_time_in_days "Days"#>
