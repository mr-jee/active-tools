import subprocess
import json


def search_by_employee_id(id):
    employee_id = id
    fetch_user_by_employeeid = f'''
       $employee_id = "{id}"
    $PaddedEmployeeID = $employee_id.PadLeft(8,'0')
    $users = Get-ADUser -Filter {{EmployeeID -eq $PaddedEmployeeID -or EmployeeID -eq $employee_id}} -Properties sAMAccountName,EmployeeId,DisplayName,Mail,Manager,Title,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
        Select-Object sAMAccountName, EmployeeId, DisplayName, Mail, Manager, Title,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
        ConvertTo-Json
    $users
    '''
    search_result = run_ps(fetch_user_by_employeeid)
    if search_result:
        return search_result
    else:
        print("ERROR")
        return None


def choose_user(result_of_search):
    if isinstance(result_of_search, list):
        print("+|++" * 50, "\nUsers: ")
        for index, item in enumerate(result_of_search, start=1):
            print(f"{index}-{item['sAMAccountName']}")
        choice = input("Which User do you wanna reset password? ")
        try:
            choice_index = int(choice)
            if 1 <= choice_index <= len(result_of_search):
                selected_option = result_of_search[choice_index - 1]
                print("You selected:", selected_option['sAMAccountName'])
                samaccountname = selected_option['sAMAccountName']
                return samaccountname
            else:
                print("Invalid Option")
        except ValueError:
            print("Inavlid Input. Please enter a number")
    else:
        samaccountname = result_of_search['sAMAccountName']
        return samaccountname


def run_ps(powershell_script):
    try:
        result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
                                capture_output=True, text=True,
                                check=True)
        print(f"result is: {result.stdout}")
        if result.stdout.strip():  # Check if the output is not empty
            convert_to_json = json.loads(result.stdout)
            return convert_to_json
        else:
            print("User not found.")
            return None
    except subprocess.CalledProcessError as e:
        print(f"Error executing PowerShell script: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return None


employee_id = input("Please enter employee id: ")
result_of_search = search_by_employee_id(employee_id)
choose_user(result_of_search)
