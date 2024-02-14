import subprocess
import json

import pandas


def run_powershell_command(powershell_script, expect_json=True):
    """ This method meant to run powershell commands, it takes powershell command as a variable, if JSON output is needed(like data of a user) expect_json
    equals True, if no json needed, set it to False"""
    # "-ExecutionPolicy", "Bypass" cause that script can run through python and expire time will update
    try:
        print("Executing PowerShell command:")
        print(powershell_script)
        result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
                                capture_output=True, text=True,
                                check=True)

        if expect_json:
            try:
                print(f"stdout is: {result.stdout}")
                return json.loads(result.stdout)
            except json.JSONDecodeError:
                # Return None if JSON parsing fails
                print("JSONDecode ERROR", result.stdout)
                print("Error decoding JSON. Output:", result.stdout)
                return None
        else:
            return result.stdout
    except subprocess.CalledProcessError as e:
        print("PowerShell Error:", e)
        print(e)
        return None


# def run_powershell_command(powershell_script, expect_json=True):
#     """ This method meant to run PowerShell commands. It takes PowerShell command as a variable.
#     If JSON output is needed (like data of a user), expect_json equals True; if no JSON needed, set it to False."""
#     # "-ExecutionPolicy", "Bypass" cause that script can run through python and expire time will update
#     try:
#         print("Executing PowerShell command:")
#         print(powershell_script)
#         result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
#                                 capture_output=True, text=True,
#                                 check=True)
#         if expect_json:
#             try:
#                 output = json.loads(result.stdout)
#                 # If the output is a list of dictionaries, parse each dictionary separately
#                 if isinstance(output, list):
#                     parsed_output = [json.loads(item) for item in output]
#                     print("Parsed output:")
#                     print(parsed_output)
#                     return parsed_output
#                 else:
#                     print("Output:")
#                     print(output)
#                     return output
#             except json.JSONDecodeError:
#                 # Return None if JSON parsing fails
#                 print("JSONDecode ERROR", result.stdout)
#                 print("Error decoding JSON. Output:", result.stdout)
#                 return None
#         else:
#             print("Output:")
#             print(result.stdout)
#             return result.stdout
#     except subprocess.CalledProcessError as e:
#         print("PowerShell Error:", e)
#         print(e)
#         return None


def search_by_employee_id(id):
    employee_id = id
    fetch_user_by_employeeid = f'''
       $employee_id = "18139"
    $PaddedEmployeeID = $employee_id.PadLeft(8,'0')
    $users = Get-ADUser -Filter {{EmployeeID -eq $PaddedEmployeeID -or EmployeeID -eq $employee_id}} -Properties sAMAccountName,EmployeeId,DisplayName,Mail,Manager,Title,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
        Select-Object sAMAccountName, EmployeeId, DisplayName, Mail, Manager, Title,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
        ConvertTo-Json
    $users
    '''
    search_result = run_ps(fetch_user_by_employeeid)
    # user_data = json.loads(search_result)
    # Check if multiple users are returned
    # if isinstance(search_result, list):
    #     print("search result is:", search_result)
    # print(f"User data is: {user_data}")
    print("SEARCH RESULT IN SEARCH FUNC IS:", search_result)
    return search_result
    # else:
    #     # Return a list with a single user object
    #     print("search result is:", search_result)
    #     return [search_result]


def run_ps(powershell_script):
    result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
                            capture_output=True, text=True,
                            check=True)
    convert_to_json = json.loads(result.stdout)
    print(f"result is: {result.stdout}")
    print(f"CONVERTED TO JSON IS: {convert_to_json}")
    print(f"Type of convert_to_json is: {type(convert_to_json)}")

def search_by_samaccountname(username):
    samaccount_name = username
    fetch_user_by_samaccountname = f'''
    $samaccountname = "{samaccount_name}"
    Get-ADUser -Filter {{sAMAccountName -eq $samaccountname}} -Properties sAMAccountName,EmployeeId,DisplayName,Mail,Title,Manager,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
    Select-Object sAMAccountName, EmployeeId, DisplayName, Mail, Title, Manager, physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
    ConvertTo-Json
    '''
    search_result = run_powershell_command(fetch_user_by_samaccountname, expect_json=False)
    return search_result




search_by_employee_id(3232)
