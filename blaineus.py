'''
                                                        Developed by Darian Moore

This program was developed to make the tedious task of generating PDFs autonomous. The user will be given 3 options on how they want to generate their
PDFs; by model, by package, or by assembly (all PDFs will be an assembly because models and packages contain assemblies). Based off what option the user 
chooses they will be prompted to enter some information such as model name, package name, and or assembly name. The user also has the option to select 
multiple models, packages, and or assemblies (which can be from different jobs) to be added in the list assemblies to generate a PDF for. Once all information 
for the models, packages, and or assemblies has been entered by the user, selenium will be used to interact with the web page. Selenium will be used to interact
with the website by logging into the website, finding all assemblies needed, and generating a PDF file for each assembly needed. Each PDF generated will be 
re-located from the deafualt download path, the downloads folder, to a specified directory. While the PDFs are being generated the user will stay updated on how 
many assemblies there are in the list, what assembly they are currently on, what model or package they are on if this option is selected, and updates on what 
the program is doing in the web browser as well as the file system. Once the program has completely finished running it will let the user know how long it took 
to complete generating all PDFs and then terminate itself after 10 seconds.

       FUN FACT: program is named "blaineus" because I specifically made this for a guy named Blaine who was complaining about downloading PDFs lol
'''

import time
import os
import shutil
import sys
import requests
from cryptography.fernet import Fernet
from pathlib import Path
from colorama import init, Fore, Back, Style
import selenium
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
init(convert=True) #problem with colorama not working properly

def main():
    # Setting up API info for data requests in dataGet() function
    apiInfo = {'url':'https://api.gtpstratus.com/','headers':{'app-key': 'key'}}

    # Check the API health
    health = (dataGet(apiInfo, 'health')).get('status')
    if health == 'Healthy':
        pass
    else:
        print(f'\n{Fore.RED}Stratus is experiencing slow connection speeds{Style.RESET_ALL}')
        print('Closing Program in 10 seconds, try again')
        time.sleep(10)
        sys.exit()

    # Getting a dictionary of all of the active job numbers, ids, and names
    proj_info = dataGet(apiInfo, 'v2/project?where=status%20%3D%3D%201')
    active_jobs = {}
    active_ids_list = []
    active_nums_list = []
    for i in proj_info.get('data'):
        active_jobs[i.get('number')] = [{'projId': i.get('id')}, {'projName': i.get('name')}]
        active_ids_list.append(i.get('id'))
        active_nums_list.append(i.get('number'))

    # Asking the user if how they want to generate their pdfs; Either by model(s), package(s), or assembly/(ies)
    userOptions(True)
    user_choice = input('Choose your option: ')

    # Using a while loop to make sure the user gives a valid choice to generate the pdfs
    options = ['1', '2', '3']
    while user_choice not in options:
        os.system('cls')
        userOptions(False, user_choice=user_choice)
        user_choice = input('Choose your option: ')
    os.system('cls')

    # Try and except for a random connection error issue on Stratus' side
    # Depending on what option the user chose, a specific funtion will run 
    # and return different data formatted in the same dictionary template
    try:
        if user_choice == '1':
            essential_data = byModel(active_jobs, apiInfo) #Gets data by model
            os.system('cls')

        elif user_choice == '2':
            essential_data = byPkg(active_jobs, apiInfo) #Gets data by package
            os.system('cls')

        elif user_choice == '3':
            essential_data = byAsmbly(active_jobs, apiInfo) #Gets data by assembly
            os.system('cls')
    except:
        print(f'\n{Fore.RED}Connection Errror Occured{Style.RESET_ALL}')
        print('Closing Program in 10 seconds, try again')
        time.sleep(10)
        sys.exit()

    # Starting the timer
    start = time.time()

    # Starting the driver and calling the generatePDFs function
    driver = webdriver.Chrome(ChromeDriverManager().install())
    os.system('cls')
    generatePDFs(essential_data, driver, user_choice)
    os.system('cls')

    # Ending the timer
    end = time.time()

    # Calculating the run time of the actual pdf generating, and displaying the run time to the user
    timeTracking(start, end)
    time.sleep(5)

    # Terminating program
    print(f'\n{Fore.RED}--- Program terminating in 10sec ---')
    time.sleep(10)
    sys.exit()

def byModel(active_jobs, apiInfo):
    '''
    Prompts the user to enter multiple or a single model name(s) and then collects the
    necessary data related to those assemblies via API call
    Inputs:
    - active_jobs, a dictionary of all active jobs where the keys are the job numbers and the values
      are a list that contains dictionaries of the project ID and the project name (dict)
      Ex: [job_num1: [{'projId': projId}, {'projName': projName}], job_num2: [{'projId': projId}, {'projName': projName}], ...]
    - apiInfo, url and API key to access the API (dict)
    Returns: a dictionary of all necessary assembly data in order to generate the PDFs
    '''
    model_data_list, filtered_models_list, model_names = ([] for i in range(3))
    model_count = 0 #the amount of models the user entered
    again = 'y'
    while again == 'y' or again == 'Y':
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Models ---{Style.RESET_ALL}\n')
        modelName = ''
        filtered_models_list = ((str(model_names).replace("['", '')).replace("']", '')).replace("'", '')

        if model_count == 0:
            job_num = input("Enter a project number: ")
        else:
            print(f'{Fore.GREEN}Entered Models: {Fore.LIGHTBLACK_EX + filtered_models_list + Style.RESET_ALL}\n')
            print(f'Is this model still in project number {job_num} ?')
            choice = input("If yes then type 'y' and if no then type 'n': ")
            if choice == 'y' or choice == 'Y':
                pass
            else:
                job_num = input("Enter a project number: ")

        while job_num not in [*active_jobs.keys()]:
            os.system('cls')
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Models ---{Style.RESET_ALL}\n')
            print(f'{Back.RED}ERROR: {job_num} IS AN INVALID JOB NUMBER{Style.RESET_ALL}\n')
            job_num = input("Enter job number again: ")
        
        os.system('cls')
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Models ---{Style.RESET_ALL}\n')

        if model_count > 0:
            os.system('cls')
            print(f'{Fore.GREEN}Entered Packages: {Fore.LIGHTBLACK_EX + filtered_models_list + Style.RESET_ALL}\n')

        projId = active_jobs.get(job_num)[0].get('projId')
        projName = active_jobs.get(job_num)[1].get('projName')
        
        #Get user input for the project model name
        modelName = input(f'Enter a model name from {job_num}: ')

        #proj_models_info = dataGet(apiInfo, (f'v1/project/{active_jobs.get(job_num)[0].get(projId)}/models'))
        proj_models_info = dataGet(apiInfo, f"v1/project/{active_jobs.get(job_num)[0].get('projId')}/models")
        modelName_check = bool([d for d in proj_models_info.get('data') if d['name'] == modelName])

        while modelName_check == False:
            modelName = model_format(model_count, filtered_models_list, model_names, modelName, False)
            modelName_check = bool([d for d in proj_models_info.get('data') if d['name'] == modelName])

        model_format(model_count, filtered_models_list, model_names, modelName, True)
        
        for d in proj_models_info.get('data'):
            if d ['name'] == modelName:
                modelId = d.get('id')

        model_assemblies_info = dataGet(apiInfo, f'v1/model/{modelId}/assemblies')
        asmblyNames = [i.get('name') for i in model_assemblies_info.get('data')]
        asmblyIds = [i.get('id') for i in model_assemblies_info.get('data')]

        model_data = {'projId':projId, 'projNum':job_num, 'projName':projName, 'modelId':modelId, 'modelName':modelName, \
                    'asmblyIds':asmblyIds, 'asmblyNames':asmblyNames}

        model_names.append(modelName) #list of all valid model names entered by the user

        model_data_list.append(model_data) #list of all model(s) info in a dictionary and the PDF generate choice chosen by the user

        #Continuing or ending the primary while loop inside of the else statement
        model_format(model_count, filtered_models_list, model_names, modelName, True)
        print('Do you want to enter another model name?')
        again = input("If yes then type 'y' and if no then type 'n': ")
        if again == 'y' or again == 'Y':
            model_count += 1
        else:
            model_data_list.append('Model')
        
        os.system('cls')
        
    return model_data_list

def model_format(model_count, filtered_models_list, model_names, modelName, x):
    '''
    Getting the model name(s) from the user and displaying that info. Also has an error output for the user
    Inputs:
    - model_count, the number of models entered (int)
    - filtered_models_list, all of the models entered by the user in a list (string)
    - model_names, raw list of all model names entered by the user (list of strings)
    - modelName, the most recent model name entered by user (string)
    - x, True or False (boolean)
    Returns: if x is False then it returns a model name entered by the user
    '''
    if x:
        os.system('cls')
        if model_count > 0:
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Models ---{Style.RESET_ALL}\n')
            print(f'{Fore.GREEN}Entered Models: {Fore.LIGHTBLACK_EX + filtered_models_list + Style.RESET_ALL} {modelName}\n')
        elif model_count == 0:
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Models ---{Style.RESET_ALL}\n')
            print(f'{Fore.GREEN}Entered Models: {Style.RESET_ALL+modelName}\n')
    else:
        os.system('cls')
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Models ---{Style.RESET_ALL}\n')
        print(f'{Back.RED}ERROR: "{modelName}" IS AN INVALID MODEL NAME{Style.RESET_ALL}\n')
        model_name = input('Enter a valid model name: ')

        return model_name

def byPkg(active_jobs, apiInfo):
    '''
    Prompts the user to enter multiple or a single package name(s) and then collects the
    necessary data related to those assemblies via API call
    Inputs:
    - active_jobs, a dictionary of all active jobs where the keys are the job numbers and the values
      are a list that contains dictionaries of the project ID and the project name (dict)
      Ex: [job_num1: [{'projId': projId}, {'projName': projName}], job_num2: [{'projId': projId}, {'projName': projName}], ...]
    - apiInfo, url and API key to access the API (dict)
    Returns: a dictionary of all necessary assembly data in order to generate the PDFs
    '''
    pkg_data_list, filtered_packages_list, package_names = ([] for i in range(3))
    package_count = 0 #the amount of packages the user entered
    again = 'y'
    while again == 'y' or again == 'Y':
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Packages ---{Style.RESET_ALL}\n')
        filtered_packages_list = ((str(package_names).replace("['", '')).replace("']", '')).replace("'", '')

        if package_count > 0:
            print(f'{Fore.GREEN}Entered Packages: {Fore.LIGHTBLACK_EX + filtered_packages_list + Style.RESET_ALL}\n')

        #Get user input for the package number
        pkg_name = input("Enter a package name: ")
        job_num = pkg_name.split("-")[0]

        while job_num not in [*active_jobs.keys()]:
            pkg_name = package_format(package_count, filtered_packages_list, package_names, pkg_name, False)
            job_num = pkg_name.split("-")[0]

        projId = active_jobs.get(job_num)[0].get('projId')
        projName = active_jobs.get(job_num)[1].get('projName')
                
        all_pkg_info = dataGet(apiInfo, f'v1/package?where=name%20%3D%3D%20%22{pkg_name}%22')
        pkg_name_check = bool([d for d in all_pkg_info.get('data') if d ['name'] == pkg_name])

        while pkg_name_check == False:
            pkg_name = package_format(package_count, filtered_packages_list, package_names, pkg_name, False)
            all_pkg_info = dataGet(apiInfo, f'v1/package?where=name%20%3D%3D%20%22{pkg_name}%22')
            pkg_name_check = bool([d for d in all_pkg_info.get('data') if d ['name'] == pkg_name])

        package_format(package_count, filtered_packages_list, package_names, pkg_name, True)

        pkg_id = all_pkg_info.get('data')[0].get('id')
        modelId = all_pkg_info.get('data')[0].get('modelId')
        modelName = (dataGet(apiInfo, f'v1/model/{modelId}')).get('name')
        
        pkg_assemblies_info = dataGet(apiInfo, f'v2/package/{pkg_id}/assemblies')
        asmblyNames = [i.get('name') for i in pkg_assemblies_info.get('data')]
        asmblyIds = [i.get('id') for i in pkg_assemblies_info.get('data')]

        pkg_data = {'projId':projId, 'projNum':job_num, 'projName':projName, 'modelId':modelId, 'modelName':modelName, \
            'asmblyIds':asmblyIds, 'asmblyNames':asmblyNames}
        
        package_names.append(pkg_name) #list of all valid assembly names entered by the user

        pkg_data_list.append(pkg_data) #list of all package(s) info in a dictionary and the PDF generate choice chosen by the user
        
        #Continuing or ending the primary while loop inside of the else statement
        package_format(package_count, filtered_packages_list, package_names, pkg_name, True)
        print('Do you want to enter another package name?')
        again = input("If yes then type 'y' and if no then type 'n': ")
        if again == 'y' or again == 'Y':
            package_count += 1
        else:
            pkg_data_list.append(pkg_name)
            pkg_data_list.append('Package')
        
        os.system('cls')

    return pkg_data_list

def package_format(package_count, filtered_packages_list, package_names, pkg_name, x):
    '''
    Getting the package name(s) from the user and displaying that info. Also has an error output for the user
    Inputs:
    - package_count, the number of packages entered (int)
    - filtered_packages_list, all of the packages entered by the user in a list (string)
    - package_names, raw list of all package names entered by the user (list of strings)
    - pkg_name, the most recent package name entered by user (string)
    - x, True or False (boolean)
    Returns: if x is False then it returns a package name entered by the user
    '''
    if x:
        os.system('cls')
        if package_count > 0:
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Packages ---{Style.RESET_ALL}\n')
            print(f'{Fore.GREEN}Entered Packages: {Fore.LIGHTBLACK_EX + filtered_packages_list + Style.RESET_ALL} {pkg_name}\n')
        elif package_count == 0:
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Packages ---{Style.RESET_ALL}\n')
            print(f'{Fore.GREEN}Entered Packages: {Style.RESET_ALL+pkg_name}\n')
    else:
        os.system('cls')
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Packages ---{Style.RESET_ALL}\n')
        print(f'{Back.RED}ERROR: "{pkg_name}" IS AN INVALID PACKAGE NAME{Style.RESET_ALL}\n')
        pkg_name = input('Enter a valid package name: ')

        return pkg_name

def byAsmbly(active_jobs, apiInfo):
    '''
    Prompts the user to enter multiple or a single assembly name(s) and then collects the
    necessary data related to those assemblies via API call
    Inputs:
    - active_jobs, a dictionary of all active jobs where the keys are the job numbers and the values
      are a list that contains dictionaries of the project ID and the project name (dict)
      Ex: [job_num1: [{'projId': projId}, {'projName': projName}], job_num2: [{'projId': projId}, {'projName': projName}], ...]
    - apiInfo, url and API key to access the API (dict)
    Returns: a dictionary of all necessary assembly data in order to generate the PDFs
    '''
    assemblies_list, assembly_names, assembly_data_list= ([] for i in range(3))
    asmbly_count = 1 #the amount of assemblies the user entered
    again = 'y'
    while again == 'y' or again == 'Y':
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Assemblies ---{Style.RESET_ALL}\n')

        asmbly_name_check = 0 #assembly name is good if value is 0

        filtered_assemblies_list = ((str(assemblies_list).replace("['", ''))\
            .replace("']", '')).replace("'", '') #all of the assemblies entered by the user in string format

        if asmbly_count > 1:
            print(f'{Fore.GREEN}Entered Assemblies: {Fore.LIGHTBLACK_EX + filtered_assemblies_list + Style.RESET_ALL}\n')

        modelId = ''
        while modelId == '':
            if asmbly_name_check == 0:
                assemblyName = input('Enter a assembly name: ') #assembly name entered by the user
                job_num = assemblyName.split('-')[0]
                assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, True)
            else:
                assemblyName = assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, False)
                job_num = assemblyName.split('-')[0]
                assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, True)

            print('Processing Data...')

            while job_num not in [*active_jobs.keys()]:
                assemblyName = assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, False)
                job_num = assemblyName.split('-')[0]
                assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, True)

            projId = active_jobs.get(job_num)[0].get('projId')
            projName = active_jobs.get(job_num)[1].get('projName')
            proj_models = dataGet(apiInfo, f'v1/project/{projId}/models')
            modelIds_list = [n[1].get('id') for n in enumerate(proj_models.get('data'))]
            
            for i in enumerate(modelIds_list):
                models_info = dataGet(apiInfo, f'v1/model/{modelIds_list[int(i[0])]}/assemblies')
                for d in enumerate(models_info.get('data')):
                    if assemblyName == models_info.get('data')[int(d[0])].get('name'):
                        assemblyId = models_info.get('data')[int(d[0])].get('id')
                        modelId = models_info.get('data')[int(d[0])].get('modelId')
                        modelName = (dataGet(apiInfo, f'v1/model/{modelId}')).get('name')
                        break
                    else:
                        pass

            assembly_names.append(assemblyName) #list of all valid assembly names entered by the user
            asmbly_name_check += 1

        assemblies_list.append(assemblyName)
        
        assembly_data = {'projId':projId, 'projNum':job_num, 'projName':projName, 'modelId':modelId, 'modelName':modelName, \
                        'asmblyIds':[assemblyId], 'asmblyNames':[assemblyName]}

        assembly_data_list.append(assembly_data) #list of all assembly(ies) info in a dictionary and the PDF generate choice chosen by the user

        # Continuing or ending the primary while loop inside of the else statement
        assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, True)
        print('Do you want to enter another assembly name?')
        again = input("If yes then type 'y' and if no then type 'n': ")
        if again == 'y' or again == 'Y':
            asmbly_count += 1
            os.system('cls')
        else:
            assembly_data_list.append('Assembly')
    
    return assembly_data_list

def assembly_format(asmbly_count, filtered_assemblies_list, assemblyName, assembly_names, x):
    '''
    Getting the assembly name(s) from the user and displaying that info. Also has an error output for the user
    Inputs:
    - asmbly_count, the number of assemblies (int)
    - filtered_assemblies_list, all of the assemblies entered by the user in a list (string)
    - assemblyName, the most recent assembly name entered by user (string)
    - assembly_names, raw list of all assembly names entered by the user (list of strings)
    - x, True or False (boolean)
    Returns: if x is False then it returns a assembly name entered by the user
    '''
    if x:
        os.system('cls')
        if asmbly_count > 1:
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Assemblies ---{Style.RESET_ALL}\n')
            print(f'{Fore.GREEN}Entered Assemblies: {Fore.LIGHTBLACK_EX }{filtered_assemblies_list}{Style.RESET_ALL} {assemblyName}\n')
        elif asmbly_count == 1:
            print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Assemblies ---{Style.RESET_ALL}\n')
            print(f'{Fore.GREEN}Entered Assemblies: {Style.RESET_ALL}{assemblyName}\n')
    else:
        os.system('cls')
        print(assembly_names)
        time.sleep(10)
        print(f'{Fore.LIGHTBLACK_EX}--- Generating PDFs by Assemblies ---{Style.RESET_ALL}\n')
        print(f'{Back.RED}ERROR: "{assembly_names[(asmbly_count-1)]}" IS AN INVALID ASSEMBLY NAME{Style.RESET_ALL}\n')
        assemblyName = input('Enter a valid assembly name: ')

        return assemblyName

def userOptions(x, user_choice=''):
    '''
    Print out the PDF generate options that the user has
    Inputs:
    - x, True or False (boolean)
    - user_choice, optional parameter only used for invalid choices (string)
    '''
    if x:
        print(f'{Style.RESET_ALL}')
        print(f'{Back.LIGHTBLACK_EX + Fore.BLACK + Style.NORMAL}\t\t\tOptions\t\t\t\t{Style.RESET_ALL}')
        print(f'Option {Fore.LIGHTBLUE_EX}1{Style.RESET_ALL} Generate PDFs by {Fore.LIGHTBLUE_EX}Models{Style.RESET_ALL}')
        print(f'Option {Fore.GREEN}2{Style.RESET_ALL} Generate PDFs by {Fore.GREEN}Packages{Style.RESET_ALL}')
        print(f'Option {Fore.RED}3{Style.RESET_ALL} Generate PDFs by {Fore.RED}Assemblies{Style.RESET_ALL}')
        print('-'*37)
    else:
        print(f'{Back.RED}ERROR: {user_choice} IS NOT AN OPTION, SEE BELOW OPTIONS{Style.RESET_ALL}\n')
        print(f'{Back.LIGHTBLACK_EX + Fore.BLACK + Style.NORMAL}\t\t\tOptions\t\t\t\t{Style.RESET_ALL}')
        print(f'Option {Fore.LIGHTBLUE_EX}1{Style.RESET_ALL} Generate PDFs by {Fore.LIGHTBLUE_EX}Models{Style.RESET_ALL}')
        print(f'Option {Fore.GREEN}2{Style.RESET_ALL} Generate PDFs by {Fore.GREEN}Packages{Style.RESET_ALL}')
        print(f'Option {Fore.RED}3{Style.RESET_ALL} Generate PDFs by {Fore.RED}Assemblies{Style.RESET_ALL}')
        print('-'*37)

def timeTracking(start, end):
    '''
    Used to see how fast the program is running
    Inputs:
    - start, the program start time (float)
    - end, the program end time (float)
    Returns: the amount of time it took the program to finish
    '''
    run_time = int(end-start)
    if run_time >= 60 and run_time < 3600:
        minute = int(run_time/60)
        sec = int(run_time%60)
        print(f'Program finished in {Fore.GREEN}{minute}min {Style.RESET_ALL}and {Fore.GREEN}{sec}sec{Style.RESET_ALL}')
    elif run_time >= 3600:
        hour = int(run_time/3600)
        minute = int((run_time%3600)/60)
        sec = int((run_time%3600)%60)
        print(f'Program finished in {Fore.YELLOW}{hour}hr {Fore.GREEN}{minute}min and {Fore.GREEN}{sec}sec{Style.RESET_ALL}')
    else:
        print(f'Program finished in {Fore.GREEN}{run_time}sec')

def dataGet(apiInfo, requestCommand):
    '''
    Requesting info from Stratus API
    Inputs:
    - apiInfo, url and API key to access the API (dict)
    - requestCommand, a specific API command (string)
    Returns: json data from the API get request
    '''
    rawData = requests.get(apiInfo.get("url") + requestCommand, headers=apiInfo.get("headers"))
    if str(rawData) == "<Response [200]>":
        data = rawData.json()
    else:
        data = []
        print(f'{Fore.RED}ERROR: {rawData}{Style.RESET_ALL}\n')
        print("Check the url request url")
    return data

def login():
    '''
    Obtaining user's email and password to Stratus from an encrypted .txt file.
    (Alternatively I could've used a system variable to store this info 
    but I wanted to mess around with encryption)
    Returns: decrypted user email and password
    '''
    #Reading key from key.key file
    with open('key.key', 'rb') as k:
        key = k.read()

    #Reading the encrypted .txt file
    with open ('settings.txt.encrypted', 'rb') as s:
        data = s.read()

    #Decrypt the .txt file
    fernet = Fernet(key)
    encrypted = fernet.decrypt(data)
    
    #Convert from bytes to string and create a list from the string
    decrypt = encrypted.decode("utf-8")
    dlist = list(decrypt.split(" "))
    email = dlist[0]
    password = dlist[1]

    return email, password

def createFolder(directory, count = 0):
    '''
    Creates folder if it does not already exist
    Inputs:
    - directory, new directory name (string)
    - count, param used for OSError purposes (int)
    '''
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        if count < 3:
            print(f'{Fore.RED}ERROR: CREATING FOLDER {directory} {count+1}/3 ATTEMPTS{Style.RESET_ALL}\n')
            createFolder(directory, count+1)

def checkElement(driver, xpath):
    '''
    Actively checking if an element exists on a webpage
    Inputs:
    - driver, webdriver to launch and interact with a webpage (driver object)
    - xpath, the target element's xpath (string)
    '''
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

def click(driver, sec, element, key, x, value = "", count = 0):
    '''
    Uses try and except to catch any errors when trying to click elements on the webpage
    Inputs:
    - driver, webdriver to launch and interact with a webpage (driver object)
    - sec, amount of seconds to wait before throwing an error (int)
    - element, the type attribute we are trying to click (attribute)
    - key, the value of an element's attribute like name, class, xpath, etc. (string)
    - x, True or False (boolean)
    - value, value to submit into a web element like a search box or email field (string)
    - count, used for error control (int)
    '''
    try:
        if x:
            WebDriverWait(driver,sec).until(EC.element_to_be_clickable((element, key))).send_keys(value)
        else:
            WebDriverWait(driver,sec).until(EC.element_to_be_clickable((element,key))).click()
    except (selenium.common.exceptions.TimeoutException, selenium.common.exceptions.ElementClickInterceptedException, selenium.common.exceptions.StaleElementReferenceException):
        if count < 3:
            print(f'{Fore.RED}ERROR: UNABLE TO CLICK ELEMENT {count+1}/3 ATTEMPTS{Style.RESET_ALL}\n')
            click(driver, sec, element, key, x, value, count+1)

def wait(driver, sec, element, key, count = 0):
    '''
    Uses try and except to catch any errors when waiting for the progress bar element to become hidden on the webpage
    Inputs:
    - driver, webdriver to launch and interact with a webpage (driver object)
    - sec, amount of seconds to wait before throwing an error (int)
    - element, the type attribute we are trying to click (attribute)
    - key, the value of an element's attribute like name, class, xpath, etc. (string)
    - count, used for error control (int)
    '''
    try:    
        WebDriverWait(driver,sec).until(EC.invisibility_of_element_located((element,key)))
    except (selenium.common.exceptions.TimeoutException, selenium.common.exceptions.ElementClickInterceptedException):
        if count < 3:
            print(f'{Fore.RED}ERROR: PROGRESS BAR NOT FINISHED {count+1}/3 ATTEMPTS{Style.RESET_ALL}\n')
            driver.refresh()
            time.sleep(2)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-tabs"]/li[3]/a')))
            time.sleep(2)
            wait(driver, sec, element, key, count+1)

def waitForStatusBar(counter, driver, x):
    '''
    Wait for the progress bar to finish loading
    Inputs:
    - counter, the current index of assembly (int)
    - driver, webdriver to launch and interact with a webpage (driver object)
    - x, True or False (boolean)
    '''
    if x:
        print('Waiting for progress bar...')
        wait(driver, 15, By.XPATH, "//*[@id='model-viewer']/div[1]/div[3]")
        print('Progress bar complete!\n')
    else:
        time.sleep(2)
        print('Waiting for progress bar...')
        wait(driver, 15, By.XPATH, "//*[@id='model-viewer']/div[1]/div[3]")
        print('Progress bar complete!\n')

def header(option, index, count, counter, asmbly_count):
    '''
    Displays model/package and assembly data at the top of the CMD. Keeps user informed on how many models/packages and assemblies there are total,
    as well as how many are left that need a PDF generated for them. For example, "Model 1/2 - Assembly 4/9" would be displayed for the user.
    Inputs:
    - option, what option the user chose by Model, by Package, or by Assembly (string)
    - index, the current index of the model or package (int)
    - count, the total number of models or packages (int)
    - counter, the current index of the assembly (int)
    - asmbly_count, the total number of assemblies (int)
    '''
    os.system('cls')
    if option == 'Model' or option == 'Package':
        print(f'{Fore.YELLOW}{option} {index+1}/{count}{Style.RESET_ALL} - {Fore.GREEN}Assembly {counter+1}/{asmbly_count}{Style.RESET_ALL}\n')
    else:
        print(f'{Fore.YELLOW}{option} {index+1}/{count}{Style.RESET_ALL}\n')

def generatePDFs(essential_data, driver, user_choice):
    '''
    With all of the assembly data gathered we can go generate a PDF file for each assembly and
    save that assembly to a specific path on our file system.
    Inputs:
    - essential_data, all needed data for all assemblies entered by the user (dict)
    - driver, webdriver to launch and interact with a webpage (driver object)
    - user_choice, optional parameter only used for invalid choices (string)
    '''
    os.system('cls')
    print(f'{Fore.LIGHTCYAN_EX}--- Loading Stratus ---{Fore.BLACK}')

    # File paths to be used for file moving
    home = str(Path.home()).replace("\\", "/")
    downloads = f'{home}/Downloads'

    # Determining how many different models, packages, or assembleis there are
    count = len(essential_data)
    option = essential_data[count-1]
    if option == 'Package':
        count = count-2
    else:
        count = count-1

    # Getting the number of assemblies
    index = 0
    asmbly_count = len(essential_data[index].get('asmblyIds'))

    # Passing variables that contain login info from the login() funtion
    email, password = login()

    # Report to run
    report = ('03 - Pipe Assemblies BOM (CSV)')

    # Using a for loop to go through each model, package, or assembly based off the count variable
    for n in range(count):
        # Getting the project and model id form the essential_data variable
        projId = essential_data[index].get('projId')
        modelId = essential_data[index].get('modelId')

        # Buffer time for website to load
        if n > 0:
            time.sleep(1)
        elif n == count:
            pass

        # Request the Stratus Assembly's Dashboard web page
        driver.get(f'https://www.gtpstratus.com/assemblies?projectId={projId}&modelId={modelId}#tab_dashboard')

        # Login to Stratus if it's the first time visting the webpage
        if n == 0: 
            click(driver, 10, By.NAME, 'email', True, email)
            click(driver, 10, By.NAME, 'password', True, password)
            click(driver, 10, By.XPATH, '//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/button/span', False)

        # Using while loop to keep generating pdfs until there are no more assemblies
        for counter in range(asmbly_count):
            assemblyId = essential_data[index].get('asmblyIds')[counter]
            assemblyName = essential_data[index].get('asmblyNames')[counter]

            # Opening a new tab to the assembly viewer based off the list of assemblies
            driver.execute_script("window.open('');")
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[1])
            driver.get(f'https://www.gtpstratus.com/assemblies?projectId={projId}&modelId={modelId}&assemblyId={assemblyId}#tab_viewer')

            # Indicates what assembly you are on out of the total number of assemblies
            # As well as what model or package you are in if there is any
            header(option, index, count, counter, asmbly_count)

            # Checking for a random sign in error
            if counter == 0:
                if checkElement(driver, '//*[@id="help-info-panel"]/div/div[2]/div/a') == True:
                    click(driver, 10, By.XPATH, '//*[@id="help-info-panel"]/div/div[2]/div/a', False)
                    click(driver, 10, By.XPATH, '//*[@id="signon-autodesk"]', False)
                    click(driver, 10, By.XPATH, '//*[@id="userName"]', True, email)
                    click(driver, 10, By.XPATH, '//*[@id="verify_user_btn"]', False)
                    click(driver, 10, By.XPATH, '//*[@id="password"]', True, password)
                    click(driver, 10, By.XPATH, '//*[@id="btnSubmit"]', False)
                    click(driver, 10, By.XPATH, '//*[@id="allow_btn"]', False)
                    driver.get(f'https://www.gtpstratus.com/assemblies?projectId={projId}&modelId={modelId}&assemblyId={assemblyId}#tab_viewer')

            # Clicking the "Cache Later" button
            if counter == 0:
                click(driver, 10, By.XPATH, '//*[@id="eFooter"]/button[2]', False)

            # Waiting for the viewer to fully load
            if counter == 0:
                waitForStatusBar(counter, driver, True)
            else:
                waitForStatusBar(counter, driver, False)

            # Setting the report
            click(driver, 10, By.XPATH, '//*[@id="parts-list"]/div[1]/div[1]/div/span/span[1]/span', False)
            click(driver, 10, By.XPATH, '/html/body/span/span/span[1]/input', True, report)
            click(driver, 10, By.XPATH, '/html/body/span/span/span[1]/input', True, Keys.ENTER)
            header(option, index, count, counter, asmbly_count)
            print(f'Correct report chosen ({report})')

            # Clicking "STRATUS Sheet" tab
            time.sleep(2)
            click(driver, 10, By.XPATH, '//*[@id="viewer-sub-tabs"]/li[2]/a', False)

            # Waiting for the viewer and sheet to load
            header(option, index, count, counter, asmbly_count)
            waitForStatusBar(counter, driver, False)

            # Clicking "Generate PDF" button
            time.sleep(2)
            click(driver, 5, By.XPATH, '//*[@id="report-workspace"]/div/div[1]/button[1]/span', False)

            # Giving the assembly some time to process
            time.sleep(5)

            # Letting the user know the pdf is being generated
            header(option, index, count, counter, asmbly_count)
            print('Generating PDF...')

            # Waiting for the pdf to generate and then close the current tab
            source = (f'{downloads}/{assemblyName}.pdf')
            while os.path.exists(source) == False:
                time.sleep(1)
            
            # Letting the user know the pdf has been created
            header(option, index, count, counter, asmbly_count)
            print(f'{assemblyName}.pdf created\n')
            time.sleep(1)
            os.system('cls')
            header(option, index, count, counter, asmbly_count)
            driver.close()

            # Switch back to the first tab
            driver.switch_to.window(driver.window_handles[0])

            # Create new project folder based off of where you want the pdfs to be saved to
            if option == 'Model':
                byModel = f"{downloads}/StratusPDFs/{essential_data[index].get('projName')}/ByModel/{essential_data[index].get('modelName')}"
                createFolder(byModel)
                shutil.copy(source, byModel)
                time.sleep(1)
                print(f'{assemblyName}.pdf relocated to ByModel folder')
                time.sleep(2)
                os.remove(source)
            elif option == 'Package':
                time.sleep(5)
                byPackage = f"{downloads}/StratusPDFs/{essential_data[index].get('projName')}/ByPackage/{essential_data[count]}"
                createFolder(byPackage)
                shutil.copy(source, byPackage)
                time.sleep(1)
                print(f'{assemblyName}.pdf relocated to ByPackage folder')
                time.sleep(2)
                os.remove(source)
            else:
                byAssembly = f"{downloads}/StratusPDFs/{essential_data[index].get('projName')}/ByAssembly"
                createFolder(byAssembly)
                shutil.copy(source, byAssembly)
                time.sleep(1)
                print(f'{assemblyName}.pdf relocated to ByAssembly folder')
                time.sleep(2)
                os.remove(source)

        index += 1

if __name__ == '__main__':
    main()