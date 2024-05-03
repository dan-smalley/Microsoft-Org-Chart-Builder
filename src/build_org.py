from omegaconf import OmegaConf
import requests
import pyinputplus as pyip
import os
from socket import gethostname
import logging
import json

# General Vars
scriptHost = gethostname()
scriptName = os.path.basename(__file__)

# Load the config and key files.
config = OmegaConf.create()
keys = OmegaConf.create()

config_files = [
    f'{os.path.split(os.path.dirname(__file__))[0]}/default_cfg.yaml',
    f'{os.path.split(os.path.dirname(__file__))[0]}/user_cfg.yaml'
]
key_files = [
    f'{os.path.split(os.path.dirname(__file__))[0]}/default_key.yaml',
    f'{os.path.split(os.path.dirname(__file__))[0]}/user_key.yaml'
]

try:
    for c in config_files:
        if os.path.exists(c):
            config = OmegaConf.merge(config, OmegaConf.load(c))
    if len(config) == 0:
        raise FileNotFoundError('No config files found in script or parent directory.  See script docstring for details.')
    for k in key_files:
        if os.path.exists(k):
            keys = OmegaConf.merge(keys, OmegaConf.load(k))
    if len(keys) == 0:
        raise FileNotFoundError('No key files found in script or parent directory.  See script docstring for details.')

except Exception as exc:
    logging.critical(f'Error loading keys: {exc}')
    exit(1)

# Load command line arguments (config only, not keys)
# Done after loading config files so command line arguments can overwrite those values
try:
    cli_args = OmegaConf.from_cli()
    config = OmegaConf.merge(config, cli_args)

except Exception as exc:
    logging.error(f'Error loading command line arguments: {exc}')

# Format logging
logging.basicConfig(format=config['logFormat'],
                    level=config['logLevel'],
                    datefmt="%H:%M:%S")

# Check if we have a token, if not ask user for one.
if keys['token'] is None:
    keys['token'] = pyip.inputStr('Enter your API token:\n')
else:
    logging.info('Using API token from key file.')


headers = {'Authorization': f'Bearer{keys["token"]}',
           'Content-Type': 'application/json'
           }

def validate_token():
    request = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
    while request.ok is False:
        keys['token'] = pyip.inputStr('Invalid token, please enter your API token:\n')
        request = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
    logging.info(f'Validated token for {request.json()["mail"]}')


def get_top_node(big_boss):
    payload = json.dumps({"requests":[{"entityTypes":["person"],"query":{"queryString":big_boss},"from":0,"size":2}]})
    request = requests.post('https://graph.microsoft.com/v1.0/search/microsoft.graph.query', payload, headers=headers)

    if request.ok is False:
        raise requests.exceptions.RequestException

    if request.json()['value'][0]['hitsContainers'][0]['total'] > 1:
        logging.error(f'Too many results returned ({request.json()["value"][0]["hitsContainers"][0]["total"]}), '
                      f'you will need to update your search.')
        exit(1)
    elif request.json()['value'][0]['hitsContainers'][0]['total'] == 0:
        logging.error('No results returned from Microsoft Graph')
        exit(1)

    return request.json()['value'][0]['hitsContainers'][0]['hits'][0]['resource']


def exclude_account(account):
    for exclusion in config['nameExclusions']:
        if exclusion in str(account['displayName']).lower():
            logging.info(f'Skipping {account["displayName"]}, {account["userPrincipalName"]}, suspected service or admin account.')
            return True
    for exclusion in config['mailExclusions']:
        if exclusion in str(account['userPrincipalName']).lower():
            logging.info(f'Skipping {account["displayName"]}, {account["userPrincipalName"]}, suspected service or admin account.')
            return True

def get_department(_id):
    try:
        request = requests.get(f'https://graph.microsoft.com/v1.0/users/{_id}/department', headers=headers)
        return request.json()['value']
    except Exception as exc:
        return None


def get_reports(_id, manager, department, depth):
    global managers
    request = requests.get(f'https://graph.microsoft.com/v1.0/users/{_id}/directReports', headers=headers)
    reports = request.json()['value']
    if len(reports) > 0:
        department = get_department(_id)
        managers += 1
        depth += 1
        logging.info(f'Processing reports for {manager}, {department}  ---  {managers} managers / {len(org)} employees')
        for report in reports:
            get_reports(report['id'], report['displayName'], department, depth)
            if not exclude_account(report):
                org.append({
                    'name': report['displayName'],
                    'email': report['userPrincipalName'],
                    'title': report['jobTitle'],
                    'manager': manager,
                    'department': department,
                    'depth': depth
    })


def sanitize_results():
    c_org = []
    for p in org:
        name, manager, email, title = None, None, None, None
        if p['name'] is not None:
            if ',' in p['name']:
                name = f'{p["name"].split(",")[0]} {p["name"].split(",")[1].strip()}'
            else:
                name = p["name"]
            name = name.replace(',', '').replace('/', '').replace('\'', '').replace('"', '').replace('\\', '').strip()
        if p['manager'] is not None:
            if ',' in p['manager']:
                manager = f'{p["manager"].split(",")[0]} {p["manager"].split(",")[1].strip()}'
            else:
                manager = p['manager']
            manager = manager.replace(',', '').replace('/', '').replace('\'', '').replace('"', '').replace('\\', '').strip()
        if p['title'] is not None:
            title = p['title'].replace(',', '').replace('/', '').replace('\'', '').replace('"', '').replace('\\', '').strip()
        if p['email'] is not None:
            email = p['email'].replace('"', '').strip()

        c_person = {
            'name': name,
            'email': email,
            'title': title,
            'manager': manager,
        }

        c_org.append(c_person)

    return c_org


def write_file(f_name, vendors, org_list):
    with open('drawio_csv_template.txt', 'r') as template:
        with open(f_name, 'w') as outfile:
            for line in template.readlines():
                outfile.writelines(line)
            for p in org_list:
                if p['title'] is None:
                    outfile.writelines(f"\n{p['name']},{p['manager']},{p['email']},{p['title']}")
                elif 'vendor' not in p['title'].lower():
                    outfile.writelines(f"\n{p['name']},{p['manager']},{p['email']},{p['title']}")
                elif 'vendor' in p['title'].lower() and vendors is True:
                    outfile.writelines(f"\n{p['name']},{p['manager']},{p['email']},{p['title']}")
    logging.info(f'Org file written to {f_name}, see instructions in file for how to upload to draw.io')


# Main body
global org
global managers

org = []
managers = 0
validate_token()

big_boss = pyip.inputStr('Enter the email address of top level person to begin the tree at (manager, director, etc): ')

boss_node = get_top_node(big_boss)
boss_id = boss_node['id'].split('@')[0]

org.append({
                    'name': boss_node['displayName'],
                    'email': boss_node['userPrincipalName'],
                    'title': boss_node['jobTitle'],
                    'manager': None,
                    'department': boss_node['department'],
                    'depth': 0
    })

get_reports(boss_id, boss_node['displayName'], boss_node['department'], 0)

logging.info(f'Done querying, sanitizing output  ---  {managers} managers / {len(org)} employees')

clean_org = sanitize_results()

f_name = pyip.inputStr('Enter the name of the file you want to save the results: ')
vendors = pyip.inputYesNo('Would you like to include vendors in the org chart (y/n): ')
write_file(f_name, vendors, clean_org)

logging.info(f'Script complete!')
