import os
import subprocess
import sys

path = r'G:\My Drive\Ninjavan Project\Send_report_Bot_id_52717207'

list_run = []

all_files = os.listdir(path)
items_not_run = ['Reco.bat', 'start_everything.bat']
bat_files = [bat for bat in all_files if bat.endswith('.bat')]



for items in bat_files:
    if 'test' in items or 'once' in items or 'quick' in items:
        items_not_run.append(items)

# remove files(not run) from run_list
for items in items_not_run:
    bat_files.remove(items)


for n in bat_files:        
    subprocess.Popen(f'{path}/{n}', creationflags=subprocess.CREATE_NEW_CONSOLE)  # , cwd=f'{path}/'
