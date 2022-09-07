# # create a lot of bat file
# list_of_command = {'t': 'test', 'o': 'once', 's': 'schedule', 'q': 'quick'}
# list_code = ['vol', 'lm20', 'lm12', 'lmc', 'gtsc']
# schedule_dict = {'vol': {'normal': [f'{str(h).zfill(2)}:00' for h in range(10, 20)],
# 						 '9': ['09:00'],
# 						 '20': ['20:00'],
# 						 'cn': ['10:00', '16:00']},
# 				 'gtsc': {'normal': ['21:00'],
# 				 		  'gtsc10': ['10:00']},
# 				 'lm20': ['16:40', '19:00'],
# 				 'lm12': ['10:00', '12:00'],
# 				 'lmc': ['21:15']}

# all_time_run = {}
# for i in list_code:
# 	if not isinstance(schedule_dict[i], dict):
# 		for j in schedule_dict[i]:
# 			all_time_run[j] = []
# 	else:
# 		for subtype in schedule_dict[i].keys():
# 			for j in schedule_dict[i][subtype]:
# 				all_time_run[j] = []

# for i in list_code:
# 	if not isinstance(schedule_dict[i], dict):
# 		for j in schedule_dict[i]:
# 			all_time_run[j].append(i)
# 	else:
# 		for subtype in schedule_dict[i].keys():
# 			for j in schedule_dict[i][subtype]:
# 				if subtype == 'normal':
# 					all_time_run[j].append(i)
# 				else:
# 					all_time_run[j].append(f'{i}_{subtype}')

# print(all_time_run)

# # Clear old bat file
# import os
# lst = [f for f in os.listdir() if '.bat' in f and 'send' in f]
# for _ in lst:
# 	os.remove(_)

# # Create not schedule code
# for code in list_code:
# 	for run_type in list_of_command.keys():
# 		if run_type != 's':
# 			if not isinstance(schedule_dict[code], dict):
# 				with open(f'send_{code}_{list_of_command[run_type]}.bat', 'w') as f:
# 					f.write(f'TITLE send_{code}_{run_type}\npy tasks.py {code} {run_type}\nPAUSE')
# 			else:
# 				for subtype in schedule_dict[code].keys():
# 					with open(f'send_{code}_{subtype}_{list_of_command[run_type]}.bat', 'w') as f:
# 						f.write(f'TITLE send_{code}_{subtype}_{run_type}\npy tasks.py {code} {subtype} {run_type}\nPAUSE')

# # Create single schedule code
# for code in schedule_dict.keys():
# 	if not isinstance(schedule_dict[code], dict):
# 		for time in schedule_dict[code]:
# 			with open(f'send_{code}_at_{time.replace(":","_")}.bat', 'w') as f:
# 				f.write(f'TITLE send_{code}_{time}\npy tasks.py {code} s {time}\nPAUSE')
# 	else:
# 		for subtype in schedule_dict[code].keys():
# 			for time in schedule_dict[code][subtype]:
# 				if subtype == 'normal':
# 					subtype = ''
# 				with open(f'send_{code}_{subtype}_at_{time.replace(":","_")}.bat', 'w') as f:
# 					f.write(f'TITLE send_{code}_{subtype}_{time}\npy tasks.py {code} {subtype} s {time}\nPAUSE')


# # Create multi schedule code
# for hour in all_time_run.keys():
# 	if len(all_time_run[hour]) == 1:
# 		continue
# 	if len(all_time_run[hour]) == 2:
# 		name1 = '_'.join(all_time_run[hour])
# 		name2 = ' '.join(all_time_run[hour])
# 		with open(f'send_{name1}_at_{hour.replace(":","_")}.bat', 'w') as f:
# 			f.write(f'TITLE send_{name1}_{hour}\npy tasks.py {name2} s {hour}\nPAUSE')

# # Manually
# with open(f'send_vol_lm12_at_10_00.bat', 'w') as f:
# 	f.write(f'TITLE send_vol_lm12_10_00\npy tasks.py vol lm12 s 10:00\nPAUSE')
# with open(f'send_vol_gtsc_at_10_00.bat', 'w') as f:
# 	f.write(f'TITLE send_vol_gtsc_10_00\npy tasks.py vol gtsc gtsc10 s 10:00\nPAUSE')
# with open(f'send_lm12_gtsc_at_10_00.bat', 'w') as f:
# 	f.write(f'TITLE send_lm12_gtsc_10_00\npy tasks.py lm12 gtsc gtsc10 s 10:00\nPAUSE')
# with open(f'send_vol_lm12_gtsc_at_10_00.bat', 'w') as f:
# 	f.write(f'TITLE send_vol_lm12_gtsc_10_00\npy tasks.py vol lm12 gtsc gtsc10 s 10:00\nPAUSE')

list_hour = []

for n in range(0, 24):
	for n1 in range(0, 60):
		list_hour.append(f'{str(n).zfill(2)}:{str(n1).zfill(2)}')

list_of_command = {'t': 'test', 'o': 'once', 'q': 'quick'}    # , 's': 'schedule'

vol = [f'{str(x).zfill(2)}:00' for x in range(8,21)]
lm20 = ['19:00']
lm12 = ['13:00']
lmfsr = ['16:30']
lmc = ['21:00']
fm = ['16:30']
# A BIG !@# NOTE: vollm dc tich hop trong vol va chi gui luc 20h


list_code = ['vol', 'lm20', 'lm12', 'lmfsr', 'lmc', 'fm']
dict_code_schedule = {'vol': vol, 'lm20': lm20, 'lm12': lm12, 'lmfsr': lmfsr, 'lmc': lmc, 'fm': fm}  # , 'gtsc': gtsc


for time_run in list_hour:
	any_code_match = 0
	file_name = r'refresh_'
	call_dataset = r'py tasks.py '
	for code in dict_code_schedule:
		for code_hour in dict_code_schedule[code]:
			if code_hour == time_run:
				any_code_match = any_code_match + 1
				file_name = file_name + code + '_'
				call_dataset = call_dataset + code + r' '

	if any_code_match >= 1:
		if time_run == '20:00':
			call_dataset = call_dataset + '20 '
		file_name = file_name + time_run[0:2] + '_' + time_run[3:]
		call_dataset = call_dataset + 's '+ time_run
		print(file_name)
		print(call_dataset)
		with open(f'{file_name}.bat', 'w') as f:
			f.write(f'TITLE {file_name}\n{call_dataset}\nPAUSE')


for code in list_code:
	for type_run in list_of_command:
		file_name = r'refresh_' + code + '_' + list_of_command[type_run]
		call_dataset = r'py tasks.py ' + code + ' ' + type_run
		print(file_name)
		print(call_dataset)
		with open(f'{file_name}.bat', 'w') as f:
			f.write(f'TITLE {file_name}\n{call_dataset}\nPAUSE')