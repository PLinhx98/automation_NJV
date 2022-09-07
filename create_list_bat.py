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
# NOTE: vollm dc tich hop trong vol va chi gui luc 20h


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