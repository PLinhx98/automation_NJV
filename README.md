# automation_NJV
how files works:
- data_collectors.py: run data from Redash, export data if necessary.
- gui.py: mostly tasks require processing via GUI - general code that will run fuction by name.
- tasks.py: receive command from bat files and call to gui.py.
- bat files: created by create_list_bat.py, call to tasks.py with commands for different purposes such as schedule code, test run...
- create_list_bat.py: create multiple bat files based on needs.
- start_everything.py: run all bat files with schedule command only.

An example of a typical flow:

TITLE refresh_vol_09_00
py tasks.py vol s 09:00
PAUSE

Those above command is an example of a bat file. This file will call to tasks.py with command "vol", "s", "09:00". In tasks.py, there are ways to identify those command like the list of code running here is only vol; type: s (schedule); 09:00: schedule time. Call_func function on gui.py gets activated and it calls to send_vol function (since code name passed in was "vol"). It refresh data from Redash via data_collectors.py, refresh Power BI (Power BI load data using api link from Redash), capture screen of reports and send data to google chat, send detail data via webhook.

TITLE refresh_lm20_test
py tasks.py lm20 t
PAUSE

same as above, but this time, commands are "lm20" and "t" which means this is only a test run and reports will be sent to test channel and run only one time. Function gets called here is send_lm20.


Report names, drive links, api_keys... are not uploaded here for security reason. Thanks!
