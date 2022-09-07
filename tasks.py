import sys, time

import schedule

sys.path.append(r'G:\My Drive\Ninjavan Project\Send_report_Bot_id_52717207')
import gui as g

list_code = ['vol', 'lm20', 'lm12', 'lmfsr', 'lmc', 'gtsc', 'fm']


def run(lst):
    _list_code = [c for c in lst if c in list_code]
    print(f'list_code = {_list_code}')
    if 's' not in lst:
        g.call_func(_list_code, command=lst)
    else:
        for _ in lst:
            if ":" in _:
                run_time = _

        schedule.every().day.at(f"{run_time}").do(g.call_func, _list_code, command=lst)

        while True:
            schedule.run_pending()
            time.sleep(1)

if __name__ == '__main__':

    run(sys.argv)
