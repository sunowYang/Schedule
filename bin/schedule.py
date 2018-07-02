#! coding=utf8

import os
import sys
import time
import win32com.client
import win32api


def get_sys_time():
    return int(time.mktime(time.localtime()))


def modify_time(_time):
    try:
        set_time = time.localtime(_time)
        win32api.SetSystemTime(set_time)
    except Exception, e:
        raise Exception('Modify time failed:%s' % e)


def check_process_exist(name, timeout=5):
    while timeout > 0:
        try:
            api = win32com.client.GetObject('winmgmts:')
            process_code = api.ExecQuery('select * from Win32_Process where Name="%s"' % name)
        except Exception, e:
            raise Exception('Get process name failed:%s' % e)
        if len(process_code) > 0:
            return True
        else:
            timeout -= 1
            time.sleep(1)
    return False


def write_result(path, result):
    if not os.path.isfile(path):
        raise IOError('The result path is not a file:%s' % path)
    try:
        with open(path, 'a') as file_open:
            file_open.write(result)
            file_open.close()
    except Exception, e:
        raise IOError('Write result failed:%s' % e)


class Daily:
    def __init__(self, data):
        if not isinstance(data, dict):
            raise IOError('Daily data is not a dic')
        self.data = data
        self.daily_mode = self.data['Schedule_Daily_Backup_Method'] \
            if 'Schedule_Daily_Backup_Method' in self.data else '2'
        self.execute_time = []


    def execute_mode(self):
        if self.daily_mode == '1':
            return  self.at_time_mode()
        elif self.daily_mode == '2':
            self.interval_mode()
        else:
            raise IOError('Daily schedule has an error mode:%s' % self.daily_mode)


    def at_time_mode(self):
        if 'Schedule_Daily_addtime' not in self.data.items():
            raise IOError('Daily schedule at_time_mode does not have parameter "Schedule_Daily_addtime"')
        if ',' in self.data['Schedule_Daily_addtime']:
            at_time = self.data['Schedule_Daily_addtime'].split(',')
        else:
            at_time = [self.data['Schedule_Daily_addtime']]
        return at_time


    def interval_mode(self):




class Schedule:
    def __init__(self, data, result_path, execute_time):
        self.name = 'TbService.exe'
        self.data = data
        self.result_path = result_path
        self.execute_time = execute_time

        assert isinstance(self.data, dict)
        if 'ScheduleType' not in self.data.items():
            raise IOError('No found "ScheduleType" in data')
        self.schedule_type = self.data['ScheduleType']

    def run_schedule(self):
        if self.schedule_type == 'Onetime':
            result = self.one_time()
        elif self.schedule_type == 'Daily':
            result = self.daily()
        elif self.schedule_type == 'Weekly':
            result = self.weekly()
        elif self.schedule_type == 'Monthly':
            result = self.monthly()
        else:
            raise IOError('Schedule type is error:%s' % self.schedule_type)
        write_result(self.result_path, result)

    def one_time(self):
        set_time = time.strptime(self.get_one_time_data(), "%Y/%m/%d %H:%M")
        modify_time(int(time.mktime(set_time)))
        return True if check_process_exist(self.name) else False

    def daily(self):
        return

    def weekly(self):
        return

    def monthly(self):
        return

    def get_one_time_data(self):
        assert 'schedule_onetime_time' in self.data.items
        assert 'schedule_onetime_date' in self.data.items
        return self.data['schedule_onetime_date'] + ' ' + self.data['schedule_onetime_time']





if __name__ == '__main__':
    print(time.localtime())
