#! coding=utf8

import os
import sys
import time
import win32com.client
import win32api
import datetime
import calendar as cal

def get_sys_time():
    return int(time.mktime(time.localtime()))


def get_date_and_time():
    struct_time = time.localtime()
    _time = struct_time.tm_hour * 3600 + struct_time.tm_min * 60 + struct_time.tm_sec
    _date = get_sys_time() - _time
    return _date, _time


def time_format(_time):
    if isinstance(_time, int):
        return time.strftime("%Y/%m/%d %H:%M:%S", time.localtime(_time))
    elif isinstance(_time, str):
        if ':' not in _time:
            raise IOError('format: time does not have key ":"')
        return int(_time.split(':')[0]) * 3600 + int(_time.split(':')[1]) * 60
    else:
        return False


def modify_time(_time):
    try:
        set_time = time.localtime(_time)
        win32api.SetSystemTime(set_time)
    except Exception, e:
        raise Exception('Modify time failed:%s' % e)


def check_process_exist(name='TbService.exe', timeout=5):
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


def get_next_month():
    print(get_sys_time())
    now_year, now_month = time.localtime().tm_year, time.localtime().tm_mon
    if now_month < 12:
        now_month += 1
    else:
        now_month = 1
        now_year += 1
    return str(now_year), str(now_month)


def get_month_last_day(year, month):
    return cal.monthrange(year, month)[1]


def get_week_of_month(year, month, day):
    end = int(datetime.datetime(year, month, day).strftime("%W"))
    begin = int(datetime.datetime(year, month, 1).strftime("%W"))
    which_week_of_system = end - begin + 1
    return which_week_of_system




def get_month_last_week(year, month):
    last_day = get_month_last_day(year, month)
    end = int(datetime.datetime(year, month, last_day).strftime("%W"))
    begin = int(datetime.datetime(year, month, 1).strftime("%W"))
    last_week = end - begin + 1
    return last_week


class Daily:
    def __init__(self, data, log):
        if not isinstance(data, dict):
            raise IOError('Daily data is not a dic')
        self.data = data
        self.log = log
        self.daily_mode = self.data['Schedule_Daily_Backup_Method'] \
            if 'Schedule_Daily_Backup_Method' in self.data else '2'
        self.execute_time = self.execute_mode()

    def execute_mode(self):
        if self.daily_mode == '1':
            return self.at_time_mode()
        else:
            return self.interval_mode()

    def at_time_mode(self):  # mode one
        self.log.logger.info('Daily mode 1:backup at static time')
        if 'Schedule_Daily_addtime' not in self.data.items():
            raise IOError('Daily schedule at_time_mode does not have parameter "Schedule_Daily_addtime"')
        if ',' in self.data['Schedule_Daily_addtime']:
            at_time = self.data['Schedule_Daily_addtime'].split(',')
        else:
            at_time = [self.data['Schedule_Daily_addtime']]
        # format specify time to stamp(example:string 08:00 to int 480)
        execute_time = []
        for _time in at_time:
            execute_time.append(time_format(_time))
        return sorted(execute_time)

    def interval_mode(self):  # mode two
        self.log.logger.info('Daily mode 2:backup at interval time between start time and end time')
        if 'Schedule_Daily_Start_time' not in self.data.items():
            raise IOError('Daily schedule interval_mode does not have parameter "Schedule_Daily_Start_time"')
        if 'Schedule_Daily_End_time' not in self.data.items():
            raise IOError('Daily schedule interval_mode does not have parameter "Schedule_Daily_End_time"')
        if 'Schedule_Daily_Every_time' not in self.data.items():
            raise IOError('Daily schedule interval_mode does not have parameter "Schedule_Daily_Every_time"')
        start_time = time_format(self.data['Schedule_Daily_Start_time'])
        end_time = time_format(self.data['Schedule_Daily_End_time'])
        interval_time = time_format(self.data['Schedule_Daily_Every_time'])
        return [_time for _time in range(start_time, end_time + 1, interval_time)]

    def run(self, times):  # parameter "times":backup times
        while times > 0:
            next_execute_time = self.get_next_execute_time()
            self.log.logger.info('Modify time to:%s' % time_format(next_execute_time))
            modify_time(next_execute_time)
            if check_process_exist():
                self.log.logger.info('Backup start,schedule set success')
                times -= 1
            else:
                self.log.logger.error('Backup not start,schedule set failed')
                return False
        return True

    def get_next_execute_time(self):
        _date, _time = get_date_and_time()
        for i in self.execute_time:
            if i >= _time:
                execute_time = _date + i
                break
        else:
            execute_time = _date + 86400 + self.execute_time[0]
        return execute_time


class Weekly:
    def __init__(self, data, log):
        if not isinstance(data, dict):
            raise IOError('Daily data is not a dic')
        self.data = data
        self.log = log
        if 'Schedule_Weekly_Time' not in self.data.items():
            raise IOError('Weekly schedule data does not have keyword:Schedule_Weekly_Time')
        if 'Schedule_Weekly_Date' not in self.data.items():
            raise IOError('Weekly schedule data does not have keyword:Schedule_Weekly_Date')
        self._time = time_format(self.data['Schedule_Weekly_Time'])
        if ',' in self.data['Schedule_Weekly_Date']:
            self._weekdays = sorted([int(i) for i in self.data['Schedule_Weekly_Date'].split(',')])
        else:
            self._weekdays = [int(self.data['Schedule_Weekly_Date'])]

    def get_next_execute_time(self):
        """
            逻辑有点复杂，还是使用中文解释
            1、先求出当天是星期几[1,2,3,4,5,6,0]
            2、再进行逻辑判断
               a.当 今天_weekday < weekday(设置的备份星期),那么下次备份日期就是weekday
               b.当 今天_weekday = weekday,那么下次备份日期还需要看具体时间，如果当前时间 _time <= self._time（设置时间）
                    那么，下次备份时间就是当天的设置时间，如果_time > self._time(当天备份时间已过，则判断下一个备份时间)
               c.当 今天_weekday > 所有weekday，那么备份时间为下周的第一次备份时间
        :return:
        """
        _date, _time = get_date_and_time()
        _weekday = time.localtime().tm_wday + 1  # _weekday:today's weekday
        _weekday = _weekday if _weekday <= 6 else 0  # format week [0,1,2,3,4,5,6] to [1,2,3,4,5,6,0]
        for weekday in self._weekdays:  # weekday:execute weekday
            if _weekday < weekday:
                next_execute_time = _date + (weekday - _weekday) * 86400 + self._time
                break
            elif _weekday == weekday and _time <= self._time:
                next_execute_time = _date + self._time
                break
        else:
            next_execute_time = _date + (self._weekdays[0] - _weekday + 7) * 86400 + self._time
        return next_execute_time

    def run(self, times):  # parameter "times":backup times
        while times > 0:
            next_execute_time = self.get_next_execute_time()
            self.log.logger.info('Modify time to:%s' % time_format(next_execute_time))
            modify_time(next_execute_time)
            if check_process_exist():
                self.log.logger.info('Backup start,schedule set success')
                times -= 1
            else:
                self.log.logger.error('Backup not start,schedule set failed')
                return False
        return True


class Monthly:
    def __init__(self, data, log):
        if not isinstance(data, dict):
            raise IOError('Daily data is not a dic')
        self.data = data
        self.log = log
        self._time = time_format(self.data['Schedule_Monthly_Which_Time'])
        self.monthly_mode = self.data['Schedule_Monthly_Backup_Method'] \
            if 'Schedule_Monthly_Backup_Method' in self.data else '2'

    def run(self, times):
        while times > 0:
            next_execute_time = self.get_next_execute_time()
            self.log.logger.info('Modify time to:%s' % time_format(next_execute_time))
            modify_time(next_execute_time)
            if check_process_exist():
                self.log.logger.info('Backup start,schedule set success')
                times -= 1
            else:
                self.log.logger.error('Backup not start,schedule set failed')
                return False
        return True

    def get_next_execute_time(self):
        if self.monthly_mode == '1':
            return self.at_days_mode()
        else:
            return self.which_week_mode()

    def at_days_mode(self):
        if 'Schedule_Monthly_At_Days' not in self.data.items():
            raise IOError('Monthly data does not have keyword:Schedule_Monthly_At_Days')
        if ',' in self.data['Schedule_Monthly_At_Days']:
            execute_days = sorted([int(i) for i in self.data['Schedule_Monthly_At_Days'].split(',')])
        else:
            execute_days = [int(self.data['Schedule_Monthly_At_Days'])]
        # 判断是否包含最后一天，并转换
        if execute_days[0] == 0:
            execute_days.append(get_month_last_day(time.localtime().tm_year, time.localtime().tm_mon))
            del execute_days[0]
        today = time.localtime().tm_mday  # today
        _date, _time = get_date_and_time()
        for day in execute_days:
            if today < day:
                next_execute_time = _date + (day - today) * 86400 + self._time
                break
            elif today == day and _time <= self._time:
                next_execute_time = _date + self._time
                break
            else:
                continue
        else:
            next_year, next_month = get_next_month()
            next_execute_time = next_year + '/' + next_month + '/' + str(execute_days[0]) + \
                                ' ' + self.data['Schedule_Monthly_Which_Time']
            next_execute_time = int(time.mktime(time.strptime(next_execute_time, "%Y/%m/%d %H:%M")))
        return next_execute_time

    def which_week_mode(self):
        struct_time = time.localtime()
        year, month, day = struct_time.tm_year, struct_time.tm_mon, struct_time.tm_mday
        this_week = get_week_of_month(year, month, day)
        last_week = get_month_last_week(year, month)
        which_week = int(self.data['Schedule_Monthly_Which_Week'])
        which_week = last_week if which_week == 6 else which_week
        if ',' in self.data['Schedule_Monthly_Which_Day']:
            execute_days = sorted([int(i) for i in self.data['Schedule_Monthly_Which_Day'].split(',')])
        else:
            execute_days = [int(self.data['Schedule_Monthly_Which_Day'])]
        _weekday = day + 1
        _weekday = _weekday if _weekday <= 6 else 0
        return self.get_week_mode_execute_time(this_week, which_week, _weekday, execute_days)

    def get_week_mode_execute_time(self, this_week, which_week, _weekday, execute_days):
        next_execute_time = 0
        _date, _time = get_date_and_time()
        if this_week < which_week:
            next_execute_time = _date + ((which_week-this_week)*7 + execute_days[0]-_weekday)*86400 + self._time
        elif this_week == which_week:
            for weekday in execute_days:
                if _weekday < weekday:
                    next_execute_time = _date + (weekday-_weekday)*86400 + self._time
                    break
                elif _weekday == weekday and _time <= self._time:
                    next_execute_time = _date + self._time
                    break
        else:
            # 如果当月无执行计划时间，则修改时间+1天
            modify_time(_date + 86400)
            time.sleep(2)
            self.which_week_mode()
        return next_execute_time


class Schedule:
    def __init__(self, data, result_path, execute_times, log):
        self.name = 'TbService.exe'
        self.data = data
        self.result_path = result_path
        self.execute_times = execute_times
        self.log = log

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
        result = 'success' if result else 'Failed'
        write_result(self.result_path, result)

    def one_time(self):
        set_time = time.strptime(self.get_one_time_data(), "%Y/%m/%d %H:%M")
        modify_time(int(time.mktime(set_time)))
        return True if check_process_exist(self.name) else False

    def daily(self):
        daily = Daily(self.data, self.log)
        return daily.run(self.execute_times)

    def weekly(self):
        weekly = Weekly(self.data, self.log)
        return weekly.run(self.execute_times)

    def monthly(self):
        monthly = Monthly(self.data, self.log)
        return monthly.run(self.execute_times)

    def get_one_time_data(self):
        assert 'schedule_onetime_time' in self.data.items
        assert 'schedule_onetime_date' in self.data.items
        return self.data['schedule_onetime_date'] + ' ' + self.data['schedule_onetime_time']


if __name__ == '__main__':
    print(time.localtime())
