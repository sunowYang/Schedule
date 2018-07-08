#! coding=utf8
from bin.schedule import Schedule


def run(result_path, data, execute_time, log):
    schedule = Schedule(data, result_path, execute_time, log)
    schedule.run_schedule()

