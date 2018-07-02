#! coding=utf8
"""
    start schedule

"""
import sys
import os
from bin.log import MyLog
from bin.main import run
from bin.parse import parsing


# ********************************Get executing path******************************
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(__file__)
# ********************************************************************************


LOG = MyLog(BASE_PATH, 'log.log')

if __name__ == '__main__':
    try:
        LOG.logger.info('========================schedule============================')
        run(BASE_PATH, LOG, parsing(sys.argv[1]), sys.argv[2])
    except Exception as e:
        LOG.logger.error(e)
        LOG.logger.info('========================schedule end========================')
