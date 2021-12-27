from win32file import CreateFile, SetFileTime, GetFileTime, CloseHandle, CreateDirectory
from win32file import GENERIC_READ, GENERIC_WRITE, OPEN_EXISTING
from pywintypes import Time
import win32timezone
from typing import List
import time
import datetime
import random
import os
import re

# 仅适用于win32
import sys
assert sys.platform == "win32"

default_time_format = "%Y-%m-%d %H:%M:%S"  # 时间格式


def set_sys_time(time_st):
    import win32api
    print("set time to", time_st)
    tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec, tm_wday, tm_yday, tm_isdst = time_st

    win32api.SetSystemTime(tm_year, tm_mon, tm_wday, tm_mday, tm_hour, tm_min, tm_sec, 0)


def timeOffsetAndStruct(times, fmt, offset):
    return time.localtime(time.mktime(time.strptime(times, fmt)) + offset)


def _mod_file_attr(src: str, create_st: str, create_ed: str,
                   mod_range_st: float, mod_range_ed: float,
                   vst_range_st: float, vst_range_ed: float,
                   time_format: str = default_time_format,
                   logger = None,
                   filter_list: List[str] = None) -> object:
    """
    用来修改任意文件的相关时间属性，时间格式：YYYY-MM-DD HH:MM:SS 例如：2019-02-02 00:01:02
    :param filePath: 文件路径名
    :param createTime: 创建时间
    :param modifyTime: 修改时间
    :param accessTime: 访问时间
    :param offset: 时间偏移的秒数,tuple格式，顺序和参数时间对应
    """
    if filter_list:
        for item in filter_list:
            try:
                if re.search(item, src):
                    return
            except:
                pass
    stamp_st = datetime.datetime.strptime(create_st, time_format)
    stamp_ed = datetime.datetime.strptime(create_ed, time_format)
    delta_time = (stamp_ed - stamp_st).total_seconds()
    createTime = stamp_st + datetime.timedelta(seconds=random.randint(0, int(delta_time)))
    modifyTime = createTime + datetime\
        .timedelta(seconds=random.randint(int(mod_range_st), int(mod_range_ed)))
    accessTime = modifyTime + datetime\
        .timedelta(seconds=random.randint(int(vst_range_st), int(vst_range_ed)))

    logger("{0} --> 创建:{1} | 修改:{2} | 访问:{3}".format(src, createTime, modifyTime, accessTime))
    if os.path.isfile(src):
        try:
            cTime_t = createTime.timetuple()
            mTime_t = modifyTime.timetuple()
            aTime_t = accessTime.timetuple()
            fh = CreateFile(src, GENERIC_READ | GENERIC_WRITE, 0, None, OPEN_EXISTING, 0, 0)
            createTimes, accessTimes, modifyTimes = GetFileTime(fh)
            createTimes = Time(time.mktime(cTime_t))
            accessTimes = Time(time.mktime(aTime_t))
            modifyTimes = Time(time.mktime(mTime_t))
            SetFileTime(fh, createTimes, accessTimes, modifyTimes)
            CloseHandle(fh)
            return createTime, modifyTime, accessTime
        except Exception as e:
            raise e
    elif os.path.isdir(src):
        N_createTime = createTime
        N_modifyTime = modifyTime
        N_accessTime = accessTime
        for item in os.listdir(src):
            bak = _mod_file_attr(src + os.sep + item, create_st, create_ed,
                                 mod_range_st, mod_range_ed,
                                 vst_range_st, vst_range_ed,
                                 time_format, logger, filter_list)
            # 统计子文件的最新时间
            if isinstance(bak, tuple):
                if bak[0] > N_createTime:
                    N_createTime = bak[0]
                if bak[1] > N_modifyTime:
                    N_modifyTime = bak[1]
                if bak[2] > N_accessTime:
                    N_accessTime = bak[2]
        os.utime(src, (N_accessTime.timestamp(), N_modifyTime.timestamp()))
        return N_createTime, N_modifyTime, N_accessTime
    else:
        if callable(logger):
            logger("not support file type: {0}".format(src))


def mod_files_attr(src: str, create_st: str, create_ed: str,
                   mod_range_st: float, mod_range_ed: float,
                   vst_range_st: float, vst_range_ed: float,
                   time_format: str = default_time_format,
                   logger = None,
                   filter_list: List[str] = None) -> None:
    _mod_file_attr(src, create_st, create_ed,
                   mod_range_st, mod_range_ed,
                   vst_range_st, vst_range_ed,
                   time_format, logger, filter_list)



# if __name__ == "__main__":
#     mod_files_attr(src='XXX', create_st='1991-01-01 10:10:10',
#                    create_ed='1992-01-01 10:10:10',
#                    mod_range_st=3*60*60, mod_range_ed=24*60*60,
#                    vst_range_st=3*60*60, vst_range_ed=24*60*60,
#                    time_format="%Y-%m-%d %H:%M:%S",
#                    logger=print)
