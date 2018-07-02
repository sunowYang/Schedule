#! coding=utf8


def parsing(data):
    dic = {}
    if ';' not in data:
        raise IOError('Schedule data does not contain ";"')
    data_list = data.split(';')
    for _data in data_list:
        if '=' not in _data:
            raise IOError('Schedule data does not contain "=": %s' % _data)
        dic[_data.split('=')[0]] = _data.split('=')[1]
    return dic
