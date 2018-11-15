#!/usr/bin/env python
# -*- coding: utf-8 -*-

import _winreg
import subprocess


def detect_edge_sid():
    """
    探测Edge浏览器sid，找到并返回sid的值
    :return: sid
    """
    reg_path = r'Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Mappings'
    key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, reg_path)

    try:
        idx = 0
        while True:
            app_sid = _winreg.EnumKey(key, idx)

            try:
                sub_idx = 0
                sub_key = _winreg.OpenKey(
                    _winreg.HKEY_CURRENT_USER, reg_path + '\\' + app_sid)
                while True:
                    name, value, _ = _winreg.EnumValue(sub_key, sub_idx)
                    sub_idx += 1

                    if r'@{Microsoft.MicrosoftEdge_' in value:
                        print 'app_sid:', app_sid
                        print '\t{0}: {1}\n'.format(name, value.encode('utf8'))
                        return app_sid
            except WindowsError:
                _winreg.CloseKey(sub_key)

            idx += 1
    except WindowsError:
        _winreg.CloseKey(key)


def enable_loopback(app_sid):
    """
    开启Edge浏览器环路IP访问
    :param app_sid: Edge浏览器的app_sid
    :return:
    """
    sid_cmd = '-p={0}'.format(app_sid)

    p1 = subprocess.Popen(['CheckNetIsolation.exe', 'loopbackexempt', '-a', sid_cmd], stdout=subprocess.PIPE)
    result = p1.communicate()

    if result[1] is None:
        if u'完成' in result[0].decode('gb2312'):
            # print u'执行结果: {0}'.format(result[0].decode('gb2312'))
            return True
        else:
            return False
    else:
        print u'执行结果: {0}'.format(result[1].decode('gb2312'))
        return False

    return False


def main():
    """
    主函数
    """
    app_sid = detect_edge_sid()
    if app_sid is None:
        print u'注册表中没有找到Edge的sid'
        return

    if enable_loopback(app_sid):
        print u'激活成功！'
    else:
        print u'激活失败。。。'


if __name__ == '__main__':
    main()
