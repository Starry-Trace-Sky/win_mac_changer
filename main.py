# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# author: Skyler
from random import randint
import sys
import winreg
import ctypes

import pysnooper

reg_path = 'SYSTEM\\CurrentControlSet\\Control\\Class\\{4D36E972-E325-11CE-BFC1-08002BE10318}\\'
choose_ifc = None
choose_path = None


def is_admin():
    """判断当前用户是否拥有管理员权限"""
    result = ctypes.windll.shell32.IsUserAnAdmin()
    return result


def Exit(ecode):
    """退出"""
    try:
        sys.exit(ecode)
    except SystemExit as e:
        print("程序已结束,结束代码为", e.code)
    except InterruptedError:
        pass


def ck():
    """检测函数"""
    while True:
        ck = input("您是否以管理员权限运行该程序? y/n\n")
        # 判断用户主动判断是否管理员
        if ck == 'y':
            return 'y'
        elif (ck == 'q') or (ck == 'n'):
            # 程序退出
            Exit(0)
        else:
            print("Not solved.")


def create():
    """创建键值"""
    mac = "".join(["%02x" % x for x in map(lambda x: randint(0, 255), range(6))])
    mac = mac.upper()
    regg = None
    try:
        # 打开键,返回handler对象
        regg = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path + choose_path, reserved=0, access=winreg.KEY_WRITE)
        # 添加字符串
        winreg.SetValueEx(regg, 'NetworkAddress', 0, winreg.REG_SZ, mac)
    except FileNotFoundError:
        print("尝试新建注册表项时FileNotFoundError")
    finally:
        if regg != None:
            regg.Close()
    return mac


@pysnooper.snoop('mac.log', prefix='main:', normalize=True)
def main():
    """主程序"""
    print("本程序需要操作注册表,请留意杀软提示")
    print("请以管理员权限运行本程序")
    print("如遇报错,请查看本软件目录下的mac.log")
    print("Input 'q' at any time to exit")
    if ck() == 'y':
        ad_check = is_admin()
        reg = None
        # 判断是否为管理员
        if ad_check == 1:
            # 路径000x处理
            interface_l = []
            path_l = []
            interface_d = {}
            reg = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, reserved=0)
            i = 0
            while True:
                try:
                    path_l.append(winreg.EnumKey(reg, i) + '\\')
                    i += 1
                except OSError:
                    break
            # 获取网卡列表
            for path in path_l:
                try:
                    reg = None
                    reg = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path + path, reserved=0)
                    interface_l.append(winreg.QueryValueEx(reg, 'DriverDesc')[0])
                    interface_d[path] = winreg.QueryValueEx(reg, 'DriverDesc')[0]
                except FileNotFoundError:
                    pass
                except InterruptedError:
                    pass
                except PermissionError:
                    pass
                finally:
                    if reg != None:
                        reg.Close()

            del interface_l
            del path_l
            print("您的网卡列表如下")
            for key, value in interface_d.items():
                print(key + ':' + value)
            while True:
                pattern = input("请输入匹配网卡前的index,包含\\\n")
                if pattern == 'q':
                    Exit(0)
                elif pattern not in interface_d:
                    print("输入不存在")
                    continue
                elif pattern in interface_d:
                    # 输入处理
                    global choose_ifc
                    global choose_path
                    choose_ifc = interface_d[pattern]
                    choose_path = pattern
                    # 检测项是否存在
                    reg = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path + choose_path, reserved=0)
                    result = None
                    try:
                        result = winreg.QueryValueEx(reg, 'NetworkAddress')
                    except FileNotFoundError:
                        mmac = create()
                        print(f"mac地址修改完成,新的mac地址为{mmac}\n请重启设备,若修改失败,请再次运行该软件重复修改")
                        tp = input("键入任意键退出")
                        Exit(0)
                        break
                    if result == None:
                        # 项不存在,创建
                        mmac = create()
                        print(f"mac地址修改完成,新的mac地址为{mmac}\n请重启设备,若修改失败,请再次运行该软件重复修改")
                        tp = input("键入任意键退出")
                        Exit(0)
                        break
                    else:
                        # 项存在,修改
                        mmac = create()
                        print(f"mac地址修改完成,新的mac地址为{mmac}\n请重启设备,若修改失败,请再次运行该软件重复修改")
                        tp = input("键入任意键退出")
                        Exit(0)
                        break
                else:
                    print("输入不合法")

        else:
            print("请检查权限,然后重新运行该程序")

if __name__ == '__main__':
    main()