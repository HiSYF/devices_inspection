#!/usr/bin/env python3
# -*- coding: UTF-8 -*-


import os
import sys
import time
import getpass
import threading
import msoffcrypto
import pandas as pd
from io import BytesIO
from netmiko import ConnectHandler
from netmiko import exceptions
from netmiko.ssh_dispatcher import CLASS_MAPPER as SUPPORTED_DEVICES # <--- 新增此行
import re

# 自定义异常类，用于处理输入密码为None情况
class PasswordRequiredError(Exception):
    """文件受密码保护，必须提供密码"""
    pass


FILENAME = input(f"\n请输入info文件名（默认为 info.xlsx）：") or "info.xlsx"  # 指定info文件名称
INFO_PATH = os.path.join(os.getcwd(), FILENAME)  # 读取info文件路径
LOCAL_TIME = time.strftime('%Y.%m.%d', time.localtime())  # 读取当前日期
LOCK = threading.Lock()  # 线程锁实例化
POOL = threading.BoundedSemaphore(200)  # 最大线程控制


# 判断info文件是否被加密，使用不同的读取方式
def read_info():
    if is_encrypted(INFO_PATH):
        return read_encrypted_file(INFO_PATH)  # 读取被加密info文件
    else:
        return read_unencrypted_file(INFO_PATH)  # 读取未加密info文件


# 检测info文件是否被加密
def is_encrypted(info_file: str) -> bool:
    try:
        with open(info_file, "rb") as f:
            return msoffcrypto.OfficeFile(f).is_encrypted()  # 检测info文件是否被加密
    except Exception:
        return False


# 读取被加密info文件
def read_encrypted_file(info_file: str, max_retry: int = 3) -> pd.DataFrame:
    retry_count = 0  # 初始化重试计数器，用于记录用户尝试输入密码的次数
    while retry_count < max_retry:  # 当重试次数小于最大允许重试次数时，继续循环
        try:
            password = getpass.getpass("\ninfo文件被加密，请输入密码：") or None  # 提示用户输入密码，隐式输入。如果用户直接按Enter键，password将为None
            if not password:  # 如果用户没有输入密码
                raise PasswordRequiredError("文件受密码保护，必须提供密码！")  # 抛出自定义异常，提示用户必须提供密码

            # 解密文件
            decrypted_data = BytesIO()  # 创建一个BytesIO对象，用于在内存中存储解密后的文件内容
            # BytesIO是一个内存中的二进制流，可以像文件一样进行读写操作
            with open(info_file, "rb") as f:  # 以二进制只读模式打开加密的info文件
                office_file = msoffcrypto.OfficeFile(f)  # 使用msoffcrypto库创建一个OfficeFile对象，表示加密的Office文件
                office_file.load_key(password=password)  # 使用用户提供的密码加载解密密钥
                office_file.decrypt(decrypted_data)  # 解密文件内容，并将解密后的数据写入decrypted_data对象中
            decrypted_data.seek(0)  # 将decrypted_data的指针重置到起始位置，以便后续读取操作
            # 由于解密后的数据已经写入decrypted_data，需要将指针重置到开头，以便后续读取

            # 读取解密后的文件
            devices_dataframe = pd.read_excel(decrypted_data, sheet_name=0, dtype=str, keep_default_na=False)
            cmds_dataframe = pd.read_excel(decrypted_data, sheet_name=1, dtype=str)

        except FileNotFoundError:  # 如果没有配置info文件或info文件名错误
            print(f'\n没有找到info文件！\n')  # 提示用户没有找到info文件或info文件名错误
            input('输入Enter退出！')  # 提示用户按Enter键退出
            sys.exit(1)  # 异常退出
        except ValueError:  # 捕获异常信息
            print(f'\ninfo文件缺失子表格信息！\n')  # 代表info文件缺失子表格信息
            input('输入Enter退出！')  # 提示用户按Enter键退出
            sys.exit(1)  # 异常退出
        except (msoffcrypto.exceptions.InvalidKeyError, PasswordRequiredError) as e:
            retry_count += 1
            if retry_count < max_retry:
                print(f"\n密码错误，请重新输入！（剩余尝试次数：{max_retry - retry_count}）")
            else:
                input("\n超过最大尝试次数，输入Enter退出！")
                sys.exit(1)
        except Exception as e:
            print(f"\n解密失败：{str(e)}")
            sys.exit(1)
        else:
            devices_dict = devices_dataframe.to_dict('records')  # 将DataFrame转换成字典
            # "records"参数规定外层为列表，内层以列标题为key，以此列的行内容为value的字典
            # 若有多列，代表字典内有多个key:value对；若有多行，每行为一个字典

            cmds_dict = cmds_dataframe.to_dict('list')  # 将DataFrame转换成字典
            # "list"参数规定外层为字典，列标题为key，列下所有行内容以list形式为value的字典
            # 若有多列，代表字典内有多个key:value对

            return devices_dict, cmds_dict


# 读取未加密info文件
def read_unencrypted_file(info_file: str) -> pd.DataFrame:
    try:
        devices_dataframe = pd.read_excel(info_file, sheet_name=0, dtype=str, keep_default_na=False)
        cmds_dataframe = pd.read_excel(info_file, sheet_name=1, dtype=str)
    except FileNotFoundError:  # 如果没有配置info文件或info文件名错误
        print(f'\n没有找到info文件！\n')  # 代表没有找到info文件或info文件名错误
        input('输入Enter退出！')  # 提示用户按Enter键退出
        sys.exit(1)  # 异常退出
    except ValueError:  # 捕获异常信息
        print(f'\ninfo文件缺失子表格信息！\n')  # 代表info文件缺失子表格信息
        input('输入Enter退出！')  # 提示用户按Enter键退出
        sys.exit(1)  # 异常退出
    else:
        devices_dict = devices_dataframe.to_dict('records')  # 将DataFrame转换成字典
        # "records"参数规定外层为列表，内层以列标题为key，以此列的行内容为value的字典
        # 若有多列，代表字典内有多个key:value对；若有多行，每行为一个字典

        cmds_dict = cmds_dataframe.to_dict('list')  # 将DataFrame转换成字典
        # "list"参数规定外层为字典，列标题为key，列下所有行内容以list形式为value的字典
        # 若有多列，代表字典内有多个key:value对

        return devices_dict, cmds_dict


def disable_paging(ssh_connection, device_info):
    """根据设备的原始功能类型，发送相应的命令来禁用分页显示。"""
    # <--- 关键变更：我们现在基于原始设备类型来选择命令
    device_type = device_info.get('original_device_type')

    PAGING_COMMANDS = {
        'cisco_ios': 'terminal length 0',
        'huawei': 'screen-length 0 temporary',
        'h3c_comware': 'screen-length disable',
        'ruijie_os': 'terminal length 0',
        # 在这里为您自定义的、不受Netmiko支持的设备类型添加命令
        # 例如: 'my_custom_firewall': 'set cli pagination off',
    }

    # 尝试从特定命令字典中获取命令，如果找不到，则尝试一个通用命令
    command = PAGING_COMMANDS.get(device_type) or 'terminal length 0'

    if command:
        try:
            # 优先使用 send_config_set
            ssh_connection.send_config_set([command])
        except Exception:
            try:
                # 备用方案
                prompt = ssh_connection.find_prompt()
                ssh_connection.send_command(command, expect_string=re.escape(prompt))
            except Exception as e:
                with LOCK:
                    print(
                        f"警告：设备 {device_info['host']} (类型: {device_type}) 无法执行禁用分页命令 '{command}'。错误: {e}")
# 巡检
def inspection(login_info, cmds_dict):
    # 使用传入的设备登录信息和巡检命令，登录设备依次输入巡检命令，如果设备登录出现异常，生成01log文件记录。
    t11 = time.time()  # 子线程执行计时起始点
    ssh = None  # 初始化ssh对象

    try:  # 尝试登录设备
        # 1. 复制一份传入的字典，以免影响原始数据
        netmiko_params = login_info.copy()

        # 2. 从这个副本中移除我们自定义的、Netmiko不认识的键
        netmiko_params.pop('original_device_type', None)  # .pop(key, None) 可以安全移除，即使键不存在也不会报错
        netmiko_params.pop('file_extension', None)  # <-- 也移除自定义的后缀名键

        # 3. 将这个“干净”的字典传递给 ConnectHandler
        ssh = ConnectHandler(**netmiko_params)
        ssh.enable()  # 进入设备Enable模式
    except Exception as ssh_error:  # 登录设备出现异常
        with LOCK:  # 线程锁
            exception_name = type(ssh_error).__name__

            if exception_name == 'AttributeError':  # 异常名称为：AttributeError
                print(f'设备 {login_info["host"]} 缺少设备管理地址！')  # CMD输出提示信息
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} 缺少设备管理地址！\n')  # 记录到log文件
            elif exception_name == 'NetmikoTimeoutException':
                print(f'设备 {login_info["host"]} 管理地址或端口不可达！')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} 管理地址或端口不可达！\n')
            elif exception_name == 'NetmikoAuthenticationException':
                print(f'设备 {login_info["host"]} 用户名或密码认证失败！')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} 用户名或密码认证失败！\n')
            elif exception_name == 'ValueError':
                print(f'设备 {login_info["host"]} Enable密码认证失败！')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} Enable密码认证失败！\n')
            elif exception_name == 'TimeoutError':
                print(f'设备 {login_info["host"]} Telnet连接超时！')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} Telnet连接超时！\n')
            elif exception_name == 'ReadTimeout':
                print(f'设备 {login_info["host"]} Enable密码认证失败！')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} Enable密码认证失败！\n')
            elif exception_name == 'ConnectionRefusedError':
                print(f'设备 {login_info["host"]} 远程登录协议错误！')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} 远程登录协议错误！\n')
            else:
                print(f'设备 {login_info["host"]} 未知错误！{type(ssh_error).__name__}: {str(ssh_error)}')
                with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'a', encoding='utf-8') as log:
                    log.write(f'设备 {login_info["host"]} 未知错误！{type(ssh_error).__name__}: {str(ssh_error)}\n')
    else:  # 如果登录正常，开始巡检
        # 1. 禁用分页 (基于原始类型)
        disable_paging(ssh, login_info)

        # 2. <--- 关键变更：使用 'original_device_type' 从Excel数据中查找命令列表
        device_type_for_cmds = login_info.get('original_device_type')

        file_ext = login_info.get('file_extension')
        if not file_ext or not file_ext.strip():
            file_ext = '.log'

        # 2. 确保后缀名以 '.' 开头
        file_ext = file_ext.strip()
        if not file_ext.startswith('.'):
            file_ext = '.' + file_ext

        # 3. 构建最终的文件名 (格式: hostname-devicetype.extension)
        output_filename = os.path.join(os.getcwd(), LOCAL_TIME, f"{login_info['host']}{file_ext}")

        # 3. 检查 cmds_dict 中是否存在该原始类型的命令列表
        if not device_type_for_cmds or device_type_for_cmds not in cmds_dict:
            with LOCK:
                print(
                    f"错误：设备 {login_info['host']} (原始类型: {device_type_for_cmds}) - 在 info.xlsx 的命令表中找不到对应的命令列。")
            with open(output_filename, 'w', encoding='utf-8') as f:
                f.write(f"错误：找不到原始设备类型 '{device_type_for_cmds}' 对应的巡检命令。\n")
        else:
            with open(output_filename, 'w',
                      encoding='utf-8') as device_log_file:
                with LOCK:
                    print(f'设备 {login_info["host"]} 正在巡检...')
                device_log_file.write('=' * 10 + ' ' + 'Local Time' + ' ' + '=' * 10 + '\n\n')
                device_log_file.write(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '\n\n')

                for cmd in cmds_dict[device_type_for_cmds]:
                    if isinstance(cmd, str) and cmd:
                        device_log_file.write('=' * 10 + ' ' + cmd + ' ' + '=' * 10 + '\n\n')
                        try:
                            show = ssh.send_command(cmd, read_timeout=120)
                        except exceptions.ReadTimeout:
                            print(f'设备 {login_info["host"]} 命令 "{cmd}" 执行超时！')
                            show = f'命令 "{cmd}" 执行超时！'
                        finally:
                            device_log_file.write(show + '\n\n')

        t12 = time.time()
        with LOCK:  # 线程锁
            print(f'设备 {login_info["host"]} 巡检完成，用时 {round(t12 - t11, 1)} 秒。')  # 打印子线程执行时长
    finally:  # 最后结束SSH连接释放线程
        if ssh is not None:  # 判断ssh对象是否被正确赋值，赋值成功不为None，即SSH连接已建立，需要关闭连接
            ssh.disconnect()  # 关闭SSH连接
        POOL.release()  # 最大线程限制，释放一个线程


if __name__ == '__main__':
    t1 = time.time()
    threading_list = []
    devices_info, cmds_info = read_info()

    print(f'\n巡检开始...')
    print(f'\n' + '>' * 40 + '\n')

    if not os.path.exists(LOCAL_TIME):
        os.makedirs(LOCAL_TIME)
    else:
        try:
            os.remove(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'))
        except FileNotFoundError:
            pass

    for device_info in devices_info:
        updated_device_info = device_info.copy()
        updated_device_info["conn_timeout"] = 40

        # =================================================================================
        # ==== 关键逻辑变更：同时保存原始类型和连接类型 ====
        # =================================================================================
        original_device_type = updated_device_info.get('device_type')

        # 1. 无论如何，都将原始类型保存下来，用于后续的命令查找
        updated_device_info['original_device_type'] = original_device_type

        # 2. 仅当连接需要时，才修改 'device_type' 的值
        if original_device_type not in SUPPORTED_DEVICES:
            print(
                f"警告：设备 {updated_device_info.get('host')} 的类型 '{original_device_type}' 不在Netmiko支持列表，将使用 'generic' 类型进行连接。")
            updated_device_info['device_type'] = 'generic'
        # =================================================================================

        pre_device = threading.Thread(target=inspection, args=(updated_device_info, cmds_info),
                                      name=device_info['host'] + '_Thread')
        threading_list.append(pre_device)
        POOL.acquire()
        pre_device.start()

    for _ in threading_list:
        _.join()
    try:  # 尝试打开01log文件
        with open(os.path.join(os.getcwd(), LOCAL_TIME, '01log.log'), 'r', encoding='utf-8') as log_file:
            file_lines = len(log_file.readlines())  # 读取01log文件共有多少行。有多少行，代表出现了多少个设备登录异常
    except FileNotFoundError:  # 如果找不到01log文件
        file_lines = 0  # 证明本次巡检没有出现巡检异常情况
    t2 = time.time()  # 程序执行计时结束点
    print(f'\n' + '<' * 40 + '\n')  # 打印一行“<”，隔开巡检报告信息
    print(f'巡检完成，共巡检 {len(threading_list)} 台设备，{file_lines} 台异常，共用时 {round(t2 - t1, 1)} 秒。\n')  # 打印巡检报告
    input('输入Enter退出！')  # 提示用户按Enter键退出
