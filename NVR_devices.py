# 检测同一 IP 对应多个设备并输出提示结果
def detect_duplicate_ips(all_channels):
    ip_count = {}
    for raw_id, info in all_channels.items():
        cam_ip = info["ip"]
        if cam_ip != "无IP":
            if cam_ip in ip_count:
                ip_count[cam_ip].append(info["name"])
            else:
                ip_count[cam_ip] = [info["name"]]

    # 输出提示结果
    for ip, names in ip_count.items():
        if len(names) > 1:
            print(f"[警告] IP {ip} 对应多个设备: {', '.join(names)}")

    return ip_count

# 在合适的位置调用该函数，例如在 check_single_nvr 函数的末尾
    detect_duplicate_ips(all_channels)
# -*- coding: utf-8 -*-
import requests
import xml.etree.ElementTree as ET
from requests.auth import HTTPDigestAuth
import csv
import os
import sys
import io
import re

# 💡 强行重定向标准输出的编码，防止 Windows 控制台编码错误
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(BASE_DIR, "NVRdevices.csv")

CSV_DIR = r"d:\sshAuto\reports_csv"
os.makedirs(CSV_DIR, exist_ok=True)
RESULT_CSV_PATH = os.path.join(CSV_DIR, "nvr_temp_report.csv")

def get_nvr_list(csv_path):
    nvr_list = []
    if not os.path.exists(csv_path):
        print(f"[-] 错误：找不到 CSV 文件 -> {csv_path}")
        return nvr_list
        
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            nvr_list.append({k.strip(): v.strip() for k, v in row.items()})
    return nvr_list

def get_base_name(name):
    """提取基础名称（强化对“主/子”、“镜头”等中文后缀的识别）"""
    if not name or name.lower().startswith("camera"):
        return ""
    # 剔除复杂的后缀，如 -1, _2, 通道1, ch2, 镜头1, 主码流, 子码流, 主, 子等
    pattern = r'[-_ ]?(通道|chan|ch|镜头|码流)?\d+$|[-_ ]?(主码流|子码流|主通道|子通道|主|子|主摄|副摄)$'
    base = re.sub(pattern, '', name, flags=re.IGNORECASE).strip()
    return base if len(base) >= 2 else ""
def check_single_nvr(nvr):
    ip = nvr.get('ip')
    user = nvr.get('username')
    pw = nvr.get('password')
    
    CLR_RESET = "\033[0m"
    CLR_RED = "\033[1;31m"      
    CLR_GREEN = "\033[1;32m"    
    CLR_YELLOW = "\033[1;33m"   
    CLR_CYAN = "\033[1;36m"     
    
    print(f"\n{CLR_CYAN}[SCAN] 正在扫描 NVR: {ip} ...{CLR_RESET}")
    
    config_url = f"http://{ip}/ISAPI/ContentMgmt/InputProxy/channels"
    status_url = f"http://{ip}/ISAPI/ContentMgmt/InputProxy/channels/status"
    
    all_channels = {} 

    # 1. 获取全量配置 (撤销了会引发丢数据的 int 强转)
    try:
        config_res = requests.get(config_url, auth=HTTPDigestAuth(user, pw), timeout=8)
        if config_res.status_code == 200:
            config_res.encoding = 'utf-8'
            c_root = ET.fromstring(config_res.text)
            for chan in c_root.findall('.//{*}InputProxyChannel'):
                id_node = chan.find('.//{*}id')
                if id_node is not None and id_node.text:
                    raw_id = id_node.text.strip() # 💡 保持字符串，绝不遗漏设备

                    name_node = chan.find('.//{*}name')
                    c_name = name_node.text.strip() if (name_node is not None and name_node.text) else f"Camera {raw_id}"
                    
                    c_ip = "无IP"
                    for ip_node in chan.findall('.//{*}ipAddress'):
                        if ip_node is not None and ip_node.text and ip_node.text.strip() not in ["0.0.0.0", ""]:
                            c_ip = ip_node.text.strip()
                            break

                    c_sn = ""
                    for sn_node in chan.findall('.//{*}serialNumber'):
                        if sn_node is not None and sn_node.text:
                            c_sn = sn_node.text.strip()
                            break

                    all_channels[raw_id] = {
                        "id": raw_id,
                        "name": c_name,
                        "ip": c_ip,
                        "sn": c_sn,
                        "online": False
                    }
    except Exception as e:
        print(f"  [!] {CLR_YELLOW}读取配置接口异常: {e}{CLR_RESET}")

    # 2. 获取状态并覆盖合并
    try:
        response = requests.get(status_url, auth=HTTPDigestAuth(user, pw), timeout=8)
        if response.status_code == 200:
            response.encoding = 'utf-8' 
            root = ET.fromstring(response.text)
            for channel in root.findall('.//{*}InputProxyChannelStatus'):
                id_node = channel.find('.//{*}id')
                if id_node is not None and id_node.text:
                    raw_id = id_node.text.strip()

                    if raw_id not in all_channels:
                        all_channels[raw_id] = {"id": raw_id, "name": f"Camera {raw_id}", "ip": "无IP", "sn": "", "online": False}

                    online_node = channel.find('.//{*}online')
                    if online_node is not None and online_node.text:
                        is_online = online_node.text.strip().lower() == 'true'
                    else:
                        status_node = channel.find('.//{*}status')
                        status_text = status_node.text.strip().lower() if status_node is not None else ""
                        is_online = (status_text == 'online')
                    
                    all_channels[raw_id]["online"] = is_online

                    ip_node = channel.find('.//{*}ipAddress')
                    if ip_node is not None and ip_node.text and ip_node.text.strip() not in ["0.0.0.0", ""]:
                        all_channels[raw_id]["ip"] = ip_node.text.strip()
                    
                    sn_node = channel.find('.//{*}serialNumber')
                    if sn_node is not None and sn_node.text:
                        all_channels[raw_id]["sn"] = sn_node.text.strip()
                        
                    dev_name_node = channel.find('.//{*}name') or channel.find('.//{*}deviceName')
                    if dev_name_node is not None and dev_name_node.text and all_channels[raw_id]["name"].startswith("Camera"):
                        all_channels[raw_id]["name"] = dev_name_node.text.strip()

    except Exception as e:
        print(f"  {CLR_RED}[EXCEPT] 连接失败: {e}{CLR_RESET}")
        return []

    if not all_channels:
        print(f"  [!] {CLR_YELLOW}未发现任何通道信息{CLR_RESET}")
        return []

    # 3. 💡 建立 IP/SN 共享池：同名双摄如果有一半断网，强行让断网的继承兄弟的IP和SN
    name_to_ip = {}
    name_to_sn = {}
    for raw_id, info in all_channels.items():
        base = get_base_name(info["name"])
        if base:
            if info["ip"] != "无IP": name_to_ip[base] = info["ip"]
            if info["sn"]: name_to_sn[base] = info["sn"]

    # 4. 字典聚类法（彻底修复 10.175.202.62 拆分与漏报）
    physical_devices = {}

    for raw_id, info in all_channels.items():
        cam_name = info["name"]
        cam_ip = info["ip"]
        cam_sn = info["sn"]
        
        base = get_base_name(cam_name)
        
        # 补全丢失的信息
        if cam_ip == "无IP" and base and base in name_to_ip:
            cam_ip = name_to_ip[base]
        if not cam_sn and base and base in name_to_sn:
            cam_sn = name_to_sn[base]
            
        # 聚类 Key 优先级调整：IP地址才是唯一真理！
        if cam_ip != "无IP": 
            device_key = f"IP_{cam_ip}"
        elif cam_sn: 
            device_key = f"SN_{cam_sn}"
        elif base: 
            device_key = f"NAME_{base}"
        else: 
            device_key = f"CHAN_{raw_id}"
            
        if device_key not in physical_devices:
            physical_devices[device_key] = {
                "name": cam_name,  # 💡 完美保留最原始的长名称，绝不随意截断！
                "ip": cam_ip,
                "channels": {}
            }
        else:
            if physical_devices[device_key]["ip"] == "无IP" and cam_ip != "无IP":
                physical_devices[device_key]["ip"] = cam_ip
                
        physical_devices[device_key]["channels"][f"D{raw_id}"] = {
            "name": cam_name,
            "online": info["online"]
        }

    # 5. 生成报告
    print(f"\n{CLR_CYAN}[ALERT] [网络状态诊断清单]：{CLR_RESET}")
    print("-" * 85)

    scan_results = []
    online_physical_count = 0
    offline_physical_count = 0

    for dev_key, dev in physical_devices.items():
        chans = dev["channels"]
        total_ch = len(chans)
        channel_type = "单通道" if total_ch == 1 else "双通道"
        
        offline_ch_ids = [ch_id for ch_id, info in chans.items() if not info["online"]]
        chans_str = ", ".join(chans.keys())

        if len(offline_ch_ids) == total_ch:
            offline_physical_count += 1
            print(f" {CLR_RED}❌ [网络不可达] {dev['name']} | IP: {dev['ip']} | 属性: {channel_type} | 对应通道: {chans_str}{CLR_RESET}")
            scan_results.append({
                "所属NVR": ip,
                "监控设备名称": dev['name'],
                "IP地址": dev['ip'],
                "通道属性": channel_type,
                "对应/异常通道": chans_str,
                "状态": "网络不可达"
            })
        elif len(offline_ch_ids) > 0:
            online_physical_count += 1
            bad_chans_str = ", ".join(offline_ch_ids)
            print(f" {CLR_YELLOW}⚠️ [部分不可达] {dev['name']} | IP: {dev['ip']} | 属性: {channel_type} | 异常通道: {bad_chans_str} 无法连接{CLR_RESET}")
            scan_results.append({
                "所属NVR": ip,
                "监控设备名称": dev['name'],
                "IP地址": dev['ip'],
                "通道属性": channel_type,
                "对应/异常通道": bad_chans_str,
                "状态": "部分不可达"
            })
        else:
            online_physical_count += 1
            print(f" {CLR_GREEN}✅ [正常在线]  {dev['name']} | IP: {dev['ip']} | 属性: {channel_type} | 对应通道: {chans_str}{CLR_RESET}")
            scan_results.append({
                "所属NVR": ip,
                "监控设备名称": dev['name'],
                "IP地址": dev['ip'],
                "通道属性": channel_type,
                "对应/异常通道": chans_str,
                "状态": "正常在线"
            })

    print("-" * 85)
    offline_summary_color = CLR_RED if offline_physical_count > 0 else CLR_RESET
    print(f"[REPORT] 扫描完成 -> {CLR_GREEN}正常在线: {online_physical_count} 台{CLR_RESET} | {offline_summary_color}完全离线: {offline_physical_count} 台{CLR_RESET}")
    print("-" * 85)
    return scan_results

if __name__ == "__main__":
    nvr_list = get_nvr_list(CSV_PATH)
    all_results = []
    if nvr_list:
        for nvr in nvr_list:
            results = check_single_nvr(nvr)
            if results:
                all_results.extend(results)
                
        if all_results:
            fieldnames = ["所属NVR", "监控设备名称", "IP地址", "通道属性", "对应/异常通道", "状态"]
            with open(RESULT_CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(all_results)
            print(f"✅ 排查完成，临时报告保存在: {RESULT_CSV_PATH}")