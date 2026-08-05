# -*- coding: utf-8 -*-
import os
import time
import re
import csv
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from netmiko import ConnectHandler
from prettytable import PrettyTable, TableStyle

GREEN = "\033[1;32m"
RED = "\033[1;31m"
YELLOW = "\033[1;33m"
BLUE = "\033[1;34m"
RESET = "\033[0m"

DEVICES = []
csv_file_path = "SWdevicesTest.csv"

if os.path.exists(csv_file_path):
    with open(csv_file_path, "r", encoding="utf-8-sig", errors="ignore") as f:
        first_line = f.readline()
        f.seek(0)
        delimiter = ';' if ';' in first_line else ','
        reader = csv.DictReader(f, delimiter=delimiter)
        for row in reader:
            clean_row = {k.strip().replace('"', '').replace('\ufeff', '').lower(): v.strip() for k, v in row.items() if k and v}
            ip = clean_row.get("ip")
            username = clean_row.get("username")
            password = clean_row.get("password")
            dev_type = clean_row.get("device_type", "huawei")
            if ip and username and password:
                DEVICES.append({
                    "device_type": dev_type, "ip": ip, "username": username, "password": password,
                    "port": 22, "conn_timeout": 15, "banner_timeout": 15
                })
else:
    print(f"{RED}× 错误：未找到 {csv_file_path} 设备列表文件！{RESET}")
    exit(1)

# ================= 严格定义绝对目标目录 =================
CSV_DIR = r"d:\sshAuto\reports_csv"
TEMP_DIR = r"d:\sshAuto\tempFlies"
BACKUP_DIR = r"d:\sshAuto\network_backups" # 配置备份也改为绝对路径，防止污染根目录

for d in [CSV_DIR, TEMP_DIR, BACKUP_DIR]:
    os.makedirs(d, exist_ok=True)

today = time.strftime("%Y%m%d")

def parse_latest_alarm(raw_log_buffer):
    if not raw_log_buffer or "Log buffer is empty" in raw_log_buffer:
        return {"text": "No Alarm", "severity": "normal"}
    lines = [line.strip() for line in raw_log_buffer.splitlines() if line.strip()]
    valid_logs = [line for line in lines if any(x in line for x in ["%%", "MSTP/", "IFNET/", "SHELL/"])]
    if not valid_logs: return {"text": "No Alarm", "severity": "normal"}
    latest_log = valid_logs[-1]
    match = re.search(r'%%[\d]*(.*?)/(.*?):(.*)', latest_log)
    if match:
        module = match.group(1).strip()
        detail = match.group(3).strip()
        if "MSTP" in module:
            port_match = re.search(r'([a-zA-Z]+[0-9/]+)', detail)
            port_info = port_match.group(1) if port_match else "UnknownPort"
            state_match = re.search(r'(discarding|forwarding|learning|backup)', detail, re.IGNORECASE)
            state_info = state_match.group(1) if state_match else "StateChange"
            return {"text": f"MSTP: {port_info} -> {state_info.upper()}", "severity": "alarm" if state_info.lower() == "discarding" else "info"}
        if "IFNET" in module:
            port_match = re.search(r'([a-zA-Z]+[0-9/]+)', detail)
            port_info = port_match.group(1) if port_match else "Port"
            return {"text": f"IFNET: {port_info} is {'UP' if 'UP' in detail.upper() else 'DOWN'}", "severity": "info"}
        return {"text": f"{module}: {detail}", "severity": "alarm"}
    return {"text": latest_log[:40], "severity": "alarm"}

def inspect_single_device(dev):
    ip = dev["ip"]
    print(f"{YELLOW}[+] 线程启动：开始异步连接 {ip} ...{RESET}")
    result = {"ip": ip, "sysname": "未知设备", "uptime": "连接失败", "cpu": "--", "mem": "--", "alarm": "无法获取", "status": "异常", "success": False}
    try:
        connection = ConnectHandler(**dev)
        prompt = connection.find_prompt()
        sysname_val = prompt.strip("<>[] ") or "Unknown-SW"
        
        ver_res = connection.send_command("display version", expect_string=r'.*[>\]]')
        uptime_match = re.search(r"uptime is ([^\n\r]+)", ver_res, re.IGNORECASE)
        uptime_val = uptime_match.group(1).replace("weeks", "w").replace("days", "d").replace("hours", "h").replace("minutes", "m").strip() if uptime_match else "未知"
        
        cpu_res = connection.send_command("display cpu-usage", expect_string=r'.*[>\]]')
        cpu_match = re.search(r"CPU [Uu]sage\s*:\s*(\d+%)", cpu_res)
        cpu_val = cpu_match.group(1) if cpu_match else "未知"
        
        mem_res = connection.send_command("display memory-usage", expect_string=r'.*[>\]]')
        mem_match = re.search(r"Memory\s+(?:Used\s+Percent|utilization|Usage)\s*:\s*(\d+%)", mem_res, re.IGNORECASE)
        mem_val = mem_match.group(1) if mem_match else "未知"
        
        log_res = connection.send_command("display logbuffer level warning size 3", expect_string=r'.*[>\]]')
        alarm_data = parse_latest_alarm(log_res)
        alarm_val, severity = alarm_data["text"], alarm_data["severity"]

        status = "正常"
        try:
            if cpu_val != "未知" and int(cpu_val.replace("%", "")) > 80: status = "异常"
            elif mem_val != "未知" and int(mem_val.replace("%", "")) > 85: status = "异常"
        except: pass
        if status != "异常":
            status = "异常" if severity == "alarm" else ("提示" if severity == "info" else "正常")

        config_res = connection.send_command("display current-configuration", expect_string=r'.*[>\]]', read_timeout=30)
        
        # 强制使用绝对路径备份配置
        backup_file = os.path.join(BACKUP_DIR, f"{ip}_{today}_config.cfg")
        with open(backup_file, "w", encoding="utf-8") as f_cfg: 
            f_cfg.write(config_res)
            
        print(f"{GREEN}[√] {ip} ({sysname_val}) 巡检完成，评定状态为: {status}！{RESET}")
        connection.disconnect()
        result.update({"sysname": sysname_val, "uptime": uptime_val, "cpu": cpu_val, "mem": mem_val, "alarm": alarm_val, "status": status, "success": True})
    except Exception as e:
        print(f"{RED}[×] {ip} 处理失败！原因: {e}{RESET}")
    return result

# 执行巡检
print(f"{BLUE}>>> 自动化并发巡检开始...{RESET}")
results = []
with ThreadPoolExecutor(max_workers=min(10, len(DEVICES))) as executor:
    future_to_device = {executor.submit(inspect_single_device, dev): dev for dev in DEVICES}
    for future in as_completed(future_to_device): 
        results.append(future.result())

ip_order = {dev["ip"]: i for i, dev in enumerate(DEVICES)}
results.sort(key=lambda x: ip_order.get(x["ip"], 999))

# 数据提取与转换
columns = ["设备IP", "设备名称", "运行时间", "CPU利用率", "内存利用率", "最新关键告警", "状态"]
table_data = [[r["ip"], r["sysname"], r["uptime"], r["cpu"], r["mem"], r["alarm"], r["status"]] for r in results]

table = PrettyTable(columns)
table.set_style(TableStyle.DOUBLE_BORDER)
table.max_width["最新关键告警"] = 45
for row in table_data: 
    table.add_row(row)

# ================= 严格管控：只往目标文件夹内写文件 =================

# 1. 写入临时文本文件 (d:\sshAuto\tempFlies\temp_table.txt)
clean_table_string = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-9?]*[mK])').sub('', table.get_string())
target_txt_path = os.path.join(TEMP_DIR, "temp_table.txt")
with open(target_txt_path, "w", encoding="utf-8") as f_temp:
    f_temp.write(clean_table_string)

# 2. 写入全局汇总CSV (d:\sshAuto\reports_csv\temp_report.csv)
target_csv_path = os.path.join(CSV_DIR, "temp_report.csv")
with open(target_csv_path, "w", encoding="utf-8-sig", newline="") as f_csv:
    writer = csv.writer(f_csv)
    writer.writerow(columns)
    for row in table_data:
        writer.writerow(row)

print(f"{GREEN}[√] 数据已精准投放至指定资产文件夹，未污染根目录！{RESET}")