# -*- coding: utf-8 -*-
import os
import re
import csv
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from netmiko import ConnectHandler
import networkx as nx
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']  
plt.rcParams['axes.unicode_minus'] = False    

DEVICES = []
csv_file_path = "SWdevices.csv"

if os.path.exists(csv_file_path):
    with open(csv_file_path, "r", encoding="utf-8-sig", errors="ignore") as f:
        first_line = f.readline()
        f.seek(0)
        reader = csv.DictReader(f, delimiter=';' if ';' in first_line else ',')
        for row in reader:
            clean_row = {k.strip().lower(): v.strip() for k, v in row.items() if k and v}
            ip = clean_row.get("ip")
            username = clean_row.get("username") or clean_row.get("user")
            password = clean_row.get("password") or clean_row.get("pass")
            dev_type = clean_row.get("device_type") or clean_row.get("type") or "huawei"
            if ip and username and password:
                DEVICES.append({
                    "device_type": dev_type, "ip": ip, "username": username, "password": password, "port": 22
                })
else:
    print("❌ 未找到 SWdevices.csv 文件！")
    exit()

def get_lldp_neighbors(dev):
    ip = dev["ip"]
    links = []
    hostname = ip
    try:
        connection = ConnectHandler(**dev)
        prompt = connection.find_prompt()
        hostname = prompt.strip("<>[] ")
        print(f"⚡ 正在获取 {hostname} ({ip}) 的 LLDP 拓扑关系...")
        lldp_res = connection.send_command("display lldp neighbor brief", expect_string=r'.*[>\]]')
        connection.disconnect()
        
        lines = lldp_res.splitlines()
        for line in lines:
            line = line.strip()
            if not line or "Local Intf" in line or "---" in line or "Neighbor" in line:
                continue
            parts = re.split(r'\s+', line)
            if len(parts) >= 3:
                links.append({
                    "local_device": hostname, "local_port": parts[0],
                    "remote_device": parts[1], "remote_port": parts[2]
                })
    except Exception as e:
        print(f"❌ 无法连接到设备 {ip}: {e}")
    return hostname, links

all_links = []
device_names = {}

print(">>> 开始多线程分析网络拓扑...")
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = {executor.submit(get_lldp_neighbors, dev): dev for dev in DEVICES}
    for future in as_completed(futures):
        host, links = future.result()
        device_names[futures[future]["ip"]] = host
        all_links.extend(links)

G = nx.Graph()
for link in all_links:
    u = link["local_device"]
    v = link["remote_device"]
    edge_label = f"{link['local_port']} ⇄ {link['remote_port']}"
    G.add_edge(u, v, label=edge_label)

for ip, host in device_names.items():
    if host not in G: G.add_node(host)

plt.figure(figsize=(10, 8))
pos = nx.spring_layout(G, k=1.0, iterations=50)
nx.draw_networkx_nodes(G, pos, node_size=2000, node_color="#3498db", edgecolors="#2980b9")
nx.draw_networkx_edges(G, pos, width=2, edge_color="#bdc3c7")
nx.draw_networkx_labels(G, pos, font_size=10, font_weight="bold", font_color="#ffffff")

edge_labels = nx.get_edge_attributes(G, 'label')
nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=7, font_color="#e74c3c")
plt.title("企业内网物理拓扑自动发现图 (基于LLDP)", fontsize=14, fontweight="bold")
plt.axis("off")

# 💡 核心修复 2：统一存储位置到 reports_imgs
IMG_DIR = r"d:\sshAuto\reports_imgs"
os.makedirs(IMG_DIR, exist_ok=True)
output_image = os.path.join(IMG_DIR, "Topology.png")

plt.savefig(output_image, format="PNG", dpi=300, bbox_inches="tight")
plt.close()
print(f"\n🎉最新的真实物理拓扑图已生成并保存至：{output_image}")