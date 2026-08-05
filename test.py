# -*- coding: utf-8 -*-
"""
路径: d:\sshAuto\test.py
"""
import sys
import os
import requests
from dotenv import load_dotenv

# 加载环境变量
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(BASE_DIR, ".env")
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)

FEISHU_APP_ID = os.getenv("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET")
DEFAULT_RECEIVER = os.getenv("TARGET_RECEIVER", "oc_b2cb8c8d58c96a96f2ba65a739c88166")

def get_feishu_tenant_access_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    try:
        res = requests.post(url, json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}, timeout=10)
        return res.json().get("tenant_access_token")
    except Exception:
        return None

def send_feishu_message(text_content, receive_id):
    token = get_feishu_tenant_access_token()
    if not token: return
    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"}
    escaped_text = text_content.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')
    payload = {"receive_id": receive_id, "msg_type": "text", "content": f'{{"text": "{escaped_text}"}}'}
    requests.post(url, headers=headers, json=payload, timeout=10)

if __name__ == "__main__":
    # 接收来自 agent_server 传过来的当前 chat_id 
    target_chat_id = DEFAULT_RECEIVER
    if len(sys.argv) > 1 and sys.argv[1].strip():
        target_chat_id = sys.argv[1].strip()

    send_feishu_message("🟢 报告！独立测试脚本（test.py）已被主服务成功拉起，运行环境一切正常！", target_chat_id)
