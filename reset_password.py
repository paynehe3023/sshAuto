# -*- coding: utf-8 -*-
import os
import sys
import getpass
from passlib.hash import bcrypt

def reset_password():
    # 1. 自动定位项目根目录下的 .env 文件
    base_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(base_dir, ".env")

    if not os.path.exists(env_path):
        print(f"❌ 错误：找不到 .env 文件，路径：{env_path}")
        return

    print("========================================")
    print("🔒 网管系统管理员密码重置工具")
    print("========================================\n")

    # 2. 交互式输入新密码
    try:
        new_pass = getpass.getpass("请输入新的管理员密码: ").strip()
        if not new_pass:
            print("❌ 密码不能为空！")
            return
            
        confirm_pass = getpass.getpass("请再次输入新密码: ").strip()
        if new_pass != confirm_pass:
            print("❌ 两次输入的密码不一致，操作已取消！")
            return
    except Exception as e:
        print(f"❌ 读取输入失败: {e}")
        return

    # 3. 生成新的 bcrypt 哈希
    print("\n⏳ 正在生成安全的 bcrypt 哈希值...")
    new_hash = bcrypt.hash(new_pass)

    # 4. 读取并更新 .env 文件内容
    with open(env_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    new_lines = []
    has_pass = False
    has_hash = False

    for line in lines:
        if line.startswith("ADMIN_PASS="):
            new_lines.append(f"ADMIN_PASS={new_pass}\n")
            has_pass = True
        elif line.startswith("ADMIN_PASS_HASH="):
            new_lines.append(f"ADMIN_PASS_HASH={new_hash}\n")
            has_hash = True
        else:
            new_lines.append(line)

    # 如果 .env 里原本没有这两个键，追加到末尾
    if not has_pass:
        new_lines.append(f"ADMIN_PASS={new_pass}\n")
    if not has_hash:
        new_lines.append(f"ADMIN_PASS_HASH={new_hash}\n")

    # 5. 写回 .env 文件
    with open(env_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)

    print("\n✅ 密码重置成功！")
    print(f"🔑 明文密码已更新为: {new_pass}")
    print(f"🔐 新 Hash 值已存入 .env: {new_hash}")
    print("\n💡 提示：修改已立即生效（若未开启热加载，重启 agent_server 即可）。")

if __name__ == "__main__":
    reset_password()