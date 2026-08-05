import csv
import os
from cryptography.fernet import Fernet

def main():
    # 1. 设定路径
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    CONFIG_DIR = os.path.join(BASE_DIR, "config")
    os.makedirs(CONFIG_DIR, exist_ok=True) # 自动创建 config 文件夹

    # 2. 生成一把新钥匙（Key）
    key = Fernet.generate_key()
    print("==================================================")
    print(f"请将以下密钥配置到系统环境变量 'NVR_SECRET_KEY' 中:\n\n{key.decode()}\n")
    print("==================================================")

    cipher = Fernet(key)

    # 3. 输入输出路径
    input_csv = "NVRdevices.csv"  # 假设原始文件还在根目录
    output_csv = os.path.join(CONFIG_DIR, "NVRdevices_secure.csv") # 存入 config 文件夹

    try:
        with open(input_csv, 'r', encoding='utf-8-sig') as f_in, \
             open(output_csv, 'w', encoding='utf-8-sig', newline='') as f_out:
            
            reader = csv.DictReader(f_in)
            writer = csv.DictWriter(f_out, fieldnames=reader.fieldnames)
            writer.writeheader()

            for row in reader:
                # 只加密密码字段
                plain_pwd = row['password'].encode('utf-8')
                encrypted_pwd = cipher.encrypt(plain_pwd).decode('utf-8')
                row['password'] = encrypted_pwd
                writer.writerow(row)

        print(f"✅ 加密完成！密文已安全存放到: {output_csv}")
        print(f"⚠️ 请务必在测试无误后，删除项目根目录下的 {input_csv}。")
        
    except FileNotFoundError:
        print(f"❌ 找不到原始文件: {input_csv}，请确认它是否在项目根目录下。")

if __name__ == "__main__":
    main()