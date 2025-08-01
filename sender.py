import smtplib
import pandas as pd
import re
import os
import time
import csv
import random
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication
from email.utils import formataddr
from email.mime.image import MIMEImage
from datetime import datetime

current_dir = os.path.dirname(os.path.abspath(__file__))

# Gửi email từ danh sách người gửi theo thứ tự ngẫu nhiên
senders = []
sender_index = 0

def load_config(config_path):
    if not os.path.exists(config_path):
        print(f"❌ Không tìm thấy file cấu hình: {config_path}")
        return None

    config = configparser.ConfigParser()
    config.read(config_path)
    return config

def check_required_files(config):
    required_files = {
        '📄 File danh sách người nhận': config['FILES']['recipients_excel'],
        '🖼 Logo chèn trong nội dung': config['FILES']['logo_path'],
        '🧾 File mẫu nội dung HTML': config['FILES']['email_template'],
    }

    print("🔍 Kiểm tra các file cần thiết:")
    all_ok = True

    for desc, path in required_files.items():
        full_path = os.path.join(current_dir, path)
        if os.path.exists(full_path):
            print(f"✅ {desc}: Tìm thấy ({full_path})")
        else:
            print(f"❌ {desc}: KHÔNG tìm thấy! ({full_path})")
            all_ok = False

    # Xử lý đặc biệt cho file đính kèm PDF có thể chứa nhiều file
    attachment_list = config.get("FILES", "attachment_pdf", fallback="")
    attachments = [f.strip() for f in attachment_list.split(",") if f.strip()]
    for file_path in attachments:
        full_path = os.path.join(current_dir, file_path)
        if os.path.exists(full_path):
            print(f"✅ 📎 File đính kèm PDF: Tìm thấy ({full_path})")
        else:
            print(f"❌ 📎 File đính kèm PDF: KHÔNG tìm thấy! ({full_path})")
            all_ok = False

    return all_ok

# Kiểm tra định dạng email
def is_valid_email(email):
    if pd.isna(email) or str(email).strip() == "":
        return False, "Địa chỉ email bị bỏ trống."

    email = str(email).strip()
    regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    if not re.match(regex, email):
        return False, f"Địa chỉ email sai định dạng: {email}"

    return True, ""

# Kiểm tra thông tin cổ đông
def is_valid_shareholder_info(hoten, maso):
    if pd.isna(hoten) or str(hoten).strip() == "":
        return False
    if pd.isna(maso) or str(maso).strip() == "":
        return False
    return True

def get_next_sender():
    global sender_index
    sender = senders[sender_index]
    sender_index = (sender_index + 1) % len(senders)
    return sender

# Ghi log với timestamp
def write_log(log_message):
    log_path = os.path.join(current_dir, "log.csv")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    file_exists = os.path.isfile(log_path)
    with open(log_path, mode="a", encoding="utf-8-sig", newline="") as log_file:
        writer = csv.writer(log_file)
        if not file_exists:
            writer.writerow(["Thời gian", "Nội dung"])
        writer.writerow([timestamp, log_message])

def send_email(sender_email, recipient_email, full_name, shareholder_id, t_holding, config):
    smtp_server = config["SMTP"]["server"]
    smtp_port = int(config["SMTP"]["port"])
    password = config["SMTP"]["password"]

    # Tạo message chính kiểu multipart/related
    msg_root = MIMEMultipart("related")
    msg_root["Subject"] = f"HAGL Group. Notice to - {full_name}"
    msg_root["From"] = formataddr(("HAGL Group", sender_email))
    msg_root["To"] = recipient_email
    msg_root["Reply-To"] = "daihoicodong@hagl.com.vn"
    msg_root.preamble = "This is a multi-part message in MIME format."

    # Tạo phần nội dung HTML lồng bên trong multipart/alternative
    msg_alternative = MIMEMultipart("alternative")
    msg_root.attach(msg_alternative)

    try:
        # Đọc nội dung HTML và thay thế các biến
        with open(os.path.join(current_dir, config["FILES"]["email_template"]), "r", encoding="utf-8") as file:
            html_content = file.read()
            html_content = html_content.replace("{ho_ten}", str(full_name))
            html_content = html_content.replace("{tt_dksh}", str(shareholder_id))
            # html_content = html_content.replace("{so_cp}", str(t_holding))
            formatted_holding = "{:,}".format(int(float(t_holding))).replace(",", ".")
            html_content = html_content.replace("{so_cp}", formatted_holding)
        # Đính nội dung HTML vào alternative
        msg_alternative.attach(MIMEText(html_content, "html"))

        # Gắn ảnh inner_img (ảnh QR code)
        inner_img = config["FILES"]["inner_img"]
        inner_img_path = os.path.join(current_dir, inner_img)
        with open(inner_img_path, "rb") as img_file:
            image = MIMEImage(img_file.read())
            image.add_header("Content-ID", "<inner_image>")
            image.add_header("Content-Disposition", "inline", filename=inner_img)
            msg_root.attach(image)

        # Gắn logo công ty
        logo_filename = config["FILES"]["logo_path"]
        logo_path = os.path.join(current_dir, logo_filename)
        with open(logo_path, "rb") as logo_file:
            logo = MIMEImage(logo_file.read())
            logo.add_header("Content-ID", "<company_logo>")
            logo.add_header("Content-Disposition", "inline", filename=logo_filename)
            msg_root.attach(logo)

    except FileNotFoundError as e:
        print(f"❌ Không tìm thấy file: {e}")
        write_log(f"❌ Không tìm thấy file: {e}")
        return

    # Đính kèm các file PDF nếu có
    try:
        attachments = config.get("FILES", "attachment_pdf", fallback="")
        attachment_paths = [f.strip() for f in attachments.split(",") if f.strip()]
        for file_path in attachment_paths:
            full_path = os.path.join(current_dir, file_path)
            with open(full_path, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="pdf")
                part.add_header("Content-Disposition", "attachment", filename=os.path.basename(file_path))
                msg_root.attach(part)
    except FileNotFoundError as e:
        print(f"❌ Không tìm thấy file PDF đính kèm: {e}")
        write_log(f"❌ Không tìm thấy file PDF đính kèm: {e}")
        return

    # Gửi email
    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()

        server.login(sender_email, password)
        server.sendmail(sender_email, recipient_email, msg_root.as_string())
        print(f"✅ [{sender_email}] Gửi đến {recipient_email} thành công!")
        write_log(f"✅ [{sender_email}] Gửi đến {recipient_email} thành công")

    except smtplib.SMTPRecipientsRefused:
        print(f"⚠️ Từ chối địa chỉ email: {recipient_email}")
        write_log(f"⚠️ Từ chối địa chỉ email: {recipient_email}")
    except smtplib.SMTPException as e:
        print(f"❌ SMTP lỗi với {recipient_email}: {e}")
        write_log(f"❌ SMTP lỗi với {recipient_email}: {e}")
    except Exception as e:
        print(f"⚠️ Lỗi khác với {recipient_email}: {e}")
        write_log(f"⚠️ Lỗi khác với {recipient_email}: {e}")
    finally:
        if 'server' in locals():
            server.quit()


def main():
    global senders
    config_file = os.path.join(current_dir, "sender.conf")
    config = load_config(config_file)
    if not config:
        return

    if not check_required_files(config):
        return

    senders = [email.strip() for email in config["SENDER"]["emails"].split(",")]
    if not senders:
        print("❌ Không tìm thấy email người gửi nào trong cấu hình.")
        return
    random.shuffle(senders)

    recipients_file = os.path.join(current_dir, config["FILES"]["recipients_excel"])
    try:
        df = pd.read_excel(recipients_file)
    except Exception as e:
        print(f"❌ Lỗi đọc file Excel người nhận: {e}")
        return
    
    required_columns = ["Email", "HoTen", "MaSoCoDong"]
    if not all(col in df.columns for col in required_columns):
        print(f"❌ Thiếu cột cần thiết: {', '.join(required_columns)}")
        return

    confirm = input("Đã đủ điều kiện gửi thư, Bạn có muốn gửi email không? (y/n): ").strip().lower()
    if confirm != "y":
        print("🛑 Đã huỷ.")
        return

    try:
        start_row = int(input("📌 Bạn muốn bắt đầu gửi từ dòng thứ mấy? (2 là dòng đầu tiên): ").strip())
        if start_row < 2 or start_row > len(df):
            print("❌ Dòng bắt đầu không hợp lệ. Mặc định bắt đầu từ dòng 2.")
            start_row = 2
    except ValueError:
        print("❌ Giá trị không hợp lệ. Mặc định bắt đầu từ dòng 2.")
        start_row = 2

    sent_count = 0
    print("🚀 Bắt đầu gửi email... Nhấn Ctrl+C để dừng lại an toàn.\n")
    try:
        for index, row in df.iloc[start_row - 2:].iterrows():
            # Kiểm tra địa chi email
            is_valid, error_msg = is_valid_email(row["Email"])
            if not is_valid:
                print(f"❌ Dòng {index + 2}: {error_msg}")
                write_log(f"❌ Dòng {index + 2}: {error_msg}")
                continue

            # Kiểm tra thông tin cổ đông
            if not is_valid_shareholder_info(row["HoTen"], row["MaSoCoDong"]):
                print(f"❌ Dòng {index + 2}: Không có đủ thông tin cổ đông.")
                write_log(f"❌ Dòng {index + 2}: Không có đủ thông tin cổ đông.")
                continue

            # OK hết rồi, gửi đi thôi
            email = str(row["Email"]).strip()
            sender_email = get_next_sender()
            send_email(sender_email, email, row["HoTen"], row["MaSoCoDong"], row["SoCP"], config)
            sent_count += 1
            time.sleep(10)

    except KeyboardInterrupt:
        print("\n🛑 Đã dừng gửi theo yêu cầu người dùng (Ctrl+C).")
        write_log("🛑 Đã dừng gửi theo yêu cầu người dùng (Ctrl+C).")

    print(f"\n✅ Đã gửi thành công {sent_count} email.")

if __name__ == "__main__":
    main()
