import os
import re
import csv
import sys
import time
import random
import smtplib
import pandas as pd
import configparser
from datetime import datetime
from email.utils import formataddr
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#current_dir = os.path.dirname(os.path.abspath(__file__))

senders = []
sender_index = 0


# ===================== COMMON HELPERS =====================

def log(message):
    """Ghi log ra file CSV."""
    log_path = os.path.join(current_dir, "log.csv")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_file = not os.path.exists(log_path)
    with open(log_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["Thời gian", "Nội dung"])
        writer.writerow([timestamp, message])


def load_config(config_dir):
    #config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_filename)
    config_path = os.path.join(config_dir, "sender.conf")

    if not os.path.exists(config_path):
        print(f"❌ Không tìm thấy file cấu hình: {config_path}")
        return None
    
    cfg = configparser.ConfigParser()
    cfg.read(config_path)
    return cfg


def file_exists(path):
    return os.path.exists(os.path.join(current_dir, path))


# ===================== VALIDATIONS =====================

def check_required_files(config):
    print("🔍 Kiểm tra các file cần thiết:")
    all_ok = True

    required = {
        "📄 File danh sách người nhận": config["FILES"]["recipients_excel"],
        "🖼 Logo chèn trong nội dung": config["FILES"]["logo_path"],
        "🧾 File mẫu nội dung HTML": config["FILES"]["email_template"],
    }

    # Kiểm tra file bắt buộc
    for desc, file_path in required.items():
        full = os.path.join(current_dir, file_path)
        if os.path.exists(full):
            print(f"✅ {desc}: Tìm thấy ({full})")
        else:
            print(f"❌ {desc}: KHÔNG tìm thấy! ({full})")
            all_ok = False

    # PDF – nếu có
    pdf_list = config.get("FILES", "attachment_pdf", fallback="").strip()
    if pdf_list:
        for f in pdf_list.split(","):
            f = f.strip()
            full = os.path.join(current_dir, f)
            if os.path.exists(full):
                print(f"✅ 📎 File đính kèm PDF: Tìm thấy ({full})")
            else:
                print(f"❌ 📎 File đính kèm PDF: KHÔNG tìm thấy! ({full})")
                all_ok = False
    else:
        print("ℹ️ Không khai báo file PDF đính kèm — bỏ qua.")

    return all_ok


def validate_email(email):
    if pd.isna(email) or not str(email).strip():
        return False, "Email bị bỏ trống."
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    if not re.match(pattern, str(email).strip()):
        return False, f"Sai định dạng email: {email}"
    return True, ""


def validate_shareholder(row, row_index):
    if pd.isna(row["HoTen"]) or pd.isna(row["MaSoCoDong"]):
        return False, f"❌ Dòng {row_index}: Thiếu thông tin cổ đông."
    return True, ""


# ===================== EMAIL BUILDING =====================

def attach_image(msg, path, cid):
    """Gắn ảnh inline nếu tồn tại."""
    if not path:
        print(f"ℹ️ Không có {cid} đính kèm — bỏ qua.")
        return

    full = os.path.join(current_dir, path)

    if file_exists(path):
        with open(full, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline", filename=path)
            msg.attach(img)
    else:
        print(f"⚠️ File ảnh không tồn tại: {full}")
        log(f"⚠️ File ảnh không tồn tại: {full}")


def attach_pdfs(msg, config):
    """Đính kèm tất cả PDF."""
    pdfs = config.get("FILES", "attachment_pdf", fallback="").strip()
    if not pdfs:
        print("ℹ️ Không có file PDF đính kèm — bỏ qua.")
        return

    for file_path in [p.strip() for p in pdfs.split(",") if p.strip()]:
        full = os.path.join(current_dir, file_path)
        if os.path.exists(full):
            with open(full, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="pdf")
                part.add_header("Content-Disposition", "attachment", filename=os.path.basename(file_path))
                msg.attach(part)
        else:
            print(f"⚠️ File PDF không tồn tại: {full}")
            log(f"⚠️ File PDF không tồn tại: {full}")


# ===================== SEND EMAIL =====================

def send_email(sender, recipient, name, code, holding, config):
    smtp_server = config["SMTP"]["server"]
    smtp_port = int(config["SMTP"]["port"])
    password = config["SMTP"]["password"]

    # Tạo message
    msg = MIMEMultipart("related")
    msg["Subject"] = f"HAGL Group. Notice to - {name}"
    msg["From"] = formataddr(("HAGL Group", sender))
    msg["To"] = recipient
    msg["Reply-To"] = "daihoicodong@hagl.com.vn"

    # Nội dung HTML
    alt = MIMEMultipart("alternative")
    msg.attach(alt)

    try:
        template_path = os.path.join(current_dir, config["FILES"]["email_template"])
        with open(template_path, "r", encoding="utf-8") as f:
            html = f.read()

        html = html.replace("{ho_ten}", str(name))
        html = html.replace("{tt_dksh}", str(code))
        html = html.replace("{so_cp}", "{:,}".format(int(float(holding))).replace(",", "."))

        alt.attach(MIMEText(html, "html"))

        # Logo luôn bắt buộc
        attach_image(msg, config["FILES"]["logo_path"], "company_logo")

        # Ảnh QR (tùy chọn)
        attach_image(msg, config["FILES"].get("inner_img", "").strip(), "inner_image")

        # PDF (tùy chọn)
        attach_pdfs(msg, config)

    except Exception as e:
        print(f"❌ Lỗi đọc template hoặc file đính kèm: {e}")
        log(f"❌ Lỗi email build: {e}")
        return

    # Gửi
    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()

        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())
        print(f"✅ [{sender}] → {recipient}")
        log(f"Sent OK: {sender} → {recipient}")

    except Exception as e:
        print(f"❌ SMTP lỗi: {e}")
        log(f"❌ SMTP lỗi gửi đến {recipient}: {e}")
    finally:
        try:
            server.quit()
        except:
            pass


# ===================== MAIN =====================

def get_next_sender():
    global sender_index
    s = senders[sender_index]
    sender_index = (sender_index + 1) % len(senders)
    return s


def main():
    global senders
    global current_dir
    if getattr(sys, 'frozen', False):   # nếu đang chạy trong .exe
        current_dir = os.path.dirname(sys.executable)
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    # Load config
    config = load_config(current_dir)
    if not config:
        return

    # Kiểm tra file
    if not check_required_files(config):
        return

    # Load danh sách email người gửi
    senders = [e.strip() for e in config["SENDER"]["emails"].split(",") if e.strip()]
    if not senders:
        print("❌ Không tìm thấy email người gửi.")
        return
    random.shuffle(senders)

    # Load Excel người nhận
    try:
        df = pd.read_excel(
            os.path.join(current_dir, config["FILES"]["recipients_excel"]),
            dtype={"MaSoCoDong": str}
        )
    except Exception as e:
        print(f"❌ Lỗi đọc file Excel: {e}")
        return

    # Kiểm tra cột
    required_cols = ["Email", "HoTen", "MaSoCoDong", "SoCP"]
    if not all(c in df.columns for c in required_cols):
        print("❌ Thiếu các cột bắt buộc:", ", ".join(required_cols))
        return

    # Xác nhận gửi
    if input("Đã đủ điều kiện gửi thư, bạn muốn gửi email không? (y/n): ").lower() != "y":
        print("🛑 Đã hủy.")
        return

    # Nhập dòng bắt đầu
    try:
        start_row = int(input("📌 Bắt đầu từ dòng số mấy? (2 = dòng đầu tiên): ").strip())
        start_row = max(2, min(start_row, len(df)))
    except:
        print("❌ Giá trị không hợp lệ. Mặc định dòng 2.")
        start_row = 2

    print("\n🚀 Bắt đầu gửi email... Nhấn Ctrl+C để dừng.\n")

    sent = 0
    try:
        for idx, row in df.iloc[start_row - 2:].iterrows():
            row_index = idx + 2

            # Validate email
            ok, msg = validate_email(row["Email"])
            if not ok:
                print(f"❌ Dòng {row_index}: {msg}")
                log(msg)
                continue

            # Validate cổ đông
            ok, msg = validate_shareholder(row, row_index)
            if not ok:
                print(msg)
                log(msg)
                continue

            sender = get_next_sender()
            send_email(sender, row["Email"], row["HoTen"], row["MaSoCoDong"], row["SoCP"], config)
            sent += 1
            time.sleep(2)

    except KeyboardInterrupt:
        print("\n🛑 Đã dừng theo yêu cầu.")

    print(f"\n✅ Hoàn tất. Đã gửi {sent} email.")


if __name__ == "__main__":
    main()
