import csv
import sys
import os
import time
import datetime
import platform
import subprocess
import smtplib
from email.message import EmailMessage
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor

# === Load Email Config from XML ===
def load_email_config(xml_path="email_config.xml"):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        return {
            "send_if_error": root.find("send_if_error").text.strip().lower() == "true",
            "sender": root.find("sender").text.strip(),
            "password": root.find("password").text.strip(),
            "recipient": root.find("recipient").text.strip(),
            "smtp_server": root.find("smtp_server").text.strip(),
            "smtp_port": int(root.find("smtp_port").text.strip()),
        }
    except Exception as e:
        print(f"[ERROR] Failed to read email_config.xml: {e}")
        return None

EMAIL_CONFIG = load_email_config()
if not EMAIL_CONFIG:
    print("[ERROR] Email configuration missing or invalid. Exiting.")
    sys.exit(1)

# === Logging Setup ===
log_entries = []

def write_log(message):
    timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    line = f"{timestamp} {message}"
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode("utf-8", errors="replace").decode("cp1252"))

# === Ping Check ===
def is_reachable(ip):
    try:
        param = "-n" if platform.system().lower() == "windows" else "-c"
        subprocess.check_output(["ping", param, "1", ip], stderr=subprocess.STDOUT, timeout=3)
        return True
    except Exception:
        return False

# === Load IP Phones from CSV ===
def read_ipphones_from_csv(file_path="devices.csv"):
    devices = []
    try:
        with open(file_path, mode='r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                devices.append({
                    "ip": row["ip"].strip(),
                    "user": row["user"].strip(),
                    "pass": row["password"].strip(),
                    "name": row["name"].strip()
                })
        return devices
    except Exception as e:
        write_log(f"[ERROR] Failed to read devices.csv: {e}")
        return []

# === Headless Chrome Setup ===
def create_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--log-level=3")
    return webdriver.Chrome(options=options)

# === Extract Status Info ===
def get_ipphone_status(driver, ip, user, passwd, name):
    uptime = callstate = "N/A"
    url = f"http://{user}:{passwd}@{ip}/DI_S_.xml"
    try:
        driver.get(url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, "html.parser")

        uptime_td = soup.find("td", string="UpTime")
        if uptime_td:
            uptime_value_td = uptime_td.find_next_sibling("td")
            uptime = uptime_value_td.text.strip() if uptime_value_td else "N/A"

        for section in soup.find_all("div", class_="title"):
            if "Service Status" in section.get_text(strip=True):
                table = section.find_next_sibling("table")
                callstate_td = table.find("td", string="CallState") if table else None
                if callstate_td:
                    val_td = callstate_td.find_next_sibling("td")
                    callstate = val_td.text.strip() if val_td else "N/A"
                break

    except Exception as e:
        write_log(f"[ERROR] Error reading DI_S_.xml from {ip} ({name}): {e}")
        return "N/A", "N/A"

    write_log(f"[INFO] {ip} ({name}) - UpTime: {uptime}, CallState: {callstate}")
    return uptime, callstate

# === Reboot Function ===
def reboot_ipphone(device):
    ip, name = device["ip"], device["name"]

    # 1️⃣ Ping check first
    if not is_reachable(ip):
        write_log(f"[ERROR] {ip} ({name}) is not reachable (ping failed). Skipping reboot.")
        log_entries.append((ip, name, "N/A", "Unreachable"))
        return

    # 2️⃣ Only fetch UpTime and CallState if ping succeeds
    driver = create_driver()
    uptime, callstate = get_ipphone_status(driver, ip, device["user"], device["pass"], name)
    log_entries.append((ip, name, uptime, callstate))

    try:
        if callstate != "0 Active Calls":
            write_log(f"[SKIPPED] Reboot for {ip} ({name}) - CallState not idle: {callstate}")
            return

        reboot_url = f"http://{device['user']}:{device['pass']}@{ip}/rebootgetconfig.htm"
        driver.get(reboot_url)
        write_log(f"[SUCCESS] Reboot triggered for {ip} ({name})")
    except Exception as e:
        write_log(f"[ERROR] Error rebooting {ip} ({name}): {e}")
    finally:
        try:
            driver.quit()
        except:
            pass

# === Build HTML Report ===
def build_html_report():
    html = """
    <html><body>
    <h3>IP Phone Reboot Status Report</h3>
    <table border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; font-family: Arial;">
        <tr style="background-color:#f2f2f2;">
            <th>S.No</th><th>IP</th><th>Name</th><th>UpTime</th><th>CallState</th><th>Status</th>
        </tr>
    """
    for idx, (ip, name, uptime, callstate) in enumerate(log_entries, start=1):
        if callstate == "Unreachable":
            row_style = ' style="background-color: #ff9999;"'
            status = "Unreachable"
        elif callstate == "N/A":
            row_style = ' style="background-color: #ffcccc;"'
            status = "Error"
        elif callstate != "0 Active Calls":
            row_style = ' style="background-color: #fff8b3;"'
            status = "Skipped"
        else:
            row_style = ' style="background-color: #ccffcc;"'
            status = "Rebooted"

        html += f"<tr{row_style}><td>{idx}</td><td>{ip}</td><td>{name}</td><td>{uptime}</td><td>{callstate}</td><td>{status}</td></tr>"

    html += "</table></body></html>"
    return html

# === Save Daily HTML Log ===
def save_daily_html_log():
    if not os.path.exists("logs"):
        os.makedirs("logs")

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    log_filename = os.path.join("logs", f"{today}_ipphone_report.html")
    try:
        with open(log_filename, "w", encoding="utf-8") as f:
            f.write(build_html_report())
        write_log(f"[INFO] Daily HTML log saved to {log_filename}")
    except Exception as e:
        write_log(f"[ERROR] Could not save daily HTML log: {e}")

# === Dynamic Email Subject ===
def get_email_subject():
    for ip, name, uptime, callstate in log_entries:
        if callstate in ["N/A", "Unreachable"]:
            return "IP Phone Reboot Report – Errors Found ⚠"
    return "IP Phone Reboot Report – All OK ✔"

# === Send Email ===
def send_error_email(retries=2, delay=300):
    html = build_html_report()
    subject = get_email_subject()

    for attempt in range(1, retries+1):
        try:
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = EMAIL_CONFIG["sender"]
            msg["To"] = EMAIL_CONFIG["recipient"]
            msg.set_content("This email contains an HTML table with IP Phone status info.")
            msg.add_alternative(html, subtype="html")

            with smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"]) as smtp:
                smtp.starttls()
                smtp.login(EMAIL_CONFIG["sender"], EMAIL_CONFIG["password"])
                smtp.send_message(msg)
            write_log(f"[SUCCESS] Email sent successfully on attempt {attempt}")
            break
        except Exception as e:
            write_log(f"[ERROR] Email sending failed (attempt {attempt}): {e}")
            if attempt < retries:
                write_log(f"[INFO] Retrying in {delay//60} minutes...")
                time.sleep(delay)

# === Main Execution ===
ip_phones = read_ipphones_from_csv()
write_log("[INFO] Starting IP Phone reboot process...")

with ThreadPoolExecutor(max_workers=len(ip_phones)) as executor:
    executor.map(reboot_ipphone, ip_phones)

write_log("[INFO] All IP Phone reboots attempted.")

# Save daily HTML log
save_daily_html_log()

# Send email if enabled
if EMAIL_CONFIG["send_if_error"]:
    send_error_email()
