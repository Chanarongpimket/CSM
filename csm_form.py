import customtkinter as ctk
from tkinter import messagebox
import datetime
import time
import sys
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

# ========= Project Mapping =========
project_map = {
    "MA IIG 2567": "206240625",
    "MA Thaipak 2567": "206250016",
    "Blockchain": "201190986"
}

# ========= GUI =========
def start_gui():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    def run_script():
        username = entry_user.get()
        password = entry_pass.get()
        task_note = entry_task_note.get()
        activity = text_activity.get("1.0", "end").strip()
        arrival = entry_arrival.get()
        finish = entry_finish.get()
        depart = entry_depart.get()
        project_code = project_map[project_dropdown.get()]

        if not all([username, password, task_note, activity, arrival, finish, depart]):
            messagebox.showerror("❌ Error", "กรุณากรอกข้อมูลให้ครบถ้วน")
            return

        datetime_fields = {
            "ArrivalDate": arrival,
            "FinishDate": finish,
            "DepartDate": depart
        }

        messagebox.showinfo("⌛", "เริ่มทำงานอัตโนมัติ กรุณาอย่าปิดหน้าต่างนี้...")
        run_automation(username, password, task_note, activity, datetime_fields, project_code)

    app = ctk.CTk()
    app.title("CSM Automation Tool")
    app.geometry("500x700")

    ctk.CTkLabel(app, text="👤 Username").pack(pady=5)
    entry_user = ctk.CTkEntry(app, width=300)
    entry_user.pack()

    ctk.CTkLabel(app, text="🔒 Password").pack(pady=5)
    entry_pass = ctk.CTkEntry(app, width=300, show="*")
    entry_pass.pack()

    ctk.CTkLabel(app, text="📌 Task Note").pack(pady=5)
    entry_task_note = ctk.CTkEntry(app, width=400)
    entry_task_note.insert(0, "INCIDENT IIG&ISP of NT1")
    entry_task_note.pack()

    ctk.CTkLabel(app, text="📝 Activity Note").pack(pady=5)
    text_activity = ctk.CTkTextbox(app, width=400, height=100)
    text_activity.pack()

    today = datetime.datetime.now().strftime("%d/%m/%Y")

    ctk.CTkLabel(app, text="📅 Arrival Date").pack(pady=5)
    entry_arrival = ctk.CTkEntry(app, width=300)
    entry_arrival.insert(0, f"{today} 08:30")
    entry_arrival.pack()

    ctk.CTkLabel(app, text="📅 Finish Date").pack(pady=5)
    entry_finish = ctk.CTkEntry(app, width=300)
    entry_finish.insert(0, f"{today} 17:30")
    entry_finish.pack()

    ctk.CTkLabel(app, text="📅 Depart Date").pack(pady=5)
    entry_depart = ctk.CTkEntry(app, width=300)
    entry_depart.insert(0, f"{today} 18:00")
    entry_depart.pack()

    ctk.CTkLabel(app, text="📁 Project").pack(pady=5)
    project_dropdown = ctk.CTkOptionMenu(app, values=list(project_map.keys()))
    project_dropdown.set("MA IIG 2567")
    project_dropdown.pack()

    ctk.CTkButton(app, text="🚀 Submit & Run", command=run_script).pack(pady=20)

    app.mainloop()

# ========= Selenium Automation =========
def wait_and_click(driver, by, identifier, timeout=10, js_fallback=True):
    try:
        element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, identifier)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(0.3)
        try:
            element.click()
        except ElementClickInterceptedException:
            if js_fallback:
                driver.execute_script("arguments[0].click();", element)
    except TimeoutException:
        print(f"❌ ไม่พบ element {identifier}")

def get_driver_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, "chromedriver.exe")
    return os.path.join(os.path.abspath("."), "chromedriver.exe")

def wait_and_type(driver, by, identifier, text, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, identifier)))
        element.clear()
        element.send_keys(text)
        print(f"✅ กรอก {identifier} แล้ว")
    except TimeoutException:
        print(f"❌ กรอก {identifier} ไม่ได้")

def run_automation(USERNAME, PASSWORD, TASK_NOTE, activity, datetime_fields, project_code):
    chrome_options = Options()
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--disable-logging")
    chrome_options.add_experimental_option("d   etach", True)

    driver = webdriver.Chrome(executable_path=get_driver_path(), options=chrome_options)
    driver.maximize_window()
    driver.get("https://csm.ait.co.th/msa/home.php")

    wait_and_type(driver, By.ID, "loginname", USERNAME)
    wait_and_type(driver, By.ID, "password", PASSWORD)
    wait_and_click(driver, By.ID, "btn-login")

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'NEW SAR')]"))
        )
        print("✅ Login สำเร็จ")
    except:
        print("❌ Login ไม่สำเร็จ — username/password ผิด")
        messagebox.showerror("Login Failed", "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง\nกรุณากรอกใหม่อีกครั้ง")
        driver.quit()
        return

    wait_and_click(driver, By.XPATH, "//button[contains(text(), 'NEW SAR')]")
    wait_and_click(driver, By.XPATH, "//a[contains(text(), 'สร้างและบันทึก')]")
    wait_and_click(driver, By.ID, "process_act_module_2")
    wait_and_click(driver, By.XPATH, f"//strong[contains(text(),'{project_code}')]")

    time.sleep(2)
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe")))

    wait_and_type(driver, By.ID, "task_note", TASK_NOTE)
    wait_and_type(driver, By.ID, "finish_text", activity)
    for field_id, dt_str in datetime_fields.items():
        wait_and_type(driver, By.ID, field_id, dt_str)

    try:
        checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "showSubmitWorkOrderButtonDD")))
        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
        driver.execute_script("arguments[0].click();", checkbox)
        print("✅ ติ๊ก checkbox 'ส่งข้อมูล' แล้ว")
    except Exception as e:
        print("❌ ติ๊ก checkbox 'ส่งข้อมูล' ไม่ได้:", e)

    try:
        submit_btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SubmitButtonWorkOrderd")))
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        driver.execute_script("arguments[0].click();", submit_btn)
        print("✅ คลิก Submit แล้ว")
    except Exception as e:
        print("❌ คลิก Submit ไม่ได้:", e)

    try:
        WebDriverWait(driver, 10).until(EC.url_changes(driver.current_url))
        print("🌐 Redirect แล้ว")
    except Exception as e:
        print("⚠️ ไม่เกิดการ redirect:", e)

    driver.switch_to.default_content()
    time.sleep(2)

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    checkbox_found = False
    for iframe in iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            checkbox_ready = driver.find_element(By.ID, "showSubmitCloseWorkOrderButton")
            driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_ready)
            driver.execute_script("arguments[0].click();", checkbox_ready)
            print("✅ ติ๊ก checkbox 'ตรวจสอบข้อมูลเรียบร้อยแล้ว' แล้ว")
            checkbox_found = True
            break
        except:
            continue

    if not checkbox_found:
        time.sleep(5)
        try:
            driver.switch_to.default_content()
            checkbox_ready = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "showSubmitCloseWorkOrderButton")))
            driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_ready)
            driver.execute_script("arguments[0].click();", checkbox_ready)
            print("✅ ติ๊ก checkbox 'ตรวจสอบข้อมูลเรียบร้อยแล้ว' แล้ว (no iframe)")
        except Exception as e:
            print("❌ ติ๊ก checkbox 'ตรวจสอบข้อมูลเรียบร้อยแล้ว' ไม่ได้:", e)

    try:
        confirm_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="SubmitButtonCloseWorkOrder"]'))
        )
        time.sleep(1.5)
        driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
        driver.execute_script("arguments[0].click();", confirm_button)
        print("✅ คลิก 'ยืนยัน' ปิดงานแล้ว")
    except Exception as e:
        print("❌ คลิกปุ่ม 'ยืนยัน' ไม่ได้:", e)

    try:
        time.sleep(3)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        checkbox = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="accordionTaxi-collapseOne"]/div/div[3]/label'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", checkbox)
        print("✅ คลิก 'ไม่มีค่าเดินทาง' แล้ว")
    except Exception as e:
        print("❌ คลิก 'ไม่มีค่าเดินทาง' ไม่ได้:", e)

    try:
        checkbox = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="action_container_form"]/div[6]/label'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", checkbox)
        print("✅ ติ๊ก checkbox สุดท้ายแล้ว")
    except Exception as e:
        print("❌ ติ๊ก checkbox สุดท้ายไม่ได้:", e)

    try:
        submit_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="SubmitButtonWorkOrderd"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", submit_btn)
        print("✅ ส่งข้อมูลเรียบร้อยแล้ว")
    except Exception as e:
        print("❌ ส่งข้อมูลไม่สำเร็จ:", e)

    time.sleep(10)
    driver.quit()

# ========= Run =========
if __name__ == "__main__":
    start_gui()
