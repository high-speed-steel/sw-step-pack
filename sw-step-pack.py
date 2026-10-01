import os
import zipfile
import datetime
import win32com.client

# ================= 設定路徑 =================
SOURCE_DIR = r"[輸入資料夾路徑]"      # 存放 .sldprt / .sldasm 的資料夾
OUTPUT_DIR = r"[輸入資料夾路徑]"      # 輸出的 STP 儲存路徑
RECURSIVE = False                   # 是否包含子資料夾（True/False）
# ===========================================

swDocPART = 1
swDocASSEMBLY = 2
swSaveAsCurrentVersion = 0
swSaveAsOptions_Silent = 1

def connect_solidworks():
    try:
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
        print("[資訊] 已成功連接至正在運行的 SOLIDWORKS。")
    except Exception:
        print("[資訊] 未檢測到開啟中的 SOLIDWORKS，正在啟動背景執行個體...")
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
    return sw_app

def convert_to_stp(sw_app, file_path, output_dir):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".sldprt":
        doc_type = swDocPART
    elif ext == ".sldasm":
        doc_type = swDocASSEMBLY
    else:
        return None

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{file_name}.stp")

    errors = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0)

    model_doc = sw_app.OpenDoc6(file_path, doc_type, 1 | 2, "", errors, warnings)
    if not model_doc:
        print(f"[-] 無法開啟檔案: {file_path} (錯誤碼: {errors.value})")
        return None

    try:
        success = model_doc.SaveAs4(output_path, swSaveAsCurrentVersion, swSaveAsOptions_Silent, errors, warnings)
        if success:
            print(f"[+] 成功輸出: {file_name}.stp")
            return output_path
        else:
            print(f"[-] 儲存失敗: {file_name} (錯誤碼: {errors.value})")
            return None
    finally:
        sw_app.CloseDoc(os.path.basename(file_path))

def create_zip(output_dir, zip_filename, files_to_zip):
    zip_path = os.path.join(output_dir, zip_filename)
    print(f"\n[開始壓縮] 正在打包檔案至: {zip_path}")
    
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
        for file in files_to_zip:
            if os.path.exists(file):
                zipf.write(file, arcname=os.path.basename(file))
                
    print(f"[壓縮完成] 已成功生成: {zip_filename} (包含 {len(files_to_zip)} 個檔案)")

def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    sw_app = connect_solidworks()

    target_files = []
    if RECURSIVE:
        for root, _, files in os.walk(SOURCE_DIR):
            for f in files:
                if f.lower().endswith((".sldprt", ".sldasm")):
                    target_files.append(os.path.join(root, f))
    else:
        for f in os.listdir(SOURCE_DIR):
            if f.lower().endswith((".sldprt", ".sldasm")):
                target_files.append(os.path.join(SOURCE_DIR, f))

    total = len(target_files)
    print(f"\n[開始轉檔] 共找到 {total} 個 CAD 模型檔案...")

    exported_stp_files = []
    for idx, file_path in enumerate(target_files, start=1):
        print(f"({idx}/{total}) 正在處理: {os.path.basename(file_path)}")
        stp_result = convert_to_stp(sw_app, file_path, OUTPUT_DIR)
        if stp_result:
            exported_stp_files.append(stp_result)

    print(f"\n[轉檔完成] 成功輸出: {len(exported_stp_files)}/{total}")

    # 轉檔完成後提示輸入壓縮檔名
    if exported_stp_files:
        default_name = f"STP_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        user_name = input(f"\n請輸入壓縮檔名稱 (直接按 Enter 使用 '{default_name}'): ").strip()
        
        zip_filename = user_name if user_name else default_name
        if not zip_filename.lower().endswith(".zip"):
            zip_filename += ".zip"

        create_zip(OUTPUT_DIR, zip_filename, exported_stp_files)

if __name__ == "__main__":
    main()