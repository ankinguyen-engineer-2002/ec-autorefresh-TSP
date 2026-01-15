"""
EC Auto Refresh - MCN Excel Refresh Logic Module
=============================================
Tự động refresh các file Excel MCN có kết nối Power Query / Data Connection.

Flow:
1. Copy files từ SOURCE sang TARGET (local disk để tránh sync issues)
2. Refresh tất cả files trong TARGET
3. Copy files đã refresh về SOURCE

Author: eCentric Team
Created: 2026-01-13
"""

import os
import time
import shutil
import traceback
import sys

# Set console encoding to utf-8 to support emojis on Windows
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass # Python versions < 3.7 might not have reconfigure

# Windows-specific imports (chỉ hoạt động trên Windows với Excel cài đặt)
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("⚠️ win32com not available - Excel refresh will not work in this environment")


# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================

# Đường dẫn mặc định (có thể override qua environment variables)
DEFAULT_SOURCE_PATH = r"C:\Users\Admin\NextCommerce\Data - General\MCN custom report"
DEFAULT_TARGET_PATH = r"C:\EC_project\EC refresh TSP\Copy MCN"

# Power Automate Webhook URL (cùng URL với ingest DAGs)
PA_URL = "https://default16f5375c95a943f0ba2ce20bd5ec28.45.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d77666e9f6e64eb4887cbdf703bf3e23/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=C113lCMcFvVvnxaThCZUyu503Z7FI_sVJjlVzVuQbiE"

def get_config():
    """Lấy cấu hình từ environment variables hoặc dùng default"""
    return {
        'source_path': os.environ.get('REFRESH_SOURCE_PATH', DEFAULT_SOURCE_PATH),
        'target_path': os.environ.get('REFRESH_TARGET_PATH', DEFAULT_TARGET_PATH),
        'file_extensions': ('.xlsx', '.xlsm'),
        'max_retries': 3,
        'refresh_wait_seconds': 3,
        'retry_delay_seconds': 5,
        'between_files_delay': 2,
        'send_notifications': True,  # Gửi thông báo qua Power Automate
    }


def send_notification(subject, html_body):
    """Gửi thông báo qua Power Automate webhook"""
    try:
        import requests
        headers = {"Content-Type": "application/json"}
        response = requests.post(
            PA_URL, 
            json={"subject": subject, "content_html": html_body}, 
            headers=headers,
            timeout=10
        )
        print(f"   📧 Notification sent: {subject}")
        return response.status_code == 200
    except Exception as e:
        print(f"   ⚠️ Failed to send notification: {e}")
        return False


# ==============================================================================
# 2. FILE OPERATIONS
# ==============================================================================

def copy_files(src_folder, dst_folder, direction="to_refresh", move_files=False):
    """
    Copy Excel files giữa 2 thư mục.
    
    Args:
        src_folder: Thư mục nguồn
        dst_folder: Thư mục đích
        direction: 'to_refresh' hoặc 'to_source' (chỉ để log)
        move_files: Nếu True, sẽ xóa file nguồn sau khi copy xong (giả lập Cut)
    
    Returns:
        tuple: (count_success, count_failed, list_failed_files)
    """
    config = get_config()
    action = "Moving" if move_files else "Copying"
    print(f"\n📂 {action} files ({direction})...")
    print(f"   From: {src_folder}")
    print(f"   To:   {dst_folder}")
    
    if not os.path.exists(src_folder):
        print(f"❌ Source folder not found: {src_folder}")
        return 0, 0, []
    
    if not os.path.exists(dst_folder):
        os.makedirs(dst_folder)
        print(f"   Created destination folder: {dst_folder}")

    count_success, count_failed = 0, 0
    failed_files = []

    for filename in os.listdir(src_folder):
        if filename.endswith(config['file_extensions']):
            src = os.path.join(src_folder, filename)
            dst = os.path.join(dst_folder, filename)
            try:
                shutil.copy2(src, dst)
                if move_files:
                    os.remove(src)
                    print(f"   ✂️ Moved: {filename}")
                else:
                    print(f"   ✅ Copied: {filename}")
                count_success += 1
            except Exception as e:
                print(f"   ⚠️ Cannot copy {filename}: {e}")
                count_failed += 1
                failed_files.append(filename)

    print(f"\n✅ Copied {count_success} files, ⚠️ Skipped {count_failed} files.")
    return count_success, count_failed, failed_files


# ==============================================================================
# 3. EXCEL REFRESH ENGINE
# ==============================================================================

def refresh_excel_file(file_path, config=None):
    """
    Refresh 1 file Excel với retry mechanism.
    
    Args:
        file_path: Đường dẫn đầy đủ tới file Excel
        config: Config dict (optional, sẽ lấy từ get_config() nếu None)
    
    Returns:
        bool: True nếu refresh thành công, False nếu thất bại
    """
    if not HAS_WIN32COM:
        print(f"❌ Cannot refresh {file_path} - win32com not available")
        return False
    
    if config is None:
        config = get_config()
    
    filename = os.path.basename(file_path)
    print(f"🔄 Refreshing: {filename}")

    for attempt in range(1, config['max_retries'] + 1):
        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False

            # Mở workbook
            wb = excel.Workbooks.Open(file_path, UpdateLinks=0)
            
            # Refresh tất cả connections/queries
            wb.RefreshAll()
            
            # Đợi cho tất cả background queries hoàn thành
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except:
                pass  # Một số phiên bản Excel không hỗ trợ
            
            time.sleep(config['refresh_wait_seconds'])
            
            # Save và close
            wb.Save()
            wb.Close(SaveChanges=True)
            wb = None

            excel.Quit()
            excel = None
            pythoncom.CoUninitialize()
            print(f"   ✅ Done: {filename}")
            return True

        except Exception as e:
            print(f"   ❌ ERROR attempt {attempt}/{config['max_retries']} on {filename}: {e}")
            traceback.print_exc()
            time.sleep(config['retry_delay_seconds'])

        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    print(f"   🚫 Failed after {config['max_retries']} attempts: {filename}")
    return False


def refresh_excel_folder(folder_path, config=None):
    """
    Refresh tất cả file Excel trong folder.
    
    Args:
        folder_path: Đường dẫn thư mục chứa files
        config: Config dict (optional)
    
    Returns:
        tuple: (success_files, failed_files)
    """
    if config is None:
        config = get_config()
    
    if not os.path.exists(folder_path):
        print(f"❌ Folder not found: {folder_path}")
        return [], []
    
    files = [f for f in os.listdir(folder_path) if f.endswith(config['file_extensions'])]
    print(f"\n🧾 Tổng số file cần refresh: {len(files)}\n")

    success_files = []
    failed_files = []

    for i, file_name in enumerate(files, start=1):
        file_path = os.path.join(folder_path, file_name)
        print(f"[{i}/{len(files)}] File: {file_name}")

        if refresh_excel_file(file_path, config):
            success_files.append(file_name)
        else:
            failed_files.append(file_name)

        time.sleep(config['between_files_delay'])

    print("\n🎯 Refresh xong toàn bộ files.")
    return success_files, failed_files


# ==============================================================================
# 4. MAIN EXECUTION FUNCTIONS
# ==============================================================================

def run_excel_refresh():
    """
    Main function - Chạy full flow refresh Excel.
    Được gọi bởi Airflow DAG hoặc chạy trực tiếp.
    
    Returns:
        dict: Summary của quá trình refresh
    """
    from datetime import datetime
    
    start_time = time.time()
    config = get_config()
    str_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    print("=" * 60)
    print("🚀 EC AUTO REFRESH - Starting Excel Refresh Job")
    print("=" * 60)
    print(f"📁 Source: {config['source_path']}")
    print(f"📁 Target: {config['target_path']}")
    print("=" * 60)

    # 📧 Notification: START
    if config['send_notifications']:
        send_notification(
            subject=f"🔄 [BẮT ĐẦU] Excel Refresh - {str_start}",
            html_body=f"""
            <h3>🔄 Excel Refresh Job Đã Khởi Động</h3>
            <p><b>Thời gian:</b> {str_start}</p>
            <p><b>Source:</b> {config['source_path']}</p>
            <p><b>Target:</b> {config['target_path']}</p>
            """
        )

    try:
        # 1️⃣ Copy files từ SOURCE qua TARGET
        copy_success_1, copy_failed_1, _ = copy_files(
            config['source_path'], 
            config['target_path'], 
            direction="to_refresh"
        )

        # 2️⃣ Refresh tất cả file trong TARGET
        success, failed = refresh_excel_folder(config['target_path'], config)

        # 3️⃣ Copy ngược lại file đã refresh về SOURCE
        copy_success_2, copy_failed_2, _ = copy_files(
            config['target_path'], 
            config['source_path'], 
            direction="to_source",
            move_files=True
        )

        # 4️⃣ Tổng kết
        elapsed = round(time.time() - start_time, 2)
        
        print("\n" + "=" * 60)
        print("📊 SUMMARY")
        print("=" * 60)
        print(f"✅ Refresh thành công: {len(success)} files")
        print(f"❌ Refresh thất bại: {len(failed)} files")
        if failed:
            print(f"   -> Files lỗi: {', '.join(failed)}")
        print(f"⏱️  Tổng thời gian: {elapsed} giây")
        print("=" * 60)
        
        summary = {
            'total_files': len(success) + len(failed),
            'success_count': len(success),
            'failed_count': len(failed),
            'success_files': success,
            'failed_files': failed,
            'elapsed_seconds': elapsed,
        }

        # 📧 Notification: SUCCESS/COMPLETE
        if config['send_notifications']:
            if len(failed) == 0:
                status = "THÀNH CÔNG"
                color = "green"
            else:
                status = "CÓ LỖI"
                color = "orange"
            
            files_html = "<ul>" + "".join([f"<li>✅ {f}</li>" for f in success]) + "</ul>"
            if failed:
                files_html += "<ul>" + "".join([f"<li>❌ {f}</li>" for f in failed]) + "</ul>"
            
            send_notification(
                subject=f"{'✅' if len(failed)==0 else '⚠️'} [{status}] Excel Refresh - {len(success)}/{len(success)+len(failed)} files",
                html_body=f"""
                <h3 style="color:{color};">📊 Báo Cáo Excel Refresh: {status}</h3>
                <table border="1" cellpadding="5" style="border-collapse:collapse;">
                    <tr><td><b>Thành công</b></td><td><b>{len(success)}</b> files</td></tr>
                    <tr><td><b>Thất bại</b></td><td><b>{len(failed)}</b> files</td></tr>
                    <tr><td><b>Thời gian</b></td><td>{elapsed}s</td></tr>
                </table>
                <h4>Chi tiết:</h4>
                {files_html}
                """
            )
        
        return summary

    except Exception as e:
        # 📧 Notification: FAILURE
        if config['send_notifications']:
            send_notification(
                subject="❌ [LỖI] Excel Refresh Thất Bại",
                html_body=f"<h3 style='color:red;'>❌ Excel Refresh Gặp Lỗi</h3><pre>{traceback.format_exc()}</pre>"
            )
        raise


# ==============================================================================
# 5. STANDALONE EXECUTION
# ==============================================================================

if __name__ == "__main__":
    summary = run_excel_refresh()
    
    # Exit với code khác 0 nếu có lỗi
    if summary['failed_count'] > 0:
        exit(1)
    exit(0)
