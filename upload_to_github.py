import os
import subprocess
import datetime

# === CẤU HÌNH ===
REPO_PATH = os.path.dirname(os.path.abspath(__file__))  # thư mục hiện tại
COMMIT_MESSAGE = f"auto update ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')})"


def run_command(cmd):
    """Chạy lệnh git và in kết quả."""
    try:
        result = subprocess.run(cmd, cwd=REPO_PATH, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            print(result.stdout.strip())
        else:
            print(f"⚠️ Lỗi khi chạy: {cmd}\n{result.stderr.strip()}")
    except Exception as e:
        print(f"❌ Lỗi: {e}")


def main():
    print("🚀 Đang đẩy dữ liệu lên GitHub...")
    run_command("git add .")
    run_command(f'git commit -m "{COMMIT_MESSAGE}"')
    run_command("git push")
    print("✅ Hoàn tất! Tất cả thay đổi đã được đẩy lên GitHub.")


if __name__ == "__main__":
    main()
