import schedule
import time
import os
from instagrapi import Client
from tiktok_uploader.upload import upload_video
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================= CONFIGURATION =================
VIDEO_PATH = r"C:\path\to\your\video.mp4"
CAPTION = "Check out my new video! #viral #fyp"

# Instagram
IG_USERNAME = "your_username"
IG_PASSWORD = "your_password"

# TikTok
# TikTok yêu cầu session_id cookie để tránh login captcha
TIKTOK_SESSION_ID = "your_tiktok_session_id_cookie"

# Facebook
FB_EMAIL = "your_email"
FB_PASSWORD = "your_password"
# =================================================

def post_to_instagram(video_path, caption):
    print(">>> Starting Instagram Upload...")
    try:
        cl = Client()
        cl.login(IG_USERNAME, IG_PASSWORD)
        media = cl.video_upload(
            video_path,
            caption=caption
        )
        print(f"✅ Instagram Upload Successful! Media PK: {media.pk}")
    except Exception as e:
        print(f"❌ Instagram Failed: {e}")

def post_to_tiktok(video_path, description):
    print(">>> Starting TikTok Upload...")
    try:
        # tiktok-uploader uses selenium under the hood usually or requests
        # Check docs for specific lib usage. This is a common pattern:
        # failed_log check is optional
        upload_video(
            filename=video_path,
            description=description,
            sessionid=TIKTOK_SESSION_ID
        )
        print("✅ TikTok Upload Successful!")
    except Exception as e:
        print(f"❌ TikTok Failed: {e}")

def post_to_facebook(video_path, caption):
    print(">>> Starting Facebook Upload (via Selenium)...")
    try:
        options = webdriver.ChromeOptions()
        # options.add_argument("--headless") # Don't use headless for FB initial login
        options.add_argument("--disable-notifications")
        
        driver = webdriver.Chrome(options=options)
        
        # 1. Login
        driver.get("https://www.facebook.com/")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "email")))
        
        driver.find_element(By.NAME, "email").send_keys(FB_EMAIL)
        driver.find_element(By.NAME, "pass").send_keys(FB_PASSWORD)
        driver.find_element(By.NAME, "login").click()
        
        time.sleep(5) # Wait for login
        
        # 2. Go to creation (This is tricky as URLs change. Simplest is to go to user profile or simplified composer)
        # Using mobile version often simpler for automation: m.facebook.com
        driver.get("https://m.facebook.com/composer/mbasic/") 
        # Note: mbasic is very strict/limited but good for simple text. Video might require full site.
        
        # Let's try standard full site composer
        driver.get("https://www.facebook.com/latest_content_creation_view")
        # This is often for Pages. For profile, it's just home.
        
        # Warning: FB automation is very brittle.
        print("⚠ FB Automation is complex and depends on your UI version. This is a placeholder for the logic.")
        
        driver.quit()
        print("✅ Facebook Upload Process Finished (Simulated)")
    except Exception as e:
        print(f"❌ Facebook Failed: {e}")

def job():
    print(f"\n⏰ Starting scheduled post job at {time.ctime()}...")
    
    if not os.path.exists(VIDEO_PATH):
        print(f"❌ Video file not found: {VIDEO_PATH}")
        return

    # Run tasks
    post_to_instagram(VIDEO_PATH, CAPTION)
    post_to_tiktok(VIDEO_PATH, CAPTION)
    post_to_facebook(VIDEO_PATH, CAPTION)
    
    print("\n🏁 All tasks completed.")

if __name__ == "__main__":
    # Schedule the job
    # Ví dụ: Post vào 8:00 mỗi ngày
    # schedule.every().day.at("08:00").do(job)
    
    # Test run ngay lập tức 1 lần
    print("Checking libraries and configuration...")
    job()
    
    # Uncomment to run scheduler
    # while True:
    #    schedule.run_pending()
    #    time.sleep(60)
