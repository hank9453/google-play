from google_play_scraper import Sort, reviews
import pandas as pd
from datetime import datetime, timedelta
import os

APP_ID = 'com.mitake.android.bk.tcb'
PLATFORM = 'google_play'
folder_path = './爬蟲結果'
os.makedirs(folder_path, exist_ok=True)

# 現在時間
now = datetime.now().replace(second=0, microsecond=0)
minute = now.minute

# 判斷目前在上半或下半時段
if minute < 30:
    # 30 分以前：抓取 25 分到整點（e.g. 12:25 ～ 13:00）
    end_time = now.replace(minute=0)
    start_time = end_time - timedelta(minutes=35)
else:
    # 30 分以後：抓取 55 分到下半點（e.g. 12:55 ～ 13:30）
    end_time = now.replace(minute=30)
    start_time = end_time - timedelta(minutes=35)

# 爬取最新評論（最多500筆）
results, _ = reviews(
    APP_ID,
    lang='zh_TW',
    country='tw',
    sort=Sort.NEWEST,
    count=200,
)

# 篩選符合時間的評論
filtered = []
for review in results:
    review_time = review['at'].replace(tzinfo=None)
    if start_time <= review_time <= end_time:
        filtered.append(review)

# 儲存為 Excel
if filtered:
    df = pd.DataFrame(filtered)
    file_name = f"{PLATFORM}_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"
    file_path = os.path.join(folder_path, file_name)
    df.to_excel(file_path, index=False)
    print(f"✔ 儲存 {len(filtered)} 筆評論於：{file_path}")
else:
    print(f"⚠ 此區間內無評論：{start_time} ~ {end_time}")
