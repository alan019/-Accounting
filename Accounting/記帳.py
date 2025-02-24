import pandas as pd
import matplotlib.pyplot as plt
import datetime
import os

# 設定支援中文的字體
plt.rcParams['font.family'] = 'Heiti TC'

# 獲取桌面路徑
桌面路徑 = os.path.expanduser("~/Desktop")

def 確保收入欄位(df):
    """確保 DataFrame 中有 '收入' 欄位，若無則新增並設置為 0"""
    if '收入' not in df.columns:
        df['收入'] = 0
    return df

def 讀取收入(df):
    """從 Excel 中讀取收入數據，如果沒有則設置為 0"""
    df = 確保收入欄位(df)
    return df['收入'].sum()

def 新增存款(excel_file):
    # 讀取現有的資料
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
    else:
        df = pd.DataFrame()

    # 確保 '收入' 欄位存在
    df = 確保收入欄位(df)

    # 輸入存款金額
    try:
        deposit = float(input("請輸入您要添加的存款金額: "))
    except ValueError:
        print("金額格式錯誤，請輸入正確的數值格式")
        return  # 如果金額格式錯誤，返回並不執行後續操作

    # 更新收入欄位，累加新收入
    df.at[0, '收入'] = df['收入'].sum() + deposit

    # 更新並保存到 Excel
    df.to_excel(excel_file, index=False)
    print(f"您的收入已增加至: {df['收入'].sum()} 元")

def 總結支出(df, target_month=None):
    """總結並顯示支出總數和收入總數"""
    df = 確保收入欄位(df)
    
    if target_month is None:
        target_month = datetime.date.today().strftime("%Y-%m")
    
    df['日期'] = pd.to_datetime(df['日期'])
    df['月份'] = df['日期'].dt.to_period('M')

    # 選擇目標月份的資料
    df_month = df[df['月份'] == target_month]
    total_expense = df_month['金額'].sum()

    # 讀取總收入
    total_income = 讀取收入(df)

    print(f"{target_month} 的總支出為: {total_expense} 元")
    print(f"{target_month} 的總收入為: {total_income} 元")
    return total_expense, total_income

def 新增支出(excel_file):
    # 讀取現有的資料
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
    else:
        df = pd.DataFrame()

    # 確保 '收入' 欄位存在
    df = 確保收入欄位(df)

    # 輸入支出詳情
    date_input = input("請輸入日期(YYYY-MM-DD): ")
    if not date_input:
        date = datetime.date.today()  # 使用當前日期
    else:
        try:
            date = datetime.datetime.strptime(date_input, "%Y-%m-%d").date()
        except ValueError:
            print("日期格式錯誤，請重新輸入正確的格式 YYYY-MM-DD")
            return  # 如果格式錯誤，返回並不執行後續操作

    # 自動計算月份
    month = date.strftime("%Y-%m")
    
    # 輸入類別、金額與描述
    category = input("請輸入類別: ")
    try:
        amount = float(input("請輸入金額: "))
    except ValueError:
        print("金額格式錯誤，請輸入正確的數值格式")
        return  # 如果金額格式錯誤，返回並不執行後續操作

    description = input("請輸入描述 (可選): ")

    # 創建新的支出記錄，並包含月份欄位
    new_expense = pd.DataFrame({
        '日期': [date],
        '月份': [month],  # 添加月份欄位
        '類別': [category],
        '金額': [amount],
        '描述': [description]
    })

    # 將新記錄追加到 Excel 文件
    df = pd.concat([df, new_expense], ignore_index=True)

    # 保存到 Excel
    df.to_excel(excel_file, index=False)
    print("支出已成功新增！")

    # 強制重新讀取 Excel 文件以確保最新支出被記錄
    df = pd.read_excel(excel_file)
    
    # 總結當前月份的支出
    總結支出(df)

def 生成月度結算圖表(excel_file, target_month=None):
    # 強制重新讀取現有的資料，確保包含最新的支出
    if not os.path.exists(excel_file):
        print("還沒有記錄任何支出。")
        return

    df = pd.read_excel(excel_file)
    df = 確保收入欄位(df)
    df['日期'] = pd.to_datetime(df['日期'])
    df['月份'] = df['日期'].dt.to_period('M')

    # 如果沒有指定月份，則自動查詢最後一次輸入的月份
    if target_month is None:
        target_month = df['月份'].max().strftime("%Y-%m")

    # 選擇目標月份的資料
    df_month = df[df['月份'] == target_month]

    if df_month.empty:
        print(f"沒有找到 {target_month} 的記錄。")
        return

    # 總結支出和收入
    total_expense, total_income = 總結支出(df, target_month)

    # 計算當月盈餘
    balance = total_income - total_expense

    print(f"當月盈餘: {balance} 元")

    # 生成圓餅圖來顯示支出分類
    monthly_summary = df_month.groupby('類別')['金額'].sum()

    # 調整圖表大小來容納更多的類別
    plt.figure(figsize=(8, 8))

    # 生成圓餅圖
    plt.pie(monthly_summary, labels=monthly_summary.index, autopct='%1.1f%%', startangle=90)
    plt.title(f'{target_month}的支出\n總收入：{total_income} 元, 當月盈餘：{balance} 元')

    # 確保圖表顯示為圓形
    plt.axis('equal')

    # 顯示圖表
    plt.tight_layout()
    plt.show()

def 主選單():
    # Excel 文件在桌面路徑處理
    excel_file = os.path.join(桌面路徑, '支出記錄.xlsx')
    
    while True:
        print("請選擇一個選項：")
        print("1. 新增支出")
        print("2. 添加存款")
        print("3. 生成月度結算圖表")
        print("4. 指定月份生成圖表")
        print("5. 退出")
        choice = input("請輸入您的選擇: ")
        if choice == '1':
            新增支出(excel_file)
        elif choice == '2':
            新增存款(excel_file)
        elif choice == '3':
            生成月度結算圖表(excel_file)
        elif choice == '4':
            target_month = input("請輸入您想查看的月份 (YYYY-MM)，或直接按下Enter查看最後輸入的月份: ")
            if not target_month:
                target_month = None
            生成月度結算圖表(excel_file, target_month)
        elif choice == '5':
            break
        else:
            print("無效的選擇，請重試。")

if __name__ == "__main__":
    主選單()
