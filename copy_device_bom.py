# -*- coding: utf-8 -*-
"""
Device BOM 查詢複製工具
從 Device-Name_對應表_鈺太芯.xlsx 查詢數據並複製到 MEMS MIC NPI tracking list.xlsm
"""

import pandas as pd
import warnings
import win32com.client as win32
import os

# 忽略 openpyxl 警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 路徑設定
SOURCE_FILE = r'\\dxfile\PE\Project Status\Andrew\Device BOM\Device-Name_對應表_鈺太芯.xlsx'
SOURCE_SHEET = '(2) 麥克風'
TARGET_FILE = r'C:\Users\aan_chen\OneDrive\6. Total MIC Test Setting\test\MEMS MIC NPI tracking list.xlsm'
TARGET_SHEET = '鈺太芯Device'


def open_source_file():
    """開啟來源檔案讓使用者自行搜尋"""
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = True
        excel.Workbooks.Open(SOURCE_FILE)
        print(f"已開啟來源檔案，請自行搜尋。")
    except Exception as e:
        print(f"開啟來源檔案失敗: {e}")


def clean_string(s):
    """清理字串：去掉換行符和多餘空白"""
    if s:
        return s.replace('\n', ' ').replace('\r', ' ').strip()
    return s


def parse_pasted_data(lines):
    """
    解析貼上的資料
    第3行: MPN
    第6行: Device Name
    第8行: ASIC
    第10行: SensorID
    """
    mpn = None
    device_name = None
    asic = None
    sensor_id = None

    if len(lines) >= 3:
        mpn = clean_string(lines[2]) or None
    if len(lines) >= 6:
        device_name = clean_string(lines[5]) or None
    if len(lines) >= 8:
        asic = clean_string(lines[7]) or None
    if len(lines) >= 10:
        sensor_id = clean_string(lines[9]) or None

    return mpn, asic, sensor_id, device_name


def search_and_copy(mpn=None, asic=None, sensor_id=None, device_name=None):
    """
    查詢並複製數據
    參數可以只填部分，會以 AND 邏輯處理
    """
    print("=" * 60)
    print("正在讀取來源檔案...")

    # 讀取來源檔案
    try:
        df = pd.read_excel(SOURCE_FILE, sheet_name=SOURCE_SHEET)
    except Exception as e:
        print(f"讀取來源檔案失敗: {e}")
        return False

    print(f"來源共 {len(df)} 筆數據")

    # 欄位列名（處理換行符）
    col_mpn = 'MPN'
    col_asic = 'ASIC'
    col_sensor = 'SensorID'
    col_device = '品名Device\nName'

    # 建立篩選條件
    mask = pd.Series([True] * len(df))

    if mpn:
        # 同時處理來源資料中的換行符
        mask &= df[col_mpn].astype(str).str.replace('\n', ' ', regex=False).str.replace('\r', ' ', regex=False).str.strip() == str(mpn).strip()
        print(f"篩選 MPN = '{mpn}'")

    if asic:
        mask &= df[col_asic].astype(str).str.replace('\n', ' ', regex=False).str.replace('\r', ' ', regex=False).str.strip() == str(asic).strip()
        print(f"篩選 ASIC = '{asic}'")

    if sensor_id:
        mask &= df[col_sensor].astype(str).str.replace('\n', ' ', regex=False).str.replace('\r', ' ', regex=False).str.strip() == str(sensor_id).strip()
        print(f"篩選 SensorID = '{sensor_id}'")

    if device_name:
        mask &= df[col_device].astype(str).str.replace('\n', ' ', regex=False).str.replace('\r', ' ', regex=False).str.strip() == str(device_name).strip()
        print(f"篩選 品名Device Name = '{device_name}'")

    # 篩選結果
    result = df[mask]
    count = len(result)

    print("-" * 60)

    if count == 0:
        print("找不到符合條件的數據！")
        print("正在開啟來源檔案讓您自行搜尋...")
        open_source_file()
        return False

    if count > 1:
        print(f"找到 {count} 筆數據，請提供更多條件來確定唯一結果：")
        print()
        # 顯示找到的資料（只顯示關鍵列）
        for idx, row in result.iterrows():
            print(f"[{idx}] MPN={row[col_mpn]}, ASIC={row[col_asic]}, SensorID={row[col_sensor]}, Device={row[col_device]}")
        print()
        print("正在開啟來源檔案讓您自行搜尋...")
        open_source_file()
        return False

    # 只有一筆，進行複製
    print("找到唯一一筆數據！")
    print()

    # 取得 A~AG 列 (index 0~32, 共33列)
    row_data = result.iloc[0, 0:33].tolist()

    # 顯示摘要資訊
    print("要複製的數據摘要:")
    print(f"  MPN: {row_data[1]}")
    print(f"  品號BOM: {row_data[2]}")
    print(f"  品名Device Name: {row_data[3]}")
    print(f"  技術負責人: {row_data[4]}")
    print(f"  ASIC: {row_data[6]}")
    print(f"  SensorID: {row_data[8]}")

    print("-" * 60)
    print("正在開啟目標檔案並寫入...")

    try:
        # 使用 win32com 開啟 Excel
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = True
        excel.DisplayAlerts = False

        # 直接開啟檔案
        wb = excel.Workbooks.Open(TARGET_FILE)

        # 找到目標工作表
        ws = wb.Worksheets(TARGET_SHEET)

        # 找最後一行（有數據的）
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
        new_row = last_row + 1

        # 複製上一行格式到新行
        ws.Rows(last_row).Copy()
        ws.Rows(new_row).Insert()
        excel.CutCopyMode = False

        # 寫入數據
        for col_idx, value in enumerate(row_data, start=1):
            if pd.isna(value):
                value = ''
            elif isinstance(value, str):
                value = value.replace('\n', ' ').replace('\r', ' ').strip()
            ws.Cells(new_row, col_idx).Value = value

        # 選取新寫入的那一行
        ws.Activate()
        ws.Rows(new_row).Select()

        print(f"已寫入到第 {new_row} 行（已複製上一行格式）！")
        print("檔案已開啟，請自行確認後決定是否儲存。")
        print("=" * 60)

        return True

    except Exception as e:
        print(f"操作失敗: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print()
    print("=" * 60)
    print("     Device BOM 查詢複製工具 - 產品料號")
    print("=" * 60)
    print()
    print("請貼上資料（貼完後按 Enter 兩次）：")
    print("格式: 第3行=MPN, 第6行=Device, 第8行=ASIC, 第10行=SensorID")
    print("-" * 60)

    lines = []
    empty_count = 0

    while True:
        line = input()
        if line == '':
            empty_count += 1
            if empty_count >= 2:
                break
        else:
            empty_count = 0
            lines.append(line)

    if not lines:
        print("未輸入任何資料！")
        input("\n按 Enter 鍵結束...")
        return

    # 解析貼上的資料
    mpn, asic, sensor_id, device_name = parse_pasted_data(lines)

    print()
    print("解析結果:")
    print(f"  MPN: {mpn}")
    print(f"  ASIC: {asic}")
    print(f"  SensorID: {sensor_id}")
    print(f"  Device Name: {device_name}")

    if not any([mpn, asic, sensor_id, device_name]):
        print("無法解析出有效條件！")
        input("\n按 Enter 鍵結束...")
        return

    print()
    search_and_copy(mpn, asic, sensor_id, device_name)

    print()
    input("按 Enter 結束...")


if __name__ == '__main__':
    main()
