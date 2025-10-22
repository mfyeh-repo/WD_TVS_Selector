#!/usr/bin/env python
# coding: utf-8

# 匯入必要套件
import os
import pandas as pd

print("=== TVS 測試資料輸入程式 ===")

# 1. 元件編號選單
part_list = ["D5094", "D5095", "D9414"]
print("\n可選擇的元件編號:")
for i, w in enumerate(part_list, 1):
    print(f"{i}. {w}")

while True:
    try:
        choice = int(input("請輸入選擇編號 (1-3): "))
        if 1 <= choice <= len(part_list):
            part_number = part_list[choice - 1]
            break
        else:
            print("⚠️ 請輸入 1 到 3 之間的數字。") 
    except ValueError:
        print("⚠️ 請輸入整數。")

# 設定資料夾路徑
data_path = "DataSheet"
# 組合完整檔案路徑
file_path = os.path.join(data_path, part_number+".xlsx")

# 檢查檔案是否存在
if os.path.isfile(file_path):
    # 嘗試讀取 Excel 檔案
    try:
        df = pd.read_excel(file_path)
        print("檔案載入成功！")
        # 顯示前幾筆資料作為確認
        print(df.head())
    except Exception as e:
        print(f"檔案讀取失敗，錯誤訊息：{e}")
else:
    print("檔案不存在，請確認輸入的檔名是否正確。")

# 2. Test level (V)
while True:
    try:
        test_level = float(input("請輸入 Test level (V): "))
        break
    except ValueError:
        print("⚠️ 請輸入有效的數值。")

# 3. Test Waveform 選單
waveforms = ["10/1000us", "1.2/50us (8/20us)", "10/700us (5/320us)"]
print("\n可選擇的 Test Waveform:")
for i, w in enumerate(waveforms, 1):
    print(f"{i}. {w}")

while True:
    try:
        choice = int(input("請輸入選擇編號 (1-3): "))
        if 1 <= choice <= len(waveforms):
            test_waveform = waveforms[choice - 1]
            break
        else:
            print("⚠️ 請輸入 1 到 3 之間的數字。")
    except ValueError:
        print("⚠️ 請輸入整數。")

# 4. Test resistance
while True:
    try:
        test_resistance = float(input("\n請輸入 Test resistance (Ω): "))
        break
    except ValueError:
        print("⚠️ 請輸入有效的數值。")


# 5. Use TVS pcs
while True:
    try:
        tvs_pcs = int(input("請輸入使用的 TVS 數量 (pcs): "))
        break
    except ValueError:
        print("⚠️ 請輸入整數。")


# 顯示輸入結果
print("\n=== 輸入結果 ===")
print(f"Part number : {part_number}")
print(f"Test level (V): {test_level}")
print(f"Test waveform : {test_waveform}")
print(f"Test resistance (Ω): {test_resistance}")
print(f"Use TVS (pcs): {tvs_pcs}")


# 計算 Test Current: Test level / Test resistance = Test current
test_current = test_level/test_resistance
print(f"Test current (A): {test_current}")

# 計算 test Ipp
#--- 根據 waveform 設定 factor ---
if test_waveform == "10/1000us":
    factor = 1
elif test_waveform == "1.2/50us (8/20us)":
    factor = 7.07
elif test_waveform == "10/700us (5/320us)":
    factor = 1.76
else:
    factor = None  # 預設安全值
print(f"{test_waveform} Ipp Factor : {factor}")

#--- test Ipp = factor * test_current / tvs_pcs ---
test_Ipp = factor * test_current / tvs_pcs
print(f"Test Ipp (A): {test_Ipp}")


# 篩選 Ipp 小於 test_Ipp 的元件編號
filtered = df[df["Ipp"] < test_Ipp]

# 取出對應的 UniNumber 欄位
result = filtered["UniNumber"]

# 顯示結果
print("Ipp < test_Ipp 對應的 UniNumber：")
print(result.to_list())