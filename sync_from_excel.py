#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PIF 難民特賣會 - Excel → 網站 同步腳本

功能：
  讀取最新的 Excel 檔，自動更新網站的：
    - 商品清單 (products.js)
    - 商品圖片 (images/)
    - 訂購須知 (info.json)

使用：
  1. 修改 Excel 檔（改價格、新增商品、貼新圖片等）
  2. 直接雙擊 sync.bat 執行同步

依賴：
  pip install openpyxl

設定：
  ▼ 如果 Excel 檔名有變更，請修改下方 EXCEL_FILE
"""

import os
import re
import sys
import json
import shutil
from pathlib import Path

# ============================================================
# 設定區（如果 Excel 檔名變了改這裡）
# ============================================================
SCRIPT_DIR = Path(__file__).parent.absolute()
PARENT_DIR = SCRIPT_DIR.parent  # PIF難民特賣會 資料夾

# 自動偵測 Excel 檔（找 .xlsx 但排除 ~$ 暫存檔）
EXCEL_CANDIDATES = [f for f in PARENT_DIR.glob("*.xlsx") if not f.name.startswith("~$")]

# 商品資料起始與結束行（在 Excel 裡）
PRODUCT_START_ROW = 13   # No.1 商品所在的行
PRODUCT_NAME_COL = 4     # 品名所在的欄 (D 欄)
PRODUCT_BARCODE_COL = 3  # 條碼欄 (C 欄)
PRODUCT_PRICE_COL = 5    # 單價欄 (E 欄)
PRODUCT_BOX_COL = 6      # 每箱入數欄 (F 欄)

# ============================================================

def main():
    print("=" * 60)
    print("  PIF 難民特賣會 - Excel 同步腳本")
    print("=" * 60)

    try:
        from openpyxl import load_workbook
    except ImportError:
        print("\n❌ 缺少 openpyxl 套件，請先在 cmd 執行：")
        print("   pip install openpyxl\n")
        input("按 Enter 結束...")
        sys.exit(1)

    if not EXCEL_CANDIDATES:
        print(f"\n❌ 在 {PARENT_DIR} 找不到任何 .xlsx 檔案")
        input("按 Enter 結束...")
        sys.exit(1)

    # 如果有多個檔案，讓使用者選
    if len(EXCEL_CANDIDATES) == 1:
        excel_path = EXCEL_CANDIDATES[0]
    else:
        print(f"\n找到 {len(EXCEL_CANDIDATES)} 個 Excel 檔，請選擇：")
        for i, f in enumerate(EXCEL_CANDIDATES, 1):
            print(f"  {i}. {f.name}")
        choice = input(f"\n輸入編號 [預設 1]: ").strip() or "1"
        excel_path = EXCEL_CANDIDATES[int(choice) - 1]

    print(f"\n📂 讀取 Excel：{excel_path.name}")

    wb = load_workbook(excel_path)
    sheet = wb.active

    # 讀取「訂購須知」
    info = {}
    for row in sheet.iter_rows(min_row=1, max_row=11, values_only=True):
        if row[0] and row[1] and isinstance(row[0], str):
            key = row[0].strip().replace("　", "").replace(" ", "")
            value = str(row[1]).strip() if row[1] else ""
            if key in ["取貨地點", "運費說明", "價格說明", "訂購單位", "截單日期", "備註"]:
                info[key] = value
    print(f"  ✓ 讀取訂購須知：{len(info)} 項")

    # 讀取商品資料
    products = []
    row_to_no = {}
    row = PRODUCT_START_ROW
    while True:
        no = sheet.cell(row=row, column=1).value
        if not isinstance(no, int):
            break
        name_raw = sheet.cell(row=row, column=PRODUCT_NAME_COL).value
        barcode = sheet.cell(row=row, column=PRODUCT_BARCODE_COL).value
        price = sheet.cell(row=row, column=PRODUCT_PRICE_COL).value
        box_qty = sheet.cell(row=row, column=PRODUCT_BOX_COL).value

        if not name_raw:
            row += 1
            continue

        name_full = str(name_raw)

        # 抓出庫存警示
        stock_match = re.search(r"⚠\s*僅剩\s*(\d+)\s*個", name_full)
        stock_left = int(stock_match.group(1)) if stock_match else None
        name_no_stock = re.sub(r"\s*⚠\s*僅剩\s*\d+\s*個\s*", "", name_full).strip()

        # 拆名稱與規格 (用 \n 分行)
        parts = name_no_stock.split("\n")
        main_name = parts[0].strip()
        spec = " ".join(p.strip() for p in parts[1:]) if len(parts) > 1 else ""

        products.append({
            "no": no,
            "name": main_name,
            "spec": spec,
            "barcode": str(barcode).strip() if barcode else "",
            "price": int(price) if price else 0,
            "box_qty": int(box_qty) if box_qty else 0,
            "stock_left": stock_left
        })
        row_to_no[row] = no
        row += 1

    print(f"  ✓ 讀取商品資料：{len(products)} 項")

    # === 處理圖片 ===
    images_dir = SCRIPT_DIR / "images"
    images_dir.mkdir(exist_ok=True)

    # 先記錄舊圖片清單，事後刪除多餘的
    old_images = set(f.name for f in images_dir.iterdir() if f.is_file())
    new_images = set()

    saved_count = 0
    for img in sheet._images:
        anchor_row = img.anchor._from.row + 1
        if anchor_row in row_to_no:
            no = row_to_no[anchor_row]
            data = img._data()
            if data[:3] == b"\xff\xd8\xff":
                ext = "jpg"
            elif data[:8] == b"\x89PNG\r\n\x1a\n":
                ext = "png"
            else:
                ext = "jpg"
            filename = f"product_{no:03d}.{ext}"

            # 移除同 no 但不同副檔名的舊檔
            for other_ext in ("jpg", "png"):
                old_path = images_dir / f"product_{no:03d}.{other_ext}"
                if old_path.exists() and old_path.name != filename:
                    old_path.unlink()

            (images_dir / filename).write_bytes(data)
            new_images.add(filename)
            saved_count += 1
            # 更新 product 中的 filename
            for p in products:
                if p["no"] == no:
                    p["filename"] = filename
                    break

    print(f"  ✓ 寫入商品圖片：{saved_count} 張")

    # 刪除多餘圖片（不在新清單中的）
    expected_filenames = set(p.get("filename") for p in products if p.get("filename"))
    redundant = old_images - expected_filenames
    deleted_count = 0
    for f in redundant:
        try:
            (images_dir / f).unlink()
            deleted_count += 1
        except Exception as e:
            print(f"  ⚠️  無法刪除多餘圖片 {f}: {e}")
    if deleted_count > 0:
        print(f"  ✓ 清除多餘圖片：{deleted_count} 張")

    # 檢查有沒有商品缺少圖片
    missing_img = [p["no"] for p in products if not p.get("filename")]
    if missing_img:
        print(f"  ⚠️  下列商品在 Excel 中沒有圖片：No.{missing_img}")
        print(f"     請在 Excel 對應行貼上圖片後重新執行此腳本")

    # === 產生 products.js ===
    js_lines = ["const PRODUCTS = ["]
    for p in products:
        if not p.get("filename"):
            continue
        name_esc = p["name"].replace("\\", "\\\\").replace('"', '\\"')
        spec_esc = p["spec"].replace("\\", "\\\\").replace('"', '\\"')
        line = f'  {{ no: {p["no"]}, name: "{name_esc}", spec: "{spec_esc}", '
        line += f'barcode: "{p["barcode"]}", price: {p["price"]}, boxQty: {p["box_qty"]}'
        if p["stock_left"] is not None:
            line += f', stockLeft: {p["stock_left"]}'
        line += f', img: "images/{p["filename"]}" }},'
        js_lines.append(line)
    js_lines.append("];")

    js_path = SCRIPT_DIR / "products.js"
    js_path.write_text("\n".join(js_lines) + "\n", encoding="utf-8")
    print(f"  ✓ 更新 products.js")

    # === 產生 info.json (供未來使用，目前 HTML 是寫死的，這份留作備份) ===
    info_path = SCRIPT_DIR / "info.json"
    info_path.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  ✓ 更新 info.json")

    # === 統計 ===
    print(f"\n{'=' * 60}")
    print(f"  ✅ 同步完成！")
    print(f"{'=' * 60}")
    print(f"  • 商品總數：{len(products)} 項")
    print(f"  • 含庫存警示：{sum(1 for p in products if p['stock_left'] is not None)} 項")
    print(f"  • 圖片總數：{len(new_images)} 張")
    print(f"\n  下一步：直接打開 index.html 測試新內容")
    print(f"  或重新部署到 GitHub Pages / Netlify")

    # === 警告：訂購須知有變動 ===
    print(f"\n⚠️  注意：訂購須知區塊目前是寫在 index.html 內，")
    print(f"   如果你修改了 Excel 上半部的取貨地點/運費說明等內容，")
    print(f"   請手動編輯 index.html 對應段落（在 '<section class=\"info-section\">' 內）")
    print(f"   或告訴 Claude 幫你同步更新")

    input("\n按 Enter 關閉視窗...")


if __name__ == "__main__":
    main()
