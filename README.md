# Schedule to Form 1

專為有 Zone / TRMZ 的 Tree Assessment Schedule 設計（例如 Queen Mary Hospital、Kowloon Hospital）。

## 網頁版（可共用）

開啟：[https://davychan188.github.io/schedule-to-form1/](https://davychan188.github.io/schedule-to-form1/)

把 Tree Assessment Schedule（`.xlsx` / `.xls`）拖進頁面，即可產生 Section 4 / 5 / 6 及 Form 1 Table (A)。檔案只在瀏覽器處理，不會上載到伺服器。可用「列印 / 存 PDF」輸出。

## 桌面版（Python）

由 Schedule 產生 Cover／Form 1 Excel（及可選 Word Cover）。排除 Removed、Duplicate／not in use、Out of LOT、Construction area、Outside boundary。

需要 Python 3，以及：

```bash
pip install pandas openpyxl python-docx
```

```bash
python Schedule_to_Form1.py
```

（Word Cover 需要 `python-docx`；沒有仍可出 Excel。）
