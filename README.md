# Schedule to Form 1

專為有 Zone / TRMZ 的 Tree Assessment Schedule 設計（例如 Queen Mary Hospital、Kowloon Hospital）。

由 Schedule 產生 Cover／Form 1 Excel（及可選 Word Cover）。排除 Removed、Duplicate／not in use、Out of LOT、Construction area、Outside boundary。

## 執行

需要 Python 3，以及：

```bash
pip install pandas openpyxl python-docx
```

```bash
python Schedule_to_Form1.py
```

（Word Cover 需要 `python-docx`；沒有仍可出 Excel。）

## 來源

Grok Bot 對話附件，2026-11-06 收入目錄。
