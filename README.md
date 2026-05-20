# 月銷售排行產生器

整合多個電子書平台的原始銷售報表，自動產生月銷售排行 Excel。

## 支援平台

Google Play、Kobo、Readmoo、博客來、亂搭、Pubu、誠品、PChome

## 安裝方式

### 1. 安裝 Python

如果還沒有安裝，前往 [python.org](https://www.python.org) 下載安裝。
建議版本：Python 3.10 以上。

### 2. 安裝依賴套件

在終端機（或命令提示字元）執行：

```
pip install -r requirements.txt
```

### 3. 啟動應用程式

```
streamlit run app.py
```

執行後，瀏覽器會自動打開網頁介面。

## 使用方式

1. 在「當月月份」欄位輸入月份，格式為 YYYYMM（例如 202504）
2. 選填：上傳上個月的排行 Excel（用來保留歷史分頁）
3. 選填：上傳書種目錄（product_fixed.xlsx）
4. 上傳各平台的原始報表
5. 點擊「開始產生」
6. 完成後下載結果 Excel
