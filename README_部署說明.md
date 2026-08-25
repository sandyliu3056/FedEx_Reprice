# FedEx Reprice Studio — 部署包

網頁版 FedEx 重新計價工具。**計價引擎直接使用桌面版的 `FedEx_RepricingTool.py`(一行未改)**,
所以網頁版和桌面版算出來的結果保證一致;桌面版更新時,覆蓋這個檔案重新部署即可。

## 包裡有什麼

| 檔案 | 用途 |
|---|---|
| `app.py` | 網頁介面(Streamlit) |
| `adapter.py` | 無頭轉接層:讓桌面版程式在伺服器上跑(不需要視窗) |
| `FedEx_RepricingTool.py` | 你的桌面版原始程式 = 計價邏輯唯一來源 |
| `requirements.txt` | Python 套件清單 |
| `.streamlit/config.toml` | 佈景主題(FedEx 紫/橘) |
| `Dockerfile` | 給 Render / 其他 Docker 平台用(方案 B) |

---

## 方案 A:Streamlit Community Cloud(免費,推薦)

1. 到 GitHub 建一個 **新的私人 repo**(例如 `fedex-reprice-web`),
   用網頁介面把這個資料夾裡的檔案全部上傳(含 `.streamlit` 資料夾)。
2. 開 <https://share.streamlit.io> → 用 GitHub 帳號登入 → **New app**。
3. 選 repo `fedex-reprice-web`、branch `main`、Main file path 填 `app.py` → **Deploy**。
4. 等 1–2 分鐘,就會拿到一個 `https://xxxx.streamlit.app` 的網址。

### 設定密碼(建議)
帳單是敏感資料,部署完請設密碼:
App 右下角 **Manage app → Settings → Secrets**,貼上:

```toml
APP_PASSWORD = "你要的密碼"
```

存檔後 app 會重啟,之後開網頁要先輸入密碼。

### 更新
改了任何檔案(例如換新的 `FedEx_RepricingTool.py`),
直接在 GitHub 上傳覆蓋,Streamlit Cloud 會自動重新部署。

> 注意:免費層閒置一段時間會休眠,第一次打開要等 30 秒左右喚醒,屬正常現象。

---

## 方案 B:Render(Docker)

1. 一樣把整個資料夾放上 GitHub repo。
2. 到 <https://render.com> → **New → Web Service** → 連 GitHub 選這個 repo。
3. Runtime 選 **Docker**(它會自己用包裡的 `Dockerfile`)→ Create。
4. 密碼:在 Render 的 **Environment** 加不了 Streamlit secrets,
   改成建立檔案 `.streamlit/secrets.toml`(內容同上)一起上傳,
   或用 Render 的 Secret File 功能放到 `/app/.streamlit/secrets.toml`。

---

## 本機測試(部署前先跑一次)

```bash
pip install -r requirements.txt
streamlit run app.py
```

瀏覽器開 <http://localhost:8501>。

---

## 使用方式

1. 左側 **匯入設定 JSON**:直接用桌面版存的 `billing_tool_config.json`;
   網頁上改的費率也可 **下載設定 JSON** 拿回桌面版用(同一格式,雙向互通)。
2. **執行 Reprice**:上傳 FedEx 帳單(CSV / XLSX)→ Generate → 下載 `_repriced.xlsx`。
3. **全域規則 / 費率表** 分頁可直接在網頁上編輯。

## 資料與隱私

- 上傳的帳單只在計算當下存在暫存檔,算完即刪,不會保存在伺服器上。
- 網頁上的設定改動只存在當前工作階段;要長期保存請下載設定 JSON。
- 請務必設 `APP_PASSWORD`,不要把網址公開。
