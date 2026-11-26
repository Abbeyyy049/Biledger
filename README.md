# Biledger 比記本

記帳的網頁 App：記錄收入和支出、管理資產（房子、車子等，選填）和存錢罐（預先存好、專門用來付某些開銷的錢），提醒繳費，並依固定帳單算出年度預算和每期該存多少。
可以用 Google、GitHub 或 Email 登入，資料存在 Firebase，iPhone、iPad、Mac 會自動同步。

網址：https://abbeyyy049.github.io/Biledger/

## 檔案

| 檔案 | 用途 |
|---|---|
| `index.html`、`styles.css` | 頁面和樣式 |
| `js/app.js` | 所有畫面和操作 |
| `js/backend.js` | Firebase 登入和雲端資料庫 |
| `js/excel.js` | 匯出和匯入 Excel |
| `sw.js`、`manifest.webmanifest`、`icons/` | 加到 iPhone 主畫面、離線也能打開 |
| `firestore.rules` | 資料庫安全規則（要貼到 Firebase 主控台） |

帳務資料**不在**這個 repository 裡，只存在 Firebase，而且只有登入的本人讀得到。

## 第一次設定（Firebase 主控台）

1. **Firestore 安全規則**：Firestore Database → 規則，把 `firestore.rules` 的內容整段貼上，按「發布」。
2. **授權網域**：Authentication → 設定 → 授權網域 → 新增 `abbeyyy049.github.io`。
3. **Google 登入**：Authentication → 登入方式 → Google → 啟用。
   **Email 登入**：Authentication → 登入方式 → 電子郵件/密碼 → 啟用（不用開「電子郵件連結」）。
4. **GitHub 登入**：
   - Authentication → 登入方式 → GitHub → 啟用，複製畫面上的回呼網址（`https://biledger.firebaseapp.com/__/auth/handler`）。
   - GitHub → Settings → Developer settings → OAuth Apps → New OAuth App：
     Homepage URL 填 `https://abbeyyy049.github.io/Biledger/`，Authorization callback URL 貼上剛剛的回呼網址。
   - 把拿到的 Client ID 和 Client Secret 貼回 Firebase 的 GitHub 設定。

## 發布到 GitHub Pages

repository 的 Settings → Pages → Source 選「Deploy from a branch」，Branch 選 `main`、資料夾選 `/ (root)`，按 Save。

## 資料結構（Firestore）

```
ledgers/{帳本 id}                      名稱、擁有者、成員 {uid: 權限}、邀請 {email: 權限}、成員名字
ledgers/{id}/meta/config               分組、資產、類別、付款方式、存錢罐、固定帳單、貨幣、財年
ledgers/{id}/expenses|moves/{id}       每一筆收支、存錢罐的存入和調整
ledgers/{id}/receipts/{開銷 id}        收據照片（壓縮後存成一份文件）
ledgers/{id}/icons|comments/{id}       自己上傳的圖示、每筆紀錄下的留言
users/{uid}/meta/prefs                 個人偏好（字體、上次開的帳本）
users/{uid}/...                        舊版資料，第一次登入新版時會自動搬進「房車帳本」
```

權限：擁有者（全部）、可以記帳、可以留言、只能看。邀請要對方用同一個 Email 登入，而且 Email 要先驗證。
`firestore.rules` 是安全規則，更新後要整段貼到 Firebase 主控台 → Firestore Database → 規則 → 發布。

## 預算怎麼算（不用 AI）

每筆固定帳單的下次金額 = 最近一年的平均 × 季節係數 × 趨勢：

- **平均**：最近 12／4／2／1 次（月繳／季繳／半年繳／年繳）的金額平均
- **季節係數**：前後一個月內同月份的平均 ÷ 最近 12 個月平均（要有 11 個月以上的資料，限制在 0.6–1.6）
- **趨勢**：有季節資料時，用最近 12 個月平均 ÷ 前 12 個月平均 − 1；資料較少時用直線迴歸的斜率 ÷ 平均（限制在每年 ±15%）
- 手動填的「預估金額」一律優先；完全沒有資料時用「初始預算」

年度預算 =（本財年已付 + 到年底的預估）×（1 + 緩衝比例），可以鎖定。
存錢計畫：建議每期金額是「每一筆帳單到期時，帳單帳戶都夠付」的最小固定金額，和「平均每期」取比較大的。
