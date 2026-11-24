# 房車帳本

記錄房子和車子的開銷、存錢罐（例如媽媽先放在我這裡的錢）和繳費提醒的網頁 App。
用 Google 或 GitHub 登入，資料存在 Firebase，iPhone、iPad、Mac 會自動同步。

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
4. **GitHub 登入**：
   - Authentication → 登入方式 → GitHub → 啟用，複製畫面上的回呼網址（`https://biledger.firebaseapp.com/__/auth/handler`）。
   - GitHub → Settings → Developer settings → OAuth Apps → New OAuth App：
     Homepage URL 填 `https://abbeyyy049.github.io/Biledger/`，Authorization callback URL 貼上剛剛的回呼網址。
   - 把拿到的 Client ID 和 Client Secret 貼回 Firebase 的 GitHub 設定。

## 發布到 GitHub Pages

repository 的 Settings → Pages → Source 選「Deploy from a branch」，Branch 選 `main`、資料夾選 `/ (root)`，按 Save。

## 資料結構（Firestore）

```
users/{登入者 uid}/meta/config     資料夾、資產、類別、付款方式、存錢罐、固定帳單
users/{uid}/expenses/{id}          每一筆開銷
users/{uid}/moves/{id}             存錢罐的存入和調整
users/{uid}/receipts/{開銷 id}     收據照片（壓縮後存成一份文件）
```
