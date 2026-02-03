# Windows EXE（不需要安裝 Python）—用 GitHub Actions 自動打包

你要「Windows 沒有安裝 Python 也能開啟」：做法就是先產出 `.exe`。
因為 PyInstaller **不能跨平台**，一定要在 Windows 環境打包。
這份專案已經幫你準備好 GitHub Actions，丟上 GitHub 後會自動在 Windows 產出 `.exe`，下載即可使用。

---

## 步驟（完全不需要在你電腦安裝 Python）

1. 到 GitHub 建一個新的 repository（例如 `gcode-checker`）
2. 把此資料夾內的所有檔案上傳到 repo 根目錄（包含 `.github/workflows/...`）
3. 進入 GitHub → **Actions** → 選 `Build Windows EXE`
4. 點右側 **Run workflow**（或 push 到 main 也會自動跑）
5. 等流程完成後，點進該次 workflow → 下載 Artifacts：
   - `gcode_checker_windows_exe`
6. 解壓後拿到：
   - `程式內容檢查工具.exe`

在任何 Windows 電腦上 **直接雙擊 exe 就能開**，不需要 Python。

---

## 如果公司擋 SmartScreen / 防毒
這是自製 exe 常見情況：
- 右鍵檔案 → 內容 → 勾選「解除封鎖」→ 套用
- 或 SmartScreen 點「更多資訊」→「仍要執行」
- 公司環境可能需要 IT 白名單

---

## 改 EXE 名稱
在 `.github/workflows/build_windows_exe.yml` 裡把：
`--name "程式內容檢查工具"` 改成你想要的名字即可。
