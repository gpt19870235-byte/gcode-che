# 程式內容檢查工具（Windows 免安裝 Python 版本）

你現在遇到的狀況是正常的：`main.py` 是 **Python 程式**，Windows 沒有安裝 Python 就無法直接開啟。

本資料夾已經準備好 **GitHub Actions（雲端 Windows）自動打包**，會產出 `程式內容檢查工具.exe`，你只要下載 exe 就能在任何 Windows 電腦上直接雙擊執行（不用裝 Python）。

---

## 你要做的事（不用安裝 Python）

1. 在 GitHub 新建一個 repository（例如：`gcode-checker`）
2. 把本資料夾的所有檔案上傳到 repo 根目錄（**包含** `.github/workflows/build_windows_exe.yml`）
3. GitHub → **Actions** → 選 **Build Windows EXE** → **Run workflow**
4. 等流程跑完後，點進該次 workflow → **Artifacts** 下載：
   - `gcode_checker_windows_exe`
5. 解壓後會得到：
   - `程式內容檢查工具.exe`

> 小提醒：第一次跑 Actions 會稍慢一些，因為要安裝 PyInstaller。

---

## 可能會遇到的 SmartScreen / 防毒提示
自製 exe 在公司環境常見：
- 右鍵檔案 → **內容** → 勾選「解除封鎖」→ 套用
- 或 SmartScreen 點「更多資訊」→「仍要執行」
- 公司環境可能需要 IT 加白名單

---

## 進階（可選）
- 想改 exe 名稱：
  - 打開 `.github/workflows/build_windows_exe.yml`，修改 `--name "程式內容檢查工具"`
