# B站排行榜爬虫 - Android APK

将 B站全站排行榜爬虫打包成 Android APK，可在手机上直接运行抓取排行榜数据。

## 功能

- 抓取 B站全站排行榜 Top 100
- 抓取实时热搜词 Top 10
- 生成格式化的 Excel 文件（.xlsx）
- 支持分享文件到微信、QQ、文件管理器等
- 深色主题，触摸友好界面

## 安装 APK

APK 文件安装到 Android 手机即可使用，首次打开会请求网络和存储权限，同意后点击「开始抓取」按钮。

生成的 Excel 文件保存在手机的 `Download/B站排行榜_时间.xlsx`

## 构建方法

### 方法一：GitHub Actions 自动构建（推荐）

无需在本机搭建环境，使用 GitHub 免费服务器自动构建：

1. 在 GitHub 创建一个新仓库
2. 将本文件夹所有文件推送上去：
   ```
   git init
   git add .
   git commit -m "添加B站排行榜爬虫APK"
   git remote add origin https://github.com/你的用户名/仓库名.git
   git push -u origin main
   ```
3. 在 GitHub 仓库页面点击 **Actions** → **Build APK** → **Run workflow**
4. 等待 20-30 分钟，构建完成后在 Action 页面下载 **APK 产物**（`B站排行榜爬虫-APK.zip`）
5. 解压得到 `.apk` 文件，传到手机安装即可

### 方法二：WSL2（Windows 用户）

如果你有 WSL2（Windows Subsystem for Linux）：

```bash
# 1. 在 WSL2 中
sudo apt update && sudo apt install -y git python3-pip
pip install buildozer cython

# 2. 将本项目复制到 WSL2 中
cd ~
git clone https://github.com/你的用户名/仓库名.git  # 或直接复制文件夹

# 3. 开始构建（首次会下载 SDK/NDK，约 2GB）
cd B站热门排行榜爬虫_APK
buildozer android debug

# 4. 构建完成后 APK 在 bin/ 目录下
```

### 方法三：Google Colab

使用在线 Colab 笔记本构建（无需本地环境）：

<a href="https://colab.research.google.com/github" target="_blank">
  <img src="https://colab.research.google.com/assets/colab-badge.svg" alt="Open In Colab">
</a>

```python
# Colab 构建脚本（在 Colab 中运行）
!git clone https://github.com/你的用户名/仓库名.git
%cd 仓库名/B站热门排行榜爬虫_APK
!pip install buildozer cython
!buildozer android debug
# 下载 bin/*.apk 文件
```

## 项目文件说明

| 文件 | 说明 |
|------|------|
| `main.py` | Kivy Android 应用主程序 |
| `buildozer.spec` | Buildozer 构建配置文件 |
| `requirements.txt` | Python 依赖清单 |
| `icon.png` | 应用图标 |
| `presplash.png` | 启动画面 |
| `.github/workflows/build-apk.yml` | GitHub Actions 自动构建配置 |
