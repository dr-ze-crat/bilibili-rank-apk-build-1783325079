[app]

# 应用名称
title = B站排行榜爬虫

# 包名（唯一标识）
package.name = bilibiliranking

# 域名（反写，用于包名）
package.domain = com.bilibili.crawler

# 源码入口
source.dir = .
source.include_exts = py,png,jpg,kv,atlas

# 版本
version = 1.0.0

# 需求
requirements = python3,kivy,requests,openpyxl,plyer,android

# 权限
android.permissions = INTERNET,ACCESS_NETWORK_STATE,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,READ_MEDIA_IMAGES

# Android API 级别
android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b
android.ndk_path =

# 是否启用 AndroidX（新版 Android 支持）
android.enable_androidx = True

# 自动接受 SDK 许可
android.accept_sdk_license = True

# 应用图标（可选，留空使用默认）
android.icon = icon.png

# 是否全屏
android.fullscreen = 0

# 允许的 orientation
android.orientation = portrait

# 存储权限（Android 11+ 使用 MANAGE_EXTERNAL_STORAGE）
android.insert_external_storage = True

# 兼容旧版 Android 存储
android.storage_path = /sdcard

# 是否使用 wakelock（后台保持运行）
android.wakelock = False

# 调试模式（发布时设为 0）
android.debug = 0

# 是否归档为 AAB（Android App Bundle）
android.release_artifact = apk

# 签名
android.keystore =
android.keystore.alias =
android.keystore.password =

# 架构
android.archs = arm64-v8a

# 日志级别
log_level = 2

# 应用需求
presplash.filename = presplash.png
presplash.color = #1E1E24

# 支持的平台
osx.package_name = BilibiliRanking
osx.bundle_identifier = com.bilibili.crawler.ranking
osx.icon =

[buildozer]
# 编译输出目录
log_level = 1
warn_on_root = 1
