import csv
import ctypes
import base64
import hashlib
import http.server
import html
import json
import os
import re
import secrets
import shutil
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from ctypes import wintypes
from dataclasses import dataclass
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Bookmarks to CSV / Google Sheets"
OUTPUT_FOLDER_NAME = "bookmarks"
GOOGLE_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
TOKEN_CACHE_FILE = "google_sheets_oauth_token.dat"
SETTINGS_FILE = "bookmark_tool_settings.json"
ICON_FILE = "app_icon.ico"
WM_DROPFILES = 0x0233
GWL_WNDPROC = -4
CRYPTPROTECT_UI_FORBIDDEN = 0x01


LANGUAGES = {
    "zh": "中文",
    "en": "English",
    "vi": "Tiếng Việt",
    "my": "မြန်မာ",
    "hi": "हिन्दी",
    "fil": "Filipino",
    "pt": "Português",
    "es": "Español",
}

BROWSER_CODES = ("chrome", "edge")
BROWSER_LABEL_KEYS = {
    "chrome": "browser_chrome",
    "edge": "browser_edge",
}


TEXT = {
    "zh": {
        "app_title": "书签导出工具",
        "language": "界面语言",
        "tab_html": "HTML 转 CSV",
        "tab_profiles": "扫描浏览器用户",
        "tab_sheets": "导出到 Google 表格",
        "output_folder": "输出文件夹",
        "desktop_bookmarks": "桌面 bookmarks 文件夹",
        "add_html": "添加 HTML 文件",
        "clear": "清空",
        "convert_html": "转换为 CSV",
        "html_hint": "可一次选择多个 Chrome 导出的书签 HTML 文件，或直接拖到下方空白列表区域；每个 HTML 会生成一个 CSV。",
        "browser": "浏览器",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "浏览器用户资料目录",
        "browse": "浏览",
        "scan": "扫描",
        "convert_profiles": "导出 CSV",
        "mode": "导出方式",
        "mode_separate": "每个用户一个 CSV",
        "mode_merged": "全部用户合并成一个 CSV",
        "profiles_found": "找到的用户资料",
        "selected": "选择",
        "select_all": "全选",
        "clear_selection": "清空选择",
        "sheet_url": "Google 表格链接",
        "open_sheet": "打开表格",
        "api_key": "Google Sheet API Key",
        "oauth_client": "OAuth 客户端 JSON",
        "sheet_name": "Sheet 名称",
        "upload": "上传到 Google 表格",
        "sheets_hint": "程序会从所选浏览器用户资料读取 Bookmarks，通过 OAuth 写入 Google 表格。每个浏览器用户会自动创建一个同名分表，并把数据写入对应分表。",
        "status_ready": "准备好了。",
        "status_working": "正在处理...",
        "status_done": "完成。",
        "no_html": "请先添加至少一个 HTML 文件。",
        "no_profiles": "没有找到 Bookmarks 文件。",
        "no_profiles_selected": "请至少勾选一个用户资料。",
        "no_sheet_url": "请输入 Google 表格链接。",
        "no_api_key": "请输入 API key。",
        "no_oauth_client": "请选择 OAuth 客户端 JSON 文件。",
        "done_html": "已转换 {count} 个 HTML 文件。\n输出：{folder}",
        "done_profiles": "已导出 {count} 个 CSV 文件。\n输出：{folder}",
        "done_merged": "已导出合并 CSV。\n输出：{file}",
        "done_upload": "已上传 {rows} 条书签到 Google 表格。",
        "done_upload_profiles": "已上传 {profiles} 个用户，共 {rows} 条书签到 Google 表格。",
        "error": "错误",
        "info": "提示",
        "warning": "注意",
        "copied_bookmarks": "已复制被占用的 Bookmarks 文件。",
        "api_key_write_error": "Google 拒绝了写入请求。通常原因是 API key 不能授权写入 Google Sheets；请使用已授权的方式，例如 OAuth、服务账号，或先导出 CSV 后手动导入。",
        "choose_output": "选择输出文件夹",
        "choose_oauth_client": "选择 OAuth 客户端 JSON",
        "oauth_client_invalid": "OAuth 客户端 JSON 无效，请确认它是 Google Cloud 下载的桌面应用客户端 JSON。",
        "oauth_browser": "请在打开的浏览器中登录 Google 并授权访问表格。",
        "drop_added": "已添加 {count} 个 HTML 文件。",
        "drop_unsupported": "请拖入 .html 或 .htm 书签文件。",
        "open_output": "打开输出文件夹",
        "selected_files": "已选择文件",
        "empty": "无",
        "rows": "书签数量",
        "profile": "用户",
    },
    "en": {
        "app_title": "Bookmark Export Tool",
        "language": "Language",
        "tab_html": "HTML to CSV",
        "tab_profiles": "Scan Browser Profiles",
        "tab_sheets": "Export to Google Sheets",
        "output_folder": "Output folder",
        "desktop_bookmarks": "Desktop bookmarks folder",
        "add_html": "Add HTML files",
        "clear": "Clear",
        "convert_html": "Convert to CSV",
        "html_hint": "Select one or more Chrome bookmark HTML exports, or drop them into the blank list area below. Each HTML file becomes one CSV.",
        "browser": "Browser",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "Browser user data folder",
        "browse": "Browse",
        "scan": "Scan",
        "convert_profiles": "Export CSV",
        "mode": "Export mode",
        "mode_separate": "One CSV per profile",
        "mode_merged": "Merge all profiles into one CSV",
        "profiles_found": "Profiles found",
        "selected": "Select",
        "select_all": "Select all",
        "clear_selection": "Clear selection",
        "sheet_url": "Google Sheet link",
        "open_sheet": "Open sheet",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Sheet name",
        "upload": "Upload to Google Sheets",
        "sheets_hint": "The app reads Bookmarks from the selected browser and writes to Google Sheets with OAuth. Each browser profile gets a separate sheet tab named after the profile.",
        "status_ready": "Ready.",
        "status_working": "Working...",
        "status_done": "Done.",
        "no_html": "Please add at least one HTML file.",
        "no_profiles": "No Bookmarks files were found.",
        "no_profiles_selected": "Please select at least one profile.",
        "no_sheet_url": "Please enter a Google Sheet link.",
        "no_api_key": "Please enter an API key.",
        "no_oauth_client": "Please select an OAuth client JSON file.",
        "done_html": "Converted {count} HTML file(s).\nOutput: {folder}",
        "done_profiles": "Exported {count} CSV file(s).\nOutput: {folder}",
        "done_merged": "Exported merged CSV.\nOutput: {file}",
        "done_upload": "Uploaded {rows} bookmarks to Google Sheets.",
        "done_upload_profiles": "Uploaded {profiles} profile(s), {rows} bookmark(s) total, to Google Sheets.",
        "error": "Error",
        "info": "Info",
        "warning": "Note",
        "copied_bookmarks": "Copied a locked Bookmarks file.",
        "api_key_write_error": "Google rejected the write request. The usual reason is that an API key cannot authorize writes to Google Sheets. Use OAuth, a service account, or export CSV and import it manually.",
        "choose_output": "Choose output folder",
        "choose_oauth_client": "Choose OAuth client JSON",
        "oauth_client_invalid": "Invalid OAuth client JSON. Please use a Desktop app client JSON downloaded from Google Cloud.",
        "oauth_browser": "Please sign in and authorize Google Sheets access in the browser window that opens.",
        "drop_added": "Added {count} HTML file(s).",
        "drop_unsupported": "Please drop .html or .htm bookmark files.",
        "open_output": "Open output folder",
        "selected_files": "Selected files",
        "empty": "None",
        "rows": "Bookmarks",
        "profile": "Profile",
    },
    "vi": {
        "app_title": "Công cụ xuất dấu trang",
        "language": "Ngôn ngữ",
        "tab_html": "HTML sang CSV",
        "tab_profiles": "Quét hồ sơ trình duyệt",
        "tab_sheets": "Xuất sang Google Sheets",
        "output_folder": "Thư mục xuất",
        "desktop_bookmarks": "Thư mục bookmarks trên Desktop",
        "add_html": "Thêm tệp HTML",
        "clear": "Xóa",
        "convert_html": "Chuyển sang CSV",
        "html_hint": "Chọn một hoặc nhiều tệp HTML dấu trang Chrome. Mỗi tệp tạo một CSV.",
        "browser": "Trình duyệt",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "Thư mục dữ liệu người dùng trình duyệt",
        "browse": "Chọn",
        "scan": "Quét",
        "convert_profiles": "Xuất CSV",
        "mode": "Kiểu xuất",
        "mode_separate": "Mỗi hồ sơ một CSV",
        "mode_merged": "Gộp tất cả vào một CSV",
        "profiles_found": "Hồ sơ tìm thấy",
        "selected": "Chọn",
        "select_all": "Chọn tất cả",
        "clear_selection": "Bỏ chọn",
        "sheet_url": "Liên kết Google Sheet",
        "open_sheet": "Mở sheet",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Tên sheet",
        "upload": "Tải lên Google Sheets",
        "sheets_hint": "Ứng dụng đọc Bookmarks từ trình duyệt đã chọn và thêm vào Google Sheets. Lưu ý: ghi vào Google Sheets thường cần OAuth hoặc service account; chỉ API key có thể bị Google từ chối.",
        "status_ready": "Sẵn sàng.",
        "status_working": "Đang xử lý...",
        "status_done": "Hoàn tất.",
        "no_html": "Vui lòng thêm ít nhất một tệp HTML.",
        "no_profiles": "Không tìm thấy tệp Bookmarks.",
        "no_profiles_selected": "Vui lòng chọn ít nhất một hồ sơ.",
        "no_sheet_url": "Vui lòng nhập liên kết Google Sheet.",
        "no_api_key": "Vui lòng nhập API key.",
        "no_oauth_client": "Vui lòng chọn tệp OAuth client JSON.",
        "done_html": "Đã chuyển {count} tệp HTML.\nXuất tại: {folder}",
        "done_profiles": "Đã xuất {count} tệp CSV.\nXuất tại: {folder}",
        "done_merged": "Đã xuất CSV gộp.\nTệp: {file}",
        "done_upload": "Đã tải {rows} dấu trang lên Google Sheets.",
        "done_upload_profiles": "Đã tải {profiles} hồ sơ, tổng cộng {rows} dấu trang, lên Google Sheets.",
        "error": "Lỗi",
        "info": "Thông tin",
        "warning": "Lưu ý",
        "copied_bookmarks": "Đã sao chép tệp Bookmarks đang bị khóa.",
        "api_key_write_error": "Google từ chối yêu cầu ghi. Thường là vì API key không thể cấp quyền ghi vào Google Sheets. Hãy dùng OAuth, service account, hoặc xuất CSV rồi nhập thủ công.",
        "choose_output": "Chọn thư mục xuất",
        "choose_oauth_client": "Chọn OAuth client JSON",
        "oauth_client_invalid": "OAuth client JSON không hợp lệ. Hãy dùng tệp Desktop app client JSON tải từ Google Cloud.",
        "oauth_browser": "Vui lòng đăng nhập và cho phép truy cập Google Sheets trong cửa sổ trình duyệt được mở.",
        "drop_added": "Đã thêm {count} tệp HTML.",
        "drop_unsupported": "Vui lòng kéo tệp bookmark .html hoặc .htm.",
        "open_output": "Mở thư mục xuất",
        "selected_files": "Tệp đã chọn",
        "empty": "Không có",
        "rows": "Dấu trang",
        "profile": "Hồ sơ",
    },
    "my": {
        "app_title": "Bookmark ထုတ်ရန်ကိရိယာ",
        "language": "ဘာသာစကား",
        "tab_html": "HTML မှ CSV",
        "tab_profiles": "Browser ပရိုဖိုင် စစ်ဆေးရန်",
        "tab_sheets": "Google Sheets သို့ ထုတ်ရန်",
        "output_folder": "ထုတ်မည့်ဖိုင်တွဲ",
        "desktop_bookmarks": "Desktop bookmarks ဖိုင်တွဲ",
        "add_html": "HTML ဖိုင်ထည့်ရန်",
        "clear": "ရှင်းရန်",
        "convert_html": "CSV သို့ပြောင်းရန်",
        "html_hint": "Chrome မှ export လုပ်ထားသော bookmark HTML ဖိုင်များကို ရွေးပါ။ HTML တစ်ခုစီ CSV တစ်ခုဖြစ်မည်။",
        "browser": "Browser",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "Browser user data ဖိုင်တွဲ",
        "browse": "ရွေးရန်",
        "scan": "စစ်ရန်",
        "convert_profiles": "CSV ထုတ်ရန်",
        "mode": "ထုတ်နည်း",
        "mode_separate": "ပရိုဖိုင်တစ်ခု CSV တစ်ခု",
        "mode_merged": "ပရိုဖိုင်အားလုံး CSV တစ်ခု",
        "profiles_found": "တွေ့ရှိသော ပရိုဖိုင်များ",
        "selected": "ရွေးရန်",
        "select_all": "အားလုံးရွေးရန်",
        "clear_selection": "ရွေးထားမှု ရှင်းရန်",
        "sheet_url": "Google Sheet လင့်ခ်",
        "open_sheet": "Sheet ဖွင့်ရန်",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Sheet အမည်",
        "upload": "Google Sheets သို့ တင်ရန်",
        "sheets_hint": "အက်ပ်သည် ရွေးထားသော browser မှ Bookmarks ကိုဖတ်ပြီး Google Sheets ထဲသို့ ထည့်ပါမည်။ သတိပြုရန် - Google Sheets တွင်ရေးရန် OAuth သို့မဟုတ် service account လိုလေ့ရှိပြီး API key သက်သက်ကို Google ငြင်းနိုင်သည်။",
        "status_ready": "အသင့်ဖြစ်ပါပြီ။",
        "status_working": "လုပ်ဆောင်နေသည်...",
        "status_done": "ပြီးပါပြီ။",
        "no_html": "HTML ဖိုင် အနည်းဆုံးတစ်ခု ထည့်ပါ။",
        "no_profiles": "Bookmarks ဖိုင် မတွေ့ပါ။",
        "no_profiles_selected": "အနည်းဆုံး ပရိုဖိုင်တစ်ခု ရွေးပါ။",
        "no_sheet_url": "Google Sheet လင့်ခ် ထည့်ပါ။",
        "no_api_key": "API key ထည့်ပါ။",
        "no_oauth_client": "OAuth client JSON ဖိုင် ရွေးပါ။",
        "done_html": "HTML ဖိုင် {count} ခု ပြောင်းပြီးပါပြီ။\nနေရာ: {folder}",
        "done_profiles": "CSV ဖိုင် {count} ခု ထုတ်ပြီးပါပြီ။\nနေရာ: {folder}",
        "done_merged": "ပေါင်းထားသော CSV ထုတ်ပြီးပါပြီ။\nဖိုင်: {file}",
        "done_upload": "Bookmark {rows} ခု Google Sheets သို့ တင်ပြီးပါပြီ။",
        "done_upload_profiles": "ပရိုဖိုင် {profiles} ခု၊ စုစုပေါင်း Bookmark {rows} ခု Google Sheets သို့ တင်ပြီးပါပြီ။",
        "error": "အမှား",
        "info": "အချက်အလက်",
        "warning": "သတိပြုရန်",
        "copied_bookmarks": "အသုံးပြုနေသော Bookmarks ဖိုင်ကို ကူးယူပြီးပါပြီ။",
        "api_key_write_error": "Google က ရေးသားမှုကို ငြင်းပါသည်။ များသောအားဖြင့် API key သည် Google Sheets ရေးခွင့် မပေးနိုင်သောကြောင့်ဖြစ်သည်။ OAuth, service account သို့မဟုတ် CSV export/import ကိုသုံးပါ။",
        "choose_output": "ထုတ်မည့်ဖိုင်တွဲ ရွေးရန်",
        "choose_oauth_client": "OAuth client JSON ရွေးရန်",
        "oauth_client_invalid": "OAuth client JSON မမှန်ပါ။ Google Cloud မှ Desktop app client JSON ကို အသုံးပြုပါ။",
        "oauth_browser": "ဖွင့်လာသော browser တွင် Google သို့ဝင်ပြီး Google Sheets ခွင့်ပြုပါ။",
        "drop_added": "HTML ဖိုင် {count} ခု ထည့်ပြီးပါပြီ။",
        "drop_unsupported": ".html သို့မဟုတ် .htm bookmark ဖိုင်များကို ဆွဲထည့်ပါ။",
        "open_output": "ထုတ်ထားသောဖိုင်တွဲ ဖွင့်ရန်",
        "selected_files": "ရွေးထားသောဖိုင်များ",
        "empty": "မရှိပါ",
        "rows": "Bookmarks",
        "profile": "ပရိုဖိုင်",
    },
    "hi": {
        "app_title": "बुकमार्क निर्यात उपकरण",
        "language": "भाषा",
        "tab_html": "HTML से CSV",
        "tab_profiles": "ब्राउज़र प्रोफाइल स्कैन",
        "tab_sheets": "Google Sheets में निर्यात",
        "output_folder": "आउटपुट फ़ोल्डर",
        "desktop_bookmarks": "Desktop bookmarks फ़ोल्डर",
        "add_html": "HTML फ़ाइलें जोड़ें",
        "clear": "साफ़ करें",
        "convert_html": "CSV में बदलें",
        "html_hint": "Chrome से निर्यात की गई एक या अधिक bookmark HTML फ़ाइलें चुनें। हर HTML से एक CSV बनेगा।",
        "browser": "ब्राउज़र",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "ब्राउज़र user data फ़ोल्डर",
        "browse": "ब्राउज़",
        "scan": "स्कैन",
        "convert_profiles": "CSV निर्यात",
        "mode": "निर्यात प्रकार",
        "mode_separate": "हर प्रोफाइल के लिए अलग CSV",
        "mode_merged": "सभी प्रोफाइल एक CSV में",
        "profiles_found": "मिले हुए प्रोफाइल",
        "selected": "चुनें",
        "select_all": "सभी चुनें",
        "clear_selection": "चयन साफ़ करें",
        "sheet_url": "Google Sheet लिंक",
        "open_sheet": "Sheet खोलें",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Sheet नाम",
        "upload": "Google Sheets में अपलोड",
        "sheets_hint": "ऐप चुने गए ब्राउज़र के Bookmarks पढ़कर Google Sheets में जोड़ता है। ध्यान दें: Google Sheets में लिखने के लिए आमतौर पर OAuth या service account चाहिए; केवल API key को Google अस्वीकार कर सकता है।",
        "status_ready": "तैयार।",
        "status_working": "काम चल रहा है...",
        "status_done": "पूरा हुआ।",
        "no_html": "कृपया कम से कम एक HTML फ़ाइल जोड़ें।",
        "no_profiles": "कोई Bookmarks फ़ाइल नहीं मिली।",
        "no_profiles_selected": "कृपया कम से कम एक प्रोफाइल चुनें।",
        "no_sheet_url": "कृपया Google Sheet लिंक दर्ज करें।",
        "no_api_key": "कृपया API key दर्ज करें।",
        "no_oauth_client": "कृपया OAuth client JSON फ़ाइल चुनें।",
        "done_html": "{count} HTML फ़ाइलें बदली गईं।\nआउटपुट: {folder}",
        "done_profiles": "{count} CSV फ़ाइलें निर्यात की गईं।\nआउटपुट: {folder}",
        "done_merged": "मर्ज CSV निर्यात हो गया।\nफ़ाइल: {file}",
        "done_upload": "{rows} bookmarks Google Sheets में अपलोड हो गए।",
        "done_upload_profiles": "{profiles} प्रोफाइल और कुल {rows} bookmarks Google Sheets में अपलोड हो गए।",
        "error": "त्रुटि",
        "info": "जानकारी",
        "warning": "ध्यान दें",
        "copied_bookmarks": "लॉक Bookmarks फ़ाइल कॉपी की गई।",
        "api_key_write_error": "Google ने लिखने का अनुरोध अस्वीकार किया। सामान्य कारण है कि API key Google Sheets में लिखने की अनुमति नहीं देती। OAuth, service account, या CSV export/import का उपयोग करें।",
        "choose_output": "आउटपुट फ़ोल्डर चुनें",
        "choose_oauth_client": "OAuth client JSON चुनें",
        "oauth_client_invalid": "OAuth client JSON अमान्य है। Google Cloud से डाउनलोड की गई Desktop app client JSON का उपयोग करें।",
        "oauth_browser": "खुली हुई browser window में Google में sign in करें और Google Sheets access allow करें।",
        "drop_added": "{count} HTML फ़ाइलें जोड़ी गईं।",
        "drop_unsupported": "कृपया .html या .htm bookmark फ़ाइलें ड्रॉप करें।",
        "open_output": "आउटपुट फ़ोल्डर खोलें",
        "selected_files": "चुनी गई फ़ाइलें",
        "empty": "कोई नहीं",
        "rows": "Bookmarks",
        "profile": "प्रोफाइल",
    },
    "fil": {
        "app_title": "Bookmark Export Tool",
        "language": "Wika",
        "tab_html": "HTML papuntang CSV",
        "tab_profiles": "I-scan ang Browser Profiles",
        "tab_sheets": "I-export sa Google Sheets",
        "output_folder": "Output folder",
        "desktop_bookmarks": "Desktop bookmarks folder",
        "add_html": "Magdagdag ng HTML files",
        "clear": "I-clear",
        "convert_html": "Gawing CSV",
        "html_hint": "Pumili ng isa o higit pang Chrome bookmark HTML exports. Bawat HTML ay magiging isang CSV.",
        "browser": "Browser",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "Browser user data folder",
        "browse": "Browse",
        "scan": "Scan",
        "convert_profiles": "I-export CSV",
        "mode": "Paraan ng export",
        "mode_separate": "Isang CSV bawat profile",
        "mode_merged": "Pagsamahin lahat sa isang CSV",
        "profiles_found": "Nahanap na profiles",
        "selected": "Piliin",
        "select_all": "Piliin lahat",
        "clear_selection": "I-clear selection",
        "sheet_url": "Google Sheet link",
        "open_sheet": "Buksan ang sheet",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Sheet name",
        "upload": "I-upload sa Google Sheets",
        "sheets_hint": "Babasahin ng app ang Bookmarks mula sa napiling browser at idaragdag sa Google Sheets. Paalala: kadalasang kailangan ng OAuth o service account para magsulat sa Google Sheets; maaaring tanggihan ng Google ang API key lang.",
        "status_ready": "Handa na.",
        "status_working": "Ginagawa...",
        "status_done": "Tapos na.",
        "no_html": "Magdagdag muna ng kahit isang HTML file.",
        "no_profiles": "Walang Bookmarks file na nahanap.",
        "no_profiles_selected": "Pumili ng kahit isang profile.",
        "no_sheet_url": "Ilagay ang Google Sheet link.",
        "no_api_key": "Ilagay ang API key.",
        "no_oauth_client": "Pumili ng OAuth client JSON file.",
        "done_html": "Na-convert ang {count} HTML file(s).\nOutput: {folder}",
        "done_profiles": "Na-export ang {count} CSV file(s).\nOutput: {folder}",
        "done_merged": "Na-export ang pinagsamang CSV.\nFile: {file}",
        "done_upload": "Na-upload ang {rows} bookmarks sa Google Sheets.",
        "done_upload_profiles": "Na-upload ang {profiles} profile(s), kabuuang {rows} bookmark(s), sa Google Sheets.",
        "error": "Error",
        "info": "Info",
        "warning": "Paalala",
        "copied_bookmarks": "Nakopya ang naka-lock na Bookmarks file.",
        "api_key_write_error": "Tinanggihan ng Google ang write request. Karaniwang dahilan: hindi nakakapagbigay ng write permission ang API key sa Google Sheets. Gumamit ng OAuth, service account, o CSV export/import.",
        "choose_output": "Piliin ang output folder",
        "choose_oauth_client": "Piliin ang OAuth client JSON",
        "oauth_client_invalid": "Hindi valid ang OAuth client JSON. Gamitin ang Desktop app client JSON na na-download mula sa Google Cloud.",
        "oauth_browser": "Mag-sign in at payagan ang Google Sheets access sa browser window na bubukas.",
        "drop_added": "Nagdagdag ng {count} HTML file(s).",
        "drop_unsupported": "Mag-drop ng .html o .htm bookmark files.",
        "open_output": "Buksan ang output folder",
        "selected_files": "Napiling files",
        "empty": "Wala",
        "rows": "Bookmarks",
        "profile": "Profile",
    },
    "pt": {
        "app_title": "Ferramenta de exportação de favoritos",
        "language": "Idioma",
        "tab_html": "HTML para CSV",
        "tab_profiles": "Ler perfis do navegador",
        "tab_sheets": "Exportar para Google Sheets",
        "output_folder": "Pasta de saída",
        "desktop_bookmarks": "Pasta bookmarks no Desktop",
        "add_html": "Adicionar arquivos HTML",
        "clear": "Limpar",
        "convert_html": "Converter para CSV",
        "html_hint": "Selecione um ou mais HTMLs de favoritos exportados do Chrome. Cada HTML gera um CSV.",
        "browser": "Navegador",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "Pasta de dados do navegador",
        "browse": "Procurar",
        "scan": "Ler",
        "convert_profiles": "Exportar CSV",
        "mode": "Modo de exportação",
        "mode_separate": "Um CSV por perfil",
        "mode_merged": "Todos os perfis em um CSV",
        "profiles_found": "Perfis encontrados",
        "selected": "Selecionar",
        "select_all": "Selecionar tudo",
        "clear_selection": "Limpar seleção",
        "sheet_url": "Link do Google Sheet",
        "open_sheet": "Abrir sheet",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Nome da sheet",
        "upload": "Enviar para Google Sheets",
        "sheets_hint": "O app lê os Bookmarks do navegador selecionado e adiciona ao Google Sheets. Observação: escrever no Google Sheets normalmente exige OAuth ou service account; apenas API key pode ser rejeitada pelo Google.",
        "status_ready": "Pronto.",
        "status_working": "Processando...",
        "status_done": "Concluído.",
        "no_html": "Adicione pelo menos um arquivo HTML.",
        "no_profiles": "Nenhum arquivo Bookmarks encontrado.",
        "no_profiles_selected": "Selecione pelo menos um perfil.",
        "no_sheet_url": "Informe o link do Google Sheet.",
        "no_api_key": "Informe a API key.",
        "no_oauth_client": "Selecione um arquivo OAuth client JSON.",
        "done_html": "{count} arquivo(s) HTML convertido(s).\nSaída: {folder}",
        "done_profiles": "{count} arquivo(s) CSV exportado(s).\nSaída: {folder}",
        "done_merged": "CSV combinado exportado.\nArquivo: {file}",
        "done_upload": "{rows} favoritos enviados ao Google Sheets.",
        "done_upload_profiles": "{profiles} perfil(is), {rows} favoritos no total, enviados ao Google Sheets.",
        "error": "Erro",
        "info": "Informação",
        "warning": "Observação",
        "copied_bookmarks": "Arquivo Bookmarks bloqueado copiado.",
        "api_key_write_error": "O Google rejeitou a escrita. Normalmente a API key não autoriza escrita no Google Sheets. Use OAuth, service account, ou exporte CSV e importe manualmente.",
        "choose_output": "Escolher pasta de saída",
        "choose_oauth_client": "Escolher OAuth client JSON",
        "oauth_client_invalid": "OAuth client JSON inválido. Use um Desktop app client JSON baixado do Google Cloud.",
        "oauth_browser": "Faça login e autorize o acesso ao Google Sheets na janela do navegador que abrir.",
        "drop_added": "{count} arquivo(s) HTML adicionado(s).",
        "drop_unsupported": "Solte arquivos de favoritos .html ou .htm.",
        "open_output": "Abrir pasta de saída",
        "selected_files": "Arquivos selecionados",
        "empty": "Nenhum",
        "rows": "Favoritos",
        "profile": "Perfil",
    },
    "es": {
        "app_title": "Herramienta de exportación de marcadores",
        "language": "Idioma",
        "tab_html": "HTML a CSV",
        "tab_profiles": "Escanear perfiles del navegador",
        "tab_sheets": "Exportar a Google Sheets",
        "output_folder": "Carpeta de salida",
        "desktop_bookmarks": "Carpeta bookmarks en el Escritorio",
        "add_html": "Agregar archivos HTML",
        "clear": "Limpiar",
        "convert_html": "Convertir a CSV",
        "html_hint": "Selecciona uno o más HTML de marcadores exportados desde Chrome. Cada HTML crea un CSV.",
        "browser": "Navegador",
        "browser_chrome": "Google Chrome",
        "browser_edge": "Microsoft Edge",
        "chrome_base": "Carpeta de datos del navegador",
        "browse": "Examinar",
        "scan": "Escanear",
        "convert_profiles": "Exportar CSV",
        "mode": "Modo de exportación",
        "mode_separate": "Un CSV por perfil",
        "mode_merged": "Todos los perfiles en un CSV",
        "profiles_found": "Perfiles encontrados",
        "selected": "Seleccionar",
        "select_all": "Seleccionar todo",
        "clear_selection": "Limpiar selección",
        "sheet_url": "Enlace de Google Sheet",
        "open_sheet": "Abrir sheet",
        "api_key": "Google Sheet API key",
        "oauth_client": "OAuth client JSON",
        "sheet_name": "Nombre de sheet",
        "upload": "Subir a Google Sheets",
        "sheets_hint": "La app lee Bookmarks del navegador seleccionado y los agrega a Google Sheets. Nota: escribir en Google Sheets suele requerir OAuth o service account; solo una API key puede ser rechazada por Google.",
        "status_ready": "Listo.",
        "status_working": "Procesando...",
        "status_done": "Terminado.",
        "no_html": "Agrega al menos un archivo HTML.",
        "no_profiles": "No se encontró ningún archivo Bookmarks.",
        "no_profiles_selected": "Selecciona al menos un perfil.",
        "no_sheet_url": "Ingresa el enlace de Google Sheet.",
        "no_api_key": "Ingresa la API key.",
        "no_oauth_client": "Selecciona un archivo OAuth client JSON.",
        "done_html": "{count} archivo(s) HTML convertido(s).\nSalida: {folder}",
        "done_profiles": "{count} archivo(s) CSV exportado(s).\nSalida: {folder}",
        "done_merged": "CSV combinado exportado.\nArchivo: {file}",
        "done_upload": "{rows} marcadores subidos a Google Sheets.",
        "done_upload_profiles": "{profiles} perfil(es), {rows} marcadores en total, subidos a Google Sheets.",
        "error": "Error",
        "info": "Información",
        "warning": "Nota",
        "copied_bookmarks": "Archivo Bookmarks bloqueado copiado.",
        "api_key_write_error": "Google rechazó la escritura. Normalmente una API key no autoriza escribir en Google Sheets. Usa OAuth, service account, o exporta CSV e impórtalo manualmente.",
        "choose_output": "Elegir carpeta de salida",
        "choose_oauth_client": "Elegir OAuth client JSON",
        "oauth_client_invalid": "OAuth client JSON no válido. Usa un Desktop app client JSON descargado desde Google Cloud.",
        "oauth_browser": "Inicia sesión y autoriza el acceso a Google Sheets en la ventana del navegador que se abre.",
        "drop_added": "{count} archivo(s) HTML agregado(s).",
        "drop_unsupported": "Arrastra archivos de marcadores .html o .htm.",
        "open_output": "Abrir carpeta de salida",
        "selected_files": "Archivos seleccionados",
        "empty": "Ninguno",
        "rows": "Marcadores",
        "profile": "Perfil",
    },
}


CSV_HEADERS = [
    "profile",
    "source",
    "folder_path",
    "title",
    "url",
    "date_added",
    "date_modified",
]


@dataclass
class BookmarkRow:
    profile: str
    source: str
    folder_path: str
    title: str
    url: str
    date_added: str = ""
    date_modified: str = ""

    def to_csv_row(self):
        return [
            self.profile,
            self.source,
            self.folder_path,
            self.title,
            self.url,
            self.date_added,
            self.date_modified,
        ]


@dataclass
class ProfileBookmarks:
    profile_id: str
    display_name: str
    path: Path
    rows: list


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(wintypes.BYTE)),
    ]


def desktop_bookmarks_folder():
    desktop = Path.home() / "Desktop"
    if not desktop.exists():
        desktop = Path.home() / "OneDrive" / "Desktop"
    output = desktop / OUTPUT_FOLDER_NAME
    output.mkdir(parents=True, exist_ok=True)
    return output


def ensure_output_folder(path_text):
    text = (path_text or "").strip()
    if not text:
        folder = desktop_bookmarks_folder()
    else:
        folder = Path(text).expanduser()
        folder.mkdir(parents=True, exist_ok=True)
    return folder


def app_storage_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(filename):
    base = Path(getattr(sys, "_MEIPASS", app_storage_dir()))
    return base / filename


def settings_path():
    return app_storage_dir() / SETTINGS_FILE


def load_settings():
    path = settings_path()
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings):
    try:
        settings_path().write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def _bytes_to_blob(data):
    buffer = ctypes.create_string_buffer(data)
    blob = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(wintypes.BYTE)))
    return blob, buffer


def _protect_with_windows_dpapi(data):
    if os.name != "nt":
        raise RuntimeError("Windows DPAPI is required for OAuth token storage.")

    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(DATA_BLOB),
        wintypes.LPCWSTR,
        ctypes.POINTER(DATA_BLOB),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(DATA_BLOB),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL

    in_blob, in_buffer = _bytes_to_blob(data)
    out_blob = DATA_BLOB()
    try:
        ok = crypt32.CryptProtectData(
            ctypes.byref(in_blob),
            "Bookmark Export Tool OAuth Token",
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            ctypes.byref(out_blob),
        )
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        _ = in_buffer
        if out_blob.pbData:
            kernel32.LocalFree(out_blob.pbData)


def _unprotect_with_windows_dpapi(data):
    if os.name != "nt":
        raise RuntimeError("Windows DPAPI is required for OAuth token storage.")

    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(DATA_BLOB),
        ctypes.c_void_p,
        ctypes.POINTER(DATA_BLOB),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(DATA_BLOB),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL

    in_blob, in_buffer = _bytes_to_blob(data)
    out_blob = DATA_BLOB()
    try:
        ok = crypt32.CryptUnprotectData(
            ctypes.byref(in_blob),
            None,
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            ctypes.byref(out_blob),
        )
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        _ = in_buffer
        if out_blob.pbData:
            kernel32.LocalFree(out_blob.pbData)


def default_chrome_user_data():
    return default_browser_user_data("chrome")


def default_edge_user_data():
    return default_browser_user_data("edge")


def default_browser_user_data(browser):
    local = os.environ.get("LOCALAPPDATA")
    parts = {
        "chrome": ("Google", "Chrome"),
        "edge": ("Microsoft", "Edge"),
    }.get(browser, ("Google", "Chrome"))
    if local:
        return Path(local) / parts[0] / parts[1] / "User Data"
    return Path.home() / "AppData" / "Local" / parts[0] / parts[1] / "User Data"


def safe_filename(name):
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", name).strip().strip(".")
    return clean or "bookmarks"


def unique_path(folder, filename):
    candidate = folder / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(2, 10000):
        next_candidate = folder / f"{stem}_{index}{suffix}"
        if not next_candidate.exists():
            return next_candidate
    raise RuntimeError("Could not create a unique filename.")


def unix_seconds_to_iso(value):
    try:
        raw = str(value).strip()
        if not raw:
            return ""
        seconds = int(float(raw))
        if seconds <= 0:
            return ""
        return datetime.fromtimestamp(seconds, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, OSError, OverflowError):
        return ""


def chrome_time_to_iso(value):
    try:
        raw = str(value).strip()
        if not raw:
            return ""
        microseconds = int(raw)
        if microseconds <= 0:
            return ""
        seconds = microseconds / 1_000_000 - 11644473600
        return datetime.fromtimestamp(seconds, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, OSError, OverflowError):
        return ""


class NetscapeBookmarkParser(HTMLParser):
    def __init__(self, source_name):
        super().__init__(convert_charrefs=True)
        self.source_name = source_name
        self.rows = []
        self.folder_stack = []
        self.dl_contexts = []
        self.capture = None
        self.capture_text = []
        self.anchor_attrs = {}
        self.folder_attrs = {}
        self.pending_folder = None

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attr_map = {key.lower(): value for key, value in attrs}
        if tag == "h3":
            self.capture = "folder"
            self.capture_text = []
            self.folder_attrs = attr_map
        elif tag == "a":
            self.capture = "anchor"
            self.capture_text = []
            self.anchor_attrs = attr_map
        elif tag == "dl":
            if self.pending_folder:
                self.folder_stack.append(self.pending_folder["title"])
                self.dl_contexts.append(True)
                self.pending_folder = None
            else:
                self.dl_contexts.append(False)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "h3" and self.capture == "folder":
            title = clean_text("".join(self.capture_text))
            if title:
                self.pending_folder = {
                    "title": title,
                    "attrs": self.folder_attrs,
                }
            self.capture = None
            self.capture_text = []
        elif tag == "a" and self.capture == "anchor":
            title = clean_text("".join(self.capture_text))
            href = html.unescape(self.anchor_attrs.get("href", "") or "").strip()
            if href:
                self.rows.append(
                    BookmarkRow(
                        profile="",
                        source=self.source_name,
                        folder_path=" / ".join(self.folder_stack),
                        title=title or href,
                        url=href,
                        date_added=unix_seconds_to_iso(self.anchor_attrs.get("add_date", "")),
                        date_modified=unix_seconds_to_iso(self.anchor_attrs.get("last_modified", "")),
                    )
                )
            self.capture = None
            self.capture_text = []
            self.anchor_attrs = {}
        elif tag == "dl":
            if self.dl_contexts:
                pushed = self.dl_contexts.pop()
                if pushed and self.folder_stack:
                    self.folder_stack.pop()

    def handle_data(self, data):
        if self.capture in {"folder", "anchor"}:
            self.capture_text.append(data)


def clean_text(value):
    return re.sub(r"\s+", " ", html.unescape(value or "")).strip()


def parse_bookmark_html(path):
    source = Path(path).stem
    parser = NetscapeBookmarkParser(source)
    data = Path(path).read_text(encoding="utf-8", errors="replace")
    parser.feed(data)
    return parser.rows


def parse_chrome_bookmarks(path, profile_name):
    data = json.loads(Path(path).read_text(encoding="utf-8", errors="replace"))
    rows = []
    roots = data.get("roots", {})

    def walk(node, folders):
        node_type = node.get("type")
        if node_type == "url":
            url = (node.get("url") or "").strip()
            if url:
                rows.append(
                    BookmarkRow(
                        profile=profile_name,
                        source=str(path),
                        folder_path=" / ".join(folders),
                        title=clean_text(node.get("name", "")) or url,
                        url=url,
                        date_added=chrome_time_to_iso(node.get("date_added", "")),
                        date_modified=chrome_time_to_iso(node.get("date_last_used", "")),
                    )
                )
        elif node_type == "folder" or "children" in node:
            name = clean_text(node.get("name", ""))
            next_folders = folders + ([name] if name else [])
            for child in node.get("children", []) or []:
                walk(child, next_folders)

    root_labels = {
        "bookmark_bar": "Bookmarks Bar",
        "other": "Other Bookmarks",
        "synced": "Mobile Bookmarks",
    }
    for key, label in root_labels.items():
        root = roots.get(key)
        if root:
            root_name = clean_text(root.get("name", "")) or label
            for child in root.get("children", []) or []:
                walk(child, [root_name])

    for key, root in roots.items():
        if key in root_labels or not isinstance(root, dict):
            continue
        root_name = clean_text(root.get("name", "")) or key
        for child in root.get("children", []) or []:
            walk(child, [root_name])
    return rows


def read_profile_display_names(chrome_user_data):
    local_state = Path(chrome_user_data) / "Local State"
    names = {}
    if not local_state.exists():
        return names
    try:
        data = json.loads(local_state.read_text(encoding="utf-8", errors="replace"))
        info_cache = data.get("profile", {}).get("info_cache", {})
        for profile_id, profile_info in info_cache.items():
            name = profile_info.get("name") or profile_info.get("gaia_name")
            if name:
                names[profile_id] = clean_text(name)
    except (OSError, json.JSONDecodeError):
        pass
    return names


def find_browser_bookmark_files(user_data_folder, browser="chrome"):
    base = Path(user_data_folder)
    if not base.exists():
        return []
    display_names = read_profile_display_names(base)
    found = []
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if d not in {"Crashpad", "GrShaderCache", "ShaderCache", "Safe Browsing"}]
        if "Bookmarks" not in files:
            continue
        bookmark_path = Path(root) / "Bookmarks"
        profile_id = bookmark_path.parent.name
        display_name = display_names.get(profile_id) or profile_id
        try:
            rows = parse_chrome_bookmarks(bookmark_path, display_name)
        except PermissionError:
            copied = copy_locked_bookmarks(bookmark_path)
            rows = parse_chrome_bookmarks(copied, display_name)
        except (OSError, json.JSONDecodeError):
            rows = []
        found.append(ProfileBookmarks(profile_id=profile_id, display_name=display_name, path=bookmark_path, rows=rows))
    found.sort(key=lambda item: item.display_name.lower())
    return found


def find_chrome_bookmark_files(chrome_user_data):
    return find_browser_bookmark_files(chrome_user_data, "chrome")


def copy_locked_bookmarks(path):
    temp_dir = desktop_bookmarks_folder() / "_temp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_path = temp_dir / f"{safe_filename(path.parent.name)}_Bookmarks"
    shutil.copy2(path, temp_path)
    return temp_path


def write_csv(path, rows):
    with Path(path).open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(CSV_HEADERS)
        for row in rows:
            writer.writerow(row.to_csv_row())


def parse_spreadsheet_id(sheet_url):
    text = (sheet_url or "").strip()
    patterns = [
        r"/spreadsheets/d/([a-zA-Z0-9-_]+)",
        r"[?&]id=([a-zA-Z0-9-_]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    if re.fullmatch(r"[a-zA-Z0-9-_]{20,}", text):
        return text
    return ""


def token_cache_path():
    return app_storage_dir() / TOKEN_CACHE_FILE


def remove_legacy_plaintext_token_caches():
    for path in app_storage_dir().glob("google_sheets_oauth_token*.json"):
        try:
            path.unlink()
        except OSError:
            pass


def load_cached_oauth_token(path):
    if not path.exists():
        return {}
    try:
        payload = _unprotect_with_windows_dpapi(path.read_bytes())
        data = json.loads(payload.decode("utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError):
        return {}


def save_cached_oauth_token(path, token):
    payload = json.dumps(token, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    encrypted = _protect_with_windows_dpapi(payload)
    path.write_bytes(encrypted)


def load_oauth_client(path):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("Invalid OAuth client JSON.") from exc
    config = data.get("installed") or data.get("web")
    if not isinstance(config, dict):
        raise ValueError("Invalid OAuth client JSON.")
    client_id = config.get("client_id")
    token_uri = config.get("token_uri") or "https://oauth2.googleapis.com/token"
    auth_uri = config.get("auth_uri") or "https://accounts.google.com/o/oauth2/v2/auth"
    if not client_id:
        raise ValueError("Invalid OAuth client JSON.")
    return {
        "client_id": client_id,
        "client_secret": config.get("client_secret", ""),
        "token_uri": token_uri,
        "auth_uri": auth_uri,
    }


def post_form(url, data):
    body = urllib.parse.urlencode(data).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    return read_json_response(request)


def read_json_response(request):
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            payload = response.read().decode("utf-8", errors="replace")
            return json.loads(payload) if payload else {}
    except urllib.error.HTTPError as exc:
        payload = exc.read().decode("utf-8", errors="replace")
        message = payload
        try:
            data = json.loads(payload)
            message = data.get("error", {}).get("message") or data.get("error_description") or payload
        except json.JSONDecodeError:
            pass
        raise RuntimeError(message) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc


def is_token_valid(token):
    access_token = token.get("access_token")
    created_at = float(token.get("created_at", 0) or 0)
    expires_in = float(token.get("expires_in", 0) or 0)
    return bool(access_token and time.time() < created_at + expires_in - 120)


def refresh_oauth_token(client, token):
    refresh_token = token.get("refresh_token")
    if not refresh_token:
        return None
    payload = {
        "client_id": client["client_id"],
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }
    if client.get("client_secret"):
        payload["client_secret"] = client["client_secret"]
    refreshed = post_form(client["token_uri"], payload)
    refreshed["refresh_token"] = refreshed.get("refresh_token") or refresh_token
    refreshed["created_at"] = time.time()
    refreshed["client_id"] = client["client_id"]
    return refreshed


def run_oauth_flow(client):
    state = secrets.token_urlsafe(24)
    code_verifier = secrets.token_urlsafe(64)
    challenge = hashlib.sha256(code_verifier.encode("ascii")).digest()
    code_challenge = base64.urlsafe_b64encode(challenge).decode("ascii").rstrip("=")

    class OAuthHandler(http.server.BaseHTTPRequestHandler):
        def log_message(self, _format, *args):
            return

        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            params = urllib.parse.parse_qs(parsed.query)
            self.server.oauth_result = {
                "code": params.get("code", [""])[0],
                "state": params.get("state", [""])[0],
                "error": params.get("error", [""])[0],
            }
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            message = "Authorization received. You can close this window."
            if self.server.oauth_result.get("error"):
                message = "Authorization failed. You can close this window."
            self.wfile.write(f"<html><body><h2>{html.escape(message)}</h2></body></html>".encode("utf-8"))

    server = http.server.HTTPServer(("127.0.0.1", 0), OAuthHandler)
    server.timeout = 300
    server.oauth_result = {}
    redirect_uri = f"http://127.0.0.1:{server.server_port}/"
    params = {
        "client_id": client["client_id"],
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": GOOGLE_SHEETS_SCOPE,
        "access_type": "offline",
        "prompt": "consent",
        "include_granted_scopes": "true",
        "state": state,
        "code_challenge": code_challenge,
        "code_challenge_method": "S256",
    }
    auth_url = client["auth_uri"] + "?" + urllib.parse.urlencode(params)
    webbrowser.open(auth_url, new=1)
    server.handle_request()
    result = server.oauth_result
    server.server_close()
    if result.get("error"):
        raise RuntimeError(result["error"])
    if result.get("state") != state or not result.get("code"):
        raise RuntimeError("Google OAuth authorization did not complete.")

    payload = {
        "client_id": client["client_id"],
        "code": result["code"],
        "code_verifier": code_verifier,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    }
    if client.get("client_secret"):
        payload["client_secret"] = client["client_secret"]
    token = post_form(client["token_uri"], payload)
    token["created_at"] = time.time()
    token["client_id"] = client["client_id"]
    return token


def get_google_access_token(oauth_client_path):
    client = load_oauth_client(oauth_client_path)
    cache_path = token_cache_path()
    remove_legacy_plaintext_token_caches()
    token = load_cached_oauth_token(cache_path)
    if token.get("client_id") != client["client_id"]:
        token = {}
    if not is_token_valid(token):
        try:
            refreshed = refresh_oauth_token(client, token)
        except RuntimeError:
            refreshed = None
        token = refreshed or run_oauth_flow(client)
        save_cached_oauth_token(cache_path, token)
    return token["access_token"]


def google_request(access_token, method, url, body=None):
    data = None
    headers = {"Authorization": f"Bearer {access_token}"}
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    return read_json_response(request)


def quote_sheet_range(sheet_title, cell_range="A1"):
    escaped_title = sheet_title.replace("'", "''")
    return urllib.parse.quote(f"'{escaped_title}'!{cell_range}", safe="")


def sanitize_sheet_title(title):
    clean = re.sub(r"[\[\]\:\*\?\/\\]+", "_", clean_text(title))
    clean = clean.strip().strip("'")
    if not clean:
        clean = "Profile"
    return clean[:100]


def profile_sheet_titles(profiles):
    used = set()
    titles = {}
    for profile in profiles:
        base = sanitize_sheet_title(profile.display_name or profile.profile_id)
        title = base
        if title.casefold() in used:
            suffix = sanitize_sheet_title(profile.profile_id)
            max_base = max(1, 99 - len(suffix))
            title = f"{base[:max_base]}_{suffix}"
        index = 2
        candidate = title
        while candidate.casefold() in used:
            suffix = f"_{index}"
            candidate = title[:100 - len(suffix)] + suffix
            index += 1
        used.add(candidate.casefold())
        titles[profile.profile_id] = candidate
    return titles


def get_existing_sheet_titles(spreadsheet_id, access_token):
    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
        "?fields=sheets(properties(title))"
    )
    data = google_request(access_token, "GET", url)
    return {
        sheet.get("properties", {}).get("title", "")
        for sheet in data.get("sheets", [])
        if sheet.get("properties", {}).get("title")
    }


def ensure_sheet_tab(spreadsheet_id, access_token, title, existing_titles):
    if title in existing_titles:
        return
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate"
    body = {"requests": [{"addSheet": {"properties": {"title": title}}}]}
    google_request(access_token, "POST", url, body)
    existing_titles.add(title)


def clear_sheet_values(spreadsheet_id, access_token, title):
    encoded_range = quote_sheet_range(title, "A:Z")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{encoded_range}:clear"
    google_request(access_token, "POST", url, {})


def write_sheet_values(spreadsheet_id, access_token, title, rows):
    encoded_range = quote_sheet_range(title, "A1")
    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/"
        f"{encoded_range}?valueInputOption=RAW"
    )
    values = [CSV_HEADERS] + [row.to_csv_row() for row in rows]
    google_request(access_token, "PUT", url, {"values": values})


def upload_profiles_to_google_sheets(sheet_url, oauth_client_path, profiles):
    spreadsheet_id = parse_spreadsheet_id(sheet_url)
    if not spreadsheet_id:
        raise ValueError("Could not find spreadsheet id in the Google Sheet link.")

    access_token = get_google_access_token(oauth_client_path)
    existing_titles = get_existing_sheet_titles(spreadsheet_id, access_token)
    sheet_titles = profile_sheet_titles(profiles)
    total_rows = 0
    uploaded_profiles = 0
    for profile in profiles:
        title = sheet_titles[profile.profile_id]
        ensure_sheet_tab(spreadsheet_id, access_token, title, existing_titles)
        clear_sheet_values(spreadsheet_id, access_token, title)
        write_sheet_values(spreadsheet_id, access_token, title, profile.rows)
        total_rows += len(profile.rows)
        uploaded_profiles += 1
    return uploaded_profiles, total_rows


class BookmarkApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.lang = "zh"
        self.html_files = []
        self.profiles = []
        self.profiles_base_path = ""
        self.profiles_browser = ""
        self.selected_profile_ids = set()
        self.profile_item_ids = {}
        self.sheet_profiles = []
        self.sheet_profiles_base_path = ""
        self.sheet_profiles_browser = ""
        self.selected_sheet_profile_ids = set()
        self.sheet_profile_item_ids = {}
        self.checkbox_images = {}
        self.widgets = {}
        self.settings = load_settings()
        self._drop_old_procs = {}
        self._drop_wndproc = None
        self._drop_enabled = False
        self._drop_pending_files = []
        self._drop_poll_id = None
        self.mode_var = tk.StringVar(value="separate")
        self.language_var = tk.StringVar(value=LANGUAGES[self.lang])
        self.output_folder_var = tk.StringVar(value=str(desktop_bookmarks_folder()))
        self.profile_browser = self.valid_browser_code(self.settings.get("profile_browser", "chrome"))
        self.sheet_browser = self.valid_browser_code(self.settings.get("sheet_browser", "chrome"))
        self.profile_base_paths = {
            code: self.settings.get(f"profile_{code}_base", str(default_browser_user_data(code)))
            for code in BROWSER_CODES
        }
        self.sheet_base_paths = {
            code: self.settings.get(f"sheet_{code}_base", str(default_browser_user_data(code)))
            for code in BROWSER_CODES
        }
        self.profile_browser_var = tk.StringVar(value=self.browser_label(self.profile_browser))
        self.sheet_browser_var = tk.StringVar(value=self.browser_label(self.sheet_browser))
        self.chrome_base_var = tk.StringVar(value=self.profile_base_paths[self.profile_browser])
        self.sheet_base_var = tk.StringVar(value=self.sheet_base_paths[self.sheet_browser])
        self.sheet_url_var = tk.StringVar(value=self.settings.get("sheet_url", ""))
        self.oauth_client_var = tk.StringVar(value=self.settings.get("oauth_client_path", ""))
        self.status_var = tk.StringVar(value=self.t("status_ready"))
        self.geometry("920x640")
        self.minsize(820, 560)
        self.title(self.t("app_title"))
        self.configure(bg="#f6f7f9")
        self.apply_window_icon()
        self.create_widgets()
        self.apply_language()
        self.bind("<Destroy>", self.on_destroy, add="+")
        self.after(200, self.enable_file_drop)
        self._drop_poll_id = self.after(250, self.poll_dropped_files)

    def t(self, key):
        return TEXT.get(self.lang, TEXT["en"]).get(key, TEXT["en"].get(key, key))

    def tr(self, key, **values):
        return self.t(key).format(**values)

    def valid_browser_code(self, code):
        return code if code in BROWSER_CODES else "chrome"

    def browser_label(self, code):
        return self.t(BROWSER_LABEL_KEYS.get(self.valid_browser_code(code), "browser_chrome"))

    def browser_values(self):
        return [self.browser_label(code) for code in BROWSER_CODES]

    def browser_code_from_label(self, label):
        for code in BROWSER_CODES:
            if label == self.browser_label(code):
                return code
        return "chrome"

    def refresh_browser_combo_labels(self):
        values = self.browser_values()
        if hasattr(self, "profile_browser_combo"):
            self.profile_browser_combo.configure(values=values)
            self.profile_browser_var.set(self.browser_label(self.profile_browser))
        if hasattr(self, "sheet_browser_combo"):
            self.sheet_browser_combo.configure(values=values)
            self.sheet_browser_var.set(self.browser_label(self.sheet_browser))

    def remember_profile_base_path(self):
        self.profile_base_paths[self.profile_browser] = self.chrome_base_var.get().strip()

    def remember_sheet_base_path(self):
        self.sheet_base_paths[self.sheet_browser] = self.sheet_base_var.get().strip()

    def current_profile_base_path(self):
        self.remember_profile_base_path()
        return self.profile_base_paths[self.profile_browser]

    def current_sheet_base_path(self):
        self.remember_sheet_base_path()
        return self.sheet_base_paths[self.sheet_browser]

    def register(self, name, widget, key=None):
        self.widgets[name] = (widget, key)
        return widget

    def apply_window_icon(self):
        icon_path = resource_path(ICON_FILE)
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass

    def create_checkbox_images(self):
        images = {}

        def make_image(checked):
            size = 24
            img = tk.PhotoImage(width=size, height=size)
            left = 4
            top = 4
            right = 20
            bottom = 20
            fill = "#0b74de" if checked else "#ffffff"
            border = "#0b74de" if checked else "#777f89"
            for y in range(top, bottom):
                img.put(fill, to=(left, y, right, y + 1))
            for offset in range(2):
                img.put(border, to=(left, top + offset, right, top + offset + 1))
                img.put(border, to=(left, bottom - 1 - offset, right, bottom - offset))
                img.put(border, to=(left + offset, top, left + offset + 1, bottom))
                img.put(border, to=(right - 1 - offset, top, right - offset, bottom))
            if checked:
                check = "#ffffff"
                points = [
                    (8, 12), (9, 13), (10, 14), (11, 15),
                    (12, 14), (13, 13), (14, 12), (15, 11), (16, 10),
                ]
                for x, y in points:
                    img.put(check, to=(x, y, x + 2, y + 2))
            return img

        images["checked"] = make_image(True)
        images["unchecked"] = make_image(False)
        return images

    def create_widgets(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("vista")
        except tk.TclError:
            pass
        self.style.configure("TFrame", background="#f6f7f9")
        self.style.configure("Panel.TFrame", background="#ffffff")
        self.style.configure("Hint.TLabel", background="#ffffff", foreground="#5f6875")
        self.style.configure("Title.TLabel", background="#ffffff", font=("Segoe UI", 11, "bold"))
        self.style.configure("Status.TLabel", background="#eef1f5", foreground="#2f3742")
        self.style.configure("Treeview", rowheight=32)
        self.checkbox_images = self.create_checkbox_images()

        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 10))
        self.language_label = self.register("language_label", ttk.Label(header), "language")
        self.language_label.pack(side="left", padx=(0, 8))
        self.language_combo = ttk.Combobox(header, textvariable=self.language_var, values=list(LANGUAGES.values()), state="readonly", width=18)
        self.language_combo.pack(side="left")
        self.language_combo.bind("<<ComboboxSelected>>", self.on_language_change)

        self.output_label = self.register("output_label", ttk.Label(header), "output_folder")
        self.output_label.pack(side="left", padx=(30, 8))
        self.output_entry = ttk.Entry(header, textvariable=self.output_folder_var)
        self.output_entry.pack(side="left", fill="x", expand=True)
        self.choose_output_button = self.register("choose_output_button", ttk.Button(header, command=self.choose_output_folder), "choose_output")
        self.choose_output_button.pack(side="right", padx=(8, 0))
        self.open_output_button = self.register("open_output_button", ttk.Button(header, command=self.open_output_folder), "open_output")
        self.open_output_button.pack(side="right", padx=(8, 0))

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        self.html_tab = ttk.Frame(self.notebook, padding=12)
        self.profile_tab = ttk.Frame(self.notebook, padding=12)
        self.sheet_tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.html_tab, text=self.t("tab_html"))
        self.notebook.add(self.profile_tab, text=self.t("tab_profiles"))
        self.notebook.add(self.sheet_tab, text=self.t("tab_sheets"))

        self.build_html_tab()
        self.build_profile_tab()
        self.build_sheet_tab()

        status_bar = ttk.Frame(root, style="Status.TLabel", padding=(10, 6))
        status_bar.pack(fill="x", pady=(10, 0))
        self.status_label = ttk.Label(status_bar, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.pack(side="left")

    def build_html_tab(self):
        panel = ttk.Frame(self.html_tab, style="Panel.TFrame", padding=14)
        panel.pack(fill="both", expand=True)

        self.html_hint = self.register("html_hint", ttk.Label(panel, style="Hint.TLabel", wraplength=820), "html_hint")
        self.html_hint.pack(anchor="w", fill="x", pady=(0, 10))

        buttons = ttk.Frame(panel, style="Panel.TFrame")
        buttons.pack(fill="x", pady=(0, 8))
        self.add_html_button = self.register("add_html_button", ttk.Button(buttons, command=self.add_html_files), "add_html")
        self.add_html_button.pack(side="left", padx=(0, 8))
        self.clear_html_button = self.register("clear_html_button", ttk.Button(buttons, command=self.clear_html_files), "clear")
        self.clear_html_button.pack(side="left", padx=(0, 8))
        self.convert_html_button = self.register("convert_html_button", ttk.Button(buttons, command=self.convert_html_files), "convert_html")
        self.convert_html_button.pack(side="right")

        self.html_list = tk.Listbox(panel, height=18, activestyle="none")
        self.html_list.pack(fill="both", expand=True)

    def build_profile_tab(self):
        panel = ttk.Frame(self.profile_tab, style="Panel.TFrame", padding=14)
        panel.pack(fill="both", expand=True)

        browser_row = ttk.Frame(panel, style="Panel.TFrame")
        browser_row.pack(fill="x", pady=(0, 8))
        self.profile_browser_label = self.register("profile_browser_label", ttk.Label(browser_row, style="Title.TLabel"), "browser")
        self.profile_browser_label.pack(side="left", padx=(0, 8))
        self.profile_browser_combo = ttk.Combobox(
            browser_row,
            textvariable=self.profile_browser_var,
            values=self.browser_values(),
            state="readonly",
            width=24,
        )
        self.profile_browser_combo.pack(side="left")
        self.profile_browser_combo.bind("<<ComboboxSelected>>", self.on_profile_browser_change)

        base_row = ttk.Frame(panel, style="Panel.TFrame")
        base_row.pack(fill="x", pady=(0, 8))
        self.chrome_base_label = self.register("chrome_base_label", ttk.Label(base_row, style="Title.TLabel"), "chrome_base")
        self.chrome_base_label.pack(anchor="w")
        path_row = ttk.Frame(panel, style="Panel.TFrame")
        path_row.pack(fill="x", pady=(0, 10))
        self.chrome_base_entry = ttk.Entry(path_row, textvariable=self.chrome_base_var)
        self.chrome_base_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.chrome_browse_button = self.register("chrome_browse_button", ttk.Button(path_row, command=lambda: self.browse_folder(self.chrome_base_var)), "browse")
        self.chrome_browse_button.pack(side="left", padx=(0, 8))
        self.scan_button = self.register("scan_button", ttk.Button(path_row, command=self.scan_profiles), "scan")
        self.scan_button.pack(side="left")

        mode_row = ttk.Frame(panel, style="Panel.TFrame")
        mode_row.pack(fill="x", pady=(0, 8))
        self.mode_label = self.register("mode_label", ttk.Label(mode_row, style="Title.TLabel"), "mode")
        self.mode_label.pack(side="left", padx=(0, 12))
        self.mode_separate = ttk.Radiobutton(mode_row, variable=self.mode_var, value="separate")
        self.mode_separate.pack(side="left", padx=(0, 18))
        self.mode_merged = ttk.Radiobutton(mode_row, variable=self.mode_var, value="merged")
        self.mode_merged.pack(side="left")
        self.export_profiles_button = self.register("export_profiles_button", ttk.Button(mode_row, command=self.export_profiles_csv), "convert_profiles")
        self.export_profiles_button.pack(side="right")
        self.clear_profile_selection_button = self.register("clear_profile_selection_button", ttk.Button(mode_row, command=self.clear_profile_selection), "clear_selection")
        self.clear_profile_selection_button.pack(side="right", padx=(0, 8))
        self.select_all_profiles_button = self.register("select_all_profiles_button", ttk.Button(mode_row, command=self.select_all_profiles), "select_all")
        self.select_all_profiles_button.pack(side="right", padx=(0, 8))

        self.profiles_label = self.register("profiles_label", ttk.Label(panel, style="Title.TLabel"), "profiles_found")
        self.profiles_label.pack(anchor="w", pady=(8, 4))

        columns = ("profile", "path", "rows")
        self.profile_tree = ttk.Treeview(panel, columns=columns, show=("tree", "headings"), height=14)
        self.profile_tree.heading("#0", text=self.t("selected"))
        self.profile_tree.heading("profile", text=self.t("profile"))
        self.profile_tree.heading("path", text="Bookmarks")
        self.profile_tree.heading("rows", text=self.t("rows"))
        self.profile_tree.column("#0", width=80, anchor="center", stretch=False)
        self.profile_tree.column("profile", width=180, anchor="w")
        self.profile_tree.column("path", width=520, anchor="w")
        self.profile_tree.column("rows", width=90, anchor="e")
        self.profile_tree.tag_configure("checked", background="#e8f2ff")
        self.profile_tree.tag_configure("unchecked", background="#ffffff")
        self.profile_tree.pack(fill="both", expand=True)
        self.profile_tree.bind("<Button-1>", self.on_profile_tree_click)

    def build_sheet_tab(self):
        panel = ttk.Frame(self.sheet_tab, style="Panel.TFrame", padding=14)
        panel.pack(fill="both", expand=True)

        sheet_source_row = ttk.Frame(panel, style="Panel.TFrame")
        sheet_source_row.pack(fill="x", pady=(0, 10))
        self.sheet_browser_label = self.register("sheet_browser_label", ttk.Label(sheet_source_row, style="Title.TLabel"), "browser")
        self.sheet_browser_label.pack(side="left", padx=(0, 8))
        self.sheet_browser_combo = ttk.Combobox(
            sheet_source_row,
            textvariable=self.sheet_browser_var,
            values=self.browser_values(),
            state="readonly",
            width=24,
        )
        self.sheet_browser_combo.pack(side="left", padx=(0, 18))
        self.sheet_browser_combo.bind("<<ComboboxSelected>>", self.on_sheet_browser_change)

        self.sheet_base_label = self.register("sheet_base_label", ttk.Label(sheet_source_row, style="Title.TLabel"), "chrome_base")
        self.sheet_base_label.pack(side="left", padx=(0, 8))
        ttk.Entry(sheet_source_row, textvariable=self.sheet_base_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.sheet_base_browse = self.register("sheet_base_browse", ttk.Button(sheet_source_row, command=lambda: self.browse_folder(self.sheet_base_var)), "browse")
        self.sheet_base_browse.pack(side="left")

        credentials_row = ttk.Frame(panel, style="Panel.TFrame")
        credentials_row.pack(fill="x", pady=(0, 12))

        sheet_url_group = ttk.Frame(credentials_row, style="Panel.TFrame")
        sheet_url_group.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.sheet_url_label = self.register("sheet_url_label", ttk.Label(sheet_url_group, style="Title.TLabel"), "sheet_url")
        self.sheet_url_label.pack(anchor="w")
        sheet_url_row = ttk.Frame(sheet_url_group, style="Panel.TFrame")
        sheet_url_row.pack(fill="x", pady=(4, 0))
        ttk.Entry(sheet_url_row, textvariable=self.sheet_url_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.open_sheet_button = self.register("open_sheet_button", ttk.Button(sheet_url_row, command=self.open_sheet_url), "open_sheet")
        self.open_sheet_button.pack(side="left")

        oauth_group = ttk.Frame(credentials_row, style="Panel.TFrame")
        oauth_group.pack(side="left", fill="x", expand=True)
        self.oauth_client_label = self.register("oauth_client_label", ttk.Label(oauth_group, style="Title.TLabel"), "oauth_client")
        self.oauth_client_label.pack(anchor="w")
        oauth_row = ttk.Frame(oauth_group, style="Panel.TFrame")
        oauth_row.pack(fill="x", pady=(4, 0))
        ttk.Entry(oauth_row, textvariable=self.oauth_client_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.oauth_client_browse = self.register("oauth_client_browse", ttk.Button(oauth_row, command=self.browse_oauth_client), "browse")
        self.oauth_client_browse.pack(side="left")

        action_row = ttk.Frame(panel, style="Panel.TFrame")
        action_row.pack(fill="x", pady=(8, 8))
        self.sheet_profiles_label = self.register("sheet_profiles_label", ttk.Label(action_row, style="Title.TLabel"), "profiles_found")
        self.sheet_profiles_label.pack(side="left")
        self.sheet_scan_button = self.register("sheet_scan_button", ttk.Button(action_row, command=self.scan_sheet_profiles), "scan")
        self.sheet_scan_button.pack(side="right", padx=(8, 0))
        self.upload_button = self.register("upload_button", ttk.Button(action_row, command=self.upload_to_sheet), "upload")
        self.upload_button.pack(side="right", padx=(8, 0))
        self.clear_sheet_selection_button = self.register("clear_sheet_selection_button", ttk.Button(action_row, command=self.clear_sheet_profile_selection), "clear_selection")
        self.clear_sheet_selection_button.pack(side="right", padx=(8, 0))
        self.select_all_sheet_button = self.register("select_all_sheet_button", ttk.Button(action_row, command=self.select_all_sheet_profiles), "select_all")
        self.select_all_sheet_button.pack(side="right", padx=(8, 0))

        columns = ("profile", "path", "rows")
        self.sheet_profile_tree = ttk.Treeview(panel, columns=columns, show=("tree", "headings"), height=10)
        self.sheet_profile_tree.heading("#0", text=self.t("selected"))
        self.sheet_profile_tree.heading("profile", text=self.t("profile"))
        self.sheet_profile_tree.heading("path", text="Bookmarks")
        self.sheet_profile_tree.heading("rows", text=self.t("rows"))
        self.sheet_profile_tree.column("#0", width=80, anchor="center", stretch=False)
        self.sheet_profile_tree.column("profile", width=180, anchor="w")
        self.sheet_profile_tree.column("path", width=520, anchor="w")
        self.sheet_profile_tree.column("rows", width=90, anchor="e")
        self.sheet_profile_tree.tag_configure("checked", background="#e8f2ff")
        self.sheet_profile_tree.tag_configure("unchecked", background="#ffffff")
        self.sheet_profile_tree.pack(fill="both", expand=True)
        self.sheet_profile_tree.bind("<Button-1>", self.on_sheet_profile_tree_click)

    def apply_language(self):
        self.title(self.t("app_title"))
        self.notebook.tab(self.html_tab, text=self.t("tab_html"))
        self.notebook.tab(self.profile_tab, text=self.t("tab_profiles"))
        self.notebook.tab(self.sheet_tab, text=self.t("tab_sheets"))
        for widget, key in self.widgets.values():
            if key:
                widget.configure(text=self.t(key))
        self.refresh_browser_combo_labels()
        self.mode_separate.configure(text=self.t("mode_separate"))
        self.mode_merged.configure(text=self.t("mode_merged"))
        self.profile_tree.heading("#0", text=self.t("selected"))
        self.profile_tree.heading("profile", text=self.t("profile"))
        self.profile_tree.heading("rows", text=self.t("rows"))
        self.sheet_profile_tree.heading("#0", text=self.t("selected"))
        self.sheet_profile_tree.heading("profile", text=self.t("profile"))
        self.sheet_profile_tree.heading("rows", text=self.t("rows"))
        self.status_var.set(self.t("status_ready"))

    def on_language_change(self, _event=None):
        selected = self.language_var.get()
        for key, label in LANGUAGES.items():
            if label == selected:
                self.lang = key
                break
        self.apply_language()

    def set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for name, (widget, _key) in self.widgets.items():
            if isinstance(widget, ttk.Button):
                widget.configure(state=state)
        self.status_var.set(self.t("status_working") if busy else self.t("status_ready"))
        self.update_idletasks()

    def run_worker(self, work, done):
        self.set_busy(True)

        def target():
            try:
                result = work()
                self.after(0, lambda: done(result, None))
            except Exception as exc:
                self.after(0, lambda: done(None, exc))

        threading.Thread(target=target, daemon=True).start()

    def on_destroy(self, event):
        if event.widget is self:
            self.save_current_settings()
            if self._drop_poll_id:
                try:
                    self.after_cancel(self._drop_poll_id)
                except tk.TclError:
                    pass
                self._drop_poll_id = None
            self.disable_file_drop()

    def save_current_settings(self):
        self.remember_profile_base_path()
        self.remember_sheet_base_path()
        self.settings["sheet_url"] = self.sheet_url_var.get().strip()
        self.settings["oauth_client_path"] = self.oauth_client_var.get().strip()
        self.settings["profile_browser"] = self.profile_browser
        self.settings["sheet_browser"] = self.sheet_browser
        for code in BROWSER_CODES:
            self.settings[f"profile_{code}_base"] = self.profile_base_paths.get(code, "")
            self.settings[f"sheet_{code}_base"] = self.sheet_base_paths.get(code, "")
        save_settings(self.settings)

    def enable_file_drop(self):
        if self._drop_enabled or not sys.platform.startswith("win"):
            return
        if ctypes.sizeof(ctypes.c_void_p) != 8:
            return
        try:
            self.update_idletasks()
            user32 = ctypes.PyDLL("user32", use_last_error=True)
            shell32 = ctypes.PyDLL("shell32", use_last_error=True)
            long_ptr = ctypes.c_longlong

            wndproc_type = ctypes.PYFUNCTYPE(
                long_ptr,
                wintypes.HWND,
                wintypes.UINT,
                wintypes.WPARAM,
                wintypes.LPARAM,
            )
            self._drop_wndproc = wndproc_type(self.windows_drop_wndproc)
            new_proc = ctypes.cast(self._drop_wndproc, ctypes.c_void_p).value

            if ctypes.sizeof(ctypes.c_void_p) == 8:
                set_window_long = user32.SetWindowLongPtrW
            else:
                set_window_long = user32.SetWindowLongW
            set_window_long.argtypes = [wintypes.HWND, ctypes.c_int, long_ptr]
            set_window_long.restype = long_ptr
            user32.CallWindowProcW.argtypes = [
                long_ptr,
                wintypes.HWND,
                wintypes.UINT,
                wintypes.WPARAM,
                wintypes.LPARAM,
            ]
            user32.CallWindowProcW.restype = long_ptr
            user32.DefWindowProcW.argtypes = [
                wintypes.HWND,
                wintypes.UINT,
                wintypes.WPARAM,
                wintypes.LPARAM,
            ]
            user32.DefWindowProcW.restype = long_ptr
            shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
            shell32.DragQueryFileW.argtypes = [
                wintypes.HANDLE,
                wintypes.UINT,
                wintypes.LPWSTR,
                wintypes.UINT,
            ]
            shell32.DragQueryFileW.restype = wintypes.UINT
            shell32.DragFinish.argtypes = [wintypes.HANDLE]

            self._drop_user32 = user32
            self._drop_shell32 = shell32
            self._drop_set_window_long = set_window_long
            self._drop_long_ptr = long_ptr

            hwnd_value = self.html_list.winfo_id()
            hwnd = wintypes.HWND(hwnd_value)
            shell32.DragAcceptFiles(hwnd, True)
            old_proc = set_window_long(hwnd, GWL_WNDPROC, long_ptr(new_proc))
            self._drop_old_procs[hwnd_value] = old_proc
            self._drop_enabled = True
        except Exception:
            self._drop_enabled = False

    def disable_file_drop(self):
        if not self._drop_enabled or not sys.platform.startswith("win"):
            return
        try:
            for hwnd_value, old_proc in list(self._drop_old_procs.items()):
                hwnd = wintypes.HWND(hwnd_value)
                self._drop_shell32.DragAcceptFiles(hwnd, False)
                if old_proc:
                    self._drop_set_window_long(hwnd, GWL_WNDPROC, self._drop_long_ptr(old_proc))
        except Exception:
            pass
        finally:
            self._drop_old_procs.clear()
            self._drop_enabled = False

    def hwnd_key(self, hwnd):
        return hwnd.value if hasattr(hwnd, "value") else int(hwnd)

    def windows_drop_wndproc(self, hwnd, msg, wparam, lparam):
        try:
            if msg == WM_DROPFILES:
                try:
                    files = self.query_dropped_files(wparam)
                finally:
                    self._drop_shell32.DragFinish(wparam)
                self._drop_pending_files.extend(files)
                return 0
        except Exception:
            pass

        old_proc = self._drop_old_procs.get(self.hwnd_key(hwnd))
        if old_proc:
            return self._drop_user32.CallWindowProcW(old_proc, hwnd, msg, wparam, lparam)
        return self._drop_user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def query_dropped_files(self, hdrop):
        count = self._drop_shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
        files = []
        for index in range(count):
            length = self._drop_shell32.DragQueryFileW(hdrop, index, None, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            self._drop_shell32.DragQueryFileW(hdrop, index, buffer, length + 1)
            if buffer.value:
                files.append(buffer.value)
        return files

    def poll_dropped_files(self):
        if self._drop_pending_files:
            files = self._drop_pending_files[:]
            self._drop_pending_files.clear()
            self.handle_dropped_files(files)
        try:
            self._drop_poll_id = self.after(250, self.poll_dropped_files)
        except tk.TclError:
            self._drop_poll_id = None

    def add_html_paths(self, paths):
        existing = {
            os.path.normcase(os.path.abspath(path))
            for path in self.html_files
        }
        added = 0
        for path_text in paths:
            path = Path(path_text)
            if path.suffix.lower() not in {".html", ".htm"} or not path.is_file():
                continue
            full_path = str(path)
            normalized = os.path.normcase(os.path.abspath(full_path))
            if normalized in existing:
                continue
            self.html_files.append(full_path)
            self.html_list.insert("end", full_path)
            existing.add(normalized)
            added += 1
        return added

    def handle_dropped_files(self, paths):
        added = self.add_html_paths(paths)
        if added:
            self.notebook.select(self.html_tab)
            self.status_var.set(self.tr("drop_added", count=added))
        else:
            messagebox.showwarning(self.t("warning"), self.t("drop_unsupported"))

    def add_html_files(self):
        files = filedialog.askopenfilenames(
            title=self.t("add_html"),
            filetypes=[
                ("Bookmark HTML", "*.html *.htm"),
                ("All files", "*.*"),
            ],
        )
        added = self.add_html_paths(files)
        if added:
            self.status_var.set(self.tr("drop_added", count=added))

    def clear_html_files(self):
        self.html_files = []
        self.html_list.delete(0, "end")

    def get_output_folder(self):
        folder = ensure_output_folder(self.output_folder_var.get())
        self.output_folder_var.set(str(folder))
        return folder

    def choose_output_folder(self):
        current = self.output_folder_var.get().strip() or str(desktop_bookmarks_folder())
        folder = filedialog.askdirectory(initialdir=current, title=self.t("choose_output"))
        if folder:
            self.output_folder_var.set(folder)

    def convert_html_files(self):
        if not self.html_files:
            messagebox.showwarning(self.t("warning"), self.t("no_html"))
            return
        try:
            output = self.get_output_folder()
        except OSError as exc:
            messagebox.showerror(self.t("error"), str(exc))
            return

        def work():
            count = 0
            for file in self.html_files:
                rows = parse_bookmark_html(file)
                filename = safe_filename(Path(file).stem) + ".csv"
                destination = unique_path(output, filename)
                write_csv(destination, rows)
                count += 1
            return count, output

        def done(result, error):
            self.set_busy(False)
            if error:
                messagebox.showerror(self.t("error"), str(error))
                return
            count, output = result
            self.status_var.set(self.t("status_done"))
            messagebox.showinfo(self.t("info"), self.tr("done_html", count=count, folder=output))

        self.run_worker(work, done)

    def browse_folder(self, var):
        folder = filedialog.askdirectory(initialdir=var.get() or str(Path.home()))
        if folder:
            var.set(folder)

    def on_profile_browser_change(self, _event=None):
        self.remember_profile_base_path()
        self.profile_browser = self.browser_code_from_label(self.profile_browser_var.get())
        self.chrome_base_var.set(self.profile_base_paths.get(self.profile_browser, str(default_browser_user_data(self.profile_browser))))
        self.profiles = []
        self.profiles_base_path = ""
        self.profiles_browser = ""
        self.selected_profile_ids.clear()
        self.refresh_profile_tree()
        self.save_current_settings()

    def on_sheet_browser_change(self, _event=None):
        self.remember_sheet_base_path()
        self.sheet_browser = self.browser_code_from_label(self.sheet_browser_var.get())
        self.sheet_base_var.set(self.sheet_base_paths.get(self.sheet_browser, str(default_browser_user_data(self.sheet_browser))))
        self.sheet_profiles = []
        self.sheet_profiles_base_path = ""
        self.sheet_profiles_browser = ""
        self.selected_sheet_profile_ids.clear()
        self.refresh_sheet_profile_tree()
        self.save_current_settings()

    def browse_oauth_client(self):
        file = filedialog.askopenfilename(
            title=self.t("choose_oauth_client"),
            filetypes=[
                ("JSON", "*.json"),
                ("All files", "*.*"),
            ],
        )
        if file:
            self.oauth_client_var.set(file)
            self.save_current_settings()

    def open_sheet_url(self):
        sheet_url = self.sheet_url_var.get().strip()
        if not sheet_url:
            messagebox.showwarning(self.t("warning"), self.t("no_sheet_url"))
            return
        self.save_current_settings()
        try:
            webbrowser.open(sheet_url, new=1)
        except Exception as exc:
            messagebox.showerror(self.t("error"), str(exc))

    def scan_profiles(self):
        def work():
            base_path = self.current_profile_base_path()
            browser = self.profile_browser
            return base_path, browser, find_browser_bookmark_files(base_path, browser)

        def done(result, error):
            self.set_busy(False)
            if error:
                messagebox.showerror(self.t("error"), str(error))
                return
            base_path, browser, profiles = result
            self.profiles_base_path = base_path
            self.profiles_browser = browser
            self.profiles = profiles
            self.selected_profile_ids = {profile.profile_id for profile in profiles}
            self.refresh_profile_tree()
            if not profiles:
                messagebox.showwarning(self.t("warning"), self.t("no_profiles"))

        self.run_worker(work, done)

    def refresh_profile_tree(self):
        for item in self.profile_tree.get_children():
            self.profile_tree.delete(item)
        self.profile_item_ids = {}
        for index, profile in enumerate(self.profiles):
            item_id = f"profile_{index}"
            self.profile_item_ids[item_id] = profile.profile_id
            is_checked = profile.profile_id in self.selected_profile_ids
            image_key = "checked" if is_checked else "unchecked"
            tag = "checked" if is_checked else "unchecked"
            self.profile_tree.insert(
                "",
                "end",
                iid=item_id,
                image=self.checkbox_images[image_key],
                values=(profile.display_name, str(profile.path), len(profile.rows)),
                tags=(tag,),
            )

    def on_profile_tree_click(self, event):
        if self.profile_tree.identify("region", event.x, event.y) not in {"cell", "tree"}:
            return
        if self.profile_tree.identify_column(event.x) != "#0":
            return
        item_id = self.profile_tree.identify_row(event.y)
        if not item_id:
            return "break"
        profile_id = self.profile_item_ids.get(item_id)
        if not profile_id:
            return "break"
        if profile_id in self.selected_profile_ids:
            self.selected_profile_ids.remove(profile_id)
        else:
            self.selected_profile_ids.add(profile_id)
        self.refresh_profile_tree()
        return "break"

    def select_all_profiles(self):
        self.selected_profile_ids = {profile.profile_id for profile in self.profiles}
        self.refresh_profile_tree()

    def clear_profile_selection(self):
        self.selected_profile_ids.clear()
        self.refresh_profile_tree()

    def scan_sheet_profiles(self):
        def work():
            base_path = self.current_sheet_base_path()
            browser = self.sheet_browser
            return base_path, browser, find_browser_bookmark_files(base_path, browser)

        def done(result, error):
            self.set_busy(False)
            if error:
                messagebox.showerror(self.t("error"), str(error))
                return
            base_path, browser, profiles = result
            self.sheet_profiles_base_path = base_path
            self.sheet_profiles_browser = browser
            self.sheet_profiles = profiles
            self.selected_sheet_profile_ids = {profile.profile_id for profile in profiles}
            self.refresh_sheet_profile_tree()
            if not profiles:
                messagebox.showwarning(self.t("warning"), self.t("no_profiles"))

        self.run_worker(work, done)

    def refresh_sheet_profile_tree(self):
        for item in self.sheet_profile_tree.get_children():
            self.sheet_profile_tree.delete(item)
        self.sheet_profile_item_ids = {}
        for index, profile in enumerate(self.sheet_profiles):
            item_id = f"sheet_profile_{index}"
            self.sheet_profile_item_ids[item_id] = profile.profile_id
            is_checked = profile.profile_id in self.selected_sheet_profile_ids
            image_key = "checked" if is_checked else "unchecked"
            tag = "checked" if is_checked else "unchecked"
            self.sheet_profile_tree.insert(
                "",
                "end",
                iid=item_id,
                image=self.checkbox_images[image_key],
                values=(profile.display_name, str(profile.path), len(profile.rows)),
                tags=(tag,),
            )

    def on_sheet_profile_tree_click(self, event):
        if self.sheet_profile_tree.identify("region", event.x, event.y) not in {"cell", "tree"}:
            return
        if self.sheet_profile_tree.identify_column(event.x) != "#0":
            return
        item_id = self.sheet_profile_tree.identify_row(event.y)
        if not item_id:
            return "break"
        profile_id = self.sheet_profile_item_ids.get(item_id)
        if not profile_id:
            return "break"
        if profile_id in self.selected_sheet_profile_ids:
            self.selected_sheet_profile_ids.remove(profile_id)
        else:
            self.selected_sheet_profile_ids.add(profile_id)
        self.refresh_sheet_profile_tree()
        return "break"

    def select_all_sheet_profiles(self):
        self.selected_sheet_profile_ids = {profile.profile_id for profile in self.sheet_profiles}
        self.refresh_sheet_profile_tree()

    def clear_sheet_profile_selection(self):
        self.selected_sheet_profile_ids.clear()
        self.refresh_sheet_profile_tree()

    def ensure_profiles(self, base_path, browser):
        profiles = find_browser_bookmark_files(base_path, browser)
        if not profiles:
            raise RuntimeError(self.t("no_profiles"))
        return profiles

    def selected_sheet_profiles_or_all_from_scan(self):
        base_path = self.current_sheet_base_path()
        browser = self.sheet_browser
        if (
            not self.sheet_profiles
            or self.sheet_profiles_base_path != base_path
            or self.sheet_profiles_browser != browser
        ):
            profiles = self.ensure_profiles(base_path, browser)
            self.sheet_profiles = profiles
            self.sheet_profiles_base_path = base_path
            self.sheet_profiles_browser = browser
            self.selected_sheet_profile_ids = {profile.profile_id for profile in profiles}
            return profiles
        selected = [
            profile for profile in self.sheet_profiles
            if profile.profile_id in self.selected_sheet_profile_ids
        ]
        if not selected:
            raise RuntimeError(self.t("no_profiles_selected"))
        return selected

    def selected_profiles_or_all_from_scan(self):
        base_path = self.current_profile_base_path()
        browser = self.profile_browser
        if (
            not self.profiles
            or self.profiles_base_path != base_path
            or self.profiles_browser != browser
        ):
            profiles = self.ensure_profiles(base_path, browser)
            self.profiles = profiles
            self.profiles_base_path = base_path
            self.profiles_browser = browser
            self.selected_profile_ids = {profile.profile_id for profile in profiles}
            return profiles
        selected = [
            profile for profile in self.profiles
            if profile.profile_id in self.selected_profile_ids
        ]
        if not selected:
            raise RuntimeError(self.t("no_profiles_selected"))
        return selected

    def export_profiles_csv(self):
        try:
            output = self.get_output_folder()
        except OSError as exc:
            messagebox.showerror(self.t("error"), str(exc))
            return

        def work():
            base_path = self.current_profile_base_path()
            browser = self.profile_browser
            use_scanned_profiles = (
                bool(self.profiles)
                and self.profiles_base_path == base_path
                and self.profiles_browser == browser
            )
            all_profiles = self.profiles if use_scanned_profiles else self.ensure_profiles(base_path, browser)
            if use_scanned_profiles:
                profiles = [
                    profile for profile in all_profiles
                    if profile.profile_id in self.selected_profile_ids
                ]
                if not profiles:
                    raise RuntimeError(self.t("no_profiles_selected"))
            else:
                profiles = all_profiles
            if self.mode_var.get() == "merged":
                rows = []
                for profile in profiles:
                    rows.extend(profile.rows)
                destination = unique_path(output, f"all_{browser}_bookmarks.csv")
                write_csv(destination, rows)
                return "merged", destination, all_profiles, len(profiles), base_path, browser, use_scanned_profiles
            count = 0
            for profile in profiles:
                destination = unique_path(output, safe_filename(profile.display_name) + ".csv")
                write_csv(destination, profile.rows)
                count += 1
            return "separate", output, all_profiles, count, base_path, browser, use_scanned_profiles

        def done(result, error):
            self.set_busy(False)
            if error:
                messagebox.showerror(self.t("error"), str(error))
                return
            mode, target, all_profiles, exported_count, base_path, browser, used_scanned_profiles = result
            self.profiles = all_profiles
            self.profiles_base_path = base_path
            self.profiles_browser = browser
            if not used_scanned_profiles:
                self.selected_profile_ids = {profile.profile_id for profile in all_profiles}
            self.refresh_profile_tree()
            self.status_var.set(self.t("status_done"))
            if mode == "merged":
                messagebox.showinfo(self.t("info"), self.tr("done_merged", file=target))
            else:
                messagebox.showinfo(self.t("info"), self.tr("done_profiles", count=exported_count, folder=target))

        self.run_worker(work, done)

    def upload_to_sheet(self):
        sheet_url = self.sheet_url_var.get().strip()
        oauth_client = self.oauth_client_var.get().strip()
        if not sheet_url:
            messagebox.showwarning(self.t("warning"), self.t("no_sheet_url"))
            return
        if not oauth_client:
            messagebox.showwarning(self.t("warning"), self.t("no_oauth_client"))
            return
        self.save_current_settings()

        def work():
            base_path = self.current_sheet_base_path()
            browser = self.sheet_browser
            use_scanned_profiles = bool(
                self.sheet_profiles
                and self.sheet_profiles_base_path == base_path
                and self.sheet_profiles_browser == browser
            )
            all_profiles = self.sheet_profiles if use_scanned_profiles else self.ensure_profiles(base_path, browser)
            if use_scanned_profiles:
                profiles = [
                    profile for profile in all_profiles
                    if profile.profile_id in self.selected_sheet_profile_ids
                ]
                if not profiles:
                    raise RuntimeError(self.t("no_profiles_selected"))
            else:
                profiles = all_profiles
            if not any(profile.rows for profile in profiles):
                raise RuntimeError(self.t("no_profiles"))
            uploaded_profiles, rows = upload_profiles_to_google_sheets(sheet_url, oauth_client, profiles)
            return uploaded_profiles, rows, all_profiles, use_scanned_profiles, base_path, browser

        def done(result, error):
            self.set_busy(False)
            if error:
                if isinstance(error, ValueError):
                    messagebox.showerror(self.t("error"), self.t("oauth_client_invalid") + f"\n\n{error}")
                else:
                    messagebox.showerror(self.t("error"), str(error))
                return
            self.status_var.set(self.t("status_done"))
            profiles, rows, all_profiles, used_scanned_profiles, base_path, browser = result
            self.sheet_profiles = all_profiles
            self.sheet_profiles_base_path = base_path
            self.sheet_profiles_browser = browser
            if not used_scanned_profiles:
                self.selected_sheet_profile_ids = {profile.profile_id for profile in all_profiles}
            self.refresh_sheet_profile_tree()
            messagebox.showinfo(self.t("info"), self.tr("done_upload_profiles", profiles=profiles, rows=rows))

        self.run_worker(work, done)

    def open_output_folder(self):
        try:
            folder = self.get_output_folder()
            os.startfile(folder)
        except OSError as exc:
            messagebox.showerror(self.t("error"), str(exc))


def main():
    app = BookmarkApp()
    app.mainloop()


if __name__ == "__main__":
    main()
