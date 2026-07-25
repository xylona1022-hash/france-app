import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Layers, Shuffle, Volume2, Keyboard, ListChecks, Plus, Upload,
  ChevronLeft, ChevronRight, RotateCw, Check, X, Bell, Flame,
  FileSpreadsheet, ClipboardPaste, Trash2, Sparkles, Coffee,
  BookOpen, Highlighter, Eraser, Image as ImageIcon, FileText, StickyNote
} from "lucide-react";

/* ---------------------------------------------------------
   Storage helpers (window.storage — no localStorage allowed)
--------------------------------------------------------- */
const CARDS_KEY = "vocab-cards";
const ENERGY_KEY = "energy-state";
const ARTICLES_KEY = "reading-articles";
const CATEGORIES_KEY = "card-categories";
const UNCATEGORIZED = "未分類";
const DEFAULT_CATEGORIES = [UNCATEGORIZED];
const BUILTIN_IMPORTED_KEY = "builtin-daily-life-imported";
const BUILTIN_ZODIAC_IMPORTED_KEY = "builtin-zodiac-imported";
const BUILTIN_READING_IMPORTED_KEY = "builtin-reading-words-imported";
const VOICE_KEY = "tts-voice-uri";

async function loadJSON(key, fallback) {
  try {
    const res = await window.storage.get(key, false);
    return res ? JSON.parse(res.value) : fallback;
  } catch {
    return fallback;
  }
}
async function saveJSON(key, value) {
  try {
    await window.storage.set(key, JSON.stringify(value), false);
  } catch (e) {
    console.error("storage save failed", e);
  }
}

const todayStr = () => new Date().toISOString().slice(0, 10);
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

const TASKS = [
  { id: "flashcards", label: "字卡複習", short: "字卡", icon: Layers, color: "#2F6F62" },
  { id: "matching", label: "配對測驗", short: "配對", icon: Shuffle, color: "#3E7CB1" },
  { id: "reading", label: "閱讀文章", short: "閱讀", icon: BookOpen, color: "#B1793E" },
  { id: "typing", label: "拼字測驗", short: "拼字", icon: Keyboard, color: "#8C5B9F" },
  { id: "multiplechoice", label: "選擇題", short: "選擇題", icon: ListChecks, color: "#C15B5B" },
];

const SEED_CARDS = [
  { id: uid(), word: "ambitious", pos: "adj.", phonetic: "/æmˈbɪʃəs/", definition: "有野心的、有抱負的", example: "She is an ambitious student who dreams of working in Paris.", category: "求職" },
  { id: uid(), word: "internship", pos: "n.", phonetic: "/ˈɪntɜːrnʃɪp/", definition: "實習", example: "Maggie is applying for a six-month paid internship in France.", category: "求職" },
  { id: uid(), word: "brand", pos: "n.", phonetic: "/brænd/", definition: "品牌", example: "Luxury brands often value candidates who understand both markets.", category: "行銷" },
  { id: uid(), word: "resilient", pos: "adj.", phonetic: "/rɪˈzɪljənt/", definition: "有韌性的、能迅速恢復的", example: "Staying resilient after rejection is part of the job search.", category: "求職" },
];
const SEED_CATEGORIES = [UNCATEGORIZED, "求職", "行銷"];

/* Built-in "Daily Life" vocabulary set (220 words), imported from the user's uploaded Excel file */
const BUILTIN_DAILY_LIFE_CARDS = [{"word": "alarm clock", "pos": "n.", "phonetic": "", "definition": "鬧鐘", "example": "I set my alarm clock for 6:30 AM every morning.", "collocation": "set an alarm clock", "category": "日常生活"}, {"word": "pillow", "pos": "n.", "phonetic": "", "definition": "枕頭", "example": "She rested her head on a soft pillow.", "collocation": "plump up a pillow", "category": "日常生活"}, {"word": "bedside", "pos": "n. / adj.", "phonetic": "", "definition": "床邊 / 床頭", "example": "He placed his glass of water on the bedside table.", "collocation": "bedside table / lamp", "category": "日常生活"}, {"word": "mattress", "pos": "n.", "phonetic": "", "definition": "床墊", "example": "A comfortable mattress is essential for good sleep.", "collocation": "firm / comfortable mattress", "category": "日常生活"}, {"word": "comforter", "pos": "n.", "phonetic": "", "definition": "棉被", "example": "She pulled the warm comforter over her shoulders.", "collocation": "cozy comforter", "category": "日常生活"}, {"word": "bedding", "pos": "n.", "phonetic": "", "definition": "寢具", "example": "We need to wash and change the bedding this weekend.", "collocation": "change the bedding", "category": "日常生活"}, {"word": "pajamas", "pos": "n. (pl.)", "phonetic": "", "definition": "睡衣", "example": "He changed into his comfortable silk pajamas before bed.", "collocation": "put on / wear pajamas", "category": "日常生活"}, {"word": "pantyhose", "pos": "n. (pl.)", "phonetic": "", "definition": "褲襪", "example": "She wore black pantyhose with her formal dress.", "collocation": "a pair of pantyhose", "category": "日常生活"}, {"word": "clothes", "pos": "n. (pl.)", "phonetic": "", "definition": "衣服", "example": "She picked out her clothes for tomorrow morning.", "collocation": "wear / change clothes", "category": "日常生活"}, {"word": "lotion", "pos": "n.", "phonetic": "", "definition": "乳液", "example": "Apply lotion after showering to keep your skin hydrated.", "collocation": "apply body lotion", "category": "日常生活"}, {"word": "eyebrow pencil", "pos": "n.", "phonetic": "", "definition": "眉筆", "example": "She used an eyebrow pencil to define her brows.", "collocation": "draw eyebrows with a pencil", "category": "日常生活"}, {"word": "powder", "pos": "n.", "phonetic": "", "definition": "蜜粉 / 粉末", "example": "She dusted some powder on her face to reduce shine.", "collocation": "apply face powder", "category": "日常生活"}, {"word": "makeup", "pos": "n.", "phonetic": "", "definition": "化妝品 / 彩妝", "example": "It takes her fifteen minutes to put on makeup.", "collocation": "put on / wear makeup", "category": "日常生活"}, {"word": "sunscreen lotion", "pos": "n.", "phonetic": "", "definition": "防曬乳液", "example": "Don't forget to apply sunscreen lotion before going outside.", "collocation": "apply sunscreen lotion", "category": "日常生活"}, {"word": "lipstick", "pos": "n.", "phonetic": "", "definition": "唇膏 / 口紅", "example": "She chose a bright red lipstick for the party.", "collocation": "apply / put on lipstick", "category": "日常生活"}, {"word": "blush", "pos": "n.", "phonetic": "", "definition": "腮紅", "example": "A touch of blush gave her cheeks a healthy glow.", "collocation": "apply blush", "category": "日常生活"}, {"word": "contact lenses", "pos": "n. (pl.)", "phonetic": "", "definition": "隱形眼鏡", "example": "He switched from glasses to contact lenses last year.", "collocation": "wear / insert contact lenses", "category": "日常生活"}, {"word": "toast", "pos": "n.", "phonetic": "", "definition": "吐司", "example": "I had two slices of toast with jam for breakfast.", "collocation": "butter / slice of toast", "category": "日常生活"}, {"word": "yesterday's leftovers", "pos": "n. (pl.)", "phonetic": "", "definition": "昨天的剩菜", "example": "We heated up yesterday's leftovers for lunch.", "collocation": "eat / heat up leftovers", "category": "日常生活"}, {"word": "supermarket flyer", "pos": "n.", "phonetic": "", "definition": "超市傳單", "example": "She checked the supermarket flyer for weekly discounts.", "collocation": "check the supermarket flyer", "category": "日常生活"}, {"word": "morning drama series", "pos": "n.", "phonetic": "", "definition": "晨間連續劇", "example": "My grandmother enjoys watching the morning drama series.", "collocation": "watch a drama series", "category": "日常生活"}, {"word": "outlet", "pos": "n.", "phonetic": "", "definition": "電源插座", "example": "He plugged his laptop charger into the wall outlet.", "collocation": "wall outlet / plug into an outlet", "category": "日常生活"}, {"word": "plug", "pos": "n.", "phonetic": "", "definition": "插頭", "example": "Insert the plug firmly into the socket.", "collocation": "pull the plug / power plug", "category": "日常生活"}, {"word": "folding umbrella", "pos": "n.", "phonetic": "", "definition": "折傘", "example": "Keep a folding umbrella in your bag just in case it rains.", "collocation": "carry a folding umbrella", "category": "日常生活"}, {"word": "sunshade umbrella", "pos": "n.", "phonetic": "", "definition": "晴雨兩用傘 / 陽傘", "example": "She used a sunshade umbrella to protect her skin from UV rays.", "collocation": "use a sunshade umbrella", "category": "日常生活"}, {"word": "crosswalk", "pos": "n.", "phonetic": "", "definition": "行人穿越道 / 斑馬線", "example": "Always look both ways before using the crosswalk.", "collocation": "pedestrian crosswalk / use the crosswalk", "category": "日常生活"}, {"word": "light", "pos": "n.", "phonetic": "", "definition": "號誌燈 (紅綠燈)", "example": "Stop when the traffic light turns red.", "collocation": "traffic light / green light", "category": "日常生活"}, {"word": "pedestrian", "pos": "n.", "phonetic": "", "definition": "行人", "example": "The street was crowded with pedestrians during rush hour.", "collocation": "pedestrian zone / crossing", "category": "日常生活"}, {"word": "underpass", "pos": "n.", "phonetic": "", "definition": "地下道", "example": "Pedestrians can safely cross the busy road via the underpass.", "collocation": "pedestrian underpass", "category": "日常生活"}, {"word": "station", "pos": "n.", "phonetic": "", "definition": "車站", "example": "Meet me at the main entrance of the train station.", "collocation": "train / subway station", "category": "日常生活"}, {"word": "ticket gate", "pos": "n.", "phonetic": "", "definition": "驗票口 / 閘門", "example": "Scan your ticket at the ticket gate to enter the platform.", "collocation": "pass through the ticket gate", "category": "日常生活"}, {"word": "prepaid railway pass", "pos": "n.", "phonetic": "", "definition": "儲值火車票卡 / 交通卡", "example": "You can tap your prepaid railway pass at the gate.", "collocation": "tap a prepaid railway pass", "category": "日常生活"}, {"word": "train", "pos": "n.", "phonetic": "", "definition": "列車 / 火車", "example": "The commuter train leaves promptly at 8:00 AM.", "collocation": "catch / board a train", "category": "日常生活"}, {"word": "line", "pos": "n.", "phonetic": "", "definition": "排隊隊伍 / 路線", "example": "Passengers stood in line waiting to buy tickets.", "collocation": "stand in line / queue line", "category": "日常生活"}, {"word": "boarding point", "pos": "n.", "phonetic": "", "definition": "乘車位置 / 候車處", "example": "Please wait for the bus at the designated boarding point.", "collocation": "designated boarding point", "category": "日常生活"}, {"word": "platform", "pos": "n.", "phonetic": "", "definition": "月台", "example": "The train to Taipei will arrive at Platform 2.", "collocation": "train platform / Platform 3", "category": "日常生活"}, {"word": "passenger", "pos": "n.", "phonetic": "", "definition": "乘客", "example": "All passengers must remain seated until the bus stops.", "collocation": "rail passenger / transit passenger", "category": "日常生活"}, {"word": "station employee", "pos": "n.", "phonetic": "", "definition": "站務人員", "example": "If you lose your card, ask a station employee for help.", "collocation": "ask a station employee", "category": "日常生活"}, {"word": "local train", "pos": "n.", "phonetic": "", "definition": "普通車 / 各站停車火車", "example": "The local train stops at every station along the route.", "collocation": "take a local train", "category": "日常生活"}, {"word": "time recorder", "pos": "n.", "phonetic": "", "definition": "打卡鐘", "example": "Employees clock in using the time recorder near the door.", "collocation": "punch in at the time recorder", "category": "日常生活"}, {"word": "ID card", "pos": "n.", "phonetic": "", "definition": "識別證", "example": "You must wear your ID card to enter the office building.", "collocation": "wear / show an ID card", "category": "日常生活"}, {"word": "telephone", "pos": "n.", "phonetic": "", "definition": "電話", "example": "The receptionist answered the ringing telephone promptly.", "collocation": "answer the telephone", "category": "日常生活"}, {"word": "handset", "pos": "n.", "phonetic": "", "definition": "電話聽筒", "example": "He picked up the handset and dialed the extension number.", "collocation": "pick up / lift the handset", "category": "日常生活"}, {"word": "extension line", "pos": "n.", "phonetic": "", "definition": "分機", "example": "Please dial extension line 402 to reach accounting.", "collocation": "call an extension line", "category": "日常生活"}, {"word": "external line", "pos": "n.", "phonetic": "", "definition": "外線", "example": "Press 9 to get an external line for outside calls.", "collocation": "dial an external line", "category": "日常生活"}, {"word": "calculator", "pos": "n.", "phonetic": "", "definition": "計算機", "example": "She used a calculator to double-check the financial figures.", "collocation": "pocket / electronic calculator", "category": "日常生活"}, {"word": "mechanical pencil", "pos": "n.", "phonetic": "", "definition": "自動鉛筆", "example": "He wrote down notes using a 0.5mm mechanical pencil.", "collocation": "refill a mechanical pencil", "category": "日常生活"}, {"word": "marker", "pos": "n.", "phonetic": "", "definition": "螢光筆 / 標記筆", "example": "Highlight key deadlines with a yellow marker.", "collocation": "highlighter marker", "category": "日常生活"}, {"word": "memo pad", "pos": "n.", "phonetic": "", "definition": "便條紙簿", "example": "She jotted down the meeting notes on a memo pad.", "collocation": "write on a memo pad", "category": "日常生活"}, {"word": "receipt", "pos": "n.", "phonetic": "", "definition": "收據", "example": "Keep your sales receipt as proof of purchase.", "collocation": "issue / request a receipt", "category": "日常生活"}, {"word": "bill / invoice", "pos": "n.", "phonetic": "", "definition": "帳單 / 發票", "example": "The company issued an invoice for the consulting services.", "collocation": "pay a bill / send an invoice", "category": "日常生活"}, {"word": "drawer", "pos": "n.", "phonetic": "", "definition": "抽屜", "example": "Keep your passport and important documents in the desk drawer.", "collocation": "desk / top drawer", "category": "日常生活"}, {"word": "swivel chair", "pos": "n.", "phonetic": "", "definition": "旋轉椅", "example": "An adjustable swivel chair helps prevent back fatigue.", "collocation": "ergonomic swivel chair", "category": "日常生活"}, {"word": "office supplies", "pos": "n. (pl.)", "phonetic": "", "definition": "辦公室用品", "example": "The administrative department ordered fresh office supplies.", "collocation": "order office supplies", "category": "日常生活"}, {"word": "material", "pos": "n.", "phonetic": "", "definition": "資料 / 材料", "example": "Please review the presentation material before tomorrow's meeting.", "collocation": "training / meeting material", "category": "日常生活"}, {"word": "file cabinet", "pos": "n.", "phonetic": "", "definition": "檔案櫃", "example": "Confidential files are stored securely in the file cabinet.", "collocation": "lockable file cabinet", "category": "日常生活"}, {"word": "bulletin board", "pos": "n.", "phonetic": "", "definition": "佈告欄", "example": "Important notices are posted on the office bulletin board.", "collocation": "post on the bulletin board", "category": "日常生活"}, {"word": "copier", "pos": "n.", "phonetic": "", "definition": "影印機", "example": "The copier is out of paper; please load a new ream.", "collocation": "photocopier / operate a copier", "category": "日常生活"}, {"word": "fax machine", "pos": "n.", "phonetic": "", "definition": "傳真機", "example": "Although outdated, some departments still use a fax machine.", "collocation": "send via fax machine", "category": "日常生活"}, {"word": "equipment", "pos": "n. (uncount.)", "phonetic": "", "definition": "設備 / 器材", "example": "The company invested in high-tech office equipment.", "collocation": "office / IT equipment", "category": "日常生活"}, {"word": "boss", "pos": "n.", "phonetic": "", "definition": "老闆", "example": "She presented the project proposal directly to her boss.", "collocation": "report to the boss", "category": "日常生活"}, {"word": "staff", "pos": "n.", "phonetic": "", "definition": "員工 / 全體職員", "example": "Our staff is highly trained and dedicated to quality service.", "collocation": "office staff / staff member", "category": "日常生活"}, {"word": "document", "pos": "n.", "phonetic": "", "definition": "文件", "example": "Please sign and return the official document by Friday.", "collocation": "signed / official document", "category": "日常生活"}, {"word": "client / customer", "pos": "n.", "phonetic": "", "definition": "客戶 ; 顧客", "example": "Building good relationships with clients is vital for sales growth.", "collocation": "client meeting / potential customer", "category": "日常生活"}, {"word": "business card", "pos": "n.", "phonetic": "", "definition": "名片", "example": "They exchanged business cards at the beginning of the meeting.", "collocation": "exchange business cards", "category": "日常生活"}, {"word": "cellphone", "pos": "n.", "phonetic": "", "definition": "手機", "example": "She checked her cellphone for unread messages.", "collocation": "smartphone / cellphone screen", "category": "日常生活"}, {"word": "stand-by screen", "pos": "n.", "phonetic": "", "definition": "待機畫面", "example": "He set a photo of his family as his stand-by screen.", "collocation": "wallpaper on stand-by screen", "category": "日常生活"}, {"word": "ringtone", "pos": "n.", "phonetic": "", "definition": "手機鈴聲", "example": "His cellphone played a lively ringtone when it rang.", "collocation": "custom ringtone", "category": "日常生活"}, {"word": "silent mode", "pos": "n.", "phonetic": "", "definition": "靜音模式", "example": "Please switch your mobile phones to silent mode during meetings.", "collocation": "switch to silent mode", "category": "日常生活"}, {"word": "phone number", "pos": "n.", "phonetic": "", "definition": "電話號碼", "example": "Could you please give me your personal phone number?", "collocation": "contact / phone number", "category": "日常生活"}, {"word": "address list", "pos": "n.", "phonetic": "", "definition": "通訊錄", "example": "Her contact details were saved directly to my address list.", "collocation": "save to address list", "category": "日常生活"}, {"word": "personal computer", "pos": "n.", "phonetic": "", "definition": "（個人）電腦", "example": "He uses his personal computer for both work and gaming.", "collocation": "PC / personal computer setup", "category": "日常生活"}, {"word": "desktop computer", "pos": "n.", "phonetic": "", "definition": "桌上型電腦", "example": "Desktop computers generally offer higher processing performance.", "collocation": "workstation / desktop computer", "category": "日常生活"}, {"word": "laptop", "pos": "n.", "phonetic": "", "definition": "筆記型電腦", "example": "She carried her laptop in a padded backpack.", "collocation": "portable laptop", "category": "日常生活"}, {"word": "keyboard", "pos": "n.", "phonetic": "", "definition": "鍵盤", "example": "He typed rapidly on his mechanical keyboard.", "collocation": "type on a keyboard", "category": "日常生活"}, {"word": "monitor", "pos": "n.", "phonetic": "", "definition": "顯示器 / 螢幕", "example": "A dual-monitor setup can significantly boost productivity.", "collocation": "computer monitor", "category": "日常生活"}, {"word": "mouse", "pos": "n.", "phonetic": "", "definition": "滑鼠", "example": "Click the left button on your mouse to select an icon.", "collocation": "wireless mouse / optical mouse", "category": "日常生活"}, {"word": "USB flash drive", "pos": "n.", "phonetic": "", "definition": "USB 隨身碟", "example": "Save a backup copy of your document onto a USB flash drive.", "collocation": "insert a USB flash drive", "category": "日常生活"}, {"word": "icon", "pos": "n.", "phonetic": "", "definition": "圖示", "example": "Double-click the desktop icon to open the application.", "collocation": "desktop icon / click an icon", "category": "日常生活"}, {"word": "picture", "pos": "n.", "phonetic": "", "definition": "圖片 / 相片", "example": "She attached a high-resolution picture to the email.", "collocation": "high-resolution picture", "category": "日常生活"}, {"word": "video", "pos": "n.", "phonetic": "", "definition": "影片", "example": "They watched an instructional video before starting the project.", "collocation": "streaming / short video", "category": "日常生活"}, {"word": "OS (operating system)", "pos": "n.", "phonetic": "", "definition": "作業系統", "example": "Ensure your OS is updated with the latest security patches.", "collocation": "install / upgrade OS", "category": "日常生活"}, {"word": "software", "pos": "n. (uncount.)", "phonetic": "", "definition": "軟體", "example": "Our team uses specialized software for data analysis.", "collocation": "software update / application", "category": "日常生活"}, {"word": "memory", "pos": "n.", "phonetic": "", "definition": "記憶體", "example": "Adding more memory helps your computer run faster.", "collocation": "RAM memory / system memory", "category": "日常生活"}, {"word": "data", "pos": "n.", "phonetic": "", "definition": "資料 / 數據", "example": "The analyst processed the survey data to draw conclusions.", "collocation": "analyze / store data", "category": "日常生活"}, {"word": "website", "pos": "n.", "phonetic": "", "definition": "網站", "example": "Visit our official website for further information.", "collocation": "visit / build a website", "category": "日常生活"}, {"word": "search engine", "pos": "n.", "phonetic": "", "definition": "搜尋引擎", "example": "Google is the most widely used search engine globally.", "collocation": "search engine optimization", "category": "日常生活"}, {"word": "e-mail", "pos": "n.", "phonetic": "", "definition": "電子郵件", "example": "I sent an e-mail to confirm our appointment time.", "collocation": "send / receive e-mail", "category": "日常生活"}, {"word": "attached file", "pos": "n.", "phonetic": "", "definition": "附檔 / 附件", "example": "Please find the requested document in the attached file.", "collocation": "download / open attached file", "category": "日常生活"}, {"word": "junk mail", "pos": "n. (uncount.)", "phonetic": "", "definition": "垃圾郵件", "example": "Check your junk mail folder if you haven't received the code.", "collocation": "spam / filter junk mail", "category": "日常生活"}, {"word": "knitting", "pos": "n.", "phonetic": "", "definition": "編織〔物〕", "example": "Knitting is a relaxing craft that she enjoys during winter.", "collocation": "knitting needles / yarn", "category": "日常生活"}, {"word": "beadwork", "pos": "n.", "phonetic": "", "definition": "串珠", "example": "She created an elegant bracelet using colorful beadwork.", "collocation": "handmade beadwork", "category": "日常生活"}, {"word": "handicraft", "pos": "n.", "phonetic": "", "definition": "手工藝品", "example": "The night market sells a variety of traditional handicrafts.", "collocation": "local handicraft", "category": "日常生活"}, {"word": "fancywork", "pos": "n.", "phonetic": "", "definition": "編織品 ; 刺繡", "example": "Her grandmother spent hours creating detailed fancywork.", "collocation": "decorative fancywork", "category": "日常生活"}, {"word": "aroma oil", "pos": "n.", "phonetic": "", "definition": "芳香精油", "example": "Diffusing aroma oil creates a soothing atmosphere in the room.", "collocation": "essential / aroma oil", "category": "日常生活"}, {"word": "burner", "pos": "n.", "phonetic": "", "definition": "爐子 / 精油爐", "example": "She lit a tea candle beneath the ceramic aroma burner.", "collocation": "oil burner / stove burner", "category": "日常生活"}, {"word": "scented candle", "pos": "n.", "phonetic": "", "definition": "香氛蠟燭", "example": "A scented candle filled the room with a pleasant lavender aroma.", "collocation": "light a scented candle", "category": "日常生活"}, {"word": "incense", "pos": "n. (uncount.)", "phonetic": "", "definition": "香", "example": "They burn incense during traditional family ceremonies.", "collocation": "burn incense stick", "category": "日常生活"}, {"word": "herbal tea", "pos": "n.", "phonetic": "", "definition": "花草茶", "example": "Drinking a warm cup of herbal tea helps relieve stress.", "collocation": "sip hot herbal tea", "category": "日常生活"}, {"word": "magazine", "pos": "n.", "phonetic": "", "definition": "雜誌", "example": "He flipped through a design magazine while waiting.", "collocation": "fashion / monthly magazine", "category": "日常生活"}, {"word": "comic series", "pos": "n.", "phonetic": "", "definition": "連載漫畫", "example": "This popular comic series has been running for over ten years.", "collocation": "read a comic series", "category": "日常生活"}, {"word": "paperback", "pos": "n.", "phonetic": "", "definition": "平裝書", "example": "He prefers buying paperback books because they are lightweight.", "collocation": "paperback edition", "category": "日常生活"}, {"word": "picture book", "pos": "n.", "phonetic": "", "definition": "圖畫書 / 繪本", "example": "Parents love reading colorful picture books to young children.", "collocation": "children's picture book", "category": "日常生活"}, {"word": "photo album", "pos": "n.", "phonetic": "", "definition": "相簿", "example": "They looked through old photo albums to reminisce about the past.", "collocation": "family photo album", "category": "日常生活"}, {"word": "television", "pos": "n.", "phonetic": "", "definition": "電視", "example": "The family gathered around the television to watch the news.", "collocation": "watch television", "category": "日常生活"}, {"word": "channel", "pos": "n.", "phonetic": "", "definition": "頻道", "example": "He used the remote to switch to the news channel.", "collocation": "change the channel", "category": "日常生活"}, {"word": "cable television", "pos": "n.", "phonetic": "", "definition": "有線電視", "example": "Cable television offers a wide selection of international channels.", "collocation": "cable TV subscriber", "category": "日常生活"}, {"word": "satellite broadcasting", "pos": "n.", "phonetic": "", "definition": "衛星播送", "example": "Satellite broadcasting allows access to remote network feeds.", "collocation": "satellite TV dish", "category": "日常生活"}, {"word": "digital broadcasting", "pos": "n.", "phonetic": "", "definition": "數位播送", "example": "Digital broadcasting delivers crisp video and crystal-clear audio.", "collocation": "high-definition digital broadcasting", "category": "日常生活"}, {"word": "cartoon", "pos": "n.", "phonetic": "", "definition": "卡通", "example": "Children enjoy watching animated cartoons on Saturday mornings.", "collocation": "animated cartoon", "category": "日常生活"}, {"word": "gossip show", "pos": "n.", "phonetic": "", "definition": "八卦節目", "example": "They chatted about the rumors revealed on the gossip show.", "collocation": "celebrity gossip show", "category": "日常生活"}, {"word": "drama series", "pos": "n.", "phonetic": "", "definition": "連續劇", "example": "The TV network aired a compelling new drama series.", "collocation": "popular drama series", "category": "日常生活"}, {"word": "video game", "pos": "n.", "phonetic": "", "definition": "電動 / 電子遊戲", "example": "He spends his weekends playing console video games with friends.", "collocation": "play video games", "category": "日常生活"}, {"word": "remote control", "pos": "n.", "phonetic": "", "definition": "遙控器", "example": "Where did you leave the remote control for the television?", "collocation": "TV remote control", "category": "日常生活"}, {"word": "board game", "pos": "n.", "phonetic": "", "definition": "棋盤遊戲 / 桌遊", "example": "Playing board games with family is a great weekend activity.", "collocation": "play a board game", "category": "日常生活"}, {"word": "cards", "pos": "n. (pl.)", "phonetic": "", "definition": "撲克牌", "example": "We played a game of cards after finishing dinner.", "collocation": "deck of cards / play cards", "category": "日常生活"}, {"word": "wait time", "pos": "n.", "phonetic": "", "definition": "等待的時間", "example": "The estimated wait time for the amusement ride was 45 minutes.", "collocation": "estimated wait time", "category": "日常生活"}, {"word": "meeting place", "pos": "n.", "phonetic": "", "definition": "會面地點", "example": "The clock tower served as our agreed meeting place.", "collocation": "designated meeting place", "category": "日常生活"}, {"word": "coffee shop", "pos": "n.", "phonetic": "", "definition": "咖啡廳", "example": "They met at a local coffee shop to catch up over lattes.", "collocation": "cozy coffee shop", "category": "日常生活"}, {"word": "Sunday best", "pos": "n.", "phonetic": "", "definition": "盛裝 / 假日最好衣服", "example": "Everyone wore their Sunday best to the traditional wedding.", "collocation": "wear one's Sunday best", "category": "日常生活"}, {"word": "date", "pos": "n. / v.", "phonetic": "", "definition": "約會", "example": "They are planning to go on a romantic dinner date this Friday.", "collocation": "go on a date", "category": "日常生活"}, {"word": "digital camera", "pos": "n.", "phonetic": "", "definition": "數位相機", "example": "He captured beautiful travel landscapes using a digital camera.", "collocation": "compact digital camera", "category": "日常生活"}, {"word": "amusement park", "pos": "n.", "phonetic": "", "definition": "遊樂園", "example": "The children were thrilled about visiting the amusement park.", "collocation": "theme / amusement park", "category": "日常生活"}, {"word": "haunted house", "pos": "n.", "phonetic": "", "definition": "鬼屋", "example": "We decided to brave the spooky haunted house at the fair.", "collocation": "spooky haunted house", "category": "日常生活"}, {"word": "roller coaster", "pos": "n.", "phonetic": "", "definition": "雲霄飛車", "example": "The high-speed roller coaster provided an exhilarating experience.", "collocation": "thrilling roller coaster", "category": "日常生活"}, {"word": "Ferris wheel", "pos": "n.", "phonetic": "", "definition": "摩天輪", "example": "Riding the giant Ferris wheel gave us a panoramic view of the city.", "collocation": "ride a Ferris wheel", "category": "日常生活"}, {"word": "attraction", "pos": "n.", "phonetic": "", "definition": "景點 / 觀光勝地", "example": "The historical castle is the city's most famous attraction.", "collocation": "tourist attraction", "category": "日常生活"}, {"word": "theater", "pos": "n.", "phonetic": "", "definition": "戲院 / 劇院", "example": "We bought tickets to watch the premiere at the local theater.", "collocation": "movie theater", "category": "日常生活"}, {"word": "movie", "pos": "n.", "phonetic": "", "definition": "電影", "example": "They decided to watch an action movie at the cinema.", "collocation": "watch a movie", "category": "日常生活"}, {"word": "reserved seat", "pos": "n.", "phonetic": "", "definition": "保留位", "example": "We had reserved seats in the middle section of the theater.", "collocation": "book a reserved seat", "category": "日常生活"}, {"word": "trailer", "pos": "n.", "phonetic": "", "definition": "預告片", "example": "The movie trailer looked exciting and generated a lot of buzz.", "collocation": "movie trailer", "category": "日常生活"}, {"word": "showing", "pos": "n.", "phonetic": "", "definition": "上映 / 放映", "example": "The next showing of the film starts at 7:00 PM.", "collocation": "latest showing", "category": "日常生活"}, {"word": "R-rated movie", "pos": "n.", "phonetic": "", "definition": "限制級電影", "example": "You must be 18 or older to watch an R-rated movie.", "collocation": "watch an R-rated movie", "category": "日常生活"}, {"word": "complimentary ticket", "pos": "n.", "phonetic": "", "definition": "招待券", "example": "She received two complimentary tickets to the premiere.", "collocation": "free / complimentary ticket", "category": "日常生活"}, {"word": "shopping center", "pos": "n.", "phonetic": "", "definition": "購物中心", "example": "The new shopping center features over a hundred popular brands.", "collocation": "large shopping center", "category": "日常生活"}, {"word": "shop", "pos": "n. / v.", "phonetic": "", "definition": "商店", "example": "We visited a charming antique shop near the central square.", "collocation": "gift shop / shop for clothes", "category": "日常生活"}, {"word": "price tag", "pos": "n.", "phonetic": "", "definition": "價格標籤", "example": "He checked the price tag before deciding to make the purchase.", "collocation": "check the price tag", "category": "日常生活"}, {"word": "sales clerk", "pos": "n.", "phonetic": "", "definition": "店員 ; 銷售員", "example": "The helpful sales clerk assisted us in finding the right size.", "collocation": "friendly sales clerk", "category": "日常生活"}, {"word": "fitting room", "pos": "n.", "phonetic": "", "definition": "試衣間", "example": "She took three dresses into the fitting room to try them on.", "collocation": "try on in the fitting room", "category": "日常生活"}, {"word": "sale", "pos": "n.", "phonetic": "", "definition": "特賣 / 促銷", "example": "The store is holding a big summer sale with great discounts.", "collocation": "clearance / seasonal sale", "category": "日常生活"}, {"word": "bargain", "pos": "n. / v.", "phonetic": "", "definition": "特價商品 / 便宜貨", "example": "Finding a designer coat at half price was a real bargain.", "collocation": "bargain price / hunt for bargains", "category": "日常生活"}, {"word": "classy restaurant", "pos": "n.", "phonetic": "", "definition": "高級餐廳", "example": "They celebrated their wedding anniversary at a classy restaurant.", "collocation": "dine at a classy restaurant", "category": "日常生活"}, {"word": "drink", "pos": "n. / v.", "phonetic": "", "definition": "飲料", "example": "Would you like something cold to drink after our meal?", "collocation": "soft drink / grab a drink", "category": "日常生活"}, {"word": "wine list", "pos": "n.", "phonetic": "", "definition": "酒單", "example": "The waiter handed us the menu and the wine list.", "collocation": "request the wine list", "category": "日常生活"}, {"word": "pitcher", "pos": "n.", "phonetic": "", "definition": "水罐 / 水壺", "example": "The waiter brought a large pitcher of iced water to our table.", "collocation": "pitcher of water / iced tea", "category": "日常生活"}, {"word": "plate", "pos": "n.", "phonetic": "", "definition": "盤子", "example": "She placed a warm dinner plate in front of each guest.", "collocation": "dinner / salad plate", "category": "日常生活"}, {"word": "cutlery", "pos": "n. (uncount.)", "phonetic": "", "definition": "刀叉餐具", "example": "The waiter placed polished stainless steel cutlery on the table.", "collocation": "set of cutlery", "category": "日常生活"}, {"word": "appetizer", "pos": "n.", "phonetic": "", "definition": "前菜", "example": "We ordered garlic bread as an appetizer while waiting for the main meal.", "collocation": "order an appetizer", "category": "日常生活"}, {"word": "main course", "pos": "n.", "phonetic": "", "definition": "主菜", "example": "For the main course, I selected grilled salmon with asparagus.", "collocation": "serve the main course", "category": "日常生活"}, {"word": "dessert", "pos": "n.", "phonetic": "", "definition": "甜點", "example": "Save room for dessert because the chocolate cake is delicious.", "collocation": "order / serve dessert", "category": "日常生活"}, {"word": "sampler", "pos": "n.", "phonetic": "", "definition": "拼盤", "example": "We ordered a cheese and meat sampler to share with everyone.", "collocation": "appetizer sampler", "category": "日常生活"}, {"word": "party", "pos": "n.", "phonetic": "", "definition": "團體聚會 / 宴會", "example": "Our department hosted a lively holiday party at the hotel.", "collocation": "dinner party", "category": "日常生活"}, {"word": "tavern", "pos": "n.", "phonetic": "", "definition": "居酒屋 ; 小酒館", "example": "They gathered at a cozy neighborhood tavern after work.", "collocation": "local / traditional tavern", "category": "日常生活"}, {"word": "food stall", "pos": "n.", "phonetic": "", "definition": "小吃攤", "example": "The night market is famous for its delicious food stalls.", "collocation": "street food stall", "category": "日常生活"}, {"word": "shop curtain", "pos": "n.", "phonetic": "", "definition": "店舖門簾", "example": "A traditional Japanese shop curtain hung at the restaurant entrance.", "collocation": "hang a shop curtain", "category": "日常生活"}, {"word": "liquor", "pos": "n. / adj.", "phonetic": "", "definition": "烈酒", "example": "The bar offers a wide selection of fine wines and spirit liquors.", "collocation": "liquor store / hard liquor", "category": "日常生活"}, {"word": "beer", "pos": "n.", "phonetic": "", "definition": "啤酒", "example": "They enjoyed a cold glass of draft beer after work.", "collocation": "draft / cold beer", "category": "日常生活"}, {"word": "mug", "pos": "n.", "phonetic": "", "definition": "啤酒杯 / 馬克杯", "example": "He clinked his glass mug with his colleague to toast.", "collocation": "beer mug", "category": "日常生活"}, {"word": "local beer", "pos": "n.", "phonetic": "", "definition": "當地啤酒", "example": "Whenever I travel, I always like to try the local beer.", "collocation": "craft / local beer", "category": "日常生活"}, {"word": "toothpick", "pos": "n.", "phonetic": "", "definition": "牙籤", "example": "He picked up a toothpick from the dispenser on the counter.", "collocation": "box of toothpicks", "category": "日常生活"}, {"word": "chopsticks", "pos": "n. (pl.)", "phonetic": "", "definition": "筷子", "example": "She used chopsticks skillfully to pick up the sushi.", "collocation": "pair of chopsticks", "category": "日常生活"}, {"word": "side dish / snack", "pos": "n.", "phonetic": "", "definition": "小菜 / 點心", "example": "Korean meals are always served with various tasty side dishes.", "collocation": "assorted side dishes", "category": "日常生活"}, {"word": "additional order", "pos": "n.", "phonetic": "", "definition": "加點 (菜)", "example": "We placed an additional order of dumplings because they were great.", "collocation": "place an additional order", "category": "日常生活"}, {"word": "sweets", "pos": "n. (pl.)", "phonetic": "", "definition": "甜食", "example": "She bought some sweet pastries and candies from the bakery.", "collocation": "traditional sweets", "category": "日常生活"}, {"word": "revolving sushi", "pos": "n.", "phonetic": "", "definition": "迴轉壽司", "example": "Children love watching plates go around at the revolving sushi restaurant.", "collocation": "revolving sushi bar", "category": "日常生活"}, {"word": "yakiniku restaurant", "pos": "n.", "phonetic": "", "definition": "燒肉餐廳", "example": "We grilled fresh beef sliced thin at the Japanese yakiniku restaurant.", "collocation": "dine at a yakiniku restaurant", "category": "日常生活"}, {"word": "all-you-can-eat buffet", "pos": "n.", "phonetic": "", "definition": "吃到飽自助餐", "example": "The hotel offers a lavish all-you-can-eat buffet for lunch.", "collocation": "buffet dinner", "category": "日常生活"}, {"word": "diner", "pos": "n.", "phonetic": "", "definition": "小吃店 ; 簡餐店", "example": "We grabbed a quick late-night breakfast at a 24-hour diner.", "collocation": "casual diner", "category": "日常生活"}, {"word": "steamed rice", "pos": "n. (uncount.)", "phonetic": "", "definition": "白飯", "example": "A hot bowl of steamed rice accompanies most traditional dishes.", "collocation": "bowl of steamed rice", "category": "日常生活"}, {"word": "scale", "pos": "n.", "phonetic": "", "definition": "體重計", "example": "He steps on the scale every morning to track his fitness progress.", "collocation": "step on the scale", "category": "日常生活"}, {"word": "chart", "pos": "n.", "phonetic": "", "definition": "圖表", "example": "The medical nurse plotted her body weight change on a chart.", "collocation": "progress / health chart", "category": "日常生活"}, {"word": "clinical thermometer", "pos": "n.", "phonetic": "", "definition": "體溫計", "example": "The nurse used a clinical thermometer to check his body temperature.", "collocation": "use a clinical thermometer", "category": "日常生活"}, {"word": "basal body temperature", "pos": "n.", "phonetic": "", "definition": "基礎體溫", "example": "Tracking basal body temperature helps monitor ovulation cycles.", "collocation": "measure basal body temperature", "category": "日常生活"}, {"word": "belly", "pos": "n.", "phonetic": "", "definition": "腹部", "example": "Doing core exercises helps tone and strengthen the belly.", "collocation": "belly fat / flat belly", "category": "日常生活"}, {"word": "waist", "pos": "n.", "phonetic": "", "definition": "腰部", "example": "She measured her waist to see if her workout was effective.", "collocation": "waist line / measurement", "category": "日常生活"}, {"word": "diet", "pos": "n. / v.", "phonetic": "", "definition": "飲食 / 節食", "example": "A balanced diet combined with regular exercise keeps you healthy.", "collocation": "balanced diet / go on a diet", "category": "日常生活"}, {"word": "calorie", "pos": "n.", "phonetic": "", "definition": "卡路里", "example": "Checking the calorie count on food labels helps control weight.", "collocation": "low calorie / count calories", "category": "日常生活"}, {"word": "weight rebound", "pos": "n.", "phonetic": "", "definition": "體重回升", "example": "Maintaining healthy eating habits prevents weight rebound after dieting.", "collocation": "prevent weight rebound", "category": "日常生活"}, {"word": "metabolism", "pos": "n.", "phonetic": "", "definition": "新陳代謝", "example": "Drinking water and staying active helps boost your metabolism.", "collocation": "fast / boost metabolism", "category": "日常生活"}, {"word": "basal metabolism", "pos": "n.", "phonetic": "", "definition": "基礎代謝", "example": "Basal metabolism accounts for most of the daily energy burned.", "collocation": "basal metabolic rate (BMR)", "category": "日常生活"}, {"word": "medical checkup", "pos": "n.", "phonetic": "", "definition": "健康檢查", "example": "It is recommended to schedule an annual medical checkup.", "collocation": "annual medical checkup", "category": "日常生活"}, {"word": "carbohydrate", "pos": "n.", "phonetic": "", "definition": "碳水化合物", "example": "Complex carbohydrates provide steady, long-lasting energy.", "collocation": "low carbohydrate diet", "category": "日常生活"}, {"word": "protein", "pos": "n.", "phonetic": "", "definition": "蛋白質", "example": "Chicken breast and eggs are excellent sources of lean protein.", "collocation": "high protein foods", "category": "日常生活"}, {"word": "fat", "pos": "n. / adj.", "phonetic": "", "definition": "脂肪", "example": "Limiting saturated fat intake reduces cardiovascular risks.", "collocation": "body fat / low fat", "category": "日常生活"}, {"word": "vitamine", "pos": "n.", "phonetic": "", "definition": "維他命", "example": "Citrus fruits like oranges are rich in Vitamin C.", "collocation": "multivitamin / essential vitamine", "category": "日常生活"}, {"word": "mineral", "pos": "n.", "phonetic": "", "definition": "礦物質", "example": "Calcium and potassium are vital minerals for bone health.", "collocation": "essential minerals", "category": "日常生活"}, {"word": "dietary fiber", "pos": "n.", "phonetic": "", "definition": "膳食纖維", "example": "Whole grains and vegetables contain plenty of dietary fiber.", "collocation": "rich in dietary fiber", "category": "日常生活"}, {"word": "unbalanced diet", "pos": "n.", "phonetic": "", "definition": "飲食不均衡", "example": "An unbalanced diet can lead to nutritional deficiencies.", "collocation": "suffer from an unbalanced diet", "category": "日常生活"}, {"word": "medicine", "pos": "n. (uncount.)", "phonetic": "", "definition": "（廣義的）藥", "example": "Take this medicine twice daily after meals as prescribed.", "collocation": "take prescription medicine", "category": "日常生活"}, {"word": "drug", "pos": "n.", "phonetic": "", "definition": "藥物 ; 毒品", "example": "The doctor prescribed a prescription drug to lower his blood pressure.", "collocation": "prescription drug", "category": "日常生活"}, {"word": "tablet", "pos": "n.", "phonetic": "", "definition": "藥片", "example": "Take one tablet with a glass of water every eight hours.", "collocation": "swallow a tablet", "category": "日常生活"}, {"word": "sit-up", "pos": "n.", "phonetic": "", "definition": "仰臥起坐", "example": "Doing sit-ups daily helps strengthen your abdominal muscles.", "collocation": "do sit-ups", "category": "日常生活"}, {"word": "push-up", "pos": "n.", "phonetic": "", "definition": "伏地挺身", "example": "Push-ups build strength in your chest, arms, and shoulders.", "collocation": "do push-ups", "category": "日常生活"}, {"word": "aerobics", "pos": "n.", "phonetic": "", "definition": "有氧舞蹈", "example": "She attends an aerobics class at the local gym twice a week.", "collocation": "do aerobics", "category": "日常生活"}, {"word": "warm-up", "pos": "n.", "phonetic": "", "definition": "暖身 ; 熱身", "example": "A light warm-up helps prevent muscle injuries during workouts.", "collocation": "do warm-up exercises", "category": "日常生活"}, {"word": "gymnastic exercise", "pos": "n.", "phonetic": "", "definition": "體操", "example": "Gymnastic exercises improve overall flexibility and balance.", "collocation": "gymnastic training", "category": "日常生活"}, {"word": "bathtub", "pos": "n.", "phonetic": "", "definition": "浴缸", "example": "She filled the bathtub with warm water for a relaxing bath.", "collocation": "soak in a bathtub", "category": "日常生活"}, {"word": "bubble bath", "pos": "n.", "phonetic": "", "definition": "泡泡浴", "example": "Children love taking a bubble bath with colorful bath toys.", "collocation": "take a bubble bath", "category": "日常生活"}, {"word": "facial mask", "pos": "n.", "phonetic": "", "definition": "面膜", "example": "She put on a hydrating facial mask before going to sleep.", "collocation": "apply / put on a facial mask", "category": "日常生活"}, {"word": "shampoo", "pos": "n.", "phonetic": "", "definition": "洗髮精", "example": "Use a gentle shampoo to keep your hair clean and healthy.", "collocation": "hair shampoo", "category": "日常生活"}, {"word": "conditioner", "pos": "n.", "phonetic": "", "definition": "潤絲精 / 護髮素", "example": "Apply conditioner to the ends of your hair after washing.", "collocation": "hair conditioner", "category": "日常生活"}, {"word": "bath powder", "pos": "n.", "phonetic": "", "definition": "泡澡粉", "example": "She added scented bath powder to the tub for relaxation.", "collocation": "soaking bath powder", "category": "日常生活"}, {"word": "shower gel", "pos": "n.", "phonetic": "", "definition": "沐浴乳", "example": "A refreshing shower gel leaves your skin clean and fragrant.", "collocation": "scented shower gel", "category": "日常生活"}, {"word": "air conditioner", "pos": "n.", "phonetic": "", "definition": "冷氣 ; 空調", "example": "Turn on the air conditioner to keep the room cool in summer.", "collocation": "turn on the air conditioner", "category": "日常生活"}, {"word": "timer", "pos": "n.", "phonetic": "", "definition": "定時器", "example": "She set the air conditioner timer to turn off automatically in two hours.", "collocation": "set a sleep timer", "category": "日常生活"}, {"word": "nightgown", "pos": "n.", "phonetic": "", "definition": "睡袍 / 女性睡衣", "example": "She wore a cozy silk nightgown on cold winter evenings.", "collocation": "silk nightgown", "category": "日常生活"}, {"word": "nightwear", "pos": "n. (uncount.)", "phonetic": "", "definition": "睡衣", "example": "Loose-fitting cotton nightwear ensures a restful night of sleep.", "collocation": "comfortable nightwear", "category": "日常生活"}, {"word": "massage", "pos": "n. / v.", "phonetic": "", "definition": "按摩", "example": "A gentle facial massage helps improve skin circulation.", "collocation": "body / facial massage", "category": "日常生活"}, {"word": "blemish", "pos": "n.", "phonetic": "", "definition": "（臉上的）斑點 / 瑕疵", "example": "Concealer helps cover up minor facial blemishes.", "collocation": "skin blemish", "category": "日常生活"}, {"word": "wrinkle", "pos": "n.", "phonetic": "", "definition": "皺紋", "example": "Using moisturizer helps prevent fine lines and wrinkles.", "collocation": "anti-wrinkle cream", "category": "日常生活"}, {"word": "pimple / spot", "pos": "n.", "phonetic": "", "definition": "痘痘 / 痣", "example": "She applied spot treatment cream on the pimple.", "collocation": "pimple cream / spot treatment", "category": "日常生活"}, {"word": "armpit", "pos": "n.", "phonetic": "", "definition": "腋下", "example": "Shaving under the armpit is part of her grooming routine.", "collocation": "under the armpit", "category": "日常生活"}, {"word": "tweezer", "pos": "n.", "phonetic": "", "definition": "鑷子 (拔毛的)", "example": "She used a pair of tweezers to shape her eyebrows.", "collocation": "pair of tweezers", "category": "日常生活"}, {"word": "hair-removing cream", "pos": "n.", "phonetic": "", "definition": "除毛膏", "example": "Follow product directions carefully when applying hair-removing cream.", "collocation": "apply hair-removing cream", "category": "日常生活"}, {"word": "makeup remover", "pos": "n.", "phonetic": "", "definition": "卸妝水 / 卸妝品", "example": "Always clean off cosmetic makeup with makeup remover at night.", "collocation": "gentle makeup remover", "category": "日常生活"}, {"word": "cleansing cream", "pos": "n.", "phonetic": "", "definition": "潔面霜", "example": "She washed her face thoroughly with a cleansing cream.", "collocation": "facial cleansing cream", "category": "日常生活"}, {"word": "nourishing cream", "pos": "n.", "phonetic": "", "definition": "滋養霜", "example": "Nourishing cream keeps the skin hydrated throughout the night.", "collocation": "apply nourishing cream", "category": "日常生活"}, {"word": "nail polish", "pos": "n.", "phonetic": "", "definition": "指甲油", "example": "She painted her nails with a vibrant red nail polish.", "collocation": "apply nail polish", "category": "日常生活"}, {"word": "nightcap", "pos": "n.", "phonetic": "", "definition": "睡前酒", "example": "He enjoyed a small glass of brandy as a relaxing nightcap.", "collocation": "have a nightcap", "category": "日常生活"}, {"word": "home security system", "pos": "n.", "phonetic": "", "definition": "住家保全系統", "example": "Set the home security system before going to bed for safety.", "collocation": "arm / install home security system", "category": "日常生活"}];

/* Built-in "Zodiac Signs" vocabulary set (16 words) */
const BUILTIN_ZODIAC_CARDS = [{"word": "star sign", "pos": "n.", "phonetic": "", "definition": "星座（口語用法）", "example": "What's your star sign? I'm a Leo.", "collocation": "ask someone's star sign", "category": "星座"}, {"word": "zodiac sign", "pos": "n.", "phonetic": "", "definition": "星座（正式用法）", "example": "Her zodiac sign is Gemini, so she loves to talk.", "collocation": "determine a zodiac sign", "category": "星座"}, {"word": "constellation", "pos": "n.", "phonetic": "", "definition": "（天文學上的）星座", "example": "Orion is one of the most recognizable constellations in the night sky.", "collocation": "observe a constellation", "category": "星座"}, {"word": "horoscope", "pos": "n.", "phonetic": "", "definition": "星座運勢", "example": "She checks her horoscope every morning before leaving home.", "collocation": "read/check a horoscope", "category": "星座"}, {"word": "Aries", "pos": "n.", "phonetic": "", "definition": "牡羊座", "example": "As an Aries, he is confident and always ready to take the lead.", "collocation": "a typical Aries", "category": "星座"}, {"word": "Taurus", "pos": "n.", "phonetic": "", "definition": "金牛座", "example": "Taurus people are known for being reliable and a little stubborn.", "collocation": "a stubborn Taurus", "category": "星座"}, {"word": "Gemini", "pos": "n.", "phonetic": "", "definition": "雙子座", "example": "My sister is a Gemini, so she gets bored easily.", "collocation": "a curious Gemini", "category": "星座"}, {"word": "Cancer", "pos": "n.", "phonetic": "", "definition": "巨蟹座", "example": "People born under Cancer tend to be caring and family-oriented.", "collocation": "a caring Cancer", "category": "星座"}, {"word": "Leo", "pos": "n.", "phonetic": "", "definition": "獅子座", "example": "He is a proud Leo who loves being the center of attention.", "collocation": "a confident Leo", "category": "星座"}, {"word": "Virgo", "pos": "n.", "phonetic": "", "definition": "處女座", "example": "As a Virgo, she pays close attention to every little detail.", "collocation": "a detail-oriented Virgo", "category": "星座"}, {"word": "Libra", "pos": "n.", "phonetic": "", "definition": "天秤座", "example": "Libras are known for trying to keep things fair and balanced.", "collocation": "a diplomatic Libra", "category": "星座"}, {"word": "Scorpio", "pos": "n.", "phonetic": "", "definition": "天蠍座", "example": "A true Scorpio, he is intense and fiercely loyal to his friends.", "collocation": "a mysterious Scorpio", "category": "星座"}, {"word": "Sagittarius", "pos": "n.", "phonetic": "", "definition": "射手座", "example": "Sagittarius people usually love traveling and exploring new places.", "collocation": "an adventurous Sagittarius", "category": "星座"}, {"word": "Capricorn", "pos": "n.", "phonetic": "", "definition": "摩羯座", "example": "Being a Capricorn, she is disciplined and very career-focused.", "collocation": "a hardworking Capricorn", "category": "星座"}, {"word": "Aquarius", "pos": "n.", "phonetic": "", "definition": "水瓶座", "example": "Aquarius people often have unique ideas and value their independence.", "collocation": "an independent Aquarius", "category": "星座"}, {"word": "Pisces", "pos": "n.", "phonetic": "", "definition": "雙魚座", "example": "As a Pisces, he is dreamy, sensitive, and very imaginative.", "collocation": "a dreamy Pisces", "category": "星座"}];

/* Built-in "Reading Article" vocabulary set (77 words, with synonyms) */
const BUILTIN_READING_WORDS = [{"word": "cuisine", "pos": "n.", "phonetic": "", "definition": "菜餚；料理", "example": "Italian cuisine is famous all over the world.", "collocation": "haute cuisine（高級料理）", "synonym": "culinary art, food, dish", "category": "閱讀文章單字"}, {"word": "contemporary", "pos": "adj.", "phonetic": "", "definition": "當代的", "example": "She is a big fan of contemporary architecture.", "collocation": "contemporary art（當代藝術）", "synonym": "modern, current, present-day", "category": "閱讀文章單字"}, {"word": "tasty", "pos": "adj.", "phonetic": "", "definition": "好吃的", "example": "This homemade pie is extremely tasty.", "collocation": "tasty meal（美味的一餐）", "synonym": "delicious, yummy, appetizing", "category": "閱讀文章單字"}, {"word": "ingredient", "pos": "n.", "phonetic": "", "definition": "原料；成分", "example": "Fresh vegetables are essential ingredients for a healthy salad.", "collocation": "key ingredient（關鍵成分 / 主要原料）", "synonym": "element, component, constituent", "category": "閱讀文章單字"}, {"word": "recipe", "pos": "n.", "phonetic": "", "definition": "烹飪法；食譜", "example": "My grandmother gave me her secret cookie recipe.", "collocation": "secret recipe（獨家秘方 / 秘密食譜）", "synonym": "formula, instructions", "category": "閱讀文章單字"}, {"word": "depend on", "pos": "v. phr.", "phonetic": "", "definition": "依靠；依賴", "example": "Children depend on their parents for care and guidance.", "collocation": "depend on somebody（依靠 / 依賴某人）", "synonym": "rely on, count on, lean on", "category": "閱讀文章單字"}, {"word": "fiery", "pos": "adj.", "phonetic": "", "definition": "火一般的", "example": "The fiery sunset painted the sky in shades of red and orange.", "collocation": "fiery passion / fiery temper（熾熱的激情 / 火爆的脾氣）", "synonym": "blazing, burning, passionate", "category": "閱讀文章單字"}, {"word": "digestion", "pos": "n.", "phonetic": "", "definition": "消化；消化作用", "example": "Walking slowly after meals can improve your digestion.", "collocation": "aid digestion（有助於消化）", "synonym": "ingestion, assimilation", "category": "閱讀文章單字"}, {"word": "slum", "pos": "n.", "phonetic": "", "definition": "貧民窟", "example": "The government plans to redevelop the urban slum.", "collocation": "slum area（貧民窟地區）", "synonym": "shantytown, ghetto", "category": "閱讀文章單字"}, {"word": "release", "pos": "v.", "phonetic": "", "definition": "發行；發表", "example": "The band will release their new album next month.", "collocation": "release a movie / album（上映電影 / 發行專輯）", "synonym": "issue, launch, publish", "category": "閱讀文章單字"}, {"word": "depict", "pos": "v.", "phonetic": "", "definition": "描繪；描寫", "example": "The novel vividly depicts life in early 20th century Europe.", "collocation": "depict a scene（描繪景象 / 場景）", "synonym": "portray, describe, illustrate", "category": "閱讀文章單字"}, {"word": "disdain", "pos": "n.", "phonetic": "", "definition": "輕蔑；鄙視", "example": "He treated the rude suggestion with silent disdain.", "collocation": "look with disdain（帶著輕蔑的眼神看）", "synonym": "scorn, contempt, disrespect", "category": "閱讀文章單字"}, {"word": "demand", "pos": "v.", "phonetic": "", "definition": "要求；請求", "example": "The workers demand higher wages and better conditions.", "collocation": "demand an answer（要求給出答覆）", "synonym": "require, request, insist on", "category": "閱讀文章單字"}, {"word": "explanation", "pos": "n.", "phonetic": "", "definition": "說明；解釋", "example": "Please provide a clear explanation for the sudden delay.", "collocation": "detailed explanation（詳細的解釋說明）", "synonym": "clarification, account, rationale", "category": "閱讀文章單字"}, {"word": "turbulent", "pos": "adj.", "phonetic": "", "definition": "騷動的", "example": "The country went through a turbulent period of political change.", "collocation": "turbulent era / times（動盪不安的時代）", "synonym": "chaotic, stormy, tumultuous", "category": "閱讀文章單字"}, {"word": "orphan", "pos": "n.", "phonetic": "", "definition": "孤兒", "example": "The charity provides housing and education for orphans.", "collocation": "orphan child（孤兒幼童）", "synonym": "foundling, parentless child", "category": "閱讀文章單字"}, {"word": "beggar", "pos": "n.", "phonetic": "", "definition": "乞丐；叫化子", "example": "The kind stranger gave some spare change to the beggar.", "collocation": "street beggar（街頭乞丐）", "synonym": "pauper, mendicant", "category": "閱讀文章單字"}, {"word": "sacrifice", "pos": "v.", "phonetic": "", "definition": "犧牲", "example": "Parents often sacrifice their leisure time for their children.", "collocation": "sacrifice time / personal interests（犧牲時間 / 個人利益）", "synonym": "surrender, give up, forgo", "category": "閱讀文章單字"}, {"word": "observe", "pos": "v.", "phonetic": "", "definition": "遵守；奉行", "example": "All visitors must observe the security guidelines.", "collocation": "observe rules / traditions（遵守規定 / 遵循傳統）", "synonym": "follow, obey, comply with", "category": "閱讀文章單字"}, {"word": "traditionally", "pos": "adv.", "phonetic": "", "definition": "傳說上；傳統上；習慣上", "example": "Red envelopes are traditionally given during Lunar New Year.", "collocation": "traditionally made（傳統手工製作的）", "synonym": "customarily, conventionally", "category": "閱讀文章單字"}, {"word": "lunar", "pos": "adj.", "phonetic": "", "definition": "陰曆的", "example": "The Mid-Autumn Festival follows the lunar calendar.", "collocation": "lunar calendar / eclipse（農曆（陰曆） / 月食）", "synonym": "moon-related, monthly", "category": "閱讀文章單字"}, {"word": "origin", "pos": "n.", "phonetic": "", "definition": "起源；由來；起因", "example": "Scientists are researching the origin of the universe.", "collocation": "point of origin（發源地 / 起始點）", "synonym": "source, root, beginning", "category": "閱讀文章單字"}, {"word": "immortality", "pos": "n.", "phonetic": "", "definition": "不死；不朽；不滅", "example": "Ancient emperors often sought potions for immortality.", "collocation": "seek immortality（追求長生不老 / 不朽）", "synonym": "eternity, everlasting life", "category": "閱讀文章單字"}, {"word": "dynasty", "pos": "n.", "phonetic": "", "definition": "王朝；朝代", "example": "The Han Dynasty left a deep cultural legacy in China.", "collocation": "Tang Dynasty（唐朝）", "synonym": "reign, empire, era", "category": "閱讀文章單字"}, {"word": "rebel", "pos": "v.", "phonetic": "", "definition": "造反；反叛；反抗", "example": "The people decided to rebel against the tyrannical rule.", "collocation": "rebel against authority（反抗權威 / 造反）", "synonym": "revolt, resist, defy", "category": "閱讀文章單字"}, {"word": "crust", "pos": "n.", "phonetic": "", "definition": "外皮；外殼", "example": "The earth's crust consists of several tectonic plates.", "collocation": "earth's crust / pie crust（地殼 / 派皮）", "synonym": "outer layer, shell, surface", "category": "閱讀文章單字"}, {"word": "promote", "pos": "v.", "phonetic": "", "definition": "宣傳；推銷", "example": "The company launched a campaign to promote healthy living.", "collocation": "promote a product（推銷 / 宣傳產品）", "synonym": "advertise, publicize, advocate", "category": "閱讀文章單字"}, {"word": "survey", "pos": "n.", "phonetic": "", "definition": "調查報告", "example": "The recent survey shows high customer satisfaction.", "collocation": "conduct a survey（進行民意調查報告）", "synonym": "poll, study, investigation", "category": "閱讀文章單字"}, {"word": "medication", "pos": "n.", "phonetic": "", "definition": "藥物治療", "example": "He takes daily medication to manage his blood pressure.", "collocation": "take medication（服藥 / 接受藥物治療）", "synonym": "medicine, remedy, drug", "category": "閱讀文章單字"}, {"word": "virus", "pos": "n.", "phonetic": "", "definition": "病毒；濾過性病毒", "example": "Wash your hands frequently to prevent virus infections.", "collocation": "computer / flu virus（電腦病毒 / 流感病毒）", "synonym": "pathogen, bug", "category": "閱讀文章單字"}, {"word": "bacteria", "pos": "n. (pl.)", "phonetic": "", "definition": "細菌 (bacterium的複數)", "example": "Proper cooking temperatures kill harmful bacteria in food.", "collocation": "harmful bacteria（有害細菌）", "synonym": "microbes, germs", "category": "閱讀文章單字"}, {"word": "absolutely", "pos": "adv.", "phonetic": "", "definition": "絕對地；完全地", "example": "You are absolutely right about this matter.", "collocation": "absolutely necessary（絕對必要的）", "synonym": "completely, totally, entirely", "category": "閱讀文章單字"}, {"word": "dose", "pos": "n.", "phonetic": "", "definition": "（藥物等的）一劑；一服", "example": "Take a measured dose of syrup before going to bed.", "collocation": "daily dose（每日劑量）", "synonym": "amount, dosage, portion", "category": "閱讀文章單字"}, {"word": "nicotine", "pos": "n.", "phonetic": "", "definition": "菸鹼；尼古丁", "example": "Nicotine is an addictive substance found in cigarettes.", "collocation": "nicotine addiction（尼古丁成癮）", "synonym": "tobacco extract", "category": "閱讀文章單字"}, {"word": "outlook", "pos": "n.", "phonetic": "", "definition": "展望；前景", "example": "The economic outlook for the next quarter remains positive.", "collocation": "economic outlook（經濟前景 / 展望）", "synonym": "prospect, forecast, viewpoint", "category": "閱讀文章單字"}, {"word": "negative", "pos": "adj.", "phonetic": "", "definition": "否定地；消極地", "example": "Try to avoid harboring negative thoughts.", "collocation": "negative attitude（消極悲觀的態度）", "synonym": "pessimistic, adverse, unfavorable", "category": "閱讀文章單字"}, {"word": "senior citizen", "pos": "n.", "phonetic": "", "definition": "老人（65歲以上）；退休老人", "example": "Senior citizens enjoy free access to public museums.", "collocation": "senior citizen discount（銀髮族 / 老人優惠）", "synonym": "elderly person, retiree", "category": "閱讀文章單字"}, {"word": "decade", "pos": "n.", "phonetic": "", "definition": "十年", "example": "The city has undergone dramatic changes in the past decade.", "collocation": "over a decade（長達十年以上）", "synonym": "ten years, decennium", "category": "閱讀文章單字"}, {"word": "devastate", "pos": "v.", "phonetic": "", "definition": "使荒蕪；破壞", "example": "The powerful earthquake threatened to devastate the town.", "collocation": "devastate the region（毀壞 / 重創該地區）", "synonym": "destroy, ruin, wreck", "category": "閱讀文章單字"}, {"word": "grieve", "pos": "v.", "phonetic": "", "definition": "使悲傷；使苦惱", "example": "They gathered to grieve the loss of a dear friend.", "collocation": "grieve for a loss（為損失或悼念而悲傷）", "synonym": "mourn, sorrow, lament", "category": "閱讀文章單字"}, {"word": "mud-brick", "pos": "n.", "phonetic": "", "definition": "泥磚", "example": "Traditional homes in the desert were built from mud-brick.", "collocation": "mud-brick wall（泥磚牆）", "synonym": "adobe, sun-dried brick", "category": "閱讀文章單字"}, {"word": "monument", "pos": "n.", "phonetic": "", "definition": "紀念碑", "example": "The marble monument was built to honor heroic fallen soldiers.", "collocation": "historic monument（歷史紀念碑 / 遺跡）", "synonym": "memorial, shrine", "category": "閱讀文章單字"}, {"word": "ancient", "pos": "adj.", "phonetic": "", "definition": "古代的；古老的", "example": "They discovered ancient pottery during the archaeological dig.", "collocation": "ancient civilization（古代文明）", "synonym": "antique, historic, age-old", "category": "閱讀文章單字"}, {"word": "crumble", "pos": "v.", "phonetic": "", "definition": "粉碎；摧毀", "example": "The old castle walls began to crumble over time.", "collocation": "crumble into dust（崩解粉碎成灰燼）", "synonym": "disintegrate, fall apart, decay", "category": "閱讀文章單字"}, {"word": "disaster", "pos": "n.", "phonetic": "", "definition": "災害；災難", "example": "Emergency response teams reacted swiftly to the natural disaster.", "collocation": "natural disaster（天然災害）", "synonym": "catastrophe, calamity, tragedy", "category": "閱讀文章單字"}, {"word": "dwelling", "pos": "n.", "phonetic": "", "definition": "住處；住宅", "example": "Ancient tribes built temporary dwellings near water sources.", "collocation": "temporary dwelling（臨時住所 / 住宅）", "synonym": "residence, abode, home", "category": "閱讀文章單字"}, {"word": "mosque", "pos": "n.", "phonetic": "", "definition": "清真寺；回教堂", "example": "The grand mosque features breathtaking marble architecture.", "collocation": "visit a mosque（參觀清真寺）", "synonym": "Islamic place of worship", "category": "閱讀文章單字"}, {"word": "blossom", "pos": "v.", "phonetic": "", "definition": "開花；生長茂盛", "example": "Cherry trees blossom beautifully every springtime.", "collocation": "blossom in spring（在春季綻放開花）", "synonym": "bloom, flower, flourish", "category": "閱讀文章單字"}, {"word": "capital", "pos": "n.", "phonetic": "", "definition": "首都；首府", "example": "Paris is the cultural and political capital of France.", "collocation": "capital city（首都城市）", "synonym": "seat of government, primary city", "category": "閱讀文章單字"}, {"word": "tragedy", "pos": "n.", "phonetic": "", "definition": "悲劇；災難", "example": "The sudden loss was a terrible tragedy for the community.", "collocation": "family tragedy（家庭悲劇）", "synonym": "disaster, misfortune, calamity", "category": "閱讀文章單字"}, {"word": "provincial", "pos": "adj.", "phonetic": "", "definition": "鄉氣的；地方性", "example": "He preferred the calm life of a small provincial town.", "collocation": "provincial town（外省小鎮 / 地方城鎮）", "synonym": "regional, local, rural", "category": "閱讀文章單字"}, {"word": "spiritual", "pos": "adj.", "phonetic": "", "definition": "精神（上）的；心靈的", "example": "Meditation promotes inner peace and spiritual well-being.", "collocation": "spiritual growth（心靈成長 / 精神層面的提升）", "synonym": "inner, mental, sacred", "category": "閱讀文章單字"}, {"word": "brew", "pos": "v.", "phonetic": "", "definition": "泡（茶）；釀造", "example": "She likes to brew fresh herbal tea every morning.", "collocation": "brew coffee / tea（沖泡咖啡 / 茶）", "synonym": "infuse, steep, ferment", "category": "閱讀文章單字"}, {"word": "supplier", "pos": "n.", "phonetic": "", "definition": "供應者；供應商", "example": "They selected a new supplier for organic coffee beans.", "collocation": "leading supplier（主要 / 領先的供應商）", "synonym": "provider, vendor, distributor", "category": "閱讀文章單字"}, {"word": "cultivation", "pos": "n.", "phonetic": "", "definition": "耕作；栽培；培養", "example": "Soil quality plays a vital role in successful crop cultivation.", "collocation": "crop cultivation（農作物栽培耕作）", "synonym": "farming, growing, tillage", "category": "閱讀文章單字"}, {"word": "bud", "pos": "n.", "phonetic": "", "definition": "芽；葉芽", "example": "Tiny flower buds began to appear on the tree branches.", "collocation": "flower bud（花蕾 / 花芽）", "synonym": "sprout, shoot", "category": "閱讀文章單字"}, {"word": "fermentation", "pos": "n.", "phonetic": "", "definition": "發酵", "example": "Fermentation transforms grape juice into fine wine.", "collocation": "fermentation process（發酵過程）", "synonym": "brew, aging, breakdown", "category": "閱讀文章單字"}, {"word": "scent", "pos": "n.", "phonetic": "", "definition": "氣味；香味", "example": "A pleasant scent of fresh roses filled the entire room.", "collocation": "floral scent（花香味）", "synonym": "fragrance, aroma, perfume", "category": "閱讀文章單字"}, {"word": "charcoal", "pos": "n.", "phonetic": "", "definition": "木炭", "example": "They used charcoal to fire up the outdoor barbecue.", "collocation": "charcoal grill（木炭烤肉架）", "synonym": "carbon, coalized wood", "category": "閱讀文章單字"}, {"word": "steam", "pos": "n.", "phonetic": "", "definition": "蒸汽；水蒸汽", "example": "Steam rose gently from the hot bowl of noodle soup.", "collocation": "rising steam（蒸騰冒出的水蒸汽）", "synonym": "vapor, condensation", "category": "閱讀文章單字"}, {"word": "distinct", "pos": "adj.", "phonetic": "", "definition": "有區別的；明顯的", "example": "The spice gives the sauce a distinct and unique flavor.", "collocation": "distinct flavor（獨特的 / 明顯風味）", "synonym": "clear, noticeable, unmistakable", "category": "閱讀文章單字"}, {"word": "stimulating", "pos": "adj.", "phonetic": "", "definition": "有刺激性的", "example": "Coffee provides a stimulating boost early in the morning.", "collocation": "stimulating drink / discussion（提神飲料 / 具啟發性的討論）", "synonym": "invigorating, inspiring, exciting", "category": "閱讀文章單字"}, {"word": "expert", "pos": "n.", "phonetic": "", "definition": "專家", "example": "She consulted an expert before investing in the stock market.", "collocation": "industry expert（業界專家）", "synonym": "specialist, professional, authority", "category": "閱讀文章單字"}, {"word": "particularly", "pos": "adv.", "phonetic": "", "definition": "尤其", "example": "This rule is particularly important for initial beginners.", "collocation": "particularly useful（特別有用 / 尤其適用）", "synonym": "especially, specifically, notably", "category": "閱讀文章單字"}, {"word": "approximately", "pos": "adv.", "phonetic": "", "definition": "大概", "example": "The flight will take approximately two and a half hours.", "collocation": "approximately 10 minutes（大約 10 分鐘）", "synonym": "roughly, about, around", "category": "閱讀文章單字"}, {"word": "civet", "pos": "n.", "phonetic": "", "definition": "麝香貓", "example": "Kopi Luwak is a famous coffee made from civet beans.", "collocation": "civet coffee（麝香貓咖啡（貓屎咖啡））", "synonym": "viverrid, nocturnal mammal", "category": "閱讀文章單字"}, {"word": "droppings", "pos": "n. (pl.)", "phonetic": "", "definition": "鳥獸的糞便", "example": "Hikers spotted animal droppings along the forest trail.", "collocation": "animal droppings（動物糞便）", "synonym": "excrement, dung, waste", "category": "閱讀文章單字"}, {"word": "waste", "pos": "n.", "phonetic": "", "definition": "排洩物；廢物", "example": "Proper treatment of chemical waste protects the local river.", "collocation": "industrial / body waste（工業廢料 / 人體排泄物）", "synonym": "refuse, trash, excrement", "category": "閱讀文章單字"}, {"word": "intact", "pos": "adj.", "phonetic": "", "definition": "完整無缺的", "example": "The package arrived in perfect condition with the seal intact.", "collocation": "remain intact（保持完整無缺）", "synonym": "undamaged, complete, unbroken", "category": "閱讀文章單字"}, {"word": "digestive system", "pos": "n. phr.", "phonetic": "", "definition": "消化系統", "example": "Eating dietary fiber supports a healthy digestive system.", "collocation": "healthy digestive system（健康的消化系統）", "synonym": "gastrointestinal tract", "category": "閱讀文章單字"}, {"word": "flavor", "pos": "n.", "phonetic": "", "definition": "味道", "example": "The soup has a rich savory flavor from slow boiling.", "collocation": "rich flavor（濃郁的風味 / 味道）", "synonym": "taste, savor, aroma", "category": "閱讀文章單字"}, {"word": "unfailing", "pos": "adj.", "phonetic": "", "definition": "可靠的", "example": "We are deeply grateful for her unfailing support.", "collocation": "unfailing support（堅定不移 / 可靠的支持）", "synonym": "reliable, dependable, steadfast", "category": "閱讀文章單字"}, {"word": "instinct", "pos": "n.", "phonetic": "", "definition": "本能；天性", "example": "Mother birds act on natural instinct to protect their young.", "collocation": "natural instinct（自然本能 / 直覺）", "synonym": "intuition, impulse, inclination", "category": "閱讀文章單字"}, {"word": "ripeness", "pos": "n.", "phonetic": "", "definition": "成熟", "example": "Harvest the fruit at peak ripeness for maximum sweetness.", "collocation": "peak ripeness（最佳成熟度）", "synonym": "maturity, readiness", "category": "閱讀文章單字"}, {"word": "undertone", "pos": "n.", "phonetic": "", "definition": "底色", "example": "The wine has a subtle undertone of vanilla and oak.", "collocation": "subtle undertone（隱約的底色 / 潛在基調）", "synonym": "undercurrent, nuance, shade", "category": "閱讀文章單字"}, {"word": "earthy", "pos": "adj.", "phonetic": "", "definition": "土味的", "example": "Beetroot has a distinct earthy taste that many enjoy.", "collocation": "earthy aroma（泥土香氣 / 樸實的風味）", "synonym": "muddy, natural, organic", "category": "閱讀文章單字"}, {"word": "gamey", "pos": "adj.", "phonetic": "", "definition": "帶明顯野味味的", "example": "Wild venison often has a stronger, more gamey flavor than beef.", "collocation": "gamey meat（具強烈野味風味的肉類）", "synonym": "wild, pungent, musky", "category": "閱讀文章單字"}];


/* Shared text-to-speech helpers.
   Browsers speak using whatever voices are installed on the device (the
   Web Speech API). If none of the "better" voices have loaded yet, some
   browsers fall back to a low-quality robotic default. We cache the voice
   list as soon as it's available and try to pick the most natural-sounding
   English voice automatically, while still letting the person override it. */
let cachedVoices = [];
if (typeof window !== "undefined" && window.speechSynthesis) {
  cachedVoices = window.speechSynthesis.getVoices() || [];
  window.speechSynthesis.onvoiceschanged = () => {
    cachedVoices = window.speechSynthesis.getVoices() || [];
  };
}

function getAvailableVoices() {
  if (typeof window === "undefined" || !window.speechSynthesis) return [];
  const fresh = window.speechSynthesis.getVoices();
  if (fresh && fresh.length) cachedVoices = fresh;
  return cachedVoices;
}

function pickBestVoice(voices, preferredURI) {
  if (!voices || voices.length === 0) return null;
  if (preferredURI) {
    const found = voices.find((v) => v.voiceURI === preferredURI);
    if (found) return found;
  }
  const english = voices.filter((v) => v.lang && v.lang.toLowerCase().startsWith("en"));
  const pool = english.length ? english : voices;
  const scored = pool.map((v) => {
    const name = v.name.toLowerCase();
    let score = 0;
    if (name.includes("natural")) score += 5;
    if (name.includes("neural")) score += 5;
    if (name.includes("premium")) score += 4;
    if (name.includes("enhanced")) score += 4;
    if (name.includes("google")) score += 3;
    if (["samantha", "alex", "daniel", "karen", "moira", "ava", "zoe"].some((n) => name.includes(n))) score += 2;
    if (v.lang === "en-US" || v.lang === "en-GB") score += 1;
    if (!v.localService) score += 1; // network-backed voices are often higher quality
    return { v, score };
  });
  scored.sort((a, b) => b.score - a.score);
  return scored[0].v;
}

function speakWord(text, voiceURI) {
  try {
    const u = new SpeechSynthesisUtterance(text);
    u.lang = "en-US";
    u.rate = 0.9;
    const voice = pickBestVoice(getAvailableVoices(), voiceURI);
    if (voice) u.voice = voice;
    window.speechSynthesis.cancel();
    window.speechSynthesis.speak(u);
    return true;
  } catch {
    return false;
  }
}

/* Hint helper for dictation / typing: first letter + underscores for the rest */
function hintDisplay(word) {
  if (!word) return "";
  return word[0].toUpperCase() + " " + Array(word.length - 1).fill("_").join(" ");
}

const ENCOURAGE_MSGS = [
  "答錯了沒關係，多練習幾次就會記住！",
  "很接近了，再試一次一定可以！",
  "錯誤是學習的一部分，繼續加油 💪",
  "沒關係，這題先記起來，下次就會了！",
  "別氣餒，每一次練習都是進步！",
];
const REVEAL_MSGS = [
  "沒關係，先記住這次的答案，多背幾次就會熟練！",
  "看過答案也是一種學習，下次再試試看！",
  "先了解正確答案，下一輪你一定記得住！",
];
function randomFrom(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

/* ---------------------------------------------------------
   Root component
--------------------------------------------------------- */
export default function App() {
  const [cards, setCards] = useState(null);
  const [energy, setEnergy] = useState(null);
  const [articles, setArticles] = useState(null);
  const [categories, setCategories] = useState(null);
  const [voicePref, setVoicePref] = useState(null);
  const [tab, setTab] = useState("cards");
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      let c = await loadJSON(CARDS_KEY, null);
      c = c && c.length ? c : [...SEED_CARDS];
      let cats = await loadJSON(CATEGORIES_KEY, null);
      cats = cats && cats.length ? cats : [...SEED_CATEGORIES];

      const alreadyImported = await loadJSON(BUILTIN_IMPORTED_KEY, false);
      if (!alreadyImported) {
        const existingWords = new Set(c.map((k) => k.word.toLowerCase()));
        const toAdd = BUILTIN_DAILY_LIFE_CARDS.filter((k) => !existingWords.has(k.word.toLowerCase())).map((k) => ({
          ...k,
          id: uid(),
        }));
        c = [...c, ...toAdd];
        if (!cats.includes("日常生活")) cats = [...cats, "日常生活"];
        await saveJSON(BUILTIN_IMPORTED_KEY, true);
      }

      const zodiacImported = await loadJSON(BUILTIN_ZODIAC_IMPORTED_KEY, false);
      if (!zodiacImported) {
        const existingWords2 = new Set(c.map((k) => k.word.toLowerCase()));
        const toAdd2 = BUILTIN_ZODIAC_CARDS.filter((k) => !existingWords2.has(k.word.toLowerCase())).map((k) => ({
          ...k,
          id: uid(),
        }));
        c = [...c, ...toAdd2];
        if (!cats.includes("星座")) cats = [...cats, "星座"];
        await saveJSON(BUILTIN_ZODIAC_IMPORTED_KEY, true);
      }

      const readingImported = await loadJSON(BUILTIN_READING_IMPORTED_KEY, false);
      if (!readingImported) {
        const existingWords3 = new Set(c.map((k) => k.word.toLowerCase()));
        const toAdd3 = BUILTIN_READING_WORDS.filter((k) => !existingWords3.has(k.word.toLowerCase())).map((k) => ({
          ...k,
          id: uid(),
        }));
        c = [...c, ...toAdd3];
        if (!cats.includes("閱讀文章單字")) cats = [...cats, "閱讀文章單字"];
        await saveJSON(BUILTIN_READING_IMPORTED_KEY, true);
      }

      setCards(c);
      setCategories(cats);

      const e = await loadJSON(ENERGY_KEY, null);
      setEnergy(e || { date: todayStr(), points: 0, done: {}, reminderTime: "20:00", restDays: 0 });
      const a = await loadJSON(ARTICLES_KEY, null);
      setArticles(a || []);
      const v = await loadJSON(VOICE_KEY, "");
      setVoicePref(v || "");
    })();
  }, []);

  useEffect(() => {
    if (cards) saveJSON(CARDS_KEY, cards);
  }, [cards]);

  useEffect(() => {
    if (energy) saveJSON(ENERGY_KEY, energy);
  }, [energy]);

  useEffect(() => {
    if (articles) saveJSON(ARTICLES_KEY, articles);
  }, [articles]);

  useEffect(() => {
    if (categories) saveJSON(CATEGORIES_KEY, categories);
  }, [categories]);

  useEffect(() => {
    if (voicePref !== null) saveJSON(VOICE_KEY, voicePref);
  }, [voicePref]);

  // roll over energy at a new day
  useEffect(() => {
    if (!energy) return;
    if (energy.date !== todayStr()) {
      setEnergy((prev) => ({
        ...prev,
        date: todayStr(),
        points: 0,
        done: {},
        restDays: prev.points >= 100 ? prev.restDays + 1 : prev.restDays,
      }));
    }
  }, [energy]);

  const showToast = useCallback((msg, kind = "info") => {
    setToast({ msg, kind, id: uid() });
    setTimeout(() => setToast((t) => (t && t.msg === msg ? null : t)), 2600);
  }, []);

  const completeTask = useCallback(
    (taskId) => {
      setEnergy((prev) => {
        if (prev.done[taskId]) return prev;
        const points = Math.min(100, prev.points + 20);
        showToast(`${TASKS.find((t) => t.id === taskId).label} 完成！+20 能量`, "success");
        return { ...prev, done: { ...prev.done, [taskId]: true }, points };
      });
    },
    [showToast]
  );

  const addCards = useCallback(
    (newCards) => {
      setCards((prev) => [...prev, ...newCards]);
      setCategories((prev) => {
        const extra = newCards
          .map((c) => (c.category || "").trim())
          .filter((cat) => cat && !prev.includes(cat));
        return extra.length ? [...prev, ...Array.from(new Set(extra))] : prev;
      });
      showToast(`已新增 ${newCards.length} 張單字卡`, "success");
    },
    [showToast]
  );

  const deleteCard = useCallback((id) => {
    setCards((prev) => prev.filter((c) => c.id !== id));
  }, []);

  const markCard = useCallback((id, know) => {
    setCards((prev) =>
      prev.map((c) =>
        c.id === id
          ? {
              ...c,
              knownCount: (c.knownCount || 0) + (know ? 1 : 0),
              unknownCount: (c.unknownCount || 0) + (know ? 0 : 1),
            }
          : c
      )
    );
  }, []);

  const moveCard = useCallback((id, category) => {
    setCards((prev) => prev.map((c) => (c.id === id ? { ...c, category } : c)));
  }, []);

  const addCategory = useCallback(
    (name) => {
      const clean = name.trim();
      if (!clean) return;
      setCategories((prev) => {
        if (prev.includes(clean)) {
          showToast("這個分類已經存在了", "error");
          return prev;
        }
        showToast(`已新增分類「${clean}」`, "success");
        return [...prev, clean];
      });
    },
    [showToast]
  );

  const removeCategory = useCallback(
    (name) => {
      if (name === UNCATEGORIZED) return;
      setCategories((prev) => prev.filter((c) => c !== name));
      setCards((prev) => prev.map((c) => (c.category === name ? { ...c, category: UNCATEGORIZED } : c)));
      showToast(`已刪除分類「${name}」，該分類的字卡移至「${UNCATEGORIZED}」`, "success");
    },
    [showToast]
  );

  if (!cards || !energy || !articles || !categories || voicePref === null) {
    return (
      <div style={styles.loadingScreen}>
        <div style={styles.loadingCard}>載入單字卡中…</div>
      </div>
    );
  }

  return (
    <div style={styles.app}>
      <GlobalStyle />
      <Header energy={energy} />
      <ReminderBanner energy={energy} cardsCount={cards.length} />
      <TabBar tab={tab} setTab={setTab} />
      <main style={styles.main}>
        {tab === "cards" && (
          <FlashcardsTab
            cards={cards}
            categories={categories}
            onComplete={() => completeTask("flashcards")}
            done={!!energy.done.flashcards}
            onMark={markCard}
            voiceURI={voicePref}
          />
        )}
        {tab === "quiz" && (
          <QuizTab cards={cards} categories={categories} energy={energy} completeTask={completeTask} showToast={showToast} voiceURI={voicePref} />
        )}
        {tab === "reading" && (
          <ReadingTab
            articles={articles}
            setArticles={setArticles}
            onComplete={() => completeTask("reading")}
            done={!!energy.done.reading}
            showToast={showToast}
          />
        )}
        {tab === "import" && (
          <ImportTab
            onAdd={addCards}
            cards={cards}
            onDelete={deleteCard}
            categories={categories}
            onMove={moveCard}
            onAddCategory={addCategory}
            onRemoveCategory={removeCategory}
          />
        )}
        {tab === "progress" && (
          <ProgressTab energy={energy} setEnergy={setEnergy} voicePref={voicePref} setVoicePref={setVoicePref} />
        )}
      </main>
      {toast && <Toast toast={toast} />}
    </div>
  );
}

/* ---------------------------------------------------------
   Global style / fonts / tokens
--------------------------------------------------------- */
function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
      * { box-sizing: border-box; }
      body, html { margin: 0; padding: 0; }
      button { font-family: inherit; cursor: pointer; }
      input, textarea, select { font-family: inherit; }
      ::selection { background: #D9A44166; }
      @keyframes flipIn { from { transform: rotateY(90deg); opacity: 0.3; } to { transform: rotateY(0deg); opacity: 1; } }
      @keyframes shake { 10%,90% { transform: translateX(-2px);} 20%,80% { transform: translateX(3px);} 30%,50%,70% { transform: translateX(-5px);} 40%,60% { transform: translateX(5px);} }
      @keyframes popIn { from { transform: scale(0.9); opacity: 0; } to { transform: scale(1); opacity: 1; } }
      @keyframes slideUp { from { transform: translateY(16px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
      .shake { animation: shake 0.4s; }
      .popIn { animation: popIn 0.25s ease-out; }
    `}</style>
  );
}

const COLORS = {
  bg: "#EDF1F0",
  paper: "#FAFBF9",
  ink: "#1E2A26",
  inkSoft: "#5B6B65",
  line: "#D7E0DB",
  accent: "#2F6F62",
  accentSoft: "#DCEAE4",
  energy: "#D9A441",
  energySoft: "#F5E7C9",
  good: "#4C8C5B",
  bad: "#C15B5B",
};

const styles = {
  loadingScreen: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: COLORS.bg, fontFamily: "Inter, sans-serif" },
  loadingCard: { padding: "20px 28px", background: COLORS.paper, borderRadius: 12, color: COLORS.inkSoft, boxShadow: "0 1px 2px rgba(0,0,0,0.06)" },
  app: { minHeight: "100vh", background: COLORS.bg, fontFamily: "Inter, sans-serif", color: COLORS.ink, paddingBottom: 40 },
  main: { maxWidth: 760, margin: "0 auto", padding: "0 16px" },
};

/* ---------------------------------------------------------
   Header + energy gauge
--------------------------------------------------------- */
function Header({ energy }) {
  const doneCount = Object.keys(energy.done).length;
  return (
    <div style={{ background: COLORS.ink, color: "#F5F7F5", padding: "22px 16px 26px" }}>
      <div style={{ maxWidth: 760, margin: "0 auto" }}>
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <div>
            <div style={{ fontFamily: "'Source Serif 4', serif", fontSize: 26, fontWeight: 700, letterSpacing: 0.2 }}>
              Vocab Box
            </div>
            <div style={{ fontSize: 13, color: "#B9C4BF", marginTop: 2 }}>你的英文單字卡盒 · 今日 {doneCount}/5 項任務</div>
          </div>
          <EnergyGauge points={energy.points} restDays={energy.restDays} />
        </div>
      </div>
    </div>
  );
}

function EnergyGauge({ points, restDays }) {
  const segments = 5;
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
      <div style={{ display: "flex", gap: 4 }}>
        {Array.from({ length: segments }).map((_, i) => {
          const filled = points >= (i + 1) * 20;
          return (
            <div
              key={i}
              style={{
                width: 20,
                height: 10,
                borderRadius: 3,
                background: filled ? COLORS.energy : "rgba(255,255,255,0.18)",
                transition: "background 0.3s",
              }}
            />
          );
        })}
      </div>
      <div style={{ fontSize: 12, color: "#D9E0DC", display: "flex", alignItems: "center", gap: 6 }}>
        <Flame size={13} color={COLORS.energy} />
        <span style={{ fontWeight: 600, color: "#F5F7F5" }}>{points}</span>/100 能量
        {points >= 100 && <span style={{ color: COLORS.energy, marginLeft: 2 }}>· 今天可以休息了 🎉</span>}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Reminder banner
--------------------------------------------------------- */
function ReminderBanner({ energy, cardsCount }) {
  const [dismissed, setDismissed] = useState(false);
  if (dismissed || energy.points >= 100) return null;
  const now = new Date();
  const [h, m] = (energy.reminderTime || "20:00").split(":").map(Number);
  const pastReminderTime = now.getHours() > h || (now.getHours() === h && now.getMinutes() >= m);
  const doneCount = Object.keys(energy.done).length;
  if (!pastReminderTime || doneCount > 0) return null;

  return (
    <div style={{ maxWidth: 760, margin: "14px auto 0", padding: "0 16px" }}>
      <div
        style={{
          background: COLORS.energySoft,
          border: `1px solid ${COLORS.energy}55`,
          borderRadius: 10,
          padding: "12px 14px",
          display: "flex",
          alignItems: "center",
          gap: 10,
          animation: "slideUp 0.3s ease-out",
        }}
      >
        <Bell size={18} color="#8A6316" />
        <div style={{ fontSize: 13.5, color: "#5C4413", flex: 1 }}>
          已經 {energy.reminderTime} 了，今天還沒開始讀書喔 — {cardsCount} 張單字卡在等你。
        </div>
        <button
          onClick={() => setDismissed(true)}
          style={{ border: "none", background: "transparent", color: "#8A6316", padding: 4 }}
        >
          <X size={16} />
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Tab bar
--------------------------------------------------------- */
function TabBar({ tab, setTab }) {
  const tabs = [
    { id: "cards", label: "字卡", icon: Layers },
    { id: "quiz", label: "測驗", icon: ListChecks },
    { id: "reading", label: "閱讀", icon: BookOpen },
    { id: "import", label: "匯入", icon: Upload },
    { id: "progress", label: "進度", icon: Flame },
  ];
  return (
    <div style={{ maxWidth: 760, margin: "16px auto 0", padding: "0 16px" }}>
      <div style={{ display: "flex", gap: 4, background: "#E2E9E5", padding: 4, borderRadius: 12 }}>
        {tabs.map((t) => {
          const Icon = t.icon;
          const active = tab === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                flex: 1,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                gap: 3,
                padding: "8px 4px",
                borderRadius: 9,
                border: "none",
                background: active ? COLORS.paper : "transparent",
                color: active ? COLORS.accent : COLORS.inkSoft,
                fontWeight: active ? 700 : 500,
                fontSize: 11,
                boxShadow: active ? "0 1px 3px rgba(0,0,0,0.08)" : "none",
                transition: "all 0.15s",
              }}
            >
              <Icon size={15} />
              {t.label}
            </button>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Flashcards tab
--------------------------------------------------------- */
function buildWeightedQueue(cardsList) {
  const slots = [];
  cardsList.forEach((c, i) => {
    const known = c.knownCount || 0;
    const unknown = c.unknownCount || 0;
    // known cards show up less; unknown cards show up more (0.15x to 3x baseline)
    const expected = Math.min(3, Math.max(0.15, 1 + unknown * 0.7 - known * 0.6));
    let repeat = Math.floor(expected);
    if (Math.random() < expected - repeat) repeat += 1;
    for (let r = 0; r < repeat; r++) slots.push(i);
  });
  if (slots.length === 0) {
    cardsList.forEach((_, i) => slots.push(i));
  }
  for (let i = slots.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [slots[i], slots[j]] = [slots[j], slots[i]];
  }
  return slots;
}

function FlashcardsTab({ cards, categories, onComplete, done, onMark, voiceURI }) {
  const [flipped, setFlipped] = useState(false);
  const [filterCat, setFilterCat] = useState("全部");
  const [completedRound, setCompletedRound] = useState(false);

  const filtered = useMemo(
    () => (filterCat === "全部" ? cards : cards.filter((c) => (c.category || UNCATEGORIZED) === filterCat)),
    [cards, filterCat]
  );

  const [session, setSession] = useState(() => ({ queue: buildWeightedQueue(filtered), idx: 0 }));
  const [sessionKey, setSessionKey] = useState(`${filterCat}:${filtered.length}`);

  // Rebuild the queue synchronously during render (not in an effect) whenever the
  // selected category or its card count changes, so we never render a stale queue
  // against a different `filtered` list (which was crashing the view with an
  // out-of-bounds / undefined card).
  let activeSession = session;
  const currentKey = `${filterCat}:${filtered.length}`;
  if (sessionKey !== currentKey) {
    setSessionKey(currentKey);
    activeSession = { queue: buildWeightedQueue(filtered), idx: 0 };
    setSession(activeSession);
    setFlipped(false);
    setCompletedRound(false);
  }

  const catSelector = (
    <select
      value={filterCat}
      onChange={(e) => setFilterCat(e.target.value)}
      style={{ ...inputStyle, fontSize: 12.5, padding: "6px 8px" }}
    >
      <option value="全部">全部類別 ({cards.length})</option>
      {categories.map((c) => (
        <option key={c} value={c}>
          {c} ({cards.filter((k) => (k.category || UNCATEGORIZED) === c).length})
        </option>
      ))}
    </select>
  );

  if (filtered.length === 0) {
    return (
      <div style={{ marginTop: 22 }}>
        <div style={{ marginBottom: 14 }}>{catSelector}</div>
        <EmptyState text={cards.length === 0 ? "還沒有單字卡，先到「匯入」新增幾張吧。" : "這個分類還沒有字卡。"} />
      </div>
    );
  }

  const { queue, idx } = activeSession;
  const card = filtered[queue[idx % queue.length]] || filtered[0];

  const next = () => {
    setFlipped(false);
    setSession((s) => {
      const wrapped = s.idx + 1 >= s.queue.length;
      if (wrapped) setCompletedRound(true);
      return { ...s, idx: (s.idx + 1) % s.queue.length };
    });
  };
  const prev = () => {
    setFlipped(false);
    setSession((s) => ({ ...s, idx: (s.idx - 1 + s.queue.length) % s.queue.length }));
  };
  const shuffle = () => {
    setFlipped(false);
    setCompletedRound(false);
    setSession({ queue: buildWeightedQueue(filtered), idx: 0 });
  };
  const mark = (know) => {
    onMark(card.id, know);
    setFlipped(false);
    setSession((s) => {
      const q = [...s.queue];
      const cardSlot = q[s.idx];
      if (know) {
        // seen as known: drop one future repeat of this card so it appears less
        for (let i = s.idx + 1; i < q.length; i++) {
          if (q[i] === cardSlot) {
            q.splice(i, 1);
            break;
          }
        }
      } else {
        // marked as unfamiliar: insert an extra upcoming repeat so it appears more
        const from = s.idx + 1;
        const range = Math.max(1, q.length - from);
        const insertPos = Math.min(from + Math.floor(Math.random() * range), q.length);
        q.splice(insertPos, 0, cardSlot);
      }
      const wrapped = s.idx + 1 >= q.length;
      if (wrapped) setCompletedRound(true);
      return { queue: q, idx: (s.idx + 1) % q.length };
    });
  };

  return (
    <div style={{ marginTop: 22 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 8, flexWrap: "wrap" }}>
        {catSelector}
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ fontSize: 13, color: COLORS.inkSoft }}>{(idx % queue.length) + 1} / {queue.length}</div>
          <button onClick={shuffle} style={ghostBtn}>
            <Shuffle size={14} /> 隨機排序
          </button>
        </div>
      </div>

      <div
        style={{
          background: COLORS.paper,
          border: `1px solid ${COLORS.line}`,
          borderRadius: 18,
          minHeight: 260,
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          padding: 28,
          position: "relative",
          boxShadow: "0 2px 10px rgba(30,42,38,0.06)",
          animation: "flipIn 0.25s ease-out",
        }}
        key={`${idx}-${flipped}`}
      >
        <div
          style={{
            position: "absolute",
            top: 16,
            left: 20,
            fontSize: 11,
            letterSpacing: 1,
            color: COLORS.inkSoft,
            textTransform: "uppercase",
          }}
        >
          {flipped ? "背面" : "正面"}
        </div>

        <button
          onClick={(e) => {
            e.stopPropagation();
            speakWord(card.word, voiceURI);
          }}
          title="聽發音"
          style={{
            position: "absolute",
            top: 12,
            right: 14,
            border: "none",
            background: COLORS.accentSoft,
            color: COLORS.accent,
            borderRadius: 999,
            width: 32,
            height: 32,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <Volume2 size={15} />
        </button>

        <div onClick={() => setFlipped((f) => !f)} style={{ cursor: "pointer", width: "100%", display: "flex", flexDirection: "column", alignItems: "center" }}>
          {!flipped ? (
            <div style={{ fontFamily: "'Source Serif 4', serif", fontSize: 40, fontWeight: 700, textAlign: "center" }}>
              {card.word}
            </div>
          ) : (
            <div style={{ textAlign: "center", display: "flex", flexDirection: "column", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                {card.pos && (
                  <span
                    style={{
                      fontSize: 11.5,
                      fontWeight: 700,
                      color: COLORS.accent,
                      background: COLORS.accentSoft,
                      borderRadius: 6,
                      padding: "2px 7px",
                    }}
                  >
                    {card.pos}
                  </span>
                )}
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.accent, fontSize: 15 }}>
                  {card.phonetic || "—"}
                </span>
              </div>
              <div style={{ fontSize: 22, fontWeight: 700 }}>{card.definition}</div>
              {card.collocation && (
                <div
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 6,
                    alignSelf: "center",
                    background: COLORS.energySoft,
                    color: "#8A6316",
                    borderRadius: 8,
                    padding: "4px 10px",
                    fontSize: 13,
                    fontWeight: 600,
                  }}
                >
                  🔗 {card.collocation}
                </div>
              )}
              {card.synonym && (
                <div
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 6,
                    alignSelf: "center",
                    background: COLORS.accentSoft,
                    color: COLORS.accent,
                    borderRadius: 8,
                    padding: "4px 10px",
                    fontSize: 13,
                    fontWeight: 600,
                  }}
                >
                  ≈ {card.synonym}
                </div>
              )}
              {card.example && (
                <div style={{ fontSize: 14, color: COLORS.inkSoft, fontStyle: "italic", maxWidth: 420 }}>
                  “{card.example}”
                </div>
              )}
            </div>
          )}
        </div>
        <div style={{ position: "absolute", bottom: 14, fontSize: 11.5, color: COLORS.inkSoft }}>點卡片翻面</div>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14, gap: 10 }}>
        <button onClick={prev} style={navBtn}>
          <ChevronLeft size={18} />
        </button>
        <button onClick={() => setFlipped((f) => !f)} style={{ ...navBtn, flex: 1, gap: 6 }}>
          <RotateCw size={15} /> 翻面
        </button>
        <button onClick={next} style={navBtn}>
          <ChevronRight size={18} />
        </button>
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 10 }}>
        <button
          onClick={() => mark(false)}
          style={{ ...ghostBtn, flex: 1, justifyContent: "center", borderColor: `${COLORS.bad}55`, color: COLORS.bad }}
        >
          🤔 不熟
        </button>
        <button
          onClick={() => mark(true)}
          style={{ ...ghostBtn, flex: 1, justifyContent: "center", borderColor: `${COLORS.good}55`, color: COLORS.good }}
        >
          😀 認識
        </button>
      </div>
      <p style={{ fontSize: 11.5, color: COLORS.inkSoft, textAlign: "center", marginTop: 8 }}>
        按「認識」這張卡之後出現的機率會降低；按「不熟」則會更常出現。
      </p>

      {!done && completedRound && (
        <button onClick={onComplete} style={{ ...primaryBtn, width: "100%", marginTop: 16 }}>
          <Check size={16} /> 完成本次複習 (+20 能量)
        </button>
      )}
      {done && (
        <div style={{ textAlign: "center", marginTop: 16, fontSize: 13, color: COLORS.good, fontWeight: 600 }}>
          ✓ 今天的字卡任務已完成
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Quiz tab (sub-nav: matching / dictation / typing / MC)
--------------------------------------------------------- */
function QuizTab({ cards, categories, energy, completeTask, showToast, voiceURI }) {
  const [mode, setMode] = useState("matching");
  const [filterCat, setFilterCat] = useState("全部");
  const filtered = useMemo(
    () => (filterCat === "全部" ? cards : cards.filter((c) => (c.category || UNCATEGORIZED) === filterCat)),
    [cards, filterCat]
  );
  const sub = [
    { id: "matching", label: "配對", icon: Shuffle, min: 4 },
    { id: "typing", label: "拼字", icon: Keyboard, min: 1 },
    { id: "multiplechoice", label: "選擇題", icon: ListChecks, min: 3 },
  ];
  return (
    <div style={{ marginTop: 22 }}>
      <div style={{ marginBottom: 12 }}>
        <select
          value={filterCat}
          onChange={(e) => setFilterCat(e.target.value)}
          style={{ ...inputStyle, fontSize: 12.5, padding: "6px 8px" }}
        >
          <option value="全部">全部類別 ({cards.length})</option>
          {categories.map((c) => (
            <option key={c} value={c}>
              {c} ({cards.filter((k) => (k.category || UNCATEGORIZED) === c).length})
            </option>
          ))}
        </select>
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        {sub.map((s) => {
          const Icon = s.icon;
          const active = mode === s.id;
          const doneToday = !!energy.done[s.id];
          return (
            <button
              key={s.id}
              onClick={() => setMode(s.id)}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                padding: "7px 12px",
                borderRadius: 999,
                border: `1px solid ${active ? COLORS.accent : COLORS.line}`,
                background: active ? COLORS.accentSoft : COLORS.paper,
                color: active ? COLORS.accent : COLORS.inkSoft,
                fontSize: 13,
                fontWeight: 600,
              }}
            >
              <Icon size={13} /> {s.label}
              {doneToday && <Check size={12} color={COLORS.good} />}
            </button>
          );
        })}
      </div>

      {filtered.length < sub.find((s) => s.id === mode).min ? (
        <EmptyState text={`此測驗至少需要 ${sub.find((s) => s.id === mode).min} 張單字卡，請先新增字卡或切換類別。`} />
      ) : (
        <>
          {mode === "matching" && <MatchingQuiz cards={filtered} onComplete={() => completeTask("matching")} done={!!energy.done.matching} />}
          {mode === "typing" && <TypingQuiz cards={filtered} onComplete={() => completeTask("typing")} done={!!energy.done.typing} showToast={showToast} voiceURI={voiceURI} />}
          {mode === "multiplechoice" && <MultipleChoiceQuiz cards={filtered} onComplete={() => completeTask("multiplechoice")} done={!!energy.done.multiplechoice} showToast={showToast} />}
        </>
      )}
    </div>
  );
}

function shuffleArr(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}
function pickRandom(cards, n) {
  return shuffleArr(cards).slice(0, n);
}

/* ---- Matching quiz ---- */
function MatchingQuiz({ cards, onComplete, done }) {
  const N = Math.min(6, cards.length);
  const [round, setRound] = useState(() => pickRandom(cards, N));
  const [words, setWords] = useState(() => shuffleArr(round));
  const [defs, setDefs] = useState(() => shuffleArr(round));
  const [selWord, setSelWord] = useState(null);
  const [selDef, setSelDef] = useState(null);
  const [matched, setMatched] = useState(new Set());
  const [wrongPair, setWrongPair] = useState(null);

  function reset() {
    const r = pickRandom(cards, N);
    setRound(r);
    setWords(shuffleArr(r));
    setDefs(shuffleArr(r));
    setMatched(new Set());
    setSelWord(null);
    setSelDef(null);
  }

  useEffect(() => {
    if (matched.size === N && N > 0) {
      onComplete();
    }
  }, [matched]);

  function trySelect(type, card) {
    if (matched.has(card.id)) return;
    if (type === "word") setSelWord(card);
    else setSelDef(card);
  }

  useEffect(() => {
    if (selWord && selDef) {
      if (selWord.id === selDef.id) {
        setMatched((m) => new Set(m).add(selWord.id));
        setSelWord(null);
        setSelDef(null);
      } else {
        setWrongPair([selWord.id, selDef.id]);
        setTimeout(() => {
          setWrongPair(null);
          setSelWord(null);
          setSelDef(null);
        }, 500);
      }
    }
  }, [selWord, selDef]);

  return (
    <div>
      <p style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 10 }}>點選一個單字，再點選對應的解釋。配對 {matched.size}/{N}</p>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {words.map((c) => (
            <MatchCell
              key={c.id}
              label={c.word}
              serif
              matched={matched.has(c.id)}
              selected={selWord?.id === c.id}
              wrong={wrongPair?.includes(c.id)}
              onClick={() => trySelect("word", c)}
            />
          ))}
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {defs.map((c) => (
            <MatchCell
              key={c.id}
              label={c.definition}
              matched={matched.has(c.id)}
              selected={selDef?.id === c.id}
              wrong={wrongPair?.includes(c.id)}
              onClick={() => trySelect("def", c)}
            />
          ))}
        </div>
      </div>
      {matched.size === N && (
        <div style={{ textAlign: "center", marginTop: 16 }}>
          {!done ? (
            <div style={{ color: COLORS.good, fontWeight: 700, marginBottom: 10 }}>全部配對成功！+20 能量</div>
          ) : (
            <div style={{ color: COLORS.good, fontWeight: 600, marginBottom: 10 }}>✓ 今天的配對任務已完成</div>
          )}
          <button onClick={reset} style={ghostBtn}>
            <RotateCw size={14} /> 再玩一輪
          </button>
        </div>
      )}
    </div>
  );
}
function MatchCell({ label, serif, matched, selected, wrong, onClick }) {
  return (
    <button
      onClick={onClick}
      disabled={matched}
      className={wrong ? "shake" : ""}
      style={{
        textAlign: "left",
        padding: "10px 12px",
        borderRadius: 10,
        border: `1.5px solid ${matched ? COLORS.good : selected ? COLORS.accent : COLORS.line}`,
        background: matched ? "#E7F2E9" : selected ? COLORS.accentSoft : COLORS.paper,
        color: matched ? COLORS.good : COLORS.ink,
        fontFamily: serif ? "'Source Serif 4', serif" : "inherit",
        fontWeight: serif ? 700 : 500,
        fontSize: 13.5,
        opacity: matched ? 0.6 : 1,
        transition: "all 0.15s",
      }}
    >
      {label}
    </button>
  );
}

function WrongAnswerPanel({ word, message, onNext }) {
  return (
    <div
      className="popIn"
      style={{
        background: "#F6E4E4",
        border: `1px solid ${COLORS.bad}55`,
        borderRadius: 12,
        padding: "14px 16px",
        marginTop: 14,
        textAlign: "center",
      }}
    >
      <div style={{ color: COLORS.bad, fontWeight: 700, fontSize: 14, marginBottom: 4 }}>你做錯了，別擔心！</div>
      <div style={{ fontSize: 13, color: "#7A3D3D", marginBottom: 8 }}>{message}</div>
      <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 10 }}>
        正確答案是：<strong style={{ color: COLORS.ink, fontFamily: "'Source Serif 4', serif", fontSize: 16 }}>{word}</strong>
      </div>
      <button onClick={onNext} style={primaryBtn}>
        下一題
      </button>
    </div>
  );
}

/* ---- Typing quiz (definition -> spell the word) ---- */
function TypingQuiz({ cards, onComplete, done, showToast, voiceURI }) {
  const N = Math.min(8, cards.length);
  const [queue, setQueue] = useState(() => pickRandom(cards, N));
  const [i, setI] = useState(0);
  const [input, setInput] = useState("");
  const [feedback, setFeedback] = useState(null);
  const [correctCount, setCorrectCount] = useState(0);
  const [showHint, setShowHint] = useState(false);
  const [revealed, setRevealed] = useState(false);
  const [stopped, setStopped] = useState(false);
  const [encourageMsg, setEncourageMsg] = useState("");

  useEffect(() => {
    setShowHint(false);
    setRevealed(false);
    setStopped(false);
    setFeedback(null);
  }, [i]);

  if (i >= queue.length) {
    const pct = Math.round((correctCount / queue.length) * 100);
    return (
      <ResultScreen
        pct={pct}
        correct={correctCount}
        total={queue.length}
        done={done}
        label="拼字測驗"
        onComplete={onComplete}
        onRetry={() => {
          setQueue(pickRandom(cards, N));
          setI(0);
          setCorrectCount(0);
          setInput("");
        }}
      />
    );
  }

  const word = queue[i].word;

  function goNext() {
    setInput("");
    setI((x) => x + 1);
  }

  function submit() {
    const correct = input.trim().toLowerCase() === word.toLowerCase();
    setFeedback(correct ? "correct" : "wrong");
    if (correct) {
      setCorrectCount((c) => c + 1);
      setTimeout(goNext, 700);
    } else {
      setEncourageMsg(randomFrom(ENCOURAGE_MSGS));
      setStopped(true);
    }
  }

  function reveal() {
    setRevealed(true);
    setFeedback("wrong");
    setEncourageMsg(randomFrom(REVEAL_MSGS));
    setStopped(true);
  }

  return (
    <div style={{ textAlign: "center" }}>
      <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 6 }}>第 {i + 1} / {queue.length} 題</div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 4 }}>
        {queue[i].pos && (
          <span style={{ fontSize: 11.5, fontWeight: 700, color: COLORS.accent, background: COLORS.accentSoft, borderRadius: 6, padding: "2px 7px" }}>
            {queue[i].pos}
          </span>
        )}
        <div style={{ fontSize: 22, fontWeight: 700 }}>{queue[i].definition}</div>
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 16 }}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.accent, fontSize: 14 }}>
          {queue[i].phonetic || ""}
        </div>
        <button onClick={() => speakWord(word, voiceURI)} title="聽發音" style={{ border: "none", background: "transparent", color: COLORS.accent, display: "flex", padding: 2 }}>
          <Volume2 size={15} />
        </button>
      </div>

      {!revealed && showHint && (
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 20, letterSpacing: 3, color: COLORS.accent, marginBottom: 10 }}>
          {hintDisplay(word)}
        </div>
      )}

      <input
        autoFocus
        disabled={stopped}
        value={input}
        onChange={(e) => setInput(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && input && !stopped && submit()}
        placeholder="拼出這個英文單字"
        style={{
          ...inputStyle,
          borderColor: feedback === "correct" ? COLORS.good : feedback === "wrong" ? COLORS.bad : COLORS.line,
          textAlign: "center",
          fontSize: 18,
          maxWidth: 320,
        }}
      />

      {!stopped && (
        <div style={{ display: "flex", gap: 8, justifyContent: "center", marginTop: 14, flexWrap: "wrap" }}>
          <button onClick={() => setShowHint((h) => !h)} style={ghostBtn}>
            💡 {showHint ? "隱藏提示" : "提示"}
          </button>
          <button onClick={reveal} style={{ ...ghostBtn, borderColor: `${COLORS.bad}55`, color: COLORS.bad }}>
            顯示答案
          </button>
          <button onClick={submit} disabled={!input} style={primaryBtn}>
            確認
          </button>
        </div>
      )}

      {stopped && <WrongAnswerPanel word={word} message={encourageMsg} onNext={goNext} />}
    </div>
  );
}

/* ---- Multiple choice quiz (sentence w/ blank -> ABC) ---- */
function buildMCQuestion(cards, card) {
  const others = cards.filter((c) => c.id !== card.id);
  const distractors = pickRandom(others, Math.min(2, others.length)).map((c) => c.word);
  const options = shuffleArr([card.word, ...distractors]);
  let sentence;
  if (card.example && card.example.toLowerCase().includes(card.word.toLowerCase())) {
    const re = new RegExp(card.word, "i");
    sentence = card.example.replace(re, "_____");
  } else {
    sentence = `This word means: “${card.definition}” → _____`;
  }
  return { card, sentence, options };
}

function MultipleChoiceQuiz({ cards, onComplete, done, showToast }) {
  const N = Math.min(8, cards.length);
  const [queue, setQueue] = useState(() => pickRandom(cards, N).map((c) => buildMCQuestion(cards, c)));
  const [i, setI] = useState(0);
  const [chosen, setChosen] = useState(null);
  const [correctCount, setCorrectCount] = useState(0);
  const [encourageMsg, setEncourageMsg] = useState("");

  if (i >= queue.length) {
    const pct = Math.round((correctCount / queue.length) * 100);
    return (
      <ResultScreen
        pct={pct}
        correct={correctCount}
        total={queue.length}
        done={done}
        label="選擇題測驗"
        onComplete={onComplete}
        onRetry={() => {
          setQueue(pickRandom(cards, N).map((c) => buildMCQuestion(cards, c)));
          setI(0);
          setCorrectCount(0);
          setChosen(null);
        }}
      />
    );
  }

  const q = queue[i];
  const letters = ["A", "B", "C"];

  function choose(opt) {
    if (chosen) return;
    setChosen(opt);
    const correct = opt === q.card.word;
    if (correct) {
      setCorrectCount((c) => c + 1);
      setTimeout(() => {
        setChosen(null);
        setI((x) => x + 1);
      }, 700);
    } else {
      setEncourageMsg(randomFrom(ENCOURAGE_MSGS));
    }
  }

  const wrongChosen = chosen != null && chosen !== q.card.word;

  return (
    <div>
      <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 10, textAlign: "center" }}>
        第 {i + 1} / {queue.length} 題
      </div>
      <div
        style={{
          background: COLORS.paper,
          border: `1px solid ${COLORS.line}`,
          borderRadius: 14,
          padding: 20,
          fontSize: 16,
          textAlign: "center",
          marginBottom: 16,
        }}
      >
        {q.sentence}
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {q.options.map((opt, idx) => {
          const isCorrect = opt === q.card.word;
          const showState = chosen != null;
          let bg = COLORS.paper;
          let border = COLORS.line;
          if (showState && opt === chosen) {
            bg = isCorrect ? "#E7F2E9" : "#F6E4E4";
            border = isCorrect ? COLORS.good : COLORS.bad;
          } else if (showState && isCorrect) {
            bg = "#E7F2E9";
            border = COLORS.good;
          }
          return (
            <button
              key={opt}
              onClick={() => choose(opt)}
              disabled={chosen != null}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: "12px 14px",
                borderRadius: 10,
                border: `1.5px solid ${border}`,
                background: bg,
                textAlign: "left",
                fontSize: 15,
              }}
            >
              <span
                style={{
                  width: 22,
                  height: 22,
                  borderRadius: 6,
                  background: COLORS.accentSoft,
                  color: COLORS.accent,
                  fontSize: 12,
                  fontWeight: 700,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                }}
              >
                {letters[idx]}
              </span>
              {opt}
            </button>
          );
        })}
      </div>

      {wrongChosen && (
        <WrongAnswerPanel
          word={q.card.word}
          message={encourageMsg}
          onNext={() => {
            setChosen(null);
            setI((x) => x + 1);
          }}
        />
      )}
    </div>
  );
}

function ResultScreen({ pct, correct, total, done, label, onComplete, onRetry }) {
  useEffect(() => {
    if (pct >= 60) onComplete();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return (
    <div style={{ textAlign: "center", padding: "30px 10px" }} className="popIn">
      <Sparkles size={30} color={COLORS.energy} />
      <div style={{ fontSize: 22, fontWeight: 700, margin: "10px 0 4px" }}>{label}完成！</div>
      <div style={{ fontSize: 15, color: COLORS.inkSoft, marginBottom: 14 }}>
        答對 {correct} / {total}（{pct}%）
      </div>
      {pct >= 60 ? (
        !done ? (
          <div style={{ color: COLORS.good, fontWeight: 600, marginBottom: 14 }}>+20 能量</div>
        ) : (
          <div style={{ color: COLORS.good, fontWeight: 600, marginBottom: 14 }}>✓ 今天此任務已完成過</div>
        )
      ) : (
        <div style={{ color: COLORS.bad, fontSize: 13, marginBottom: 14 }}>答對率需達 60% 才能獲得能量，再試一次吧！</div>
      )}
      <button onClick={onRetry} style={ghostBtn}>
        <RotateCw size={14} /> 再測一次
      </button>
    </div>
  );
}

/* ---------------------------------------------------------
   Import tab
--------------------------------------------------------- */
const HEADER_MAP = {
  source: ["source", "來源", "頁碼", "page"],
  word: ["word", "單字", "英文", "vocab", "term"],
  pos: ["pos", "詞性", "part of speech", "word class", "詞類"],
  phonetic: ["phonetic", "音標", "發音", "ipa", "pronunciation"],
  definition: ["definition", "翻譯", "解釋", "意思", "translation", "meaning", "中文"],
  collocation: ["collocation", "搭配", "常用詞", "phrase"],
  collocationMeaning: ["collocation meaning", "搭配詞中文解釋", "搭配意思", "搭配詞解釋"],
  synonym: ["synonym", "同義詞", "相似詞", "近義詞"],
  example: ["example", "例句", "sentence", "句子"],
  category: ["category", "分類", "類別", "class", "group"],
};

// Column order used by the most recently uploaded reference file:
// 頁碼/來源 (Source/Page) | 英文單字 (English) | 中文翻譯 (Chinese) | 詞性 (POS)
// | 常用搭配 (Collocation) | 搭配詞中文解釋 (Collocation Meaning) | 相似詞/同義詞 (Synonyms) | 例句 (Example Sentence)
const REFERENCE_COLUMN_ORDER = {
  source: 0,
  word: 1,
  definition: 2,
  pos: 3,
  collocation: 4,
  collocationMeaning: 5,
  synonym: 6,
  example: 7,
  phonetic: -1,
  category: -1,
};

function detectColumns(header) {
  const lower = header.map((h) => String(h || "").toLowerCase().trim());
  const map = {};
  Object.keys(HEADER_MAP).forEach((field) => {
    // for fields that can share keywords (e.g. collocation vs collocation meaning),
    // check the more specific field first so it isn't accidentally matched by the broader one
    const idx = lower.findIndex((h) => HEADER_MAP[field].some((kw) => h.includes(kw)));
    map[field] = idx;
  });
  // avoid collocation accidentally grabbing the "collocation meaning" column
  if (map.collocation !== -1 && map.collocation === map.collocationMeaning) {
    const nextIdx = lower.findIndex(
      (h, i) => i !== map.collocationMeaning && HEADER_MAP.collocation.some((kw) => h.includes(kw))
    );
    map.collocation = nextIdx;
  }
  // fallback to the reference file's column order if nothing matched at all
  const anyMatched = Object.values(map).some((v) => v !== -1);
  if (!anyMatched) {
    return { ...REFERENCE_COLUMN_ORDER };
  }
  return map;
}

function rowsToCards(rows, colMap, hasHeader, defaultCategory) {
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const out = [];
  dataRows.forEach((r) => {
    const word = colMap.word >= 0 ? r[colMap.word] : "";
    if (!word || !String(word).trim()) return;
    const cat = colMap.category >= 0 ? String(r[colMap.category] || "").trim() : "";
    let collocation = colMap.collocation >= 0 ? String(r[colMap.collocation] || "").trim() : "";
    const collocationMeaning = colMap.collocationMeaning >= 0 ? String(r[colMap.collocationMeaning] || "").trim() : "";
    if (collocationMeaning) {
      collocation = collocation ? `${collocation}（${collocationMeaning}）` : collocationMeaning;
    }
    out.push({
      id: uid(),
      word: String(word).trim(),
      pos: colMap.pos >= 0 ? String(r[colMap.pos] || "").trim() : "",
      phonetic: colMap.phonetic >= 0 ? String(r[colMap.phonetic] || "").trim() : "",
      definition: colMap.definition >= 0 ? String(r[colMap.definition] || "").trim() : "",
      example: colMap.example >= 0 ? String(r[colMap.example] || "").trim() : "",
      collocation,
      synonym: colMap.synonym >= 0 ? String(r[colMap.synonym] || "").trim() : "",
      category: cat || defaultCategory || UNCATEGORIZED,
    });
  });
  return out;
}

function ImportTab({ onAdd, cards, onDelete, categories, onMove, onAddCategory, onRemoveCategory }) {
  const [form, setForm] = useState({ word: "", pos: "", phonetic: "", definition: "", example: "", collocation: "", synonym: "", category: UNCATEGORIZED });
  const [pasteText, setPasteText] = useState("");
  const [importCategory, setImportCategory] = useState(UNCATEGORIZED);
  const [newCatName, setNewCatName] = useState("");
  const [listFilterCat, setListFilterCat] = useState("全部");
  const fileRef = useRef(null);
  const [fileMsg, setFileMsg] = useState("");
  const [builtinAdded, setBuiltinAdded] = useState(false);
  const [zodiacAdded, setZodiacAdded] = useState(false);
  const [readingAdded, setReadingAdded] = useState(false);

  function addManual() {
    if (!form.word.trim() || !form.definition.trim()) return;
    onAdd([{
      id: uid(),
      word: form.word.trim(),
      pos: form.pos.trim(),
      phonetic: form.phonetic.trim(),
      definition: form.definition.trim(),
      example: form.example.trim(),
      collocation: form.collocation.trim(),
      synonym: form.synonym.trim(),
      category: form.category || UNCATEGORIZED,
    }]);
    setForm({ word: "", pos: "", phonetic: "", definition: "", example: "", collocation: "", synonym: "", category: form.category });
  }

  function importBuiltinReading() {
    const existingWords = new Set(cards.map((c) => c.word.toLowerCase()));
    const toAdd = BUILTIN_READING_WORDS.filter((c) => !existingWords.has(c.word.toLowerCase())).map((c) => ({
      ...c,
      id: uid(),
    }));
    if (toAdd.length === 0) {
      setReadingAdded(true);
      return;
    }
    onAdd(toAdd);
    setReadingAdded(true);
  }

  function importBuiltinDailyLife() {
    const existingWords = new Set(cards.map((c) => c.word.toLowerCase()));
    const toAdd = BUILTIN_DAILY_LIFE_CARDS.filter((c) => !existingWords.has(c.word.toLowerCase())).map((c) => ({
      ...c,
      id: uid(),
    }));
    if (toAdd.length === 0) {
      setBuiltinAdded(true);
      return;
    }
    onAdd(toAdd);
    setBuiltinAdded(true);
  }

  function importBuiltinZodiac() {
    const existingWords = new Set(cards.map((c) => c.word.toLowerCase()));
    const toAdd = BUILTIN_ZODIAC_CARDS.filter((c) => !existingWords.has(c.word.toLowerCase())).map((c) => ({
      ...c,
      id: uid(),
    }));
    if (toAdd.length === 0) {
      setZodiacAdded(true);
      return;
    }
    onAdd(toAdd);
    setZodiacAdded(true);
  }

  function importPaste() {
    if (!pasteText.trim()) return;
    const lines = pasteText.trim().split("\n").filter((l) => l.trim());
    const delimiter = lines[0].includes("\t") ? "\t" : ",";
    const rows = lines.map((l) => l.split(delimiter).map((c) => c.trim()));
    const header = rows[0];
    const looksLikeHeader = header.some((h) => Object.values(HEADER_MAP).flat().some((kw) => String(h).toLowerCase().includes(kw)));
    const colMap = detectColumns(header);
    const newCards = rowsToCards(rows, colMap, looksLikeHeader, importCategory);
    if (newCards.length === 0) {
      setFileMsg("找不到可用的資料，請確認格式為：單字, 音標, 解釋, 例句（以逗號或 Tab 分隔）。");
      return;
    }
    onAdd(newCards);
    setPasteText("");
    setFileMsg("");
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        if (!rows.length) {
          setFileMsg("這個檔案是空的。");
          return;
        }
        const header = rows[0];
        const looksLikeHeader = header.some((h) => Object.values(HEADER_MAP).flat().some((kw) => String(h).toLowerCase().includes(kw)));
        const colMap = detectColumns(header);
        const newCards = rowsToCards(rows, colMap, looksLikeHeader, importCategory);
        if (newCards.length === 0) {
          setFileMsg("找不到可用的欄位，請確認欄位名稱包含 word/單字、definition/解釋 等關鍵字。");
          return;
        }
        onAdd(newCards);
        setFileMsg(`成功匯入 ${newCards.length} 筆資料。`);
      } catch (err) {
        setFileMsg("讀取檔案失敗，請確認是 .xlsx 或 .csv 檔案。");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  const visibleCards = listFilterCat === "全部" ? cards : cards.filter((c) => (c.category || UNCATEGORIZED) === listFilterCat);

  return (
    <div style={{ marginTop: 22, display: "flex", flexDirection: "column", gap: 22 }}>
      {/* Built-in dataset */}
      <section
        style={{
          background: COLORS.accentSoft,
          border: `1px solid ${COLORS.accent}44`,
          borderRadius: 14,
          padding: 16,
        }}
      >
        <SectionTitle icon={Sparkles} text="內建題庫：日常生活單字" />
        <p style={{ fontSize: 12.5, color: "#3A5850", marginTop: -6, marginBottom: 10 }}>
          你上傳的 Excel 檔案（{BUILTIN_DAILY_LIFE_CARDS.length} 個單字，含詞性、中文解釋、搭配詞與例句）已經自動匯入到「日常生活」分類了，可以直接到「字卡」分頁選這個分類開始複習。如果之後刪除了想重新加回來，也可以按下面的按鈕再匯入一次（重複的單字會自動略過）。
        </p>
        <button onClick={importBuiltinDailyLife} style={primaryBtn}>
          <Upload size={15} /> 重新匯入日常生活單字（{BUILTIN_DAILY_LIFE_CARDS.length} 個）
        </button>
        {builtinAdded && (
          <div style={{ fontSize: 12.5, color: COLORS.good, fontWeight: 600, marginTop: 8 }}>
            ✓ 已匯入完成，可以到「字卡」分頁選擇「日常生活」分類開始複習。
          </div>
        )}
      </section>

      {/* Built-in dataset: zodiac */}
      <section
        style={{
          background: COLORS.energySoft,
          border: `1px solid ${COLORS.energy}55`,
          borderRadius: 14,
          padding: 16,
        }}
      >
        <SectionTitle icon={Sparkles} text="內建題庫：星座 (Zodiac Signs)" />
        <p style={{ fontSize: 12.5, color: "#7A5A16", marginTop: -6, marginBottom: 10 }}>
          已經自動幫你加入 {BUILTIN_ZODIAC_CARDS.length} 個星座相關單字（12 星座＋star sign / zodiac sign / constellation / horoscope），全部歸類在「星座」分類。如果之後刪除了想重新加回來，可以按下面的按鈕（重複的單字會自動略過）。
        </p>
        <button onClick={importBuiltinZodiac} style={primaryBtn}>
          <Upload size={15} /> 重新匯入星座單字（{BUILTIN_ZODIAC_CARDS.length} 個）
        </button>
        {zodiacAdded && (
          <div style={{ fontSize: 12.5, color: COLORS.good, fontWeight: 600, marginTop: 8 }}>
            ✓ 已匯入完成，可以到「字卡」分頁選擇「星座」分類開始複習。
          </div>
        )}
      </section>

      {/* Built-in dataset: reading article words */}
      <section
        style={{
          background: "#EDE7F6",
          border: "1px solid #8C5B9F55",
          borderRadius: 14,
          padding: 16,
        }}
      >
        <SectionTitle icon={Sparkles} text="內建題庫：閱讀文章單字" />
        <p style={{ fontSize: 12.5, color: "#4E3A63", marginTop: -6, marginBottom: 10 }}>
          已經自動幫你加入 {BUILTIN_READING_WORDS.length} 個單字（含詞性、中文解釋、搭配詞、同義詞與例句），全部歸類在「閱讀文章單字」分類。如果之後刪除了想重新加回來，可以按下面的按鈕（重複的單字會自動略過）。
        </p>
        <button onClick={importBuiltinReading} style={primaryBtn}>
          <Upload size={15} /> 重新匯入閱讀文章單字（{BUILTIN_READING_WORDS.length} 個）
        </button>
        {readingAdded && (
          <div style={{ fontSize: 12.5, color: COLORS.good, fontWeight: 600, marginTop: 8 }}>
            ✓ 已匯入完成，可以到「字卡」分頁選擇「閱讀文章單字」分類開始複習。
          </div>
        )}
      </section>

      {/* Category management */}
      <section>
        <SectionTitle icon={Layers} text="分類管理" />
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: -6 }}>
          新增或刪除分類；刪除分類時，裡面的字卡會自動移到「{UNCATEGORIZED}」。
        </p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
          {categories.map((c) => (
            <div
              key={c}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                background: COLORS.accentSoft,
                color: COLORS.accent,
                borderRadius: 999,
                padding: "5px 6px 5px 12px",
                fontSize: 12.5,
                fontWeight: 600,
              }}
            >
              {c}
              {c !== UNCATEGORIZED && (
                <button
                  onClick={() => onRemoveCategory(c)}
                  style={{ border: "none", background: "transparent", color: COLORS.accent, display: "flex", padding: 2 }}
                  title="刪除分類"
                >
                  <X size={12} />
                </button>
              )}
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            placeholder="新增分類名稱，例如：面試常用語"
            value={newCatName}
            onChange={(e) => setNewCatName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && newCatName.trim()) {
                onAddCategory(newCatName);
                setNewCatName("");
              }
            }}
            style={{ ...inputStyle, flex: 1 }}
          />
          <button
            onClick={() => {
              if (!newCatName.trim()) return;
              onAddCategory(newCatName);
              setNewCatName("");
            }}
            style={ghostBtn}
          >
            <Plus size={14} /> 新增分類
          </button>
        </div>
      </section>

      {/* Manual entry */}
      <section>
        <SectionTitle icon={Plus} text="手動新增" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
          <input placeholder="單字 (word)" value={form.word} onChange={(e) => setForm({ ...form, word: e.target.value })} style={inputStyle} />
          <input placeholder="詞性 (adj./n./v.)" value={form.pos} onChange={(e) => setForm({ ...form, pos: e.target.value })} style={inputStyle} />
          <input placeholder="音標 (phonetic)" value={form.phonetic} onChange={(e) => setForm({ ...form, phonetic: e.target.value })} style={inputStyle} />
        </div>
        <input placeholder="解釋 (definition)" value={form.definition} onChange={(e) => setForm({ ...form, definition: e.target.value })} style={{ ...inputStyle, width: "100%", marginTop: 8 }} />
        <input placeholder="例句 (example)" value={form.example} onChange={(e) => setForm({ ...form, example: e.target.value })} style={{ ...inputStyle, width: "100%", marginTop: 8 }} />
        <input placeholder="搭配詞 (collocation)，例如：set an alarm clock" value={form.collocation} onChange={(e) => setForm({ ...form, collocation: e.target.value })} style={{ ...inputStyle, width: "100%", marginTop: 8 }} />
        <input placeholder="同義詞 (synonym)，例如：rely on, count on" value={form.synonym} onChange={(e) => setForm({ ...form, synonym: e.target.value })} style={{ ...inputStyle, width: "100%", marginTop: 8 }} />
        <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} style={{ ...inputStyle, width: "100%", marginTop: 8 }}>
          {categories.map((c) => (
            <option key={c} value={c}>{c}</option>
          ))}
        </select>
        <button onClick={addManual} style={{ ...primaryBtn, marginTop: 10 }}>
          <Plus size={15} /> 新增這張字卡
        </button>
      </section>

      {/* Import default category */}
      <section>
        <SectionTitle icon={FileSpreadsheet} text="匯入設定" />
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: -6 }}>
          若匯入的檔案 / 文字沒有「分類」欄位，會自動歸類到這裡選的分類。
        </p>
        <select value={importCategory} onChange={(e) => setImportCategory(e.target.value)} style={{ ...inputStyle, maxWidth: 240 }}>
          {categories.map((c) => (
            <option key={c} value={c}>{c}</option>
          ))}
        </select>
      </section>

      {/* Excel upload */}
      <section>
        <SectionTitle icon={FileSpreadsheet} text="上傳 Excel / CSV 檔案" />
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: -6 }}>
          系統會依照欄位標題自動配對，建議格式（跟妳上傳過的檔案一致）：
          <br />
          頁碼/來源（會被忽略）→ 英文單字 → 中文翻譯 → 詞性 → 常用搭配 → 搭配詞中文解釋 → 相似詞/同義詞 → 例句
          <br />
          也支援 word / phonetic / definition / example / category 等英文欄位標題。
        </p>
        <button onClick={() => fileRef.current?.click()} style={ghostBtn}>
          <Upload size={14} /> 選擇檔案 (.xlsx / .csv)
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
      </section>

      {/* Paste text (also covers text copied from a PDF) */}
      <section>
        <SectionTitle icon={ClipboardPaste} text="貼上文字匯入（也適用於從 PDF 複製的內容）" />
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: -6 }}>
          每行一筆資料，用逗號或 Tab 分開。若有標題列會自動偵測欄位；若沒有標題列，預設欄位順序為：來源(略過), 單字, 中文翻譯, 詞性, 搭配詞, 搭配詞解釋, 同義詞, 例句。瀏覽器無法直接讀取 PDF 檔案，請先從 PDF 複製文字後貼在這裡。
        </p>
        <textarea
          value={pasteText}
          onChange={(e) => setPasteText(e.target.value)}
          placeholder={"Image 18 (P.1), cuisine, 菜餚；料理, n., haute cuisine, 高級料理, culinary art food dish, Italian cuisine is famous all over the world.\nImage 18 (P.1), contemporary, 當代的, adj., contemporary art, 當代藝術, modern current, She is a big fan of contemporary architecture."}
          rows={5}
          style={{ ...inputStyle, width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}
        />
        <button onClick={importPaste} style={{ ...primaryBtn, marginTop: 8 }}>
          <Upload size={14} /> 匯入貼上的文字
        </button>
        {fileMsg && <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 6 }}>{fileMsg}</p>}
      </section>

      {/* Card list / manage */}
      <section>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
          <SectionTitle icon={Layers} text={`目前的單字卡 (${visibleCards.length}/${cards.length})`} />
          <select value={listFilterCat} onChange={(e) => setListFilterCat(e.target.value)} style={{ ...inputStyle, fontSize: 12.5, padding: "6px 8px" }}>
            <option value="全部">全部類別</option>
            {categories.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 360, overflowY: "auto" }}>
          {visibleCards.map((c) => (
            <div
              key={c.id}
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "8px 12px",
                background: COLORS.paper,
                border: `1px solid ${COLORS.line}`,
                borderRadius: 8,
                gap: 8,
                flexWrap: "wrap",
              }}
            >
              <div style={{ fontSize: 13.5, flex: 1, minWidth: 160 }}>
                <strong>{c.word}</strong>
                {c.pos && <span style={{ color: COLORS.accent, fontSize: 11.5, fontWeight: 700 }}> {c.pos}</span>}
                <span style={{ color: COLORS.inkSoft }}> — {c.definition}</span>
              </div>
              <select
                value={c.category || UNCATEGORIZED}
                onChange={(e) => onMove(c.id, e.target.value)}
                title="搬移到其他分類"
                style={{ ...inputStyle, fontSize: 12, padding: "4px 6px" }}
              >
                {categories.map((cat) => (
                  <option key={cat} value={cat}>{cat}</option>
                ))}
              </select>
              <button onClick={() => onDelete(c.id)} style={{ border: "none", background: "transparent", color: COLORS.bad }}>
                <Trash2 size={15} />
              </button>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function SectionTitle({ icon: Icon, text }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
      <Icon size={16} color={COLORS.accent} />
      <div style={{ fontWeight: 700, fontSize: 15 }}>{text}</div>
    </div>
  );
}

/* ---------------------------------------------------------
   Reading tab — upload an article (text / PDF / image),
   highlight passages, and keep notes
--------------------------------------------------------- */
const HIGHLIGHT_COLORS = [
  { name: "黃色", value: "#FCE68A" },
  { name: "綠色", value: "#B7E4C7" },
  { name: "粉色", value: "#F6C6CE" },
];

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function ReadingTab({ articles, setArticles, onComplete, done, showToast }) {
  const [openId, setOpenId] = useState(null);
  const [addMode, setAddMode] = useState("text");
  const [title, setTitle] = useState("");
  const [textBody, setTextBody] = useState("");
  const fileRef = useRef(null);
  const [uploading, setUploading] = useState(false);

  const openArticle = articles.find((a) => a.id === openId);

  function addTextArticle() {
    if (!textBody.trim()) return;
    const paragraphs = textBody
      .trim()
      .split(/\n+/)
      .map((p) => `<p>${escapeHtml(p)}</p>`)
      .join("");
    const newArticle = {
      id: uid(),
      title: title.trim() || "未命名文章",
      type: "text",
      contentHtml: paragraphs,
      notes: "",
      createdAt: Date.now(),
    };
    setArticles((prev) => [newArticle, ...prev]);
    setTitle("");
    setTextBody("");
    setOpenId(newArticle.id);
    showToast("文章已新增", "success");
  }

  async function handleFile(e, kind) {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > 4.5 * 1024 * 1024) {
      showToast("檔案太大了（上限約 4.5MB），請換一個較小的檔案。", "error");
      e.target.value = "";
      return;
    }
    setUploading(true);
    try {
      const dataUrl = await fileToDataUrl(file);
      const newArticle = {
        id: uid(),
        title: file.name,
        type: kind,
        fileData: dataUrl,
        notes: "",
        createdAt: Date.now(),
      };
      setArticles((prev) => [newArticle, ...prev]);
      setOpenId(newArticle.id);
      showToast("已上傳", "success");
    } catch {
      showToast("上傳失敗，請再試一次。", "error");
    } finally {
      setUploading(false);
      e.target.value = "";
    }
  }

  function updateArticle(id, patch) {
    setArticles((prev) => prev.map((a) => (a.id === id ? { ...a, ...patch } : a)));
  }

  function deleteArticle(id) {
    setArticles((prev) => prev.filter((a) => a.id !== id));
    if (openId === id) setOpenId(null);
  }

  if (openArticle) {
    return (
      <ArticleReader
        article={openArticle}
        onBack={() => setOpenId(null)}
        onUpdate={(patch) => updateArticle(openArticle.id, patch)}
        onDelete={() => deleteArticle(openArticle.id)}
        onComplete={onComplete}
        done={done}
      />
    );
  }

  return (
    <div style={{ marginTop: 22, display: "flex", flexDirection: "column", gap: 22 }}>
      <section>
        <SectionTitle icon={Plus} text="新增文章" />
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          {[
            { id: "text", label: "貼上文字", icon: FileText },
            { id: "pdf", label: "上傳 PDF", icon: FileSpreadsheet },
            { id: "image", label: "上傳圖片", icon: ImageIcon },
          ].map((m) => {
            const Icon = m.icon;
            const active = addMode === m.id;
            return (
              <button
                key={m.id}
                onClick={() => setAddMode(m.id)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  padding: "7px 12px",
                  borderRadius: 999,
                  border: `1px solid ${active ? COLORS.accent : COLORS.line}`,
                  background: active ? COLORS.accentSoft : COLORS.paper,
                  color: active ? COLORS.accent : COLORS.inkSoft,
                  fontSize: 12.5,
                  fontWeight: 600,
                }}
              >
                <Icon size={13} /> {m.label}
              </button>
            );
          })}
        </div>

        {addMode === "text" && (
          <div>
            <input placeholder="文章標題" value={title} onChange={(e) => setTitle(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 8 }} />
            <textarea
              placeholder="貼上或輸入英文文章內容…"
              value={textBody}
              onChange={(e) => setTextBody(e.target.value)}
              rows={7}
              style={{ ...inputStyle, width: "100%" }}
            />
            <button onClick={addTextArticle} style={{ ...primaryBtn, marginTop: 8 }}>
              <Plus size={15} /> 新增文章
            </button>
          </div>
        )}

        {addMode === "pdf" && (
          <div>
            <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 0 }}>
              上傳 PDF 後會在下方直接顯示閱讀，並可以在旁邊寫筆記（PDF 內文暫不支援螢光筆畫記）。
            </p>
            <button onClick={() => fileRef.current?.click()} style={ghostBtn} disabled={uploading}>
              <Upload size={14} /> {uploading ? "上傳中…" : "選擇 PDF 檔案"}
            </button>
            <input ref={fileRef} type="file" accept="application/pdf" onChange={(e) => handleFile(e, "pdf")} style={{ display: "none" }} />
          </div>
        )}

        {addMode === "image" && (
          <div>
            <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 0 }}>
              上傳文章截圖或照片，可以在旁邊寫筆記（圖片暫不支援螢光筆畫記）。
            </p>
            <button onClick={() => fileRef.current?.click()} style={ghostBtn} disabled={uploading}>
              <Upload size={14} /> {uploading ? "上傳中…" : "選擇圖片"}
            </button>
            <input ref={fileRef} type="file" accept="image/*" onChange={(e) => handleFile(e, "image")} style={{ display: "none" }} />
          </div>
        )}
      </section>

      <section>
        <SectionTitle icon={BookOpen} text={`我的文章 (${articles.length})`} />
        {articles.length === 0 ? (
          <EmptyState text="還沒有文章，先新增一篇吧。" />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {articles.map((a) => {
              const Icon = a.type === "text" ? FileText : a.type === "pdf" ? FileSpreadsheet : ImageIcon;
              return (
                <div
                  key={a.id}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    padding: "10px 12px",
                    background: COLORS.paper,
                    border: `1px solid ${COLORS.line}`,
                    borderRadius: 10,
                  }}
                >
                  <Icon size={16} color={COLORS.accent} />
                  <button
                    onClick={() => setOpenId(a.id)}
                    style={{ flex: 1, textAlign: "left", background: "transparent", border: "none", fontSize: 14, fontWeight: 600, color: COLORS.ink, padding: 0 }}
                  >
                    {a.title}
                  </button>
                  {a.notes && <StickyNote size={14} color={COLORS.energy} />}
                  <button onClick={() => deleteArticle(a.id)} style={{ border: "none", background: "transparent", color: COLORS.bad }}>
                    <Trash2 size={15} />
                  </button>
                </div>
              );
            })}
          </div>
        )}
      </section>
    </div>
  );
}

function ArticleReader({ article, onBack, onUpdate, onDelete, onComplete, done }) {
  const editorRef = useRef(null);
  const [notes, setNotes] = useState(article.notes || "");

  useEffect(() => {
    if (editorRef.current && article.type === "text") {
      editorRef.current.innerHTML = article.contentHtml || "";
    }
    setNotes(article.notes || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [article.id]);

  function applyHighlight(color) {
    editorRef.current?.focus();
    try {
      document.execCommand("hiliteColor", false, color);
    } catch {
      /* no-op if unsupported */
    }
    saveContent();
  }

  function clearHighlight() {
    editorRef.current?.focus();
    try {
      document.execCommand("removeFormat");
    } catch {
      /* no-op */
    }
    saveContent();
  }

  function saveContent() {
    if (editorRef.current) {
      onUpdate({ contentHtml: editorRef.current.innerHTML });
    }
  }

  function saveNotes(v) {
    setNotes(v);
    onUpdate({ notes: v });
  }

  return (
    <div style={{ marginTop: 22 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <button onClick={onBack} style={ghostBtn}>
          <ChevronLeft size={14} /> 返回文章列表
        </button>
        <button onClick={onDelete} style={{ ...ghostBtn, borderColor: `${COLORS.bad}55`, color: COLORS.bad }}>
          <Trash2 size={14} /> 刪除
        </button>
      </div>

      <div style={{ fontFamily: "'Source Serif 4', serif", fontSize: 20, fontWeight: 700, marginBottom: 12 }}>{article.title}</div>

      {article.type === "text" && (
        <>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
            {HIGHLIGHT_COLORS.map((c) => (
              <button
                key={c.value}
                onClick={() => applyHighlight(c.value)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  border: `1px solid ${COLORS.line}`,
                  background: COLORS.paper,
                  borderRadius: 8,
                  padding: "6px 10px",
                  fontSize: 12.5,
                }}
              >
                <Highlighter size={13} color="#8A6316" />
                <span style={{ width: 12, height: 12, borderRadius: 3, background: c.value, display: "inline-block" }} />
                {c.name}
              </button>
            ))}
            <button onClick={clearHighlight} style={ghostBtn}>
              <Eraser size={13} /> 清除標記
            </button>
          </div>
          <p style={{ fontSize: 11.5, color: COLORS.inkSoft, marginTop: -4, marginBottom: 10 }}>
            選取文章中的文字，再點上面的顏色來畫螢光筆。
          </p>
          <div
            ref={editorRef}
            contentEditable
            suppressContentEditableWarning
            onBlur={saveContent}
            onMouseUp={saveContent}
            style={{
              background: COLORS.paper,
              border: `1px solid ${COLORS.line}`,
              borderRadius: 14,
              padding: 20,
              fontSize: 15.5,
              lineHeight: 1.8,
              minHeight: 220,
              outline: "none",
            }}
          />
        </>
      )}

      {article.type === "pdf" && (
        <iframe
          title={article.title}
          src={article.fileData}
          style={{ width: "100%", height: 480, border: `1px solid ${COLORS.line}`, borderRadius: 14, background: "#fff" }}
        />
      )}

      {article.type === "image" && (
        <div style={{ background: COLORS.paper, border: `1px solid ${COLORS.line}`, borderRadius: 14, padding: 12, textAlign: "center" }}>
          <img src={article.fileData} alt={article.title} style={{ maxWidth: "100%", borderRadius: 8 }} />
        </div>
      )}

      <div style={{ marginTop: 16 }}>
        <SectionTitle icon={StickyNote} text="我的筆記" />
        <textarea
          value={notes}
          onChange={(e) => saveNotes(e.target.value)}
          placeholder="寫下單字、心得或疑問…"
          rows={5}
          style={{ ...inputStyle, width: "100%" }}
        />
      </div>

      {!done ? (
        <button onClick={onComplete} style={{ ...primaryBtn, width: "100%", marginTop: 18 }}>
          <Check size={16} /> 完成本次閱讀 (+20 能量)
        </button>
      ) : (
        <div style={{ textAlign: "center", marginTop: 18, fontSize: 13, color: COLORS.good, fontWeight: 600 }}>
          ✓ 今天的閱讀任務已完成
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Progress tab
--------------------------------------------------------- */
function VoiceSettings({ voicePref, setVoicePref }) {
  const [voices, setVoices] = useState(() => getAvailableVoices());

  useEffect(() => {
    if (voices.length === 0) {
      const t = setTimeout(() => setVoices(getAvailableVoices()), 400);
      const onChange = () => setVoices(getAvailableVoices());
      if (window.speechSynthesis) window.speechSynthesis.addEventListener?.("voiceschanged", onChange);
      return () => {
        clearTimeout(t);
        if (window.speechSynthesis) window.speechSynthesis.removeEventListener?.("voiceschanged", onChange);
      };
    }
  }, [voices.length]);

  const englishVoices = voices.filter((v) => v.lang && v.lang.toLowerCase().startsWith("en"));
  const listToShow = englishVoices.length ? englishVoices : voices;
  const recommended = pickBestVoice(voices, null);

  return (
    <section>
      <SectionTitle icon={Volume2} text="發音語音設定" />
      <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: -6, marginBottom: 10 }}>
        單字發音是用瀏覽器內建的語音功能播放，音色好不好聽取決於妳的裝置裡安裝了哪些語音——不同手機/電腦聽起來會不一樣。可以在這裡試聽並選一個比較自然的聲音；沒特別選的話系統會自動挑一個聽起來較好的。
      </p>
      {listToShow.length === 0 ? (
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft }}>正在讀取這台裝置可用的語音…如果一直沒有出現，可能是瀏覽器不支援語音合成。</p>
      ) : (
        <>
          <select
            value={voicePref || ""}
            onChange={(e) => setVoicePref(e.target.value)}
            style={{ ...inputStyle, width: "100%", marginBottom: 10 }}
          >
            <option value="">自動選擇（推薦：{recommended ? recommended.name : "系統預設"}）</option>
            {listToShow.map((v) => (
              <option key={v.voiceURI} value={v.voiceURI}>
                {v.name} ({v.lang}){v.localService ? "" : " · 網路語音，音質通常較好"}
              </option>
            ))}
          </select>
          <button onClick={() => speakWord("Hello, this is what your selected voice sounds like.", voicePref)} style={ghostBtn}>
            <Volume2 size={14} /> 試聽
          </button>
        </>
      )}
    </section>
  );
}

function ProgressTab({ energy, setEnergy, voicePref, setVoicePref }) {
  return (
    <div style={{ marginTop: 22, display: "flex", flexDirection: "column", gap: 22 }}>
      <section
        style={{
          background: COLORS.paper,
          border: `1px solid ${COLORS.line}`,
          borderRadius: 16,
          padding: 20,
          textAlign: "center",
        }}
      >
        <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 6 }}>今日能量</div>
        <div style={{ fontFamily: "'Source Serif 4', serif", fontSize: 46, fontWeight: 700, color: energy.points >= 100 ? COLORS.good : COLORS.ink }}>
          {energy.points}
          <span style={{ fontSize: 18, color: COLORS.inkSoft }}>/100</span>
        </div>
        {energy.points >= 100 ? (
          <div style={{ display: "inline-flex", alignItems: "center", gap: 6, marginTop: 10, background: "#E7F2E9", color: COLORS.good, padding: "6px 14px", borderRadius: 999, fontSize: 13, fontWeight: 600 }}>
            <Coffee size={14} /> 今天達成目標，可以安心休息！
          </div>
        ) : (
          <div style={{ fontSize: 13, color: COLORS.inkSoft, marginTop: 8 }}>再完成 {5 - Object.keys(energy.done).length} 項任務即可休息一天</div>
        )}
      </section>

      <section>
        <SectionTitle icon={ListChecks} text="今日任務" />
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {TASKS.map((t) => {
            const Icon = t.icon;
            const done = !!energy.done[t.id];
            return (
              <div
                key={t.id}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  padding: "10px 14px",
                  borderRadius: 10,
                  background: done ? "#E7F2E9" : COLORS.paper,
                  border: `1px solid ${done ? COLORS.good : COLORS.line}`,
                }}
              >
                <Icon size={16} color={done ? COLORS.good : t.color} />
                <div style={{ flex: 1, fontSize: 14, fontWeight: 500 }}>{t.label}</div>
                <div style={{ fontSize: 13, color: done ? COLORS.good : COLORS.inkSoft, fontWeight: 600 }}>
                  {done ? "+20 已完成" : "+20"}
                </div>
              </div>
            );
          })}
        </div>
      </section>

      <VoiceSettings voicePref={voicePref} setVoicePref={setVoicePref} />

      <section>
        <SectionTitle icon={Bell} text="每日提醒時間" />
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: -6 }}>
          打開 App 時，如果超過這個時間還沒開始任何任務，會顯示提醒訊息。
        </p>
        <input
          type="time"
          value={energy.reminderTime}
          onChange={(e) => setEnergy((prev) => ({ ...prev, reminderTime: e.target.value }))}
          style={{ ...inputStyle, maxWidth: 160 }}
        />
      </section>

      <section
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          padding: "12px 14px",
          background: COLORS.accentSoft,
          borderRadius: 10,
        }}
      >
        <Flame size={16} color={COLORS.accent} />
        <div style={{ fontSize: 13.5, color: COLORS.accent }}>
          累積休息天數：<strong>{energy.restDays}</strong> 天
        </div>
      </section>
    </div>
  );
}

/* ---------------------------------------------------------
   Shared bits
--------------------------------------------------------- */
function EmptyState({ text }) {
  return (
    <div style={{ textAlign: "center", padding: "50px 20px", color: COLORS.inkSoft, fontSize: 14 }}>
      {text}
    </div>
  );
}

function Toast({ toast }) {
  return (
    <div
      style={{
        position: "fixed",
        bottom: 20,
        left: "50%",
        transform: "translateX(-50%)",
        background: toast.kind === "success" ? COLORS.good : COLORS.ink,
        color: "#fff",
        padding: "10px 18px",
        borderRadius: 999,
        fontSize: 13.5,
        fontWeight: 500,
        boxShadow: "0 4px 16px rgba(0,0,0,0.18)",
        animation: "slideUp 0.25s ease-out",
        zIndex: 50,
      }}
    >
      {toast.msg}
    </div>
  );
}

const inputStyle = {
  padding: "9px 12px",
  borderRadius: 8,
  border: `1px solid ${COLORS.line}`,
  fontSize: 13.5,
  outline: "none",
  background: "#fff",
};
const primaryBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  justifyContent: "center",
  background: COLORS.accent,
  color: "#fff",
  border: "none",
  borderRadius: 10,
  padding: "10px 16px",
  fontSize: 13.5,
  fontWeight: 600,
};
const ghostBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  background: "transparent",
  color: COLORS.accent,
  border: `1px solid ${COLORS.accent}55`,
  borderRadius: 8,
  padding: "7px 12px",
  fontSize: 12.5,
  fontWeight: 600,
};
const navBtn = {
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  background: COLORS.paper,
  border: `1px solid ${COLORS.line}`,
  borderRadius: 10,
  padding: "10px 14px",
  color: COLORS.ink,
};
