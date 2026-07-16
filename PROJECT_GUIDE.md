# Word-Game 项目说明

这份文档用于解释这个背单词网页项目的文件结构、核心代码和主要功能实现方式，方便后续继续加功能。

## 1. 网页为什么电脑关机后打不开

先区分两种访问方式：

1. 本地预览地址  
   例如：

   ```text
   http://127.0.0.1:8768/
   http://localhost:8768/
   ```

   这种地址只在你自己的电脑上有效。它依赖本机正在运行的 `python -m http.server` 服务。电脑关机、服务关闭、端口变化后，别人和你自己都打不开。

2. GitHub Pages 地址  
   当前项目的线上地址是：

   ```text
   https://icexuan-2707.github.io/Word-Game/
   ```

   这个地址由 GitHub 托管，理论上不依赖你的电脑。你的电脑关机后也应该能打开。

如果 GitHub Pages 地址打不开，常见原因是：

- 仓库还没有在 `Settings -> Pages` 开启发布。
- Pages 刚开启，仍在构建，通常等 1 到 10 分钟。
- 访问的是旧的本地地址，而不是 `github.io` 地址。
- 国内网络访问 `github.io` 不稳定。
- 浏览器 PWA 缓存还在使用旧版本，可以强制刷新或清理站点数据。

## 2. 项目整体流程

这个项目分成两层：

1. 数据处理和构建脚本  
   位于：

   ```text
   D:\暑假备课\scripts\
   ```

   这些脚本负责从 Excel 单词表提取词库、生成 HTML 页面、生成 PWA 图标。

2. 可部署网页文件  
   位于：

   ```text
   D:\暑假备课\outputs\word_cards\
   ```

   这里是最终推送到 GitHub Pages 的静态网站目录。

核心生成流程是：

```text
5 个 Excel 单词表
  -> scripts/extract_words.py
  -> outputs/word_cards/words.json
  -> scripts/build_word_cards_html.py
  -> outputs/word_cards/index.html
  -> GitHub Pages
```

## 3. 根目录原始表格文件

项目最开始读取了根目录下 5 个英语单词表：

```text
人教版七年级上册英语单词背诵默写表.xls
人教版七年级下册英语单词背诵默写表.xlsx
人教版八年级上册英语单词背诵默写表.xls
人教版八年级下册英语单词背诵默写表.xlsx
人教版九年级全一册英语单词背诵默写表.xls
```

它们不是网页运行时必须文件，只是生成词库的原始数据来源。网页线上运行时，词库已经被写入 `index.html`，不需要用户再下载 Excel。

## 4. scripts/extract_words.py

这个脚本负责读取 Excel 文件并生成词库。

输出文件：

```text
outputs/word_cards/words.json
outputs/word_cards/extract_summary.json
```

### 4.1 读取 `.xlsx`

`.xlsx` 本质是 zip 包，里面是 XML 文件。脚本用 Python 标准库读取：

```python
import zipfile
from xml.etree import ElementTree as ET
```

核心逻辑：

- 打开 `.xlsx` 压缩包。
- 读取 `xl/sharedStrings.xml`，拿到共享字符串表。
- 读取 `xl/workbook.xml` 和 workbook relationships，找到每个 sheet 的 XML 文件。
- 解析每个单元格，恢复成二维表格。

相关函数：

```python
read_xlsx(path)
_xlsx_shared_strings(zf)
_xlsx_sheets(zf)
```

### 4.2 读取 `.xls`

`.xls` 是旧版二进制 Excel，不是 zip，也不是普通文本。当前环境没有安装 `xlrd/openpyxl/pandas`，所以脚本自己实现了够用的读取器。

相关类和函数：

```python
class CfbFile
read_xls(path)
_biff_records(data)
_parse_xls_string(reader, ansi_encoding)
```

实现思路：

- `.xls` 是 OLE Compound File Binary 格式。
- `CfbFile` 负责读取 OLE 容器里的 `Workbook` 或 `Book` 数据流。
- `_biff_records` 按 BIFF 记录格式拆出每条记录。
- `read_xls` 读取 sheet 名、共享字符串表和单元格内容。
- `_decode_xls_8bit` 用 UTF-8、GBK、CP1252 等方式容错解码，避免中文乱码。

这部分代码比较底层，后续一般不需要改。除非新增的 `.xls` 文件结构差异很大，才需要调整。

### 4.3 提取英文和中文

核心函数：

```python
extract_pairs(workbook_name, sheets)
```

它逐行扫描单元格：

- 找到包含英文字母的单元格，认为可能是英文词条。
- 在右侧几个单元格里寻找含中文的释义。
- 过滤表头、音标、过长文本、含中文的英文候选。
- 保存英文、中文、教材来源、sheet 名和行号。

音标过滤靠：

```python
IPA_RE = re.compile(...)
```

这样 `[sɪŋ]` 这类音标不会被当成单词。

### 4.4 去重和生成 id

脚本把相同英文词条按小写去重：

```python
key = item["english"].lower()
```

并生成 URL/DOM 友好的 id：

```python
item["id"] = re.sub(r"[^a-z0-9]+", "-", key).strip("-")
```

最终得到约 2504 个去重词条。

## 5. scripts/build_word_cards_html.py

这个脚本是网页生成器。它读取 `words.json`，然后生成最终的 `index.html`。

核心代码：

```python
words = json.loads((OUT / "words.json").read_text(encoding="utf-8"))
payload = json.dumps(words, ensure_ascii=False).replace("</", "<\\/")
```

这里把词库 JSON 直接嵌入 HTML：

```javascript
const WORDS = [...];
```

这样部署时不依赖额外接口，也不需要线上读取 Excel。

### 5.1 PWA 头部配置

HTML 的 `<head>` 里有：

```html
<meta name="theme-color" content="#fbf8f4">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="单词卡">
<link rel="manifest" href="./manifest.webmanifest">
<link rel="apple-touch-icon" href="./icons/icon-192.png">
```

作用：

- 告诉浏览器这是一个可安装网页应用。
- 指定主题色。
- iPhone 添加到主屏幕时显示应用标题和图标。
- 关联 `manifest.webmanifest`。

### 5.2 页面结构

主要 HTML 结构：

```html
<div class="app">
  <header class="topbar">...</header>
  <main class="stage">
    <div class="toast" id="toast"></div>
    <div class="deck" id="deck"></div>
    <div class="side">...</div>
  </main>
</div>
```

含义：

- `.topbar`：顶部标题、进度、按钮。
- `.stage`：主学习区域。
- `#deck`：卡片舞台。
- `.side`：右侧操作按钮，类似短视频侧边按钮。
- `#toast`：临时提示，例如“已加入重点单词本”。

所有单词面板：

```html
<section class="panel" id="panel">
  ...
</section>
```

新手教程：

```html
<section class="tutorial" id="tutorial">
  ...
</section>
```

### 5.3 样式设计

主要风格：

- 简洁卡片风。
- 浅色背景。
- 8px 左右圆角。
- 卡片翻转动画。
- 顶部毛玻璃效果。
- 右侧圆形按钮。

卡片翻转靠 CSS 3D：

```css
.card {
  transform-style: preserve-3d;
  transition: transform .55s cubic-bezier(.2,.8,.2,1);
}

.card.flipped {
  transform: rotateY(180deg);
}

.face {
  backface-visibility: hidden;
}

.back {
  transform: rotateY(180deg);
}
```

点击卡片时给 `.card` 加或去掉 `.flipped`，就完成翻牌。

### 5.4 学习队列：每个词出现 5 遍

核心变量：

```javascript
const repeatTimes = 5;
let deck = [];
let current = 0;
```

构建队列：

```javascript
function buildDeck(random = false) {
  const base = [];
  for (let round = 1; round <= repeatTimes; round++) {
    for (const word of WORDS) base.push({...word, round});
  }
  deck = random ? shuffle(base) : base;
  current = 0;
  renderVirtualDeck();
  updateStats();
}
```

含义：

- `WORDS` 是去重后的原始词库。
- 外层循环 5 次。
- 每个词复制进学习队列 5 遍。
- 每条记录多一个 `round` 字段，用于显示第几轮。
- `random=true` 时会打乱顺序。

### 5.5 性能优化：虚拟渲染

最早版本的问题是一次性渲染全部学习队列：

```text
2504 个词 × 5 遍 = 12520 张卡片
```

这会导致手机和普通电脑滚动卡顿。

现在改成虚拟渲染：

```javascript
function renderVirtualDeck() {
  const slides = [];
  for (let offset = -2; offset <= 2; offset++) {
    const index = current + offset;
    if (index < 0 || index >= deck.length) continue;
    slides.push(slideMarkup(deck[index], index, offset));
  }
  els.deck.innerHTML = slides.join('');
  ...
}
```

含义：

- 真实学习队列还是 12520 条。
- DOM 里最多只放 5 张卡片。
- 当前卡片 `offset = 0`。
- 上一张、下一张 `offset = -1 / 1`。
- 再远一点的是 `offset = -2 / 2`。

这样浏览器只需要维护很少的 DOM，操作会流畅很多。

### 5.6 单张卡片模板

核心函数：

```javascript
function slideMarkup(word, index, offset) {
  const state = offset === 0 ? 'active' : Math.abs(offset) === 1 ? 'near' : 'far';
  return `
    <section class="slide ${state}" data-index="${index}" data-offset="${offset}">
      ...
    </section>`;
}
```

前面显示英文：

```html
<div class="word">${escapeHtml(word.english)}</div>
```

背面显示中文：

```html
<div class="meaning">${escapeHtml(word.chinese)}</div>
```

`escapeHtml` 用于防止特殊字符破坏 HTML：

```javascript
function escapeHtml(text) {
  return String(text).replace(/[&<>"']/g, ch => (...));
}
```

### 5.7 上下滑动和动画

当前卡片切换入口：

```javascript
function go(delta) {
  ...
}
```

逻辑：

- `delta > 0`：下一张。
- `delta < 0`：上一张。
- 如果正在动画中，直接忽略，避免连续触发导致状态错乱。
- 先移动当前 DOM 里的几张卡片。
- 260ms 后更新 `current`，重新渲染附近 5 张卡片。

动画位移由：

```javascript
transformForOffset(offset)
opacityForOffset(offset)
```

配合 CSS：

```css
.slide[data-offset="0"] { transform: translate3d(0, 0, 0) scale(1); }
.slide[data-offset="1"] { transform: translate3d(0, 104%, 0) scale(.94); }
.slide[data-offset="-1"] { transform: translate3d(0, -104%, 0) scale(.94); }
```

实现类似上下刷卡的感觉。

### 5.8 鼠标滚轮和触摸滑动

鼠标滚轮：

```javascript
els.deck.addEventListener('wheel', event => {
  event.preventDefault();
  if (wheelLock || Math.abs(event.deltaY) < 18) return;
  wheelLock = true;
  go(event.deltaY > 0 ? 1 : -1);
  setTimeout(() => wheelLock = false, 260);
}, { passive: false });
```

说明：

- 阻止页面默认滚动。
- 小幅滚轮变化忽略。
- `wheelLock` 防止滚轮一次触发很多张。

触摸滑动：

```javascript
touchstart -> 记录起始 Y
touchmove  -> 记录滑动距离
touchend   -> 超过 48px 才切换
```

核心：

```javascript
if (Math.abs(touchDeltaY) > 48) go(touchDeltaY < 0 ? 1 : -1);
```

向上滑，进入下一张；向下滑，回上一张。

### 5.9 键盘快捷键

键盘监听：

```javascript
document.addEventListener('keydown', event => {
  if (event.key === 'ArrowDown' || event.key === 'PageDown') go(1);
  if (event.key === 'ArrowUp' || event.key === 'PageUp') go(-1);
  if (event.key === 'Enter') {
    event.preventDefault();
    markCurrentForgot();
  }
  if (event.key === ' ') {
    event.preventDefault();
    const card = currentCard();
    if (card) card.classList.toggle('flipped');
  }
  if (event.key.toLowerCase() === 'f') toggleFav();
});
```

当前快捷键：

```text
Space       翻牌
ArrowDown   下一张
ArrowUp     上一张
Enter       记录忘记
F           加入/移出重点词
```

### 5.10 重点单词本

重点词存在浏览器本地：

```javascript
const store = {
  fav: new Set(JSON.parse(localStorage.getItem('wordcard:fav') || '[]')),
  ...
};
```

切换重点：

```javascript
function toggleFav() {
  const word = currentWord();
  if (store.fav.has(word.id)) {
    store.fav.delete(word.id);
  } else {
    store.fav.add(word.id);
  }
  persist();
  updateStats();
  renderWordList();
}
```

保存：

```javascript
localStorage.setItem('wordcard:fav', JSON.stringify([...store.fav]));
```

特点：

- 数据只保存在当前浏览器。
- 换设备不会同步。
- 清理浏览器数据后会丢失。

### 5.11 忘记记录

忘记次数也存在本地：

```javascript
forgot: JSON.parse(localStorage.getItem('wordcard:forgot') || '{}')
```

记录逻辑：

```javascript
function markForgot(id, card) {
  store.forgot[id] = (store.forgot[id] || 0) + 1;
  persist();
  card.classList.remove('forget');
  void card.offsetWidth;
  card.classList.add('forget');
  showToast('已记录：忘记了');
  renderWordList();
}
```

其中：

```javascript
void card.offsetWidth;
```

是为了强制浏览器重新计算布局，让同一个动画可以重复触发。

忘记动画：

```css
.forget {
  animation: forgetPulse .55s ease both;
}
```

### 5.12 所有单词面板

打开按钮：

```javascript
document.getElementById('allBtn').addEventListener('click', () => {
  els.panel.classList.add('open');
});
```

渲染列表：

```javascript
function renderWordList() {
  const q = els.search.value.trim().toLowerCase();
  const book = els.bookFilter.value;
  const items = WORDS.filter(word => {
    ...
  });
  els.wordList.innerHTML = items.map(word => `...`).join('');
}
```

支持：

- 搜索英文。
- 搜索中文。
- 搜索 Unit。
- 按教材筛选。
- 只看重点单词本。
- 在列表里直接加重点。

注意：这个列表目前仍然是一次性渲染匹配项。如果后面单词量继续变大，可以把这个列表也改成分页或虚拟列表。

### 5.13 新手教程

教程是否看过存在：

```javascript
localStorage.setItem('wordcard:tutorial', 'done');
```

如果已经看过：

```javascript
if (store.tutorial) els.tutorial.style.display = 'none';
```

顶部 `?` 按钮可以再次打开教程。

### 5.14 Service Worker 注册

页面底部：

```javascript
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  });
}
```

作用：

- 注册 `sw.js`。
- 让网页支持离线缓存。
- 让手机浏览器更像 App。

## 6. outputs/word_cards/index.html

这是最终网页文件，也是 GitHub Pages 主要访问的文件。

它包含：

- HTML 结构。
- CSS 样式。
- JavaScript 交互逻辑。
- 内嵌词库 `const WORDS = [...]`。

特点：

- 用户只访问这个文件就能运行。
- 不依赖后端服务器。
- 不依赖数据库。
- 不依赖 Excel 文件。

后续如果要改页面功能，不建议直接改 `index.html`，而是改：

```text
scripts/build_word_cards_html.py
```

然后重新运行：

```powershell
python scripts/build_word_cards_html.py
```

因为 `index.html` 是生成物。

## 7. outputs/word_cards/words.json

这是提取后的词库文件。

每个词条结构类似：

```json
{
  "english": "good",
  "chinese": "adj.好的",
  "book": "人教版七年级上册英语单词背诵默写表.xls",
  "sheet": "Starter Unit 1",
  "row": 3,
  "id": "good"
}
```

字段说明：

- `english`：英文单词或短语。
- `chinese`：中文释义。
- `book`：来源文件。
- `sheet`：来源工作表。
- `row`：来源行号。
- `id`：前端使用的稳定标识。

后续如果要加“按年级学习”“按 Unit 学习”，可以直接使用 `book` 和 `sheet` 字段。

## 8. outputs/word_cards/extract_summary.json

这是提取摘要，用于检查词库来源。

包含：

- 每个 Excel 文件读到了哪些 sheet。
- 每个文件提取了多少词。
- 每个文件的样例词条。

它不是网页运行必需文件，可以保留用于排查问题。

## 9. outputs/word_cards/manifest.webmanifest

这是 PWA 应用清单。

核心字段：

```json
{
  "name": "初中英语单词卡",
  "short_name": "单词卡",
  "start_url": "./",
  "scope": "./",
  "display": "standalone",
  "orientation": "portrait",
  "background_color": "#fbf8f4",
  "theme_color": "#fbf8f4"
}
```

含义：

- `name`：完整应用名。
- `short_name`：手机桌面短名称。
- `start_url`：从桌面图标打开时进入的地址。
- `scope`：PWA 控制范围。
- `display: standalone`：以类似 App 的方式打开，减少浏览器 UI。
- `orientation: portrait`：偏向竖屏使用。
- `icons`：安装时使用的图标。

## 10. outputs/word_cards/sw.js

这是 Service Worker，负责离线缓存。

当前缓存名：

```javascript
const CACHE_NAME = 'word-cards-v2';
```

缓存文件列表：

```javascript
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];
```

安装时缓存：

```javascript
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)));
  self.skipWaiting();
});
```

激活时删除旧缓存：

```javascript
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))
    )
  );
  self.clients.claim();
});
```

请求时优先读缓存：

```javascript
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});
```

注意：

- 如果你修改了 `index.html`，最好同步把 `CACHE_NAME` 升级，例如 `word-cards-v3`。
- 否则部分浏览器可能继续使用旧缓存。

## 11. outputs/word_cards/icons/

包含 PWA 图标：

```text
icons/icon-192.png
icons/icon-512.png
```

它们由：

```text
scripts/make_pwa_icons.py
```

生成。

## 12. scripts/make_pwa_icons.py

这个脚本用 Python 标准库手动生成 PNG。

没有依赖 Pillow 等图片库。

核心思路：

- 用 RGBA 像素数组画背景、卡片、线条和圆点。
- 用 zlib 压缩 PNG 数据。
- 手动写 PNG chunk。

相关函数：

```python
chunk(kind, data)
write_png(path, size)
```

正常情况下不需要改。要换图标风格时可以改 `write_png`。

## 13. outputs/word_cards/README.md

这是 GitHub 仓库首页说明。

包含：

- 项目名称。
- 本地预览命令。
- GitHub Pages 设置方法。
- 发布地址。

它适合给别人快速了解仓库，不承担技术细节说明。技术细节主要看本文件 `PROJECT_GUIDE.md`。

## 14. Git 和部署

部署仓库位于：

```text
D:\暑假备课\outputs\word_cards\
```

远程仓库：

```text
https://github.com/ICEXUAN-2707/Word-Game.git
```

常用命令：

```powershell
cd D:\暑假备课\outputs\word_cards
git status
git add index.html sw.js manifest.webmanifest README.md PROJECT_GUIDE.md
git commit -m "你的提交说明"
git -c http.version=HTTP/1.1 push
```

这里使用：

```powershell
git -c http.version=HTTP/1.1 push
```

是因为之前 GitHub HTTPS 推送偶尔出现 TLS 中断，HTTP/1.1 模式更稳定。

## 15. 想加新功能时怎么改

### 15.1 加一个新的快捷键

修改：

```text
scripts/build_word_cards_html.py
```

找到：

```javascript
document.addEventListener('keydown', event => {
  ...
});
```

增加分支即可，例如按 `R` 打乱：

```javascript
if (event.key.toLowerCase() === 'r') {
  buildDeck(true);
}
```

然后重新生成：

```powershell
python scripts/build_word_cards_html.py
```

### 15.2 加“只背重点词”

思路：

- 新增一个按钮。
- 构建队列时选择 `WORDS.filter(word => store.fav.has(word.id))`。
- 如果重点词为空，显示 toast 提示。

适合改的位置：

```javascript
buildDeck(random = false)
```

可以抽成：

```javascript
let mode = 'all'; // all 或 fav
```

### 15.3 加“按年级 / Unit 选择”

已有数据字段：

```javascript
word.book
word.sheet
```

可以基于它们生成筛选器。

构建队列前先过滤：

```javascript
const source = WORDS.filter(word => word.book === selectedBook && word.sheet === selectedUnit);
```

### 15.4 加“掌握 / 未掌握统计”

目前已有：

```javascript
store.forgot
store.fav
```

可以新增：

```javascript
store.known
```

并保存：

```javascript
localStorage.setItem('wordcard:known', JSON.stringify([...store.known]));
```

### 15.5 加“导出学习记录”

本地记录都在 `localStorage`：

```text
wordcard:fav
wordcard:forgot
wordcard:tutorial
```

可以新增按钮，把这些数据打包成 JSON 下载。

### 15.6 加“云同步”

当前项目是纯静态页面，没有账号系统和服务器。

如果要云同步，需要新增后端或第三方服务，例如：

- Supabase
- Firebase
- 自建 API
- GitHub Gist

这会让项目复杂度明显上升，建议先把本地功能做好。

## 16. 常见问题

### 16.1 修改后网页没变化

可能原因：

- 忘记运行 `python scripts/build_word_cards_html.py`。
- 忘记提交并推送。
- GitHub Pages 还没更新。
- PWA 缓存没刷新。

处理：

1. 修改 `scripts/build_word_cards_html.py`。
2. 运行构建脚本。
3. 升级 `sw.js` 的 `CACHE_NAME`。
4. 提交并推送。
5. 浏览器强制刷新或清理站点缓存。

### 16.2 手机打开不是 App 样式

需要：

- 使用 HTTPS 地址。
- `manifest.webmanifest` 可访问。
- `sw.js` 注册成功。
- 浏览器支持 PWA 安装。

iPhone 需要使用 Safari 的“添加到主屏幕”。

### 16.3 国内访问慢

GitHub Pages 在国内网络不稳定。

可选方案：

- 继续用 GitHub Pages，适合小范围分享。
- 部署到 Vercel，全球 CDN 体验可能更好，但国内也不保证。
- 部署到国内静态托管或对象存储，稳定性更好。

## 17. 当前功能清单

已实现：

- 从 5 个 Excel 表格提取词库。
- 英文卡片展示。
- 点击或空格翻牌显示中文。
- 每个词在学习队列中出现 5 遍。
- 上下滑动切换卡片。
- 鼠标滚轮切换卡片。
- 方向键切换卡片。
- 回车记录忘记。
- 双击背面记录忘记。
- 忘记动画和 toast 提示。
- 加入重点单词本。
- 查看所有单词。
- 搜索单词。
- 按教材筛选。
- 只看重点词。
- 新手教程。
- PWA 安装。
- 离线缓存。
- 虚拟渲染优化。

## 18. 推荐开发习惯

以后加功能建议按这个顺序：

1. 先改 `scripts/build_word_cards_html.py`。
2. 运行：

   ```powershell
   python scripts/build_word_cards_html.py
   ```

3. 本地预览：

   ```powershell
   cd D:\暑假备课\outputs\word_cards
   python -m http.server 8768 --bind 127.0.0.1
   ```

4. 打开：

   ```text
   http://127.0.0.1:8768/
   ```

5. 验证没问题后提交推送。

## 19. 最重要的维护原则

`index.html` 是生成结果，`scripts/build_word_cards_html.py` 是源头。

如果只改 `index.html`，下次运行构建脚本时改动会丢失。

所以功能开发应优先改：

```text
scripts/build_word_cards_html.py
```

数据问题优先改：

```text
scripts/extract_words.py
```

图标问题优先改：

```text
scripts/make_pwa_icons.py
```
