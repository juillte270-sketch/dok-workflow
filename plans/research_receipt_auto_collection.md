# Research: Receipt Image Auto-Collection from Suppliers

**Date**: 2026-03-02
**Status**: Research Complete
**Context**: ~15 suppliers send receipts (POS photos, ERP URLs, handwritten notes) via KakaoTalk personal chat. Need to auto-process into price sheet (Excel).

---

## Current System

Already built:
- `execution/parse_receipt.py` — Gemini Vision OCR for images, HTML scraping for URLs, text parsing
- `execution/fill_receipt_prices.py` — writes parsed prices to Excel master file Col I
- `execution/receipt_watcher.py` — watchdog daemon that monitors KakaoTalk download folder + clipboard
- `execution/kakao_order_server.py` — Flask server with `/kakao/receipt` endpoint for chatbot image/URL/text receipt processing

The **gap**: receipts arrive in KakaoTalk personal chats. The user must manually click each image to download it, or copy each URL. The `receipt_watcher.py` daemon watches the download folder, but KakaoTalk PC has **no auto-download setting for images**.

---

## Method-by-Method Analysis

### 1. KakaoTalk PC Auto-Download Setting

**Finding**: KakaoTalk PC does **NOT** have an auto-download images setting.

- Images in chat are shown as thumbnails/previews only
- To save an image, users must: click to view full-size -> right-click -> save, OR use the batch download feature in the photo gallery drawer
- The batch download (일괄저장) allows selecting up to 1000 images at once per month since late 2023, but still requires manual initiation
- Download folder: `%USERPROFILE%\Documents\카카오톡 받은 파일\` (confirmed exists on this machine)
- Download folder is configurable in Settings -> Chat -> Download folder

**Verdict**: Not viable as fully automatic. But combined with `receipt_watcher.py`, the user workflow is: batch-select today's receipt images -> download all -> watcher picks them up. This is **semi-auto** (1 manual batch per day).

**Supplier friction**: Zero (they change nothing)
**Reliability**: High (manual trigger, deterministic download)
**Cost**: Free
**Implementation**: Already built (`receipt_watcher.py`)
**Automation level**: Semi-auto (1 batch action per day, ~30 seconds)

---

### 2. KakaoTalk PC Local File Caching (AppData)

**Finding**: KakaoTalk PC caches chat images locally, but they are **encrypted**.

Confirmed on this machine:
```
%LOCALAPPDATA%\Kakao\KakaoTalk\users\{user_hash}\chat_data\cli\
  -> 7,495 .cng files (encrypted chat image cache)
  -> cli/thumbnail/ -> 18,715 thumbnail files (also encrypted)
```

The `.cng` files are **not simple images** -- file headers show varying byte patterns (not JPEG `FF D8 FF` or PNG `89 50 4E 47`), suggesting proprietary encryption. Each file has a different first-byte pattern, ruling out simple XOR with a constant key.

Other cache directories:
- `cli_http_v2/` — 7 files (URL link preview images, also .cng)
- `url_image_v2/` — URL image cache
- `oci_v2/`, `ocii_v2/`, `mci_v2/` — all empty

**Verdict**: Not viable. The encryption is proprietary and undocumented. Reverse-engineering it would be fragile (breaks on updates) and potentially violates KakaoTalk ToS.

**Supplier friction**: Zero
**Reliability**: Very low (encrypted, undocumented, breaks on updates)
**Cost**: Free
**Implementation**: Very high complexity (reverse engineering encryption)
**Automation level**: Would be fully auto if decryption worked, but impractical

---

### 3. SMS/MMS Gateway Services in Korea

**Finding**: Korean SMS/MMS services are **send-only**. None offer inbound MMS image receiving via webhook.

Researched providers:
| Provider | Send SMS | Send MMS | Receive SMS | Receive MMS+Image | Webhook |
|----------|----------|----------|-------------|-------------------|---------|
| Solapi (솔라피) | 13원/건 | 60원/건 | X | X | X |
| CoolSMS | Similar | Similar | X | X | X |
| NHN Cloud | Yes | Yes | X | X | X |
| Aligo (알리고) | Yes | Yes | X | X | X |
| Ppurio (뿌리오) | Yes | Yes | X | X | X |

Korean SMS APIs are designed for **outbound notifications** (send alimtok, send marketing SMS). They do not support:
- Renting a virtual number that receives messages
- Inbound MMS with image attachment forwarding
- Webhook callbacks for incoming messages

**International alternatives** (Twilio, Telnyx, Telerivet):
- These DO support receiving MMS with images via webhook
- But they don't provide **Korean mobile numbers** for MMS receiving
- Korean telco regulations restrict virtual number issuance
- Cost: ~$1/month for US number + $0.01-0.02/received MMS — but suppliers can't send MMS to a US number practically

**Verdict**: Not viable for this use case in Korea. The domestic SMS ecosystem is send-only, and international services can't provide Korean numbers for MMS receiving.

**Supplier friction**: Very high (would need to send MMS to a foreign number)
**Reliability**: N/A
**Cost**: N/A
**Implementation**: Not possible with Korean numbers
**Automation level**: N/A

---

### 4. Email Approach (Gmail API)

**Concept**: Ask suppliers to email receipt photos; monitor Gmail via API.

**Technical feasibility**: Easy. Gmail API with OAuth can watch for new emails, download attachments, and trigger the parse pipeline.

**Practical reality**:
- These are 가락시장 상인들 (market vendors), mostly 50-60+ years old
- They are comfortable with KakaoTalk, not email
- Asking them to open a camera app, take photo, open email app, compose email, attach photo, send — this is a 6-step process vs their current 2-step (open chat, send photo)
- Many may not have email set up on their phone
- Some send ERP URLs (itanet, marketbom) which are easier to forward via chat

**Verdict**: Technically trivial, practically impossible due to supplier demographics.

**Supplier friction**: Extremely high
**Reliability**: High if they use it
**Cost**: Free (Gmail API)
**Implementation**: Low
**Automation level**: Fully auto once email arrives

---

### 5. Android Phone Automation (Tasker / MacroDroid)

**Concept**: Run Tasker/MacroDroid on a dedicated Android phone logged into KakaoTalk. When a notification arrives from a supplier chat, automatically save the image and upload to server.

**Technical details**:
- Tasker + AutoNotification can intercept KakaoTalk notifications
- The notification includes sender name and message preview text
- For images: notification just says "사진" (photo), not the actual image data
- To actually get the image, you'd need Accessibility Service to open the chat, long-press the image, save it
- Or use KakaoTalk's auto-save to gallery feature (mobile only) + monitor the gallery folder

**KakaoTalk Mobile auto-save to gallery**:
- Settings -> Chat -> Save to Gallery (사진 자동 저장) — exists on mobile!
- When enabled, ALL images from ALL chats are auto-saved to the phone's gallery
- Path: `/storage/emulated/0/KakaoTalk/` or similar
- Can then use Tasker/MacroDroid to monitor this folder and upload new images

**Architecture**:
1. Dedicated old Android phone, always on, plugged in
2. KakaoTalk logged in (same account or a dedicated account?)
3. "Save to Gallery" enabled -> all images auto-saved
4. Tasker/Automate watches the KakaoTalk media folder
5. New image detected -> HTTP POST to `receipt_watcher.py` or local REST API
6. Problem: ALL images from ALL chats are saved (personal photos, memes, etc.)

**Challenges**:
- Same KakaoTalk account can't be logged in on two phones simultaneously (KakaoTalk limitation)
- Would need a separate phone number/account dedicated to supplier communication
- Migrating suppliers to a new number would cause friction
- Filtering supplier receipts from personal images requires OCR or sender identification
- Phone needs to stay on 24/7, keep-alive for KakaoTalk notifications

**Verdict**: Partially viable but requires dedicated hardware and account management. The "save all images" approach creates noise that needs filtering.

**Supplier friction**: Zero (if same number) or moderate (if new number)
**Reliability**: Medium (phone can crash, KakaoTalk can disconnect)
**Cost**: ~50,000-100,000원 for old Android phone
**Implementation**: Medium
**Automation level**: Fully auto but noisy (needs image filtering)

---

### 6. KakaoTalk Business Channel Chatbot (카카오 비즈채널)

**Current state**: `kakao_order_server.py` already handles this!

**How it works now**:
- KakaoTalk OpenBuilder chatbot with `/kakao/receipt` skill endpoint
- When user sends image to the **business channel**, the image CDN URL is included in `userRequest.utterance`
- The `_extract_image_url()` function detects `kakaocdn.net` URLs and `IMAGE_UPLOAD` trigger type
- Server downloads image, runs Gemini Vision OCR, fills Excel

**The core problem**: Suppliers send receipts to **personal chat**, not to the business channel.

**What it would take**:
- Suppliers would need to: Open KakaoTalk -> Search for business channel -> Open it -> Send photo there instead of personal chat
- This is a behavior change: they're used to sending to their personal contact (the user's personal KakaoTalk)
- Could create a dedicated "영수증 접수" channel
- Channel creation: Free (카카오 비즈채널 개설 무료)
- Chatbot: Free (카카오 챗봇 서비스 무료화 announced)
- No per-message cost for receiving (only sending alimtok/friendtok has cost)

**Important limitation**: When a user sends multiple images at once, only the first image URL is included in the webhook payload. Suppliers would need to send images one at a time or the system would miss images.

**Verdict**: Technically already implemented. The only gap is getting suppliers to use the channel instead of personal chat. This is a **training/adoption problem**, not a technical one.

**Supplier friction**: Moderate (need to learn to use business channel)
**Reliability**: Very high (KakaoTalk's own infrastructure)
**Cost**: Free (channel + chatbot free)
**Implementation**: Already done
**Automation level**: Fully automatic

---

### 7. Simple Web Upload Form

**Concept**: A lightweight web page where suppliers tap a link -> camera opens -> take photo -> auto-upload.

**Technical approach**:
```html
<input type="file" accept="image/*" capture="environment">
```
This single HTML element on mobile opens the camera directly. Combined with JavaScript, the image is auto-uploaded on selection.

**Architecture**:
1. Static web page hosted on: Cloudflare Pages (free), Vercel (free), or local network
2. Supplier opens a bookmarked URL or taps a link saved on home screen
3. Camera opens -> take photo -> image auto-uploads to server
4. Server: Flask endpoint that saves image, calls `parse_receipt.py`, fills Excel
5. Can include supplier selection dropdown (or auto-detect from saved bookmark URL with supplier param)

**Supplier experience** (2 taps):
1. Tap bookmark/link on phone home screen
2. Camera opens, take photo, tap "OK"
3. Done (auto-upload happens)

**Advantages**:
- No app installation required
- Works on any phone with a browser
- Can add supplier identification via unique URL per supplier (e.g., `upload.example.com/?s=건영농산`)
- Can show confirmation + preview after upload
- Can be made as a PWA (Progressive Web App) that works offline
- No account/login needed

**Challenges**:
- Requires internet hosting (either local network when suppliers are at 가락시장, or cloud)
- Suppliers need to save a bookmark (one-time setup)
- Not as natural as "send photo in chat" workflow
- Need to convince suppliers to use a different workflow

**Verdict**: Very clean solution technically. Low friction if the bookmark can be set up once. Could be the **best alternative** if business channel adoption fails.

**Supplier friction**: Low-moderate (one-time bookmark setup, then 2 taps)
**Reliability**: Very high
**Cost**: Free (Cloudflare Pages/Vercel) or minimal
**Implementation**: Low (simple Flask + HTML)
**Automation level**: Fully automatic

---

### 8. Google Forms with Image Upload

**Finding**: Google Forms supports file upload, but **requires Google account login**.

- Adding a "File upload" question type is straightforward
- However, respondents MUST be logged into a Google account to upload files
- This is a hard requirement from Google (files go to the form creator's Drive)
- Market vendors (가락시장 상인) very likely do not have Google accounts on their phones

**Workaround**: FormFacade addon converts Google Forms file upload to standard HTML upload without login. But this adds complexity and the form still looks like "Google Forms" which may confuse suppliers.

**Verdict**: The Google account requirement is a dealbreaker. The web upload form (Method 7) is strictly better.

**Supplier friction**: Very high (Google account required)
**Reliability**: High
**Cost**: Free
**Implementation**: Low
**Automation level**: Fully auto (Apps Script can trigger on submission)

---

### 9. LINE / Other Messaging Platforms

**LINE Messaging API**:
- LINE Official Account: Free to create, free Communication plan allows 500 messages/month
- Webhook receives image messages with `type: "image"` and `messageId`
- Call `GET /v2/bot/message/{messageId}/content` to download binary image data
- Content auto-deleted after a period, must download promptly
- Fully documented, well-supported API

**Problem**: LINE is barely used in Korea. KakaoTalk has ~95% market share. Asking 가락시장 상인 to install and use LINE is essentially impossible.

**Naver Works (LINE WORKS)**:
- Business messaging with Bot API
- Supports image message receiving via callback
- But requires organizational account (not for external suppliers)

**Telegram**:
- Excellent Bot API with image receiving
- But even less popular than LINE in Korea

**Verdict**: Not viable. The messaging platform must be KakaoTalk in Korea. Any alternative would fail on adoption.

**Supplier friction**: Extreme (install new app, create account)
**Reliability**: High (LINE API is excellent)
**Cost**: Free
**Implementation**: Low
**Automation level**: Fully auto

---

### 10. Windows Clipboard/Screen Monitoring

**Concept**: More aggressive monitoring of the KakaoTalk PC window directly.

**Approaches**:
a) **Clipboard monitoring** (already in `receipt_watcher.py`): When user copies text/URL from KakaoTalk, detect and process. Works for ERP URLs (건영농산 itanet, 오복상회 marketbom) and text receipts.

b) **Screen capture automation**: Use pyautogui or similar to:
   - Detect KakaoTalk window
   - Identify chat rooms with supplier names
   - Scroll through messages
   - Screenshot images
   - OCR the screenshots

c) **Win32 API window monitoring**: Hook into KakaoTalk's window messages to detect when new messages arrive.

**Reality check**:
- Screen capture is fragile (UI layout changes break it)
- KakaoTalk has anti-automation measures
- Performance: scanning each supplier chat one by one is slow
- Cannot reliably distinguish receipt images from other images
- Would need to open each of 15 supplier chats periodically

**Verdict**: Too fragile and slow. The clipboard monitoring in `receipt_watcher.py` is already the practical maximum for this approach.

**Supplier friction**: Zero
**Reliability**: Very low (fragile, breaks on UI updates)
**Cost**: Free
**Implementation**: Very high
**Automation level**: Fully auto but unreliable

---

### 11. KakaoTalk PC File System Hooks

**Finding**: KakaoTalk stores chat data in encrypted `.cng` files and `.edb` (ESE database) files.

Confirmed structure:
```
%LOCALAPPDATA%\Kakao\KakaoTalk\users\{hash}\chat_data\
  chatLogs_{chat_id}.edb     <- ESE database with chat messages
  cli\cli_{hash}.cng         <- encrypted image cache (7,495 files)
  cli\thumbnail\             <- encrypted thumbnails (18,715 files)
  cli_http_v2\               <- encrypted link preview images
```

The `.edb` files are Extensible Storage Engine (ESE/JET Blue) databases. These CAN be read with Python `pyesedb` or `libesedb`. The chat logs might contain image URLs or references.

**Potential approach**:
1. Monitor `.edb` files for new entries (file modification time)
2. Parse chat messages to find image references
3. Download original images from KakaoTalk CDN using the URLs in the database

**Challenges**:
- ESE database format is complex but readable
- Chat message format is proprietary (may contain serialized data)
- Image URLs in the database may require authentication/session tokens
- KakaoTalk CDN URLs have expiration (images expire after ~2 weeks)
- Reverse-engineering the database schema is required
- Violates KakaoTalk ToS

**Verdict**: Theoretically possible but requires significant reverse engineering. The `.edb` chat databases are readable, but extracting image URLs and downloading them with proper authentication is complex. Not recommended.

**Supplier friction**: Zero
**Reliability**: Low (undocumented, breaks on updates)
**Cost**: Free
**Implementation**: Very high
**Automation level**: Fully auto if it works

---

## Comparative Summary

| # | Method | Supplier Friction | Reliability | Cost | Implementation | Automation |
|---|--------|-------------------|-------------|------|----------------|------------|
| 1 | KakaoTalk PC batch download + watcher | **Zero** | High | Free | **Already built** | Semi-auto |
| 2 | KakaoTalk PC cache decryption | Zero | Very low | Free | Very high | N/A |
| 3 | SMS/MMS gateway (Korea) | N/A | N/A | N/A | **Not possible** | N/A |
| 4 | Email (Gmail API) | Extreme | High | Free | Low | Full |
| 5 | Android phone + Tasker | Zero~Moderate | Medium | 50-100K원 | Medium | Full (noisy) |
| 6 | KakaoTalk Business Channel | **Moderate** | **Very high** | **Free** | **Already built** | **Full** |
| 7 | Web upload form | Low-Moderate | Very high | Free | Low | Full |
| 8 | Google Forms | Very high | High | Free | Low | Full |
| 9 | LINE / other platforms | Extreme | High | Free | Low | Full |
| 10 | Screen/clipboard monitoring | Zero | Very low | Free | Very high | Full |
| 11 | KakaoTalk file system hooks | Zero | Low | Free | Very high | Full |

---

## Recommended Strategy: Hybrid Approach

### Tier 1: Immediate (no supplier change) -- TODAY
**Keep using `receipt_watcher.py` as-is.**
- User opens KakaoTalk PC each morning
- Goes to each supplier chat, selects today's receipt images
- Uses batch download (일괄저장) to download all at once
- `receipt_watcher.py` detects new images in download folder, auto-processes
- Clipboard monitoring handles ERP URLs (건영농산, 오복상회)
- **Effort: ~2-3 minutes per day for ~15 suppliers**

### Tier 2: Short-term (gradual migration) -- 1-2 WEEKS
**Deploy the KakaoTalk Business Channel for receipt collection.**
- Create a dedicated "도크 영수증" business channel
- The `/kakao/receipt` endpoint in `kakao_order_server.py` already handles everything
- Onboard 2-3 tech-savvy suppliers first (건영농산, 오복상회 who already use ERP)
- Provide a simple guide: "이 채널에 영수증 사진/URL 보내주세요"
- As suppliers adopt, the manual batch-download workload decreases
- **Key message to suppliers**: "이 채널로 보내시면 바로 자동처리됩니다"

### Tier 3: Medium-term fallback -- IF TIER 2 ADOPTION IS LOW
**Build a simple web upload form as alternative.**
- Single-page PWA: supplier taps link -> camera -> auto-upload
- Unique URL per supplier (e.g., `receipt.example.com/건영`)
- Can be saved as home screen shortcut
- For suppliers who resist the business channel but are willing to try something new
- Host on free tier (Cloudflare Pages + serverless function to call receipt API)

### Why NOT the other methods:
- **SMS/MMS**: Not possible in Korea (send-only ecosystem)
- **Email**: Demographics make this impossible
- **LINE/other**: Nobody uses it in Korea
- **Android automation**: Too complex for marginal benefit, reliability issues
- **File system hooks**: Reverse engineering, fragile, ToS violation
- **Screen automation**: Fragile, unreliable

---

## Implementation Notes for Tier 2 (Business Channel)

### Already working:
- `kakao_order_server.py` `/kakao/receipt` endpoint
- Image CDN URL extraction from `userRequest.utterance`
- Background Gemini Vision OCR
- Receipt parsing + Excel filling
- Operator notification via KakaoTalk self-message

### Needed:
1. Create "도크 영수증" business channel on Kakao Business (business.kakao.com)
2. Set up OpenBuilder chatbot with receipt skill block
3. Connect skill block to existing `/kakao/receipt` endpoint
4. Write supplier onboarding guide (1-page with screenshots)
5. Handle multi-image limitation (only first image in single send -- document "한 장씩 보내주세요")

### Needed for Tier 3 (Web Form):
1. Simple Flask route or static HTML + serverless function
2. `<input type="file" accept="image/*" capture="environment">` for direct camera
3. Supplier dropdown or URL-based identification
4. POST to existing `POST /receipt/fill` REST API
5. Confirmation page after upload

---

## Sources

- [KakaoTalk PC download folder location](https://cs.kakao.com/helps_html/1073181649?locale=ko)
- [KakaoTalk PC download folder settings](https://itons.net/%EC%B9%B4%EC%B9%B4%EC%98%A4%ED%86%A1-pc%EB%B2%84%EC%A0%84-%EB%8B%A4%EC%9A%B4%EB%A1%9C%EB%93%9C-%ED%8F%B4%EB%8D%94-%EC%9C%84%EC%B9%98-%EB%B3%80%EA%B2%BD%ED%95%98%EA%B8%B0/)
- [KakaoTalk PC auto-save not available (DevTalk)](https://devtalk.kakao.com/t/pc/43138)
- [KakaoTalk PC cache structure](https://geekorea.com/how-to-clean-up-kakaotalk-pc-storage/)
- [KakaoTalk PC cache & data paths](https://jab-guyver.co.kr/2169)
- [Solapi pricing](https://solapi.com/pricing)
- [Solapi standard pricing table](https://guide.solapi.com/pricing)
- [NHN Cloud SMS API](https://docs.nhncloud.com/ko/Notification/SMS/ko/api-guide/)
- [Aligo SMS API](https://smartsms.aligo.in/admin/api/spec.html)
- [KakaoTalk chatbot image handling (DevTalk)](https://devtalk.kakao.com/t/topic/143494)
- [KakaoTalk channel webhook](https://developers.kakao.com/docs/latest/ko/kakaotalk-channel/callback)
- [Kakao Business chatbot skill guide](https://kakaobusiness.gitbook.io/main/tool/chatbot/skill_guide/make_skill)
- [KakaoTalk channel free chatbot announcement](https://www.kakaocorp.com/page/detail/9756)
- [Kakao consult talk pricing (100원/상담방/일)](https://kakaobusiness.gitbook.io/main/ad/cstalk)
- [Happytalk pricing](https://happytalk.io/price)
- [Happytalk webhook for consult talk](https://developer-center.happytalk.io/KakaoWebhook/)
- [LINE Messaging API - receiving messages](https://developers.line.biz/en/docs/messaging-api/receiving-messages/)
- [LINE Messaging API pricing](https://developers.line.biz/en/docs/messaging-api/pricing/)
- [Google Forms file upload requires login](https://www.jotform.com/google-forms/how-to-add-upload-button-in-google-form/)
- [FormFacade: upload without Google login](https://formfacade.com/file-upload/google-forms-upload-file-without-google-account.html)
- [PWA camera/image upload](https://daviddalbusco.medium.com/take-photo-and-access-the-picture-library-in-your-pwa-without-plugins-876dc92989b)
- [Python watchdog library](https://pypi.org/project/watchdog/)
- [Tasker KakaoTalk notification forwarding (Clien)](https://www.clien.net/service/board/cm_andro/12777552)
- [MacroDroid notification forwarding](https://surplstimes.com/1733/macrodroid-forward-andorid-notification-to-sms/)
- [Twilio Korea SMS pricing](https://www.twilio.com/en-us/sms/pricing/kr)
