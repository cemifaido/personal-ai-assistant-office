# Personal AI Assistant

> 🇬🇧 **English** | 🇮🇹 **[Italiano](README.md)**

**Native AI assistant integrated directly inside Microsoft Word, Excel, PowerPoint, and classic Outlook for Windows** — **beta 0.2.1**

## What it is and how it works

**Personal AI Assistant** is an open-source add-in that brings AI **directly into the Microsoft Office ribbon** (*Personal AI* tab), eliminating copy-paste between Office and external chats and using **zero intermediary servers**.

**How it works in 3 steps:**
1. **Work inside your Office document or inbox** (Word, Excel, PowerPoint, or classic Outlook) and select the text, cells, slide, or emails you want to process.
2. **Click an action in the *Personal AI* task pane** (rewrite, proofread, translate, summarize, generate formulas/charts, create slides, or draft replies).
3. **Your PC calls your chosen AI provider directly** (using your own personal API key, or a local OpenAI-compatible model such as Ollama / `llama-server` / LM Studio) and presents the result for you to review before applying it or keeping it as a draft.

> [!IMPORTANT]
> **Requires your own personal API key (BYOK model).** Personal AI Assistant is not model-free and requires at least one API key from **Google Gemini, OpenAI, Anthropic Claude, Qwen, MiniMax, DeepSeek, or OpenRouter** (classic Outlook also supports local models without an API key; experimental in the Office task pane).
> No API key is bundled with the installer: you obtain it from your provider under **your own account and terms** (free tier or pay-as-you-go). The easiest free way to start is a [Google AI Studio](https://aistudio.google.com/apikey) key.

## Features & Connectors

- **Word**: rewrite with a chosen tone, proofread, summarize, translate, generate text, and ask questions about the entire document.
- **Excel**: generate formulas from natural language, analyze selected data (with exact statistics), create charts and sample tables.
- **PowerPoint**: rewrite, proofread, shorten, or translate selected text, generate slides from an outline, and improve existing slide text.
- **Classic Outlook**: summarize multiple selected emails and create reply drafts (confirm, decline, ask for clarification, or custom), including POP and IMAP accounts. Never sends emails automatically.

**Architecture & connectors:**
- **Office integration:** **Office.js** web task pane (running inside Office WebView2 via a local XML manifest + static GitHub Pages) for **Word, Excel, and PowerPoint**; native **VSTO (.NET Framework 4.8)** add-in on the *Home* ribbon for **classic Outlook**. Does not handle PDF/Acrobat files.
- **AI integration:** direct **HTTPS REST calls** from your PC to **Google Gemini, OpenAI, Anthropic Claude, Qwen (Alibaba), MiniMax, DeepSeek, or OpenRouter**, or HTTP/HTTPS calls to a **local OpenAI-compatible server** (**Ollama, `llama-server`, LM Studio**). No project backend server exists; keys stay on your PC.

## Comparison with `mcp4office`

| Aspect | Personal AI Assistant (this project) | [`mcp4office`](https://github.com/pai-office-automation/mcp4office) |
| --- | --- | --- |
| **Where you use it** | **Inside Office**, via the *Personal AI* ribbon/pane | **Outside Office**, from any MCP-compatible AI chat client (Antigravity, Claude Desktop/Code, Codex, etc.) |
| **AI model & billing** | **BYOK direct APIs** (pay-as-you-go or free tier) or local model (Ollama / `llama-server` / LM Studio) | **Model-free**: no API keys needed; uses your MCP client's flat subscription |
| **AI connector** | Direct **HTTPS REST** to cloud providers or local OpenAI-compatible endpoint | **MCP (JSON-RPC 2.0 over `stdio`)** |
| **Office connector** | **Office.js** (WebView2) for Word/Excel/PowerPoint + **VSTO (.NET 4.8)** for classic Outlook (no PDF/Acrobat) | Local **Windows COM** on open Word, Excel, PowerPoint, classic Outlook, Acrobat + direct **local PDF** reading |
| **Best for** | Fast, button-driven actions inside Office without leaving your document | Agentic, cross-app workflows (document/sheet diffing, PDF reading, multi-table edits) |
| **Installer privileges** | Per-user for Outlook; **1 UAC prompt** on first install to create the read-only shared manifest folder for Word/Excel/PowerPoint | 100% per-user, **no admin / UAC required** |

## Installation

Follow the **[illustrated guide (IT)](https://pai-office-automation.github.io/personal-ai-assistant-office/installa.html)**. In short:

1. Download `PersonalAIAssistant-<version>-Setup.exe` from [Releases](../../releases/latest).
2. The installer is currently unsigned: right-click → *Properties* → **Unblock**; if Windows SmartScreen appears → *More info* → **Run anyway**.
3. Run the installer. For Word, Excel, and PowerPoint, Windows asks for one administrator (UAC) confirmation on first install to create a local read-only shared folder for the manifest.
4. Once in Word, Excel, and PowerPoint: *Home → Add-ins → More Add-ins → Shared Folder → Personal AI Assistant → Add*. The **Personal AI** tab will appear.
5. Enter your personal API key on first launch.

Verify download integrity:

```powershell
Get-FileHash .\PersonalAIAssistant-0.2.1-Setup.exe -Algorithm SHA256
```

**Requirements:** Windows 10/11, Microsoft Office desktop for Windows, .NET Framework 4.8 (for Outlook).
**Not supported:** New Outlook for Windows, Outlook on the web, Office for Mac.

## Privacy

Read the **[privacy policy (IT)](https://pai-office-automation.github.io/personal-ai-assistant-office/privacy.html)**:
- Selected text, documents (only when requested), and selected emails are sent directly to your chosen provider under **its** terms. Check your provider's plan terms before processing confidential data.
- **Qwen (Alibaba), MiniMax, and DeepSeek are based in China**: the add-in asks for explicit confirmation before saving those keys. With a local model, data stays on your machine/network.
- Zero telemetry, zero tracking, zero automatic email sending (Outlook always opens a draft).

## Known Beta Limitations

- AI output can contain mistakes: always review before keeping or sending.
- On locked-down corporate PCs without admin rights, Word/Excel manifest setup may require IT; Outlook installs per-user anyway.
- Tested on Windows 11 + Office LTSC 2024 32-bit (admin & standard user); 64-bit / Microsoft 365 / Office 2016–2021 are not yet tested.

## Development & License

- Architecture, build, and release docs (IT): [docs/SVILUPPO.md](docs/SVILUPPO.md) · [CHANGELOG.md](CHANGELOG.md).
- Licensed under the **[MIT License](LICENSE)**. Independent open-source project, not affiliated with Microsoft, Google, OpenAI, or Anthropic.
