metadata
title: Text2PPTX
emoji: 📊
colorFrom: indigo
colorTo: blue
sdk: streamlit
sdk_version: 1.48.1
app_file: app.py
pinned: false
Text → PowerPoint (Template-Aware)
Public Streamlit app that turns bulk text/markdown/prose into a properly formatted PowerPoint deck matching an uploaded template (.pptx/.potx). No AI image generation; the app reuses images embedded in the template. Licensed under MIT.

Features
Paste long text; optional one-line guidance (e.g., “investor pitch deck”).
Bring your own LLM key: OpenAI, Anthropic, Gemini, OpenRouter, or a Custom OpenAI-compatible endpoint.
Upload a PowerPoint template/presentation; generated slides inherit its theme (colors, fonts, layouts).
Reuse images embedded in the template’s ppt/media/ folder.
Preview slide plan and download a ready .pptx deck.
No API keys or content stored server-side; only in memory during your session.
Requirements
Python 3.10+ recommended
See requirements.txt
Run locally
python -m venv .venv && source .venv/bin/activate  # .venv\Scripts\activate on Windows
pip install -r requirements.txt
streamlit run app.py
Deploy on Hugging Face Spaces (public demo link)
Create a new Space → Streamlit → Public.
Upload app.py and requirements.txt (and optionally README.md, LICENSE).
Open the Space URL. That URL is your hosted demo link (Deliverable #2).
Keys are pasted by users at runtime and are not stored. No server-side secrets needed.

Notes
Only reuses images from the template; no image generation.
Intelligent slide count/structure via LLM JSON plan with safeguards.
Reasonable file size limit (default: 25 MB); retry logic for API calls.
Layout matching is best-effort (exact fidelity not required).
Repository Structure
.
├── app.py
├── requirements.txt
├── README.md
├── LICENSE
└── WRITEUP.md
License
MIT — see LICENSE.
