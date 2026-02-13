# Work Report Analyzer

Minimal instructions to set up and run the analyzer.

Prerequisites

- Python 3.9+ (macOS: `python3`)
- Git (optional)

Setup

1. Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

Configuration

- Set environment variables before running, for example:

```bash
export OPENAI_API_KEY="sk-..."
export MODEL_NAME="gpt-5-mini"  # or your compatible model name
export OPENAI_BASE_URL="https://api.openai.com/v1"  # optional when using OpenAI
```

Usage

```bash
python work_report_analyzer.py
```
