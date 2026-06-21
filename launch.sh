#!/bin/bash
# ╔══════════════════════════════════════════════════╗
# ║  SYNTHGRID — Launch Script                       ║
# ║  AI-Powered Excel Automation Node                ║
# ╚══════════════════════════════════════════════════╝

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "╔══════════════════════════════════════════════════════════════╗"
echo "║  SYNTHGRID LAUNCHER                                          ║"
echo "╚══════════════════════════════════════════════════════════════╝"

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "ERROR: python3 not found. Install Python 3.10+"
    exit 1
fi

# Install deps if needed
echo "◈ Checking dependencies..."
pip install -q PyQt6 openpyxl pandas groq ollama openai anthropic --break-system-packages 2>/dev/null || \
pip install -q PyQt6 openpyxl pandas groq ollama openai anthropic 2>/dev/null || true

# Set API keys from args (order: GROQ, OPENAI, ANTHROPIC)
if [ ! -z "$1" ]; then
    export GROQ_API_KEY="$1"
    echo "◈ Groq API key loaded from argument"
fi
if [ ! -z "$2" ]; then
    export OPENAI_API_KEY="$2"
    echo "◈ OpenAI API key loaded from argument"
fi
if [ ! -z "$3" ]; then
    export ANTHROPIC_API_KEY="$3"
    echo "◈ Anthropic API key loaded from argument"
fi

echo "◈ Launching SynthGrid..."
python3 synthgrid.py
