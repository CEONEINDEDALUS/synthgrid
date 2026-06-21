#!/bin/bash
set -e

B=$(printf '\033[1m'); N=$(printf '\033[0m'); G=$(printf '\033[0;32m'); Y=$(printf '\033[0;33m'); R=$(printf '\033[0;31m'); C=$(printf '\033[0;36m')

echo "${C}╔══════════════════════════════════════════════════════════════╗${N}"
echo "${C}║  SYNTHGRID — Auto-Installer                                 ║${N}"
echo "${C}╚══════════════════════════════════════════════════════════════╝${N}"

# ── Python check ────────────────────────────────────────────────────────────
PY=$(command -v python3 || command -v python || true)
if [ -z "$PY" ]; then
    echo "${R}✗ Python not found. Install Python 3.10+ first.${N}"; exit 1
fi
PYVER=$($PY -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "${G}✓ Python ${PYVER} at ${PY}${N}"

if [ "$(echo "$PYVER" | cut -d. -f1)" -lt 3 ] || { [ "$(echo "$PYVER" | cut -d. -f1)" -eq 3 ] && [ "$(echo "$PYVER" | cut -d. -f2)" -lt 10 ]; }; then
    echo "${R}✗ Need Python 3.10+ (got ${PYVER})${N}"; exit 1
fi

# ── OS detection ────────────────────────────────────────────────────────────
OS=""
PKG_MGR=""
INSTALL_CMD=""
PYQT_SYSTEM=""
if [[ "$OSTYPE" == "darwin"* ]]; then
    OS="macos"
    PKG_MGR="brew"
    INSTALL_CMD="brew install"
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    OS="linux"
    if command -v apt &>/dev/null; then
        PKG_MGR="apt"; INSTALL_CMD="sudo apt install -y"
        PYQT_SYSTEM="python3-pyqt6"
    elif command -v dnf &>/dev/null; then
        PKG_MGR="dnf"; INSTALL_CMD="sudo dnf install -y"
        PYQT_SYSTEM="python3-qt6"
    elif command -v pacman &>/dev/null; then
        PKG_MGR="pacman"; INSTALL_CMD="sudo pacman -S --noconfirm"
        PYQT_SYSTEM="python-pyqt6"
    elif command -v zypper &>/dev/null; then
        PKG_MGR="zypper"; INSTALL_CMD="sudo zypper install -y"
        PYQT_SYSTEM="python3-qt6"
    fi
fi
echo "${G}✓ Detected: ${OS} (${PKG_MGR:-none})${N}"

# ── Install system PyQt if available (avoids pip Qt build failures) ────────
if [ -n "$PYQT_SYSTEM" ]; then
    echo "${Y}◈ Installing system PyQt6: ${PYQT_SYSTEM}${N}"
    $INSTALL_CMD $PYQT_SYSTEM 2>/dev/null || echo "${Y}  (PyQt6 system package not found, will use pip)${N}"
fi

# ── Pip install Python deps ────────────────────────────────────────────────
DIR="$(cd "$(dirname "$0")" && pwd)"
echo "${Y}◈ Installing Python packages...${N}"
$PY -m pip install --upgrade pip -q 2>/dev/null || true
if $PY -m pip install -r "$DIR/requirements.txt" -q 2>/dev/null; then
    echo "${G}✓ Python packages installed${N}"
else
    echo "${Y}  Pip install failed, trying --break-system-packages...${N}"
    $PY -m pip install -r "$DIR/requirements.txt" -q --break-system-packages 2>/dev/null || {
        echo "${R}✗ Pip install failed. Try:${N}"; echo "  ${Y}python3 -m pip install -r requirements.txt --user${N}"; exit 1
    }
    echo "${G}✓ Python packages installed${N}"
fi

# ── Ollama ──────────────────────────────────────────────────────────────────
if ! command -v ollama &>/dev/null && [ "$OS" = "linux" ]; then
    echo
    echo "${Y}◈ Ollama not detected. Install? [Y/n]${N} "
    read -r REPLY
    if [[ "$REPLY" =~ ^[Yy]?$ ]] || [ -z "$REPLY" ]; then
        if command -v curl &>/dev/null; then
            curl -fsSL https://ollama.com/install.sh | sh
            echo "${G}✓ Ollama installed${N}"
            echo "${Y}  Run 'ollama pull llama3.1' after install${N}"
        else
            echo "${R}✗ curl not found, install manually: https://ollama.com${N}"
        fi
    fi
fi

# ── Config dir + .env ───────────────────────────────────────────────────────
CFG_DIR="${XDG_CONFIG_HOME:-$HOME/.config}/synthgrid"
mkdir -p "$CFG_DIR"
if [ ! -f "$DIR/.env" ]; then
    cat > "$DIR/.env" << 'ENVEOF'
# ── SynthGrid API Keys ─────────────────────────────────────────────────────
# Set via GUI or here. GUI-saved keys go to ~/.config/synthgrid/keys.json.
GROQ_API_KEY=
OPENAI_API_KEY=
ANTHROPIC_API_KEY=
ENVEOF
    echo "${G}✓ .env created${N}"
fi

# ── .gitignore ──────────────────────────────────────────────────────────────
if [ ! -f "$DIR/.gitignore" ]; then
    echo ".env" > "$DIR/.gitignore"
    echo "${G}✓ .gitignore created${N}"
fi

echo
echo "${C}╔══════════════════════════════════════════════════════════════╗${N}"
echo "${C}║  ${B}INSTALL COMPLETE${N}${C}                                         ║${N}"
echo "${C}║  Run:  ${B}python3 synthgrid.py${N}${C}                                  ║${N}"
echo "${C}║  Or:   ${B}bash launch.sh${N}${C}                                         ║${N}"
echo "${C}╚══════════════════════════════════════════════════════════════╝${N}"
