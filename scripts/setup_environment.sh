#!/bin/bash
# ======================================================
# AI OS RAG Workshop — Environment Setup Script
# PPLUS Visions Co., Ltd.
# ======================================================

set -e  # Exit on error

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

log() { echo -e "${GREEN}✅ $1${NC}"; }
warn() { echo -e "${YELLOW}⚠️  $1${NC}"; }
error() { echo -e "${RED}❌ $1${NC}"; exit 1; }
info() { echo -e "${BLUE}ℹ️  $1${NC}"; }

echo -e "${BLUE}"
echo "========================================================"
echo "  🤖 AI OS RAG Workshop - Setup"
echo "  PPLUS Visions Co., Ltd."
echo "========================================================"
echo -e "${NC}"

# --------------------------------------------------------
# 1. Check prerequisites
# --------------------------------------------------------
info "Checking prerequisites..."

command -v python3 >/dev/null 2>&1 || error "Python 3 not found. Install from https://python.org"
command -v docker >/dev/null 2>&1 || error "Docker not found. Install from https://docker.com"
command -v ollama >/dev/null 2>&1 || error "Ollama not found. Install from https://ollama.ai"

PYTHON_VERSION=$(python3 --version | grep -o '[0-9]\+\.[0-9]\+')
log "Python $PYTHON_VERSION found"

# --------------------------------------------------------
# 2. Pull Ollama models
# --------------------------------------------------------
info "Pulling Ollama models (this may take a while)..."

echo "Starting Ollama service..."
ollama serve &>/dev/null &
OLLAMA_PID=$!
sleep 3

# Pull LLM
echo "  Pulling qwen3:8b (LLM)..."
ollama pull qwen3:8b || warn "qwen3:8b pull failed. Try: ollama pull qwen2:7b"

# Pull embedding models
echo "  Pulling bge-m3 (multilingual embeddings)..."
ollama pull bge-m3 || warn "bge-m3 pull failed. Try: ollama pull nomic-embed-text"

echo "  Pulling nomic-embed-text (backup embeddings)..."
ollama pull nomic-embed-text || warn "nomic-embed-text pull failed"

log "Ollama models ready"

# --------------------------------------------------------
# 3. Start Docker services
# --------------------------------------------------------
info "Starting Docker services..."

# Qdrant Vector Database
echo "  Starting Qdrant..."
if docker ps -a | grep -q "qdrant"; then
    docker start qdrant 2>/dev/null || true
else
    docker run -d \
        --name qdrant \
        -p 6333:6333 \
        -p 6334:6334 \
        -v $(pwd)/qdrant_storage:/qdrant/storage \
        qdrant/qdrant:latest
fi
log "Qdrant started on port 6333"

# Neo4j Graph Database
echo "  Starting Neo4j..."
if docker ps -a | grep -q "neo4j-workshop"; then
    docker start neo4j-workshop 2>/dev/null || true
else
    docker run -d \
        --name neo4j-workshop \
        -p 7474:7474 \
        -p 7687:7687 \
        -e NEO4J_AUTH=neo4j/workshop2026 \
        -e NEO4J_PLUGINS='["apoc"]' \
        -v $(pwd)/neo4j_data:/data \
        neo4j:latest
fi
log "Neo4j started on port 7474 (browser) and 7687 (bolt)"

# Wait for services
echo "  Waiting for services to be ready..."
sleep 10

# --------------------------------------------------------
# 4. Install Python dependencies
# --------------------------------------------------------
info "Installing Python dependencies..."

# Create virtual environment
if [ ! -d "venv" ]; then
    python3 -m venv venv
    log "Virtual environment created"
fi

# Activate
source venv/bin/activate 2>/dev/null || source venv/Scripts/activate 2>/dev/null

# Upgrade pip
pip install --upgrade pip -q

# Install requirements
pip install -r requirements.txt -q
log "Python dependencies installed"

# Install Jupyter
pip install jupyter jupyterlab -q
log "Jupyter installed"

# --------------------------------------------------------
# 5. Run connection test
# --------------------------------------------------------
info "Testing connections..."
python3 scripts/test_connection.py

# --------------------------------------------------------
# Summary
# --------------------------------------------------------
echo ""
echo -e "${GREEN}"
echo "========================================================"
echo "  ✅ Setup Complete!"
echo "========================================================"
echo -e "${NC}"
echo ""
echo "Services:"
echo "  🔵 Ollama:  http://localhost:11434"
echo "  🟢 Qdrant:  http://localhost:6333"
echo "  🔴 Neo4j:   http://localhost:7474  (neo4j / workshop2026)"
echo ""
echo "Start Workshop:"
echo "  source venv/bin/activate"
echo "  jupyter lab"
echo ""
echo "Open Day 1:"
echo "  day1/day1_data_engineering.ipynb"
echo ""
