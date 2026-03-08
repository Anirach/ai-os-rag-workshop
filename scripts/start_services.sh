#!/bin/bash
# Start all required services for the workshop
# Run this every time you resume the workshop

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo "🚀 Starting Workshop Services..."

# Ollama
if ! pgrep -x "ollama" > /dev/null; then
    echo "  Starting Ollama..."
    ollama serve &>/dev/null &
    sleep 2
    echo -e "  ${GREEN}✅ Ollama started${NC}"
else
    echo -e "  ${GREEN}✅ Ollama already running${NC}"
fi

# Qdrant
if docker ps | grep -q "qdrant"; then
    echo -e "  ${GREEN}✅ Qdrant already running${NC}"
else
    docker start qdrant 2>/dev/null || \
    docker run -d --name qdrant -p 6333:6333 qdrant/qdrant:latest
    echo -e "  ${GREEN}✅ Qdrant started${NC}"
fi

# Neo4j
if docker ps | grep -q "neo4j"; then
    echo -e "  ${GREEN}✅ Neo4j already running${NC}"
else
    docker start neo4j-workshop 2>/dev/null || \
    docker run -d --name neo4j-workshop \
        -p 7474:7474 -p 7687:7687 \
        -e NEO4J_AUTH=neo4j/workshop2026 \
        neo4j:latest
    echo -e "  ${GREEN}✅ Neo4j started${NC}"
fi

echo ""
echo "Services ready:"
echo "  Ollama:  http://localhost:11434"
echo "  Qdrant:  http://localhost:6333"
echo "  Neo4j:   http://localhost:7474"
echo ""
echo "Start Jupyter:"
echo "  source venv/bin/activate && jupyter lab"
