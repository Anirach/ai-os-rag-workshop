#!/usr/bin/env python3
"""
Test connections to all required services for the workshop.
Run: python scripts/test_connection.py
"""

import sys
from typing import Tuple

# Color output
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
RESET = "\033[0m"

def ok(msg): print(f"{GREEN}✅ {msg}{RESET}")
def fail(msg): print(f"{RED}❌ {msg}{RESET}")
def warn(msg): print(f"{YELLOW}⚠️  {msg}{RESET}")
def info(msg): print(f"{BLUE}ℹ️  {msg}{RESET}")


def test_ollama() -> Tuple[bool, str]:
    """Test Ollama connection and required models"""
    try:
        import ollama
        client = ollama.Client(host="http://localhost:11434")
        models = client.list()
        model_names = [m.model for m in models.models]
        
        has_llm = any("qwen" in m or "llama" in m or "mistral" in m for m in model_names)
        has_embed = any("bge" in m or "nomic" in m for m in model_names)
        
        if not model_names:
            return False, "No models found. Run: ollama pull qwen3:8b && ollama pull bge-m3"
        
        status = f"Connected | {len(model_names)} models"
        if not has_llm:
            status += " | ⚠️ No LLM (run: ollama pull qwen3:8b)"
        if not has_embed:
            status += " | ⚠️ No embeddings (run: ollama pull bge-m3)"
        
        return True, status
    
    except ImportError:
        return False, "ollama package not installed. Run: pip install ollama"
    except Exception as e:
        return False, f"Connection failed: {e}\nStart with: ollama serve"


def test_ollama_embedding() -> Tuple[bool, str]:
    """Test embedding generation"""
    try:
        import ollama
        
        # Try bge-m3 first, then nomic
        for model in ["bge-m3", "nomic-embed-text"]:
            try:
                resp = ollama.embeddings(model=model, prompt="ทดสอบ embedding ภาษาไทย")
                emb = resp["embedding"]
                return True, f"Model: {model} | Dims: {len(emb)}"
            except:
                continue
        
        return False, "No embedding model available. Run: ollama pull bge-m3"
    
    except Exception as e:
        return False, str(e)


def test_qdrant() -> Tuple[bool, str]:
    """Test Qdrant connection"""
    try:
        from qdrant_client import QdrantClient
        
        client = QdrantClient(host="localhost", port=6333)
        collections = client.get_collections()
        
        return True, f"Connected | {len(collections.collections)} collections"
    
    except ImportError:
        return False, "qdrant-client not installed. Run: pip install qdrant-client"
    except Exception as e:
        return False, f"Connection failed: {e}\nStart with: docker run -d --name qdrant -p 6333:6333 qdrant/qdrant:latest"


def test_neo4j() -> Tuple[bool, str]:
    """Test Neo4j connection"""
    try:
        from neo4j import GraphDatabase
        
        driver = GraphDatabase.driver(
            "bolt://localhost:7687",
            auth=("neo4j", "workshop2026")
        )
        driver.verify_connectivity()
        
        # Get node count
        with driver.session() as session:
            result = session.run("MATCH (n) RETURN count(n) as count")
            count = result.single()["count"]
        
        driver.close()
        return True, f"Connected | {count} nodes in database"
    
    except ImportError:
        return False, "neo4j package not installed. Run: pip install neo4j"
    except Exception as e:
        return False, f"Connection failed: {e}\nStart with: docker run -d --name neo4j -p 7474:7474 -p 7687:7687 -e NEO4J_AUTH=neo4j/workshop2026 neo4j:latest"


def test_pythainlp() -> Tuple[bool, str]:
    """Test PyThaiNLP"""
    try:
        from pythainlp.tokenize import word_tokenize
        
        tokens = word_tokenize("ทดสอบการตัดคำภาษาไทย", engine="newmm")
        return True, f"tokenize OK | tokens: {tokens}"
    
    except ImportError:
        return False, "pythainlp not installed. Run: pip install pythainlp"
    except Exception as e:
        return False, str(e)


def test_pymupdf() -> Tuple[bool, str]:
    """Test PyMuPDF for PDF processing"""
    try:
        import fitz
        return True, f"Version: {fitz.version[0]}"
    
    except ImportError:
        return False, "pymupdf not installed. Run: pip install pymupdf"


def main():
    print(f"\n{BLUE}{'='*60}")
    print("  🔍 Workshop Environment Test")
    print(f"  AI OS RAG Workshop — X-Company")
    print(f"{'='*60}{RESET}\n")
    
    tests = [
        ("Ollama Service", test_ollama),
        ("Ollama Embeddings", test_ollama_embedding),
        ("Qdrant Vector DB", test_qdrant),
        ("Neo4j Graph DB", test_neo4j),
        ("PyThaiNLP", test_pythainlp),
        ("PyMuPDF (PDF)", test_pymupdf),
    ]
    
    passed = 0
    failed = 0
    
    for name, test_fn in tests:
        try:
            success, message = test_fn()
            if success:
                ok(f"{name}: {message}")
                passed += 1
            else:
                fail(f"{name}: {message}")
                failed += 1
        except Exception as e:
            fail(f"{name}: Unexpected error: {e}")
            failed += 1
    
    print(f"\n{BLUE}{'='*60}{RESET}")
    print(f"Results: {GREEN}{passed} passed{RESET} | {RED}{failed} failed{RESET}")
    
    if failed == 0:
        print(f"\n{GREEN}🎉 All systems ready! Start Jupyter: jupyter lab{RESET}")
    else:
        print(f"\n{YELLOW}⚠️  Fix failed services before starting the workshop{RESET}")
        print(f"   Run: bash scripts/setup_environment.sh")
    
    print()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
