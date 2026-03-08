#!/usr/bin/env python3
"""
AI Workshop Handout Generator
Creates 6 DOCX handout files (3 days x 2 languages) for the AI OS RAG Workshop
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
import os

OUTPUT_DIR = "/Users/pplus/.openclaw/workspace/temp/ai-os-rag-workshop/handouts"
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_BG = RGBColor(0xF2, 0xF2, 0xF2)
LIGHT_NAVY = RGBColor(0xD6, 0xDC, 0xE4)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_footer(doc, text):
    section = doc.sections[0]
    footer = section.footer
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

def setup_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    # Normal style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    return doc

def add_heading(doc, text, level=1):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(6)
    run = para.add_run(text)
    run.font.name = 'Calibri'
    run.bold = True
    if level == 1:
        run.font.size = Pt(18)
        run.font.color.rgb = NAVY
    elif level == 2:
        run.font.size = Pt(14)
        run.font.color.rgb = NAVY
    elif level == 3:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x2E, 0x4A, 0x7A)
    return para

def add_body(doc, text):
    para = doc.add_paragraph(text)
    para.paragraph_format.space_after = Pt(6)
    for run in para.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
    return para

def add_bullet(doc, text, level=0):
    para = doc.add_paragraph(style='List Bullet')
    para.paragraph_format.left_indent = Cm(0.5 + level * 0.5)
    para.paragraph_format.space_after = Pt(3)
    run = para.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    return para

def add_code_block(doc, code_text):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.left_indent = Cm(0.5)
    # Set gray background
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)
    run = para.add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(9)
    return para

def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header row
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = h
        set_cell_bg(cell, '1B2A4A')
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in para.runs:
            run.font.bold = True
            run.font.color.rgb = WHITE
            run.font.name = 'Calibri'
            run.font.size = Pt(10)
    # Data rows
    for ri, row_data in enumerate(rows):
        row = table.rows[ri + 1]
        if ri % 2 == 1:
            bg = 'EEF1F5'
        else:
            bg = 'FFFFFF'
        for ci, cell_text in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(cell_text)
            set_cell_bg(cell, bg)
            for run in cell.paragraphs[0].runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(10)
    # Set column widths if provided
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)
    doc.add_paragraph()
    return table

def add_cover(doc, day_num, title_en, title_th, lang):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(60)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("AI Operating System Workshop")
    run.font.name = 'Calibri'
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = NAVY

    para2 = doc.add_paragraph()
    para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = para2.add_run(f"Day {day_num}: {title_en}")
    run2.font.name = 'Calibri'
    run2.font.size = Pt(18)
    run2.font.bold = True
    run2.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

    if lang == 'th':
        para3 = doc.add_paragraph()
        para3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run3 = para3.add_run(title_th)
        run3.font.name = 'Calibri'
        run3.font.size = Pt(16)
        run3.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    para4 = doc.add_paragraph()
    para4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run4 = para4.add_run("\nX-Company\n2026")
    run4.font.name = 'Calibri'
    run4.font.size = Pt(12)
    run4.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# DAY 1 ENGLISH
# ─────────────────────────────────────────────────────────────
def create_day1_en():
    doc = setup_document()
    add_footer(doc, "AI Operating System Workshop | X-Company")
    add_cover(doc, 1, "Data Engineering & RAG Foundation", "", "en")

    # Section 1: RAG Overview
    add_heading(doc, "Section 1: RAG Overview", 2)
    add_body(doc, "Retrieval-Augmented Generation (RAG) is an AI architecture that combines information retrieval with large language model generation. Instead of relying solely on the model's internal parameters, RAG fetches relevant context from an external knowledge base at query time, then passes it to the LLM for grounded, accurate answers.")
    add_body(doc, "Architecture Flow:")
    add_bullet(doc, "User Query → Embedding → Vector Search → Retrieved Chunks")
    add_bullet(doc, "Retrieved Chunks + Query → LLM Prompt → Generated Answer")

    add_heading(doc, "RAG vs Fine-tuning", 3)
    add_table(doc,
        ["Criterion", "RAG", "Fine-tuning"],
        [
            ["Knowledge update", "Real-time (update vector DB)", "Requires retraining"],
            ["Cost", "Low (no GPU training)", "High (GPU hours)"],
            ["Accuracy on new data", "High", "Low (catastrophic forgetting)"],
            ["Transparency", "Citable sources", "Black-box weights"],
            ["Best for", "Dynamic, document-heavy QA", "Style/format transfer"],
            ["Latency", "Higher (retrieval step)", "Lower (single forward pass)"],
        ],
        [4, 5, 5]
    )

    add_heading(doc, "When to Use RAG vs Fine-tuning", 3)
    add_bullet(doc, "Use RAG: When knowledge changes frequently, when source citation matters, when data is proprietary/confidential")
    add_bullet(doc, "Use Fine-tuning: When task requires a specific response style, when domain vocabulary is highly specialized")
    add_bullet(doc, "Use Both: For best results in enterprise settings — fine-tune for format, RAG for facts")

    # Section 2: Document Processing
    add_heading(doc, "Section 2: Document Processing", 2)
    add_heading(doc, "PDF/DOCX Text Extraction with PyMuPDF", 3)
    add_code_block(doc,
        "import fitz  # PyMuPDF\n\n"
        "def extract_pdf(path: str) -> str:\n"
        "    doc = fitz.open(path)\n"
        "    text = ''\n"
        "    for page in doc:\n"
        "        text += page.get_text('text')\n"
        "    return text\n\n"
        "# DOCX extraction\n"
        "from docx import Document\n\n"
        "def extract_docx(path: str) -> str:\n"
        "    doc = Document(path)\n"
        "    return '\\n'.join([p.text for p in doc.paragraphs])\n\n"
        "text = extract_pdf('document.pdf')\n"
        "print(text[:500])"
    )

    add_heading(doc, "Thai NLP: Word Segmentation Challenges", 3)
    add_body(doc, "Thai language has no spaces between words, making tokenization a critical and non-trivial step. Challenges include:")
    add_bullet(doc, "No whitespace delimiters — word boundaries must be inferred")
    add_bullet(doc, "Ambiguous segmentation: 'ตากลม' can mean 'round eyes' or 'Tak Lom (city)'")
    add_bullet(doc, "New words and loanwords not in dictionary")
    add_bullet(doc, "Context-dependent meaning shifts")

    add_heading(doc, "PyThaiNLP Word Tokenization", 3)
    add_code_block(doc,
        "from pythainlp.tokenize import word_tokenize\n\n"
        "text = 'การประมวลผลภาษาธรรมชาติมีความสำคัญมาก'\n\n"
        "# Default engine: newmm\n"
        "tokens_newmm = word_tokenize(text, engine='newmm')\n"
        "print('newmm:', tokens_newmm)\n\n"
        "# Longest match\n"
        "tokens_longest = word_tokenize(text, engine='longest')\n"
        "print('longest:', tokens_longest)\n\n"
        "# attacut (deep learning)\n"
        "tokens_attacut = word_tokenize(text, engine='attacut')\n"
        "print('attacut:', tokens_attacut)"
    )

    add_heading(doc, "Thai Tokenizer Comparison", 3)
    add_table(doc,
        ["Tokenizer", "Algorithm", "Speed", "Accuracy", "Best For"],
        [
            ["newmm", "Dict + MaxMatch", "Fast", "Good", "General purpose"],
            ["longest", "Longest match", "Fast", "Moderate", "Simple texts"],
            ["attacut", "Deep Learning (CNN)", "Medium", "High", "Complex/formal text"],
            ["icu", "Unicode rules", "Very Fast", "Low", "Quick tokenization"],
        ],
        [3, 4, 2.5, 2.5, 4]
    )

    # Section 3: Chunking
    add_heading(doc, "Section 3: Chunking Strategies", 2)
    add_heading(doc, "Fixed-Size Chunking", 3)
    add_body(doc, "Split text into equal-length chunks with optional overlap. Simple and fast.")
    add_code_block(doc,
        "def fixed_chunk(text: str, size: int = 512, overlap: int = 50) -> list[str]:\n"
        "    chunks = []\n"
        "    start = 0\n"
        "    while start < len(text):\n"
        "        end = start + size\n"
        "        chunks.append(text[start:end])\n"
        "        start += size - overlap\n"
        "    return chunks\n\n"
        "chunks = fixed_chunk(text, size=512, overlap=50)\n"
        "print(f'Total chunks: {len(chunks)}')"
    )

    add_heading(doc, "Recursive Chunking", 3)
    add_body(doc, "Splits on natural boundaries (paragraphs → sentences → words) until target size is reached.")
    add_code_block(doc,
        "from langchain.text_splitter import RecursiveCharacterTextSplitter\n\n"
        "splitter = RecursiveCharacterTextSplitter(\n"
        "    chunk_size=512,\n"
        "    chunk_overlap=50,\n"
        "    separators=['\\n\\n', '\\n', '.', ' ', '']\n"
        ")\n"
        "chunks = splitter.split_text(text)\n"
        "print(f'Chunks: {len(chunks)}, Sample: {chunks[0][:100]}')"
    )

    add_heading(doc, "Semantic Chunking", 3)
    add_body(doc, "Uses embedding similarity to split at meaningful topic boundaries.")
    add_code_block(doc,
        "from langchain_experimental.text_splitter import SemanticChunker\n"
        "from langchain_community.embeddings import OllamaEmbeddings\n\n"
        "embeddings = OllamaEmbeddings(model='bge-m3')\n"
        "chunker = SemanticChunker(\n"
        "    embeddings,\n"
        "    breakpoint_threshold_type='percentile',\n"
        "    breakpoint_threshold_amount=95\n"
        ")\n"
        "chunks = chunker.split_text(text)"
    )

    add_heading(doc, "Chunking Strategy Comparison", 3)
    add_table(doc,
        ["Strategy", "Pros", "Cons", "Best Use Case"],
        [
            ["Fixed-size", "Simple, fast, predictable", "May split mid-sentence", "Large uniform documents"],
            ["Recursive", "Respects natural boundaries", "Slightly slower", "General-purpose RAG"],
            ["Semantic", "Coherent topic chunks", "Slow, needs embeddings", "Research papers, reports"],
        ],
        [3, 4.5, 4.5, 4]
    )

    # Section 4: Embeddings & Vector DB
    add_heading(doc, "Section 4: Embeddings & Vector Database", 2)
    add_body(doc, "Embeddings are dense vector representations of text that capture semantic meaning. Similar texts produce similar vectors, enabling semantic search.")

    add_heading(doc, "Embedding Model Comparison", 3)
    add_table(doc,
        ["Model", "Dimensions", "Languages", "Size", "Best For"],
        [
            ["bge-m3", "1024", "100+", "570MB", "Multilingual, Thai-friendly"],
            ["nomic-embed-text", "768", "English primary", "274MB", "English docs, fast"],
            ["multilingual-e5-large", "1024", "100+", "560MB", "Cross-lingual retrieval"],
        ],
        [4.5, 2.5, 3, 2, 4]
    )

    add_heading(doc, "Qdrant Vector DB Setup", 3)
    add_code_block(doc,
        "# Start Qdrant with Docker\n"
        "docker run -d --name qdrant \\\n"
        "  -p 6333:6333 -p 6334:6334 \\\n"
        "  -v $(pwd)/qdrant_storage:/qdrant/storage \\\n"
        "  qdrant/qdrant:latest"
    )

    add_heading(doc, "Indexing Documents", 3)
    add_code_block(doc,
        "from qdrant_client import QdrantClient\n"
        "from qdrant_client.models import Distance, VectorParams, PointStruct\n"
        "import ollama\n\n"
        "client = QdrantClient('localhost', port=6333)\n\n"
        "# Create collection\n"
        "client.create_collection(\n"
        "    collection_name='workshop_docs',\n"
        "    vectors_config=VectorParams(size=1024, distance=Distance.COSINE)\n"
        ")\n\n"
        "# Index chunks\n"
        "for i, chunk in enumerate(chunks):\n"
        "    embedding = ollama.embeddings(model='bge-m3', prompt=chunk)\n"
        "    client.upsert('workshop_docs', points=[\n"
        "        PointStruct(id=i, vector=embedding['embedding'],\n"
        "                    payload={'text': chunk})\n"
        "    ])"
    )

    add_heading(doc, "Semantic Search", 3)
    add_code_block(doc,
        "def search(query: str, top_k: int = 5) -> list[str]:\n"
        "    query_vec = ollama.embeddings(model='bge-m3', prompt=query)['embedding']\n"
        "    results = client.search(\n"
        "        collection_name='workshop_docs',\n"
        "        query_vector=query_vec,\n"
        "        limit=top_k\n"
        "    )\n"
        "    return [r.payload['text'] for r in results]\n\n"
        "context = search('What is the refund policy?')"
    )

    # Section 5: Hybrid Search
    add_heading(doc, "Section 5: Hybrid Search", 2)
    add_body(doc, "Hybrid search combines dense vector similarity (semantic) with sparse keyword matching (BM25) for better retrieval coverage.")

    add_heading(doc, "Dense vs BM25", 3)
    add_table(doc,
        ["Approach", "Method", "Strength", "Weakness"],
        [
            ["Dense (Vector)", "Embedding similarity", "Semantic understanding", "Misses exact keywords"],
            ["BM25 (Sparse)", "TF-IDF based", "Exact keyword matching", "No semantic understanding"],
            ["Hybrid", "Weighted combination", "Best of both", "More complex pipeline"],
        ],
        [3.5, 4, 4.5, 4]
    )

    add_heading(doc, "Hybrid Search Formula", 3)
    add_code_block(doc,
        "# Reciprocal Rank Fusion (RRF)\n"
        "def rrf_score(rank: int, k: int = 60) -> float:\n"
        "    return 1.0 / (k + rank)\n\n"
        "# Combine dense + sparse results\n"
        "def hybrid_search(query, dense_results, sparse_results, alpha=0.7):\n"
        "    scores = {}\n"
        "    for rank, doc in enumerate(dense_results):\n"
        "        scores[doc.id] = alpha * rrf_score(rank)\n"
        "    for rank, doc in enumerate(sparse_results):\n"
        "        scores[doc.id] = scores.get(doc.id, 0) + (1-alpha) * rrf_score(rank)\n"
        "    return sorted(scores.items(), key=lambda x: x[1], reverse=True)"
    )

    add_heading(doc, "Reranking Strategies", 3)
    add_bullet(doc, "Cross-encoder reranking: Use a fine-tuned model (e.g., bge-reranker-v2-m3) to score query-document pairs")
    add_bullet(doc, "LLM-as-reranker: Ask LLM to rank retrieved passages by relevance")
    add_bullet(doc, "Cohere Rerank API: Commercial API for production reranking")

    # Section 6: Basic RAG Pipeline
    add_heading(doc, "Section 6: Basic RAG Pipeline", 2)
    add_code_block(doc,
        "import ollama\n"
        "from qdrant_client import QdrantClient\n\n"
        "client = QdrantClient('localhost', port=6333)\n\n"
        "def rag_pipeline(question: str) -> str:\n"
        "    # 1. Embed the question\n"
        "    q_vec = ollama.embeddings(model='bge-m3', prompt=question)['embedding']\n\n"
        "    # 2. Retrieve top-5 relevant chunks\n"
        "    results = client.search(\n"
        "        collection_name='workshop_docs',\n"
        "        query_vector=q_vec, limit=5\n"
        "    )\n"
        "    context = '\\n\\n'.join([r.payload['text'] for r in results])\n\n"
        "    # 3. Build prompt\n"
        "    prompt = f'Context:\\n{context}\\n\\nQuestion: {question}\\nAnswer:'\n\n"
        "    # 4. Generate answer\n"
        "    response = ollama.chat(\n"
        "        model='llama3.2',\n"
        "        messages=[{'role': 'user', 'content': prompt}]\n"
        "    )\n"
        "    return response['message']['content']\n\n"
        "answer = rag_pipeline('What documents are required for reimbursement?')\n"
        "print(answer)"
    )
    add_body(doc, "Expected Output: A concise, grounded answer based only on retrieved context. The model will cite relevant sections rather than hallucinating.")

    # Section 7: Exercises
    add_heading(doc, "Section 7: Hands-On Exercises", 2)
    add_heading(doc, "Exercise 1: Document Ingestion Pipeline", 3)
    add_body(doc, "Task: Build a pipeline to extract text from 5 PDFs and store chunks in Qdrant.")
    add_bullet(doc, "Hint: Use PyMuPDF for extraction, RecursiveCharacterTextSplitter for chunking")
    add_bullet(doc, "Bonus: Add metadata (filename, page number) to each chunk")

    add_heading(doc, "Exercise 2: Thai Text Processing", 3)
    add_body(doc, "Task: Compare tokenization results of newmm vs attacut on a Thai product description.")
    add_bullet(doc, "Hint: Use pythainlp.tokenize.word_tokenize with engine parameter")

    add_heading(doc, "Exercise 3: Hybrid Search", 3)
    add_body(doc, "Task: Implement RRF to combine dense and BM25 results, compare with dense-only.")
    add_bullet(doc, "Hint: Use rank_bm25 library for sparse retrieval")

    add_heading(doc, "Exercise 4: Reranking", 3)
    add_body(doc, "Task: Add a cross-encoder reranker to your RAG pipeline and measure NDCG improvement.")
    add_bullet(doc, "Hint: sentence-transformers CrossEncoder with 'BAAI/bge-reranker-v2-m3'")

    add_heading(doc, "Exercise 5: End-to-End RAG", 3)
    add_body(doc, "Task: Build a complete Q&A chatbot over a company policy document with conversation history.")
    add_bullet(doc, "Hint: Use a list to maintain chat_history, prepend context to each turn")

    # Section 8: Reference
    add_heading(doc, "Section 8: Reference & Resources", 2)
    add_heading(doc, "Libraries", 3)
    add_bullet(doc, "PyMuPDF: https://pymupdf.readthedocs.io")
    add_bullet(doc, "PyThaiNLP: https://pythainlp.github.io")
    add_bullet(doc, "Qdrant: https://qdrant.tech/documentation")
    add_bullet(doc, "LangChain Text Splitters: https://python.langchain.com/docs/modules/data_connection/document_transformers")
    add_bullet(doc, "Ollama: https://ollama.ai")

    add_heading(doc, "Recommended Reading", 3)
    add_bullet(doc, "\"Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks\" — Lewis et al. (2020)")
    add_bullet(doc, "\"BEIR: A Heterogeneous Benchmark for Zero-shot Evaluation\" — Thakur et al.")
    add_bullet(doc, "Qdrant Hybrid Search Documentation")
    add_bullet(doc, "BGE-M3 Technical Report (BAAI)")

    path = os.path.join(OUTPUT_DIR, "day1-handout-en.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path

# ─────────────────────────────────────────────────────────────
# DAY 1 THAI
# ─────────────────────────────────────────────────────────────
def create_day1_th():
    doc = setup_document()
    add_footer(doc, "AI Operating System Workshop | X-Company")
    add_cover(doc, 1, "Data Engineering & RAG Foundation", "วิศวกรรมข้อมูลและพื้นฐาน RAG", "th")

    # Section 1
    add_heading(doc, "ส่วนที่ 1: ภาพรวม RAG", 2)
    add_body(doc, "Retrieval-Augmented Generation (RAG) คือสถาปัตยกรรม AI ที่รวมการดึงข้อมูล (Retrieval) เข้ากับการสร้างข้อความ (Generation) ของ LLM โดยแทนที่จะพึ่งพาความรู้จากพารามิเตอร์ของโมเดลเพียงอย่างเดียว RAG จะค้นหาบริบทที่เกี่ยวข้องจากฐานความรู้ภายนอกก่อน แล้วส่งให้ LLM สร้างคำตอบที่แม่นยำและอ้างอิงได้")
    add_body(doc, "ขั้นตอนการทำงาน:")
    add_bullet(doc, "คำถามผู้ใช้ → Embedding → ค้นหา Vector → ดึง Chunks")
    add_bullet(doc, "Chunks + คำถาม → Prompt LLM → สร้างคำตอบ")

    add_heading(doc, "RAG vs Fine-tuning", 3)
    add_table(doc,
        ["เกณฑ์", "RAG", "Fine-tuning"],
        [
            ["อัพเดทความรู้", "Real-time (อัพเดท Vector DB)", "ต้องเทรนใหม่"],
            ["ค่าใช้จ่าย", "ต่ำ (ไม่ต้องใช้ GPU training)", "สูง (GPU hours)"],
            ["ความแม่นยำข้อมูลใหม่", "สูง", "ต่ำ (catastrophic forgetting)"],
            ["ความโปร่งใส", "อ้างอิงแหล่งข้อมูลได้", "Black-box weights"],
            ["เหมาะกับ", "Q&A เอกสารที่เปลี่ยนบ่อย", "ปรับ style/format การตอบ"],
            ["Latency", "สูงกว่า (มีขั้น retrieval)", "ต่ำกว่า"],
        ],
        [4, 5, 5]
    )

    add_heading(doc, "เมื่อไหร่ควรใช้ RAG หรือ Fine-tuning", 3)
    add_bullet(doc, "ใช้ RAG: เมื่อข้อมูลเปลี่ยนบ่อย, ต้องการอ้างอิงแหล่งข้อมูล, ข้อมูลเป็นความลับ")
    add_bullet(doc, "ใช้ Fine-tuning: เมื่อต้องการ response style เฉพาะ, คำศัพท์ domain พิเศษมาก")
    add_bullet(doc, "ใช้ทั้งคู่: เพื่อผลลัพธ์ดีที่สุดในองค์กร — fine-tune สำหรับรูปแบบ, RAG สำหรับข้อเท็จจริง")

    # Section 2
    add_heading(doc, "ส่วนที่ 2: การประมวลผลเอกสาร", 2)
    add_heading(doc, "การดึงข้อความจาก PDF/DOCX ด้วย PyMuPDF", 3)
    add_code_block(doc,
        "import fitz  # PyMuPDF\n\n"
        "def extract_pdf(path: str) -> str:\n"
        "    doc = fitz.open(path)\n"
        "    text = ''\n"
        "    for page in doc:\n"
        "        text += page.get_text('text')\n"
        "    return text\n\n"
        "# ดึงข้อความจาก DOCX\n"
        "from docx import Document\n\n"
        "def extract_docx(path: str) -> str:\n"
        "    doc = Document(path)\n"
        "    return '\\n'.join([p.text for p in doc.paragraphs])\n\n"
        "text = extract_pdf('document.pdf')\n"
        "print(text[:500])"
    )

    add_heading(doc, "ความท้าทายของ NLP ภาษาไทย: การตัดคำ", 3)
    add_body(doc, "ภาษาไทยไม่มีช่องว่างระหว่างคำ ทำให้การตัดคำเป็นขั้นตอนที่สำคัญและยากมาก ปัญหาที่พบบ่อย:")
    add_bullet(doc, "ไม่มีตัวคั่นคำ — ต้องอนุมานขอบเขตของคำ")
    add_bullet(doc, "การตัดคำที่คลุมเครือ: 'ตากลม' อาจหมายถึง 'ตา (กลม)' หรือ 'ตาก (ลม)'")
    add_bullet(doc, "คำใหม่และคำทับศัพท์ที่ไม่อยู่ใน dictionary")
    add_bullet(doc, "ความหมายเปลี่ยนตาม context")

    add_heading(doc, "การตัดคำด้วย PyThaiNLP", 3)
    add_code_block(doc,
        "from pythainlp.tokenize import word_tokenize\n\n"
        "text = 'การประมวลผลภาษาธรรมชาติมีความสำคัญมาก'\n\n"
        "# เครื่องมือเริ่มต้น: newmm\n"
        "tokens_newmm = word_tokenize(text, engine='newmm')\n"
        "print('newmm:', tokens_newmm)\n\n"
        "# Longest match\n"
        "tokens_longest = word_tokenize(text, engine='longest')\n"
        "print('longest:', tokens_longest)\n\n"
        "# attacut (deep learning)\n"
        "tokens_attacut = word_tokenize(text, engine='attacut')\n"
        "print('attacut:', tokens_attacut)"
    )

    add_heading(doc, "เปรียบเทียบเครื่องมือตัดคำภาษาไทย", 3)
    add_table(doc,
        ["เครื่องมือ", "อัลกอริทึม", "ความเร็ว", "ความแม่นยำ", "เหมาะกับ"],
        [
            ["newmm", "Dict + MaxMatch", "เร็ว", "ดี", "ทั่วไป"],
            ["longest", "Longest match", "เร็ว", "ปานกลาง", "ข้อความง่าย"],
            ["attacut", "Deep Learning (CNN)", "ปานกลาง", "สูง", "ข้อความซับซ้อน/ทางการ"],
            ["icu", "Unicode rules", "เร็วมาก", "ต่ำ", "ตัดคำเบื้องต้น"],
        ],
        [3, 4, 2.5, 2.5, 4]
    )

    # Section 3
    add_heading(doc, "ส่วนที่ 3: กลยุทธ์การแบ่ง Chunk", 2)
    add_heading(doc, "Fixed-Size Chunking", 3)
    add_body(doc, "แบ่งข้อความเป็น chunk ขนาดเท่ากัน พร้อม overlap ง่ายและเร็ว")
    add_code_block(doc,
        "def fixed_chunk(text: str, size: int = 512, overlap: int = 50) -> list[str]:\n"
        "    chunks = []\n"
        "    start = 0\n"
        "    while start < len(text):\n"
        "        end = start + size\n"
        "        chunks.append(text[start:end])\n"
        "        start += size - overlap\n"
        "    return chunks\n\n"
        "chunks = fixed_chunk(text, size=512, overlap=50)\n"
        "print(f'จำนวน chunks: {len(chunks)}')"
    )

    add_heading(doc, "Recursive Chunking", 3)
    add_body(doc, "แบ่งตามขอบเขตธรรมชาติ (ย่อหน้า → ประโยค → คำ) จนถึงขนาดที่กำหนด")
    add_code_block(doc,
        "from langchain.text_splitter import RecursiveCharacterTextSplitter\n\n"
        "splitter = RecursiveCharacterTextSplitter(\n"
        "    chunk_size=512,\n"
        "    chunk_overlap=50,\n"
        "    separators=['\\n\\n', '\\n', '.', ' ', '']\n"
        ")\n"
        "chunks = splitter.split_text(text)"
    )

    add_heading(doc, "Semantic Chunking", 3)
    add_body(doc, "ใช้ embedding similarity เพื่อแบ่งตามหัวข้อที่มีความหมาย")
    add_code_block(doc,
        "from langchain_experimental.text_splitter import SemanticChunker\n"
        "from langchain_community.embeddings import OllamaEmbeddings\n\n"
        "embeddings = OllamaEmbeddings(model='bge-m3')\n"
        "chunker = SemanticChunker(embeddings,\n"
        "    breakpoint_threshold_type='percentile')\n"
        "chunks = chunker.split_text(text)"
    )

    add_heading(doc, "เปรียบเทียบกลยุทธ์การแบ่ง Chunk", 3)
    add_table(doc,
        ["กลยุทธ์", "ข้อดี", "ข้อเสีย", "เหมาะกับ"],
        [
            ["Fixed-size", "ง่าย เร็ว คาดเดาได้", "อาจตัดกลางประโยค", "เอกสารขนาดใหญ่สม่ำเสมอ"],
            ["Recursive", "เคารพขอบเขตธรรมชาติ", "ช้ากว่าเล็กน้อย", "RAG ทั่วไป"],
            ["Semantic", "chunk สอดคล้องกับหัวข้อ", "ช้า ต้องใช้ embeddings", "บทความวิจัย รายงาน"],
        ],
        [3, 4.5, 4.5, 4]
    )

    # Section 4
    add_heading(doc, "ส่วนที่ 4: Embeddings และ Vector Database", 2)
    add_body(doc, "Embeddings คือการแปลงข้อความเป็น vector ที่จับความหมายได้ ข้อความที่มีความหมายใกล้เคียงกันจะให้ vector ที่ใกล้กัน ช่วยให้ค้นหาแบบ semantic ได้")

    add_heading(doc, "เปรียบเทียบโมเดล Embedding", 3)
    add_table(doc,
        ["โมเดล", "มิติ", "ภาษา", "ขนาด", "เหมาะกับ"],
        [
            ["bge-m3", "1024", "100+", "570MB", "หลายภาษา รองรับไทย"],
            ["nomic-embed-text", "768", "อังกฤษหลัก", "274MB", "เอกสารอังกฤษ เร็ว"],
            ["multilingual-e5-large", "1024", "100+", "560MB", "ค้นหาข้ามภาษา"],
        ],
        [4.5, 2.5, 3, 2, 4]
    )

    add_heading(doc, "ติดตั้ง Qdrant ด้วย Docker", 3)
    add_code_block(doc,
        "# เริ่มต้น Qdrant ด้วย Docker\n"
        "docker run -d --name qdrant \\\n"
        "  -p 6333:6333 -p 6334:6334 \\\n"
        "  -v $(pwd)/qdrant_storage:/qdrant/storage \\\n"
        "  qdrant/qdrant:latest"
    )

    add_heading(doc, "การ Index เอกสาร", 3)
    add_code_block(doc,
        "from qdrant_client import QdrantClient\n"
        "from qdrant_client.models import Distance, VectorParams, PointStruct\n"
        "import ollama\n\n"
        "client = QdrantClient('localhost', port=6333)\n"
        "client.create_collection(\n"
        "    collection_name='workshop_docs',\n"
        "    vectors_config=VectorParams(size=1024, distance=Distance.COSINE)\n"
        ")\n\n"
        "for i, chunk in enumerate(chunks):\n"
        "    embedding = ollama.embeddings(model='bge-m3', prompt=chunk)\n"
        "    client.upsert('workshop_docs', points=[\n"
        "        PointStruct(id=i, vector=embedding['embedding'],\n"
        "                    payload={'text': chunk})\n"
        "    ])"
    )

    add_heading(doc, "การค้นหา Semantic", 3)
    add_code_block(doc,
        "def search(query: str, top_k: int = 5) -> list[str]:\n"
        "    query_vec = ollama.embeddings(model='bge-m3', prompt=query)['embedding']\n"
        "    results = client.search(\n"
        "        collection_name='workshop_docs',\n"
        "        query_vector=query_vec,\n"
        "        limit=top_k\n"
        "    )\n"
        "    return [r.payload['text'] for r in results]"
    )

    # Section 5
    add_heading(doc, "ส่วนที่ 5: Hybrid Search", 2)
    add_body(doc, "Hybrid Search รวม dense vector similarity (semantic) กับ sparse keyword matching (BM25) เพื่อให้ได้ผลการค้นหาที่ครบถ้วนกว่า")

    add_heading(doc, "Dense vs BM25", 3)
    add_table(doc,
        ["วิธีการ", "Algorithm", "จุดแข็ง", "จุดอ่อน"],
        [
            ["Dense (Vector)", "Embedding similarity", "เข้าใจความหมาย", "พลาดคำที่แน่นอน"],
            ["BM25 (Sparse)", "TF-IDF based", "จับคีย์เวิร์ดได้แม่น", "ไม่เข้าใจ semantic"],
            ["Hybrid", "Weighted combination", "รวมข้อดีทั้งคู่", "pipeline ซับซ้อนขึ้น"],
        ],
        [3.5, 4, 4.5, 4]
    )

    add_heading(doc, "สูตร Hybrid Search (RRF)", 3)
    add_code_block(doc,
        "# Reciprocal Rank Fusion\n"
        "def rrf_score(rank: int, k: int = 60) -> float:\n"
        "    return 1.0 / (k + rank)\n\n"
        "def hybrid_search(dense_results, sparse_results, alpha=0.7):\n"
        "    scores = {}\n"
        "    for rank, doc in enumerate(dense_results):\n"
        "        scores[doc.id] = alpha * rrf_score(rank)\n"
        "    for rank, doc in enumerate(sparse_results):\n"
        "        scores[doc.id] = scores.get(doc.id, 0) + (1-alpha) * rrf_score(rank)\n"
        "    return sorted(scores.items(), key=lambda x: x[1], reverse=True)"
    )

    add_heading(doc, "กลยุทธ์ Reranking", 3)
    add_bullet(doc, "Cross-encoder reranking: ใช้โมเดล (เช่น bge-reranker-v2-m3) ให้คะแนนคู่ query-document")
    add_bullet(doc, "LLM-as-reranker: ให้ LLM จัดอันดับ passage ตามความเกี่ยวข้อง")
    add_bullet(doc, "Cohere Rerank API: API เชิงพาณิชย์สำหรับ production")

    # Section 6
    add_heading(doc, "ส่วนที่ 6: RAG Pipeline พื้นฐาน", 2)
    add_code_block(doc,
        "import ollama\n"
        "from qdrant_client import QdrantClient\n\n"
        "client = QdrantClient('localhost', port=6333)\n\n"
        "def rag_pipeline(question: str) -> str:\n"
        "    # 1. แปลง query เป็น embedding\n"
        "    q_vec = ollama.embeddings(model='bge-m3', prompt=question)['embedding']\n\n"
        "    # 2. ดึง chunks ที่เกี่ยวข้องมากที่สุด 5 อัน\n"
        "    results = client.search(\n"
        "        collection_name='workshop_docs',\n"
        "        query_vector=q_vec, limit=5\n"
        "    )\n"
        "    context = '\\n\\n'.join([r.payload['text'] for r in results])\n\n"
        "    # 3. สร้าง prompt\n"
        "    prompt = f'บริบท:\\n{context}\\n\\nคำถาม: {question}\\nคำตอบ:'\n\n"
        "    # 4. สร้างคำตอบ\n"
        "    response = ollama.chat(\n"
        "        model='llama3.2',\n"
        "        messages=[{'role': 'user', 'content': prompt}]\n"
        "    )\n"
        "    return response['message']['content']\n\n"
        "answer = rag_pipeline('เอกสารที่ต้องใช้ในการเบิกค่าใช้จ่ายมีอะไรบ้าง?')\n"
        "print(answer)"
    )
    add_body(doc, "ผลลัพธ์ที่คาดหวัง: คำตอบที่กระชับ อ้างอิงจาก context ที่ดึงมาเท่านั้น โมเดลจะอ้างถึงส่วนที่เกี่ยวข้องแทนการ hallucinate")

    # Section 7
    add_heading(doc, "ส่วนที่ 7: แบบฝึกหัด", 2)
    add_heading(doc, "แบบฝึกหัดที่ 1: Document Ingestion Pipeline", 3)
    add_body(doc, "งาน: สร้าง pipeline ดึงข้อความจาก PDF 5 ไฟล์ และเก็บ chunks ใน Qdrant")
    add_bullet(doc, "Hint: ใช้ PyMuPDF ดึงข้อความ, RecursiveCharacterTextSplitter แบ่ง chunk")
    add_bullet(doc, "Bonus: เพิ่ม metadata (ชื่อไฟล์, หน้า) ให้แต่ละ chunk")

    add_heading(doc, "แบบฝึกหัดที่ 2: การประมวลผลข้อความภาษาไทย", 3)
    add_body(doc, "งาน: เปรียบเทียบผลการตัดคำ newmm vs attacut บนคำบรรยายสินค้าภาษาไทย")
    add_bullet(doc, "Hint: ใช้ pythainlp.tokenize.word_tokenize กับ engine parameter ต่างๆ")

    add_heading(doc, "แบบฝึกหัดที่ 3: Hybrid Search", 3)
    add_body(doc, "งาน: ใช้ RRF รวมผล dense และ BM25 แล้วเปรียบเทียบกับ dense เพียงอย่างเดียว")
    add_bullet(doc, "Hint: ใช้ library rank_bm25 สำหรับ sparse retrieval")

    add_heading(doc, "แบบฝึกหัดที่ 4: Reranking", 3)
    add_body(doc, "งาน: เพิ่ม cross-encoder reranker ใน RAG pipeline และวัดการปรับปรุง NDCG")
    add_bullet(doc, "Hint: sentence-transformers CrossEncoder กับ 'BAAI/bge-reranker-v2-m3'")

    add_heading(doc, "แบบฝึกหัดที่ 5: RAG ครบวงจร", 3)
    add_body(doc, "งาน: สร้าง Q&A chatbot บนเอกสาร policy ของบริษัท พร้อม conversation history")
    add_bullet(doc, "Hint: ใช้ list เก็บ chat_history และ prepend context ทุก turn")

    # Section 8
    add_heading(doc, "ส่วนที่ 8: เอกสารอ้างอิงและแหล่งเรียนรู้", 2)
    add_heading(doc, "Libraries", 3)
    add_bullet(doc, "PyMuPDF: https://pymupdf.readthedocs.io")
    add_bullet(doc, "PyThaiNLP: https://pythainlp.github.io")
    add_bullet(doc, "Qdrant: https://qdrant.tech/documentation")
    add_bullet(doc, "Ollama: https://ollama.ai")

    add_heading(doc, "บทความแนะนำ", 3)
    add_bullet(doc, "\"Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks\" — Lewis et al. (2020)")
    add_bullet(doc, "รายงานทางเทคนิค BGE-M3 (BAAI)")
    add_bullet(doc, "เอกสาร Qdrant Hybrid Search")

    path = os.path.join(OUTPUT_DIR, "day1-handout-th.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path

# ─────────────────────────────────────────────────────────────
# DAY 2 ENGLISH
# ─────────────────────────────────────────────────────────────
def create_day2_en():
    doc = setup_document()
    add_footer(doc, "AI Operating System Workshop | X-Company")
    add_cover(doc, 2, "Knowledge Graph & Advanced RAG", "", "en")

    # Section 1
    add_heading(doc, "Section 1: Knowledge Graph Fundamentals", 2)
    add_body(doc, "A Knowledge Graph (KG) is a structured representation of real-world entities and their relationships. It organizes information as a network of nodes and edges, enabling complex multi-hop reasoning that vector databases cannot perform.")
    add_bullet(doc, "Nodes: Entities — Person, Organization, Product, Concept")
    add_bullet(doc, "Edges: Relationships — 'works_at', 'part_of', 'related_to'")
    add_bullet(doc, "Properties: Attributes on nodes/edges — name, date, confidence_score")

    add_heading(doc, "KG vs Relational DB vs Vector DB", 3)
    add_table(doc,
        ["Feature", "Knowledge Graph", "Relational DB", "Vector DB"],
        [
            ["Data model", "Graph (nodes/edges)", "Tables/rows", "High-dim vectors"],
            ["Query type", "Traversal, pattern match", "SQL joins", "Similarity search"],
            ["Multi-hop", "Native, fast", "Complex JOINs", "Not supported"],
            ["Schema", "Flexible (schema-less)", "Rigid schema", "No schema"],
            ["Best for", "Relationships, reasoning", "Transactions, CRUD", "Semantic search"],
            ["Example", "Neo4j, Amazon Neptune", "PostgreSQL, MySQL", "Qdrant, Pinecone"],
        ],
        [4, 4, 3.5, 4]
    )

    # Section 2
    add_heading(doc, "Section 2: Entity Extraction", 2)
    add_heading(doc, "NER with spaCy & Stanza", 3)
    add_code_block(doc,
        "import spacy\n\n"
        "nlp = spacy.load('en_core_web_sm')\n"
        "doc = nlp('Apple CEO Tim Cook announced new products at WWDC in San Francisco.')\n\n"
        "for ent in doc.ents:\n"
        "    print(f'{ent.text:20} | {ent.label_:10} | {spacy.explain(ent.label_)}')\n\n"
        "# Output:\n"
        "# Apple                | ORG        | Companies, agencies\n"
        "# Tim Cook             | PERSON     | People, including fictional\n"
        "# WWDC                 | EVENT      | Named hurricanes, battles\n"
        "# San Francisco        | GPE        | Countries, cities, states"
    )

    add_heading(doc, "LLM-Based Entity Extraction", 3)
    add_body(doc, "Use LLMs with structured output prompts for more flexible, context-aware extraction:")
    add_code_block(doc,
        "import ollama, json\n\n"
        "prompt = '''\n"
        "Extract entities from the text. Return JSON format:\n"
        "{\"entities\": [{\"name\": str, \"type\": str, \"description\": str}],\n"
        " \"relations\": [{\"source\": str, \"relation\": str, \"target\": str}]}\n\n"
        "Text: {text}\n"
        "'''\n\n"
        "response = ollama.chat(\n"
        "    model='llama3.2',\n"
        "    messages=[{'role': 'user', 'content': prompt.format(text=doc_text)}],\n"
        "    format='json'\n"
        ")\n"
        "data = json.loads(response['message']['content'])"
    )

    add_heading(doc, "Entity Types Reference", 3)
    add_table(doc,
        ["Entity Type", "Examples", "Extraction Method"],
        [
            ["PERSON", "Tim Cook, Elon Musk", "NER / LLM"],
            ["ORG", "Apple, X-Company, OpenAI", "NER / LLM"],
            ["LOCATION", "Bangkok, Silicon Valley", "NER / LLM"],
            ["PRODUCT", "iPhone, ChatGPT, Qdrant", "LLM (context-aware)"],
            ["CONCEPT", "Machine Learning, RAG", "LLM (domain-specific)"],
            ["EVENT", "WWDC, CES, conference", "NER / LLM"],
        ],
        [4, 5, 7]
    )

    # Section 3
    add_heading(doc, "Section 3: Graph Construction Pipeline", 2)
    add_body(doc, "Building a knowledge graph involves a multi-step pipeline:")
    add_bullet(doc, "Step 1: Document ingestion and chunking")
    add_bullet(doc, "Step 2: Entity extraction from each chunk")
    add_bullet(doc, "Step 3: Relation extraction between entities")
    add_bullet(doc, "Step 4: Entity resolution (deduplication)")
    add_bullet(doc, "Step 5: Graph storage in Neo4j")

    add_code_block(doc,
        "# Step 1-2: Extract entities from documents\n"
        "entities = []\n"
        "relations = []\n"
        "for chunk in chunks:\n"
        "    result = extract_entities_llm(chunk)  # uses LLM\n"
        "    entities.extend(result['entities'])\n"
        "    relations.extend(result['relations'])\n\n"
        "# Step 4: Entity resolution\n"
        "from collections import defaultdict\n"
        "entity_map = defaultdict(list)\n"
        "for e in entities:\n"
        "    entity_map[e['name'].lower()].append(e)\n\n"
        "# Step 5: Store in Neo4j\n"
        "from neo4j import GraphDatabase\n"
        "driver = GraphDatabase.driver('bolt://localhost:7687',\n"
        "                              auth=('neo4j', 'password'))\n"
        "with driver.session() as session:\n"
        "    for e in entities:\n"
        "        session.run('MERGE (n:{type} {{name: $name}})',\n"
        "                    type=e['type'], name=e['name'])\n"
        "    for r in relations:\n"
        "        session.run(\n"
        "            'MATCH (a {{name: $src}}), (b {{name: $tgt}})'\n"
        "            'MERGE (a)-[:{rel}]->(b)',\n"
        "            src=r['source'], tgt=r['target'], rel=r['relation'].upper())"
    )

    # Section 4
    add_heading(doc, "Section 4: Neo4j & Cypher", 2)
    add_heading(doc, "Neo4j Setup", 3)
    add_code_block(doc,
        "# Start Neo4j with Docker\n"
        "docker run -d --name neo4j \\\n"
        "  -p 7474:7474 -p 7687:7687 \\\n"
        "  -e NEO4J_AUTH=neo4j/password \\\n"
        "  -v $(pwd)/neo4j_data:/data \\\n"
        "  neo4j:5.13"
    )

    add_heading(doc, "Cypher Query Language Basics", 3)
    add_code_block(doc,
        "// CREATE node\n"
        "CREATE (p:Person {name: 'Alice', role: 'Engineer'})\n\n"
        "// MATCH and RETURN\n"
        "MATCH (p:Person) WHERE p.name = 'Alice' RETURN p\n\n"
        "// CREATE relationship\n"
        "MATCH (a:Person {name: 'Alice'}), (b:Company {name: 'X-Company'})\n"
        "CREATE (a)-[:WORKS_AT {since: 2020}]->(b)\n\n"
        "// Traverse 2 hops\n"
        "MATCH (p:Person)-[:WORKS_AT]->(c:Company)-[:USES]->(t:Technology)\n"
        "RETURN p.name, c.name, t.name\n\n"
        "// Shortest path\n"
        "MATCH path = shortestPath((a:Person {name:'Alice'})-[*]-(b:Person {name:'Bob'}))\n"
        "RETURN path"
    )

    add_heading(doc, "10 Example Cypher Queries", 3)
    add_table(doc,
        ["#", "Use Case", "Query"],
        [
            ["1", "Find all persons", "MATCH (p:Person) RETURN p.name"],
            ["2", "Find employees of org", "MATCH (p)-[:WORKS_AT]->(c:Company {name:'X-Company'}) RETURN p"],
            ["3", "Count nodes by type", "MATCH (n) RETURN labels(n), count(n)"],
            ["4", "Find 2-hop connections", "MATCH (a)-[*2]-(b) WHERE a.name='Alice' RETURN b"],
            ["5", "Find common colleagues", "MATCH (a)-[:WORKS_AT]->(c)<-[:WORKS_AT]-(b) RETURN a,b,c"],
            ["6", "Delete node", "MATCH (n:Person {name:'Alice'}) DETACH DELETE n"],
            ["7", "Update property", "MATCH (n {name:'Alice'}) SET n.role='Manager' RETURN n"],
            ["8", "Find isolated nodes", "MATCH (n) WHERE NOT (n)--() RETURN n"],
            ["9", "Aggregate relations", "MATCH ()-[r]->() RETURN type(r), count(r) ORDER BY count(r) DESC"],
            ["10", "Path between nodes", "MATCH p=shortestPath((a {name:'Alice'})-[*]-(b {name:'Bob'})) RETURN p"],
        ],
        [1, 4, 11]
    )

    # Section 5
    add_heading(doc, "Section 5: GraphRAG", 2)
    add_body(doc, "GraphRAG combines vector similarity search with graph traversal to answer complex questions requiring multi-document, multi-hop reasoning.")
    add_body(doc, "Architecture:")
    add_bullet(doc, "Query → Embed → Vector search (retrieve relevant nodes by embedding)")
    add_bullet(doc, "Retrieved nodes → Graph traversal (expand context via relationships)")
    add_bullet(doc, "Expanded context → LLM → Answer")

    add_heading(doc, "GraphRAG-rs Configuration", 3)
    add_code_block(doc,
        "# graphrag-rs config.toml\n"
        "[storage]\n"
        "  [storage.neo4j]\n"
        "  uri = 'bolt://localhost:7687'\n"
        "  username = 'neo4j'\n"
        "  password = 'password'\n\n"
        "  [storage.qdrant]\n"
        "  url = 'http://localhost:6333'\n"
        "  collection = 'workshop_docs'\n\n"
        "[embedding]\n"
        "  model = 'bge-m3'\n"
        "  dimensions = 1024\n\n"
        "[retrieval]\n"
        "  top_k = 5\n"
        "  graph_hops = 2\n"
        "  community_level = 2"
    )

    # Section 6
    add_heading(doc, "Section 6: Advanced Retrieval", 2)
    add_heading(doc, "Multi-Hop Reasoning", 3)
    add_body(doc, "Graph-based reasoning allows answering questions that require connecting information across multiple documents:")
    add_body(doc, "Example: \"Which team members worked on projects that used Python and were delivered in Q4?\"")
    add_bullet(doc, "Step 1: Find all projects delivered in Q4 (graph query)")
    add_bullet(doc, "Step 2: Filter those using Python (graph property)")
    add_bullet(doc, "Step 3: Find team members linked to those projects")

    add_heading(doc, "Community Detection", 3)
    add_body(doc, "Identifies clusters of related entities in the knowledge graph. Enables topic-level summarization.")
    add_code_block(doc,
        "# Community detection with NetworkX\n"
        "import networkx as nx\n"
        "from networkx.algorithms.community import greedy_modularity_communities\n\n"
        "G = nx.Graph()\n"
        "# Add edges from Neo4j\n"
        "for rel in relations:\n"
        "    G.add_edge(rel['source'], rel['target'])\n\n"
        "communities = list(greedy_modularity_communities(G))\n"
        "print(f'Found {len(communities)} communities')"
    )

    add_heading(doc, "Graph-Enhanced Context", 3)
    add_body(doc, "Enrich RAG context with graph neighborhood information:")
    add_code_block(doc,
        "def graph_enhanced_retrieval(query: str) -> str:\n"
        "    # 1. Vector search\n"
        "    vector_chunks = vector_search(query, top_k=3)\n\n"
        "    # 2. Extract entities from query\n"
        "    entities = extract_entities_llm(query)\n\n"
        "    # 3. Graph expansion\n"
        "    with driver.session() as session:\n"
        "        for entity in entities:\n"
        "            result = session.run(\n"
        "                'MATCH (n {name:$name})-[r*1..2]-(m) RETURN m.name, type(r[-1]) LIMIT 10',\n"
        "                name=entity['name']\n"
        "            )\n"
        "            # Add graph context to prompt\n"
        "    return combined_context"
    )

    # Section 7
    add_heading(doc, "Section 7: Exercises & Reference", 2)
    add_heading(doc, "Exercise 1: Build a Company Knowledge Graph", 3)
    add_body(doc, "Task: Extract entities from 10 company documents and build a Neo4j graph.")
    add_bullet(doc, "Include: Person, Organization, Project, Technology node types")
    add_bullet(doc, "Hint: Use LLM extraction with JSON output format")

    add_heading(doc, "Exercise 2: Cypher Query Challenge", 3)
    add_body(doc, "Task: Write Cypher queries to answer: Who are the top 3 most connected employees? Which technologies appear in the most projects?")
    add_bullet(doc, "Hint: Use degree() function or count relationships in MATCH")

    add_heading(doc, "Exercise 3: GraphRAG Pipeline", 3)
    add_body(doc, "Task: Implement a GraphRAG pipeline that answers multi-hop questions using your knowledge graph.")
    add_bullet(doc, "Hint: Combine Qdrant search with Neo4j 2-hop expansion")

    add_heading(doc, "Reference", 3)
    add_bullet(doc, "Neo4j Documentation: https://neo4j.com/docs")
    add_bullet(doc, "GraphRAG paper: https://arxiv.org/abs/2404.16130")
    add_bullet(doc, "spaCy NER: https://spacy.io/usage/linguistic-features#named-entities")
    add_bullet(doc, "NetworkX: https://networkx.org/documentation")

    path = os.path.join(OUTPUT_DIR, "day2-handout-en.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path

# ─────────────────────────────────────────────────────────────
# DAY 2 THAI
# ─────────────────────────────────────────────────────────────
def create_day2_th():
    doc = setup_document()
    add_footer(doc, "AI Operating System Workshop | X-Company")
    add_cover(doc, 2, "Knowledge Graph & Advanced RAG", "Knowledge Graph และ RAG ขั้นสูง", "th")

    # Section 1
    add_heading(doc, "ส่วนที่ 1: Knowledge Graph พื้นฐาน", 2)
    add_body(doc, "Knowledge Graph (KG) คือการแสดงข้อมูลเชิงโครงสร้างของ entities ในโลกจริงและความสัมพันธ์ระหว่างกัน จัดระเบียบข้อมูลเป็นเครือข่าย nodes และ edges ช่วยให้ทำ multi-hop reasoning ที่ Vector DB ไม่สามารถทำได้")
    add_bullet(doc, "Nodes: Entities — บุคคล, องค์กร, สินค้า, แนวคิด")
    add_bullet(doc, "Edges: ความสัมพันธ์ — 'ทำงานที่', 'เป็นส่วนหนึ่งของ', 'เกี่ยวข้องกับ'")
    add_bullet(doc, "Properties: คุณสมบัติของ nodes/edges — ชื่อ, วันที่, คะแนนความเชื่อมั่น")

    add_heading(doc, "KG vs Relational DB vs Vector DB", 3)
    add_table(doc,
        ["คุณลักษณะ", "Knowledge Graph", "Relational DB", "Vector DB"],
        [
            ["โมเดลข้อมูล", "กราฟ (nodes/edges)", "ตาราง/แถว", "Vector หลายมิติ"],
            ["ประเภท Query", "Traversal, pattern match", "SQL joins", "Similarity search"],
            ["Multi-hop", "Native, เร็ว", "JOIN ซับซ้อน", "ไม่รองรับ"],
            ["Schema", "ยืดหยุ่น (schema-less)", "Schema แข็งกร้าว", "ไม่มี schema"],
            ["เหมาะกับ", "ความสัมพันธ์, reasoning", "Transaction, CRUD", "Semantic search"],
            ["ตัวอย่าง", "Neo4j, Amazon Neptune", "PostgreSQL, MySQL", "Qdrant, Pinecone"],
        ],
        [4, 4, 3.5, 4]
    )

    # Section 2
    add_heading(doc, "ส่วนที่ 2: การดึง Entity", 2)
    add_heading(doc, "NER ด้วย spaCy", 3)
    add_code_block(doc,
        "import spacy\n\n"
        "nlp = spacy.load('en_core_web_sm')\n"
        "doc = nlp('Apple CEO Tim Cook announced new products at WWDC in San Francisco.')\n\n"
        "for ent in doc.ents:\n"
        "    print(f'{ent.text:20} | {ent.label_:10} | {spacy.explain(ent.label_)}')\n\n"
        "# ผลลัพธ์:\n"
        "# Apple                | ORG        | บริษัท, หน่วยงาน\n"
        "# Tim Cook             | PERSON     | บุคคล\n"
        "# WWDC                 | EVENT      | งานอีเวนต์\n"
        "# San Francisco        | GPE        | ประเทศ, เมือง"
    )

    add_heading(doc, "การดึง Entity ด้วย LLM", 3)
    add_body(doc, "ใช้ LLM กับ prompt แบบ structured output เพื่อการดึง entity ที่ยืดหยุ่นและเข้าใจ context:")
    add_code_block(doc,
        "import ollama, json\n\n"
        "prompt = '''\n"
        "ดึง entities จากข้อความ ส่งกลับในรูปแบบ JSON:\n"
        "{\"entities\": [{\"name\": str, \"type\": str, \"description\": str}],\n"
        " \"relations\": [{\"source\": str, \"relation\": str, \"target\": str}]}\n\n"
        "ข้อความ: {text}\n"
        "'''\n\n"
        "response = ollama.chat(\n"
        "    model='llama3.2',\n"
        "    messages=[{'role': 'user', 'content': prompt.format(text=doc_text)}],\n"
        "    format='json'\n"
        ")\n"
        "data = json.loads(response['message']['content'])"
    )

    add_heading(doc, "ประเภท Entity อ้างอิง", 3)
    add_table(doc,
        ["ประเภท Entity", "ตัวอย่าง", "วิธีการดึง"],
        [
            ["PERSON", "สมชาย, Tim Cook, Elon Musk", "NER / LLM"],
            ["ORG", "Apple, X-Company, OpenAI", "NER / LLM"],
            ["LOCATION", "กรุงเทพ, Silicon Valley", "NER / LLM"],
            ["PRODUCT", "iPhone, ChatGPT, Qdrant", "LLM (เข้าใจ context)"],
            ["CONCEPT", "Machine Learning, RAG", "LLM (domain-specific)"],
            ["EVENT", "WWDC, CES, สัมมนา", "NER / LLM"],
        ],
        [4, 5, 7]
    )

    # Section 3
    add_heading(doc, "ส่วนที่ 3: การสร้าง Knowledge Graph", 2)
    add_body(doc, "กระบวนการสร้าง Knowledge Graph ประกอบด้วยหลายขั้นตอน:")
    add_bullet(doc, "ขั้นที่ 1: รับเอกสารและแบ่ง chunk")
    add_bullet(doc, "ขั้นที่ 2: ดึง entities จากแต่ละ chunk")
    add_bullet(doc, "ขั้นที่ 3: ดึงความสัมพันธ์ระหว่าง entities")
    add_bullet(doc, "ขั้นที่ 4: Entity resolution (ลบซ้ำ)")
    add_bullet(doc, "ขั้นที่ 5: เก็บ graph ใน Neo4j")

    add_code_block(doc,
        "# ขั้นที่ 1-2: ดึง entities จากเอกสาร\n"
        "entities, relations = [], []\n"
        "for chunk in chunks:\n"
        "    result = extract_entities_llm(chunk)\n"
        "    entities.extend(result['entities'])\n"
        "    relations.extend(result['relations'])\n\n"
        "# ขั้นที่ 5: เก็บใน Neo4j\n"
        "from neo4j import GraphDatabase\n"
        "driver = GraphDatabase.driver('bolt://localhost:7687',\n"
        "                              auth=('neo4j', 'password'))\n"
        "with driver.session() as session:\n"
        "    for e in entities:\n"
        "        session.run('MERGE (n:{type} {{name: $name}})',\n"
        "                    type=e['type'], name=e['name'])\n"
        "    for r in relations:\n"
        "        session.run(\n"
        "            'MATCH (a {{name: $src}}), (b {{name: $tgt}})'\n"
        "            'MERGE (a)-[:{rel}]->(b)',\n"
        "            src=r['source'], tgt=r['target'],\n"
        "            rel=r['relation'].upper())"
    )

    # Section 4
    add_heading(doc, "ส่วนที่ 4: Neo4j และ Cypher", 2)
    add_heading(doc, "ติดตั้ง Neo4j", 3)
    add_code_block(doc,
        "# เริ่มต้น Neo4j ด้วย Docker\n"
        "docker run -d --name neo4j \\\n"
        "  -p 7474:7474 -p 7687:7687 \\\n"
        "  -e NEO4J_AUTH=neo4j/password \\\n"
        "  -v $(pwd)/neo4j_data:/data \\\n"
        "  neo4j:5.13\n\n"
        "# เข้า Neo4j Browser: http://localhost:7474"
    )

    add_heading(doc, "พื้นฐาน Cypher", 3)
    add_code_block(doc,
        "// สร้าง node\n"
        "CREATE (p:Person {name: 'สมชาย', role: 'วิศวกร'})\n\n"
        "// MATCH และ RETURN\n"
        "MATCH (p:Person) WHERE p.name = 'สมชาย' RETURN p\n\n"
        "// สร้าง relationship\n"
        "MATCH (a:Person {name: 'สมชาย'}), (b:Company {name: 'X-Company'})\n"
        "CREATE (a)-[:WORKS_AT {since: 2020}]->(b)\n\n"
        "// ค้นหา 2 ขั้น\n"
        "MATCH (p:Person)-[:WORKS_AT]->(c:Company)-[:USES]->(t:Technology)\n"
        "RETURN p.name, c.name, t.name"
    )

    add_heading(doc, "10 ตัวอย่าง Cypher Queries", 3)
    add_table(doc,
        ["#", "กรณีใช้งาน", "Query"],
        [
            ["1", "หาบุคคลทั้งหมด", "MATCH (p:Person) RETURN p.name"],
            ["2", "หาพนักงานขององค์กร", "MATCH (p)-[:WORKS_AT]->(c {name:'X-Company'}) RETURN p"],
            ["3", "นับ nodes ตามประเภท", "MATCH (n) RETURN labels(n), count(n)"],
            ["4", "ค้นหาเชื่อมต่อ 2 ขั้น", "MATCH (a)-[*2]-(b) WHERE a.name='Alice' RETURN b"],
            ["5", "เพื่อนร่วมงาน", "MATCH (a)-[:WORKS_AT]->(c)<-[:WORKS_AT]-(b) RETURN a,b"],
            ["6", "ลบ node", "MATCH (n:Person {name:'Alice'}) DETACH DELETE n"],
            ["7", "อัพเดทคุณสมบัติ", "MATCH (n {name:'Alice'}) SET n.role='Manager' RETURN n"],
            ["8", "หา isolated nodes", "MATCH (n) WHERE NOT (n)--() RETURN n"],
            ["9", "รวม relations", "MATCH ()-[r]->() RETURN type(r), count(r) ORDER BY count(r) DESC"],
            ["10", "เส้นทางสั้นสุด", "MATCH p=shortestPath((a {name:'Alice'})-[*]-(b {name:'Bob'})) RETURN p"],
        ],
        [1, 4, 11]
    )

    # Section 5
    add_heading(doc, "ส่วนที่ 5: GraphRAG", 2)
    add_body(doc, "GraphRAG รวม vector similarity search กับ graph traversal เพื่อตอบคำถามซับซ้อนที่ต้องการ reasoning ข้ามเอกสารหลายไฟล์")
    add_body(doc, "สถาปัตยกรรม:")
    add_bullet(doc, "Query → Embed → Vector search (ค้นหา nodes ที่ใกล้เคียง)")
    add_bullet(doc, "Nodes ที่ดึงมา → Graph traversal (ขยาย context ผ่านความสัมพันธ์)")
    add_bullet(doc, "Context ที่ขยายแล้ว → LLM → คำตอบ")

    add_heading(doc, "การกำหนดค่า GraphRAG-rs", 3)
    add_code_block(doc,
        "# graphrag-rs config.toml\n"
        "[storage]\n"
        "  [storage.neo4j]\n"
        "  uri = 'bolt://localhost:7687'\n"
        "  username = 'neo4j'\n"
        "  password = 'password'\n\n"
        "  [storage.qdrant]\n"
        "  url = 'http://localhost:6333'\n"
        "  collection = 'workshop_docs'\n\n"
        "[embedding]\n"
        "  model = 'bge-m3'\n"
        "  dimensions = 1024\n\n"
        "[retrieval]\n"
        "  top_k = 5\n"
        "  graph_hops = 2\n"
        "  community_level = 2"
    )

    # Section 6
    add_heading(doc, "ส่วนที่ 6: Advanced Retrieval", 2)
    add_heading(doc, "Multi-Hop Reasoning", 3)
    add_body(doc, "การ reasoning บนกราฟช่วยตอบคำถามที่ต้องเชื่อมข้อมูลจากหลายเอกสาร:")
    add_body(doc, "ตัวอย่าง: \"สมาชิกทีมคนไหนทำงานในโปรเจกต์ที่ใช้ Python และส่งมอบใน Q4?\"")
    add_bullet(doc, "ขั้น 1: หาโปรเจกต์ทั้งหมดที่ส่งใน Q4 (graph query)")
    add_bullet(doc, "ขั้น 2: กรองเฉพาะที่ใช้ Python (graph property)")
    add_bullet(doc, "ขั้น 3: หาสมาชิกทีมที่เชื่อมกับโปรเจกต์เหล่านั้น")

    add_heading(doc, "Community Detection", 3)
    add_body(doc, "ระบุกลุ่ม entities ที่เกี่ยวข้องกันในกราฟ ช่วยสรุปข้อมูลระดับหัวข้อ")
    add_code_block(doc,
        "import networkx as nx\n"
        "from networkx.algorithms.community import greedy_modularity_communities\n\n"
        "G = nx.Graph()\n"
        "for rel in relations:\n"
        "    G.add_edge(rel['source'], rel['target'])\n\n"
        "communities = list(greedy_modularity_communities(G))\n"
        "print(f'พบ {len(communities)} กลุ่ม')"
    )

    add_heading(doc, "Graph-Enhanced Context", 3)
    add_code_block(doc,
        "def graph_enhanced_retrieval(query: str) -> str:\n"
        "    # 1. Vector search\n"
        "    vector_chunks = vector_search(query, top_k=3)\n\n"
        "    # 2. ดึง entities จาก query\n"
        "    entities = extract_entities_llm(query)\n\n"
        "    # 3. ขยาย context จากกราฟ\n"
        "    with driver.session() as session:\n"
        "        for entity in entities:\n"
        "            result = session.run(\n"
        "                'MATCH (n {name:$name})-[r*1..2]-(m)'\n"
        "                'RETURN m.name, type(r[-1]) LIMIT 10',\n"
        "                name=entity['name']\n"
        "            )\n"
        "    return combined_context"
    )

    # Section 7
    add_heading(doc, "ส่วนที่ 7: แบบฝึกหัดและเอกสารอ้างอิง", 2)
    add_heading(doc, "แบบฝึกหัดที่ 1: สร้าง Company Knowledge Graph", 3)
    add_body(doc, "งาน: ดึง entities จากเอกสาร 10 ไฟล์ และสร้าง Neo4j graph")
    add_bullet(doc, "รวม: ประเภท node — Person, Organization, Project, Technology")
    add_bullet(doc, "Hint: ใช้ LLM extraction พร้อม JSON output format")

    add_heading(doc, "แบบฝึกหัดที่ 2: Cypher Query Challenge", 3)
    add_body(doc, "งาน: เขียน Cypher queries เพื่อตอบ: พนักงาน 3 คนที่เชื่อมต่อมากที่สุดคือใคร? เทคโนโลยีใดปรากฏในโปรเจกต์มากที่สุด?")
    add_bullet(doc, "Hint: ใช้ฟังก์ชัน degree() หรือนับความสัมพันธ์ใน MATCH")

    add_heading(doc, "แบบฝึกหัดที่ 3: GraphRAG Pipeline", 3)
    add_body(doc, "งาน: สร้าง GraphRAG pipeline ที่ตอบคำถาม multi-hop โดยใช้ knowledge graph")
    add_bullet(doc, "Hint: รวม Qdrant search กับ Neo4j 2-hop expansion")

    add_heading(doc, "เอกสารอ้างอิง", 3)
    add_bullet(doc, "Neo4j Documentation: https://neo4j.com/docs")
    add_bullet(doc, "GraphRAG paper: https://arxiv.org/abs/2404.16130")
    add_bullet(doc, "spaCy NER: https://spacy.io/usage/linguistic-features")
    add_bullet(doc, "NetworkX: https://networkx.org/documentation")

    path = os.path.join(OUTPUT_DIR, "day2-handout-th.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path

# ─────────────────────────────────────────────────────────────
# DAY 3 ENGLISH
# ─────────────────────────────────────────────────────────────
def create_day3_en():
    doc = setup_document()
    add_footer(doc, "AI Operating System Workshop | X-Company")
    add_cover(doc, 3, "AI Agents & AI Operating System", "", "en")

    # Section 1
    add_heading(doc, "Section 1: AI Agent Architecture", 2)
    add_body(doc, "An AI Agent is an autonomous system that perceives its environment, reasons about it, and takes actions to achieve goals. Unlike a simple LLM chatbot, an agent has persistent memory, access to tools, and can plan multi-step tasks.")
    add_heading(doc, "Core Components", 3)
    add_table(doc,
        ["Component", "Role", "Examples"],
        [
            ["LLM (Brain)", "Reasoning, planning, language understanding", "LLaMA3, Mistral, GPT-4"],
            ["Tools", "Interact with external world", "Search, calculator, file I/O, APIs"],
            ["Memory", "Store context across sessions", "Vector DB, Redis, conversation history"],
            ["Planning", "Break goals into steps", "ReAct, Chain-of-Thought, Tree-of-Thought"],
        ],
        [3.5, 7, 5.5]
    )

    add_heading(doc, "Agent Loop: Perceive → Think → Act", 3)
    add_code_block(doc,
        "# Agent loop (simplified)\n"
        "while not goal_achieved:\n"
        "    # 1. PERCEIVE: observe current state\n"
        "    observation = get_observation()  # user input, tool result, etc.\n\n"
        "    # 2. THINK: reason about next action\n"
        "    thought = llm.think(observation, memory, available_tools)\n"
        "    action = parse_action(thought)  # e.g., {'tool': 'search', 'args': {'q': 'Python docs'}}\n\n"
        "    # 3. ACT: execute chosen action\n"
        "    result = execute_tool(action)\n"
        "    memory.append({'observation': observation, 'action': action, 'result': result})\n\n"
        "    # Check if done\n"
        "    goal_achieved = check_goal(result)"
    )

    # Section 2
    add_heading(doc, "Section 2: Tool-Use & Function Calling", 2)
    add_heading(doc, "Ollama Function Calling Setup", 3)
    add_code_block(doc,
        "import ollama\n\n"
        "# Define tools\n"
        "tools = [\n"
        "    {\n"
        "        'type': 'function',\n"
        "        'function': {\n"
        "            'name': 'calculator',\n"
        "            'description': 'Performs arithmetic calculations',\n"
        "            'parameters': {\n"
        "                'type': 'object',\n"
        "                'properties': {\n"
        "                    'expression': {'type': 'string',\n"
        "                                   'description': 'Math expression to evaluate'}\n"
        "                },\n"
        "                'required': ['expression']\n"
        "            }\n"
        "        }\n"
        "    }\n"
        "]\n\n"
        "response = ollama.chat(\n"
        "    model='llama3.2',\n"
        "    messages=[{'role': 'user', 'content': 'What is 15 * 24 + 7?'}],\n"
        "    tools=tools\n"
        ")"
    )

    add_heading(doc, "Example: Calculator + Search Tools", 3)
    add_code_block(doc,
        "import math, requests\n\n"
        "def calculator(expression: str) -> str:\n"
        "    try:\n"
        "        result = eval(expression, {'__builtins__': {}}, vars(math))\n"
        "        return str(result)\n"
        "    except Exception as e:\n"
        "        return f'Error: {e}'\n\n"
        "def web_search(query: str) -> str:\n"
        "    # Example using DuckDuckGo\n"
        "    resp = requests.get(f'https://api.duckduckgo.com/?q={query}&format=json')\n"
        "    data = resp.json()\n"
        "    return data.get('AbstractText', 'No results found')\n\n"
        "TOOL_MAP = {'calculator': calculator, 'web_search': web_search}\n\n"
        "def handle_tool_call(tool_call):\n"
        "    fn = TOOL_MAP[tool_call['function']['name']]\n"
        "    args = tool_call['function']['arguments']\n"
        "    return fn(**args)"
    )

    # Section 3
    add_heading(doc, "Section 3: RAG Agent", 2)
    add_body(doc, "A RAG Agent combines the knowledge retrieval power of RAG with the planning and tool-use capabilities of an agent. The agent can decide WHEN to query the knowledge base, WHAT to search for, and HOW to combine results.")
    add_code_block(doc,
        "import ollama\n"
        "from qdrant_client import QdrantClient\n\n"
        "client = QdrantClient('localhost', port=6333)\n\n"
        "def rag_tool(query: str) -> str:\n"
        "    q_vec = ollama.embeddings(model='bge-m3', prompt=query)['embedding']\n"
        "    results = client.search('workshop_docs', query_vector=q_vec, limit=5)\n"
        "    return '\\n'.join([r.payload['text'] for r in results])\n\n"
        "tools = [{\n"
        "    'type': 'function',\n"
        "    'function': {\n"
        "        'name': 'rag_tool',\n"
        "        'description': 'Search the company knowledge base for relevant information',\n"
        "        'parameters': {\n"
        "            'type': 'object',\n"
        "            'properties': {\n"
        "                'query': {'type': 'string', 'description': 'Search query'}\n"
        "            },\n"
        "            'required': ['query']\n"
        "        }\n"
        "    }\n"
        "}]\n\n"
        "TOOL_MAP = {'rag_tool': rag_tool}\n\n"
        "# Agent loop with RAG\n"
        "messages = [{'role': 'user', 'content': user_question}]\n"
        "while True:\n"
        "    response = ollama.chat(model='llama3.2', messages=messages, tools=tools)\n"
        "    if response['message'].get('tool_calls'):\n"
        "        for tc in response['message']['tool_calls']:\n"
        "            result = handle_tool_call(tc)\n"
        "            messages.append({'role': 'tool', 'content': result})\n"
        "    else:\n"
        "        print(response['message']['content'])\n"
        "        break"
    )

    # Section 4
    add_heading(doc, "Section 4: Multi-Agent Systems", 2)
    add_heading(doc, "Patterns Overview", 3)
    add_table(doc,
        ["Pattern", "Description", "Use Case", "Example"],
        [
            ["Sequential", "A → B → C (pipeline)", "Data processing chains", "Extract → Summarize → Translate"],
            ["Parallel", "A + B + C (concurrent)", "Independent sub-tasks", "Search multiple sources simultaneously"],
            ["Routing", "Router → A or B or C", "Specialized agents", "Classify query → route to specialist"],
            ["Orchestrator", "Manager → spawns workers", "Complex long-horizon tasks", "Project manager agent delegating"],
        ],
        [3, 5, 4, 4]
    )

    add_heading(doc, "Orchestrator Pattern", 3)
    add_code_block(doc,
        "class OrchestratorAgent:\n"
        "    def __init__(self, workers: dict):\n"
        "        self.workers = workers  # {'research': ResearchAgent, 'write': WriterAgent}\n\n"
        "    def run(self, task: str) -> str:\n"
        "        # 1. Plan the task\n"
        "        plan = self.llm_plan(task)\n"
        "        # plan = [{'agent': 'research', 'input': 'AI trends 2026'},\n"
        "        #         {'agent': 'write', 'input': 'Draft report from research'}]\n\n"
        "        context = {}\n"
        "        for step in plan:\n"
        "            agent_name = step['agent']\n"
        "            agent_input = step['input'].format(**context)\n"
        "            result = self.workers[agent_name].run(agent_input)\n"
        "            context[f'{agent_name}_result'] = result\n\n"
        "        return context.get('write_result', 'Task complete')"
    )

    # Section 5
    add_heading(doc, "Section 5: AI Operating System Concepts", 2)
    add_body(doc, "An AI Operating System (AI OS) is a platform that coordinates multiple AI agents, skills, data sources, and automation schedules — much like a traditional OS coordinates processes, I/O, and scheduling for software programs.")
    add_heading(doc, "Core Components", 3)
    add_table(doc,
        ["Component", "Function", "Analogy"],
        [
            ["Agents", "Task executors with LLM + tools", "OS processes"],
            ["Skills", "Reusable capability modules", "OS libraries / drivers"],
            ["Memory", "Persistent state storage", "File system + RAM"],
            ["Heartbeat", "Periodic background monitoring", "OS cron / system daemon"],
            ["Cron", "Scheduled task execution", "OS cron scheduler"],
            ["Message Bus", "Agent-to-agent communication", "IPC / message queue"],
        ],
        [4, 7, 5]
    )

    add_heading(doc, "Architecture Diagram Description", 3)
    add_body(doc, "The AI OS architecture consists of three layers:")
    add_bullet(doc, "Infrastructure Layer: LLM servers (Ollama), Vector DB (Qdrant), Graph DB (Neo4j), Object storage")
    add_bullet(doc, "Platform Layer: Agent runtime, Skill registry, Memory manager, Event bus, Scheduler")
    add_bullet(doc, "Application Layer: Domain agents (HR, Finance, Sales), User interfaces (Telegram, Web), APIs")

    # Section 6
    add_heading(doc, "Section 6: Evaluation with RAGAS", 2)
    add_body(doc, "RAGAS (Retrieval-Augmented Generation Assessment) provides automated metrics for evaluating RAG system quality without human labels.")

    add_heading(doc, "4 Core Metrics", 3)
    add_table(doc,
        ["Metric", "What It Measures", "Formula", "Target"],
        [
            ["Faithfulness", "Is the answer grounded in context?", "Claims supported / Total claims", "> 0.8"],
            ["Answer Relevancy", "Does the answer address the question?", "Cosine sim(answer, question)", "> 0.8"],
            ["Context Precision", "Are retrieved chunks relevant?", "Relevant chunks / Total retrieved", "> 0.7"],
            ["Context Recall", "Are all key facts retrieved?", "Key facts retrieved / Total key facts", "> 0.7"],
        ],
        [3.5, 5.5, 4.5, 2.5]
    )

    add_heading(doc, "Running RAGAS", 3)
    add_code_block(doc,
        "from ragas import evaluate\n"
        "from ragas.metrics import faithfulness, answer_relevancy, context_precision, context_recall\n"
        "from datasets import Dataset\n\n"
        "# Prepare evaluation dataset\n"
        "data = {\n"
        "    'question': ['What is the refund policy?'],\n"
        "    'answer': ['Refunds are processed within 7 business days.'],\n"
        "    'contexts': [['Customers may request refunds within 30 days...']],\n"
        "    'ground_truth': ['Refunds take 7 business days to process.']\n"
        "}\n"
        "dataset = Dataset.from_dict(data)\n\n"
        "result = evaluate(\n"
        "    dataset,\n"
        "    metrics=[faithfulness, answer_relevancy, context_precision, context_recall]\n"
        ")\n"
        "print(result)"
    )

    add_heading(doc, "LLM-as-Judge Approach", 3)
    add_body(doc, "Use an LLM to evaluate response quality when ground truth is unavailable:")
    add_code_block(doc,
        "judge_prompt = '''\n"
        "Rate this RAG response on a scale of 1-5 for each criterion:\n"
        "- Faithfulness: Is every claim supported by the context?\n"
        "- Relevance: Does it answer the question?\n"
        "- Completeness: Are all aspects of the question addressed?\n\n"
        "Question: {question}\n"
        "Context: {context}\n"
        "Answer: {answer}\n\n"
        "Return JSON: {{\"faithfulness\": int, \"relevance\": int, \"completeness\": int, \"reasoning\": str}}\n"
        "'''"
    )

    add_heading(doc, "A/B Testing Methodology", 3)
    add_bullet(doc, "Define metric baseline (current system RAGAS scores)")
    add_bullet(doc, "Implement variant B (different chunking, model, prompt, etc.)")
    add_bullet(doc, "Run evaluation on same test set for both A and B")
    add_bullet(doc, "Statistical significance: use t-test or bootstrap (n > 30 samples)")
    add_bullet(doc, "Decision threshold: > 5% improvement on primary metric")

    # Section 7
    add_heading(doc, "Section 7: Capstone Project", 2)
    add_heading(doc, "Build an Enterprise AI Assistant", 3)
    add_body(doc, "Project Goal: Build a production-ready AI assistant for your organization that integrates all workshop concepts.")
    add_heading(doc, "Requirements", 3)
    add_bullet(doc, "Document ingestion: Support PDF, DOCX, with Thai language processing")
    add_bullet(doc, "Hybrid search: Dense + BM25 with RRF fusion")
    add_bullet(doc, "Knowledge graph: Entity extraction and Neo4j storage")
    add_bullet(doc, "Multi-tool agent: At least 3 tools (RAG, search, calculator)")
    add_bullet(doc, "Evaluation: RAGAS scores on 50-question test set")
    add_bullet(doc, "UI: Telegram bot or simple web interface")

    add_heading(doc, "Evaluation Criteria", 3)
    add_table(doc,
        ["Criterion", "Weight", "Minimum Score"],
        [
            ["Faithfulness (RAGAS)", "25%", "0.75"],
            ["Answer Relevancy (RAGAS)", "25%", "0.75"],
            ["Context Precision (RAGAS)", "20%", "0.70"],
            ["Code quality & documentation", "15%", "Pass"],
            ["Demo & presentation", "15%", "Pass"],
        ],
        [6, 3, 4]
    )

    # Section 8
    add_heading(doc, "Section 8: Exercises & Reference", 2)
    add_heading(doc, "Exercise 1: Tool-Equipped Agent", 3)
    add_body(doc, "Task: Build an agent with 3 tools: web search, calculator, and RAG. Handle tool chaining.")
    add_bullet(doc, "Hint: Use while loop until no more tool_calls in response")

    add_heading(doc, "Exercise 2: RAGAS Evaluation", 3)
    add_body(doc, "Task: Create 20 Q&A pairs from your RAG system, evaluate with RAGAS, identify weak areas.")
    add_bullet(doc, "Hint: Focus on Faithfulness — most RAG systems fail here first")

    add_heading(doc, "Exercise 3: Multi-Agent Orchestration", 3)
    add_body(doc, "Task: Build a 2-agent system: Research agent + Summary agent. Orchestrate them sequentially.")
    add_bullet(doc, "Hint: Pass research_result as context to the summary agent")

    add_heading(doc, "Reference", 3)
    add_bullet(doc, "Ollama Tools API: https://ollama.ai/blog/tool-support")
    add_bullet(doc, "RAGAS: https://docs.ragas.io")
    add_bullet(doc, "LangGraph (Multi-Agent): https://langchain-ai.github.io/langgraph")
    add_bullet(doc, "OpenClaw (AI OS): https://openclaw.ai")
    add_bullet(doc, "ReAct Paper: https://arxiv.org/abs/2210.03629")

    path = os.path.join(OUTPUT_DIR, "day3-handout-en.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path

# ─────────────────────────────────────────────────────────────
# DAY 3 THAI
# ─────────────────────────────────────────────────────────────
def create_day3_th():
    doc = setup_document()
    add_footer(doc, "AI Operating System Workshop | X-Company")
    add_cover(doc, 3, "AI Agents & AI Operating System", "AI Agents และ AI Operating System", "th")

    # Section 1
    add_heading(doc, "ส่วนที่ 1: สถาปัตยกรรม AI Agent", 2)
    add_body(doc, "AI Agent คือระบบอัตโนมัติที่รับรู้สภาพแวดล้อม ใช้เหตุผล และดำเนินการเพื่อบรรลุเป้าหมาย ต่างจาก LLM chatbot ทั่วไป ตรงที่ agent มี memory ถาวร เข้าถึง tools ได้ และวางแผนงานหลายขั้นตอนได้")

    add_heading(doc, "องค์ประกอบหลัก", 3)
    add_table(doc,
        ["องค์ประกอบ", "บทบาท", "ตัวอย่าง"],
        [
            ["LLM (สมอง)", "Reasoning, planning, เข้าใจภาษา", "LLaMA3, Mistral, GPT-4"],
            ["Tools", "โต้ตอบกับโลกภายนอก", "ค้นหาเว็บ, คำนวณ, ไฟล์, APIs"],
            ["Memory", "เก็บ context ข้ามเซสชัน", "Vector DB, Redis, conversation history"],
            ["Planning", "แบ่งเป้าหมายเป็นขั้นตอน", "ReAct, Chain-of-Thought, Tree-of-Thought"],
        ],
        [3.5, 7, 5.5]
    )

    add_heading(doc, "วงจร Agent: รับรู้ → คิด → กระทำ", 3)
    add_code_block(doc,
        "# วงจร Agent (simplified)\n"
        "while not goal_achieved:\n"
        "    # 1. รับรู้: สังเกตสถานะปัจจุบัน\n"
        "    observation = get_observation()  # input ผู้ใช้, ผล tool ฯลฯ\n\n"
        "    # 2. คิด: วิเคราะห์การกระทำถัดไป\n"
        "    thought = llm.think(observation, memory, available_tools)\n"
        "    action = parse_action(thought)  # เช่น {'tool': 'search', 'args': {'q': 'Python docs'}}\n\n"
        "    # 3. กระทำ: ดำเนินการ tool ที่เลือก\n"
        "    result = execute_tool(action)\n"
        "    memory.append({'observation': observation, 'action': action, 'result': result})\n\n"
        "    goal_achieved = check_goal(result)"
    )

    # Section 2
    add_heading(doc, "ส่วนที่ 2: Tool-Use และ Function Calling", 2)
    add_heading(doc, "การตั้งค่า Function Calling ของ Ollama", 3)
    add_code_block(doc,
        "import ollama\n\n"
        "# กำหนด tools\n"
        "tools = [\n"
        "    {\n"
        "        'type': 'function',\n"
        "        'function': {\n"
        "            'name': 'calculator',\n"
        "            'description': 'คำนวณทางคณิตศาสตร์',\n"
        "            'parameters': {\n"
        "                'type': 'object',\n"
        "                'properties': {\n"
        "                    'expression': {'type': 'string',\n"
        "                                   'description': 'นิพจน์คณิตศาสตร์ที่ต้องการคำนวณ'}\n"
        "                },\n"
        "                'required': ['expression']\n"
        "            }\n"
        "        }\n"
        "    }\n"
        "]\n\n"
        "response = ollama.chat(\n"
        "    model='llama3.2',\n"
        "    messages=[{'role': 'user', 'content': '15 * 24 + 7 เท่ากับเท่าไหร่?'}],\n"
        "    tools=tools\n"
        ")"
    )

    add_heading(doc, "ตัวอย่าง: Calculator + Search Tools", 3)
    add_code_block(doc,
        "import math, requests\n\n"
        "def calculator(expression: str) -> str:\n"
        "    try:\n"
        "        result = eval(expression, {'__builtins__': {}}, vars(math))\n"
        "        return str(result)\n"
        "    except Exception as e:\n"
        "        return f'Error: {e}'\n\n"
        "def web_search(query: str) -> str:\n"
        "    resp = requests.get(f'https://api.duckduckgo.com/?q={query}&format=json')\n"
        "    return resp.json().get('AbstractText', 'ไม่พบผลลัพธ์')\n\n"
        "TOOL_MAP = {'calculator': calculator, 'web_search': web_search}\n\n"
        "def handle_tool_call(tool_call):\n"
        "    fn = TOOL_MAP[tool_call['function']['name']]\n"
        "    args = tool_call['function']['arguments']\n"
        "    return fn(**args)"
    )

    # Section 3
    add_heading(doc, "ส่วนที่ 3: RAG Agent", 2)
    add_body(doc, "RAG Agent รวมพลังการดึงความรู้ของ RAG เข้ากับความสามารถในการวางแผนและใช้ tools ของ agent โดย agent จะตัดสินใจว่า เมื่อไหร่จะค้นฐานความรู้ ค้นหาอะไร และรวมผลลัพธ์อย่างไร")
    add_code_block(doc,
        "import ollama\n"
        "from qdrant_client import QdrantClient\n\n"
        "client = QdrantClient('localhost', port=6333)\n\n"
        "def rag_tool(query: str) -> str:\n"
        "    q_vec = ollama.embeddings(model='bge-m3', prompt=query)['embedding']\n"
        "    results = client.search('workshop_docs', query_vector=q_vec, limit=5)\n"
        "    return '\\n'.join([r.payload['text'] for r in results])\n\n"
        "TOOL_MAP = {'rag_tool': rag_tool}\n\n"
        "# วงจร Agent กับ RAG\n"
        "messages = [{'role': 'user', 'content': user_question}]\n"
        "while True:\n"
        "    response = ollama.chat(model='llama3.2', messages=messages, tools=tools)\n"
        "    if response['message'].get('tool_calls'):\n"
        "        for tc in response['message']['tool_calls']:\n"
        "            result = handle_tool_call(tc)\n"
        "            messages.append({'role': 'tool', 'content': result})\n"
        "    else:\n"
        "        print(response['message']['content'])\n"
        "        break"
    )

    # Section 4
    add_heading(doc, "ส่วนที่ 4: ระบบ Multi-Agent", 2)
    add_heading(doc, "ภาพรวม Patterns", 3)
    add_table(doc,
        ["Pattern", "คำอธิบาย", "กรณีใช้งาน", "ตัวอย่าง"],
        [
            ["Sequential", "A → B → C (pipeline)", "สายการประมวลผล", "ดึง → สรุป → แปล"],
            ["Parallel", "A + B + C (concurrent)", "งานอิสระหลายอย่าง", "ค้นหาหลายแหล่งพร้อมกัน"],
            ["Routing", "Router → A หรือ B หรือ C", "Agents เฉพาะทาง", "จัดประเภทคำถาม → ส่งไปผู้เชี่ยวชาญ"],
            ["Orchestrator", "Manager → สร้าง workers", "งานระยะยาวซับซ้อน", "Project manager agent มอบหมายงาน"],
        ],
        [3, 5, 4, 4]
    )

    add_heading(doc, "Orchestrator Pattern", 3)
    add_code_block(doc,
        "class OrchestratorAgent:\n"
        "    def __init__(self, workers: dict):\n"
        "        self.workers = workers\n\n"
        "    def run(self, task: str) -> str:\n"
        "        # 1. วางแผนงาน\n"
        "        plan = self.llm_plan(task)\n"
        "        # plan = [{'agent': 'research', 'input': 'เทรนด์ AI 2026'},\n"
        "        #         {'agent': 'write', 'input': 'ร่างรายงานจากงานวิจัย'}]\n\n"
        "        context = {}\n"
        "        for step in plan:\n"
        "            agent_name = step['agent']\n"
        "            result = self.workers[agent_name].run(step['input'])\n"
        "            context[f'{agent_name}_result'] = result\n\n"
        "        return context.get('write_result', 'งานเสร็จสิ้น')"
    )

    # Section 5
    add_heading(doc, "ส่วนที่ 5: AI Operating System", 2)
    add_body(doc, "AI Operating System (AI OS) คือแพลตฟอร์มที่ประสาน AI agents, skills, แหล่งข้อมูล และการทำงานอัตโนมัติตามเวลา — คล้ายกับ OS ทั่วไปที่ประสาน processes, I/O และ scheduling")

    add_heading(doc, "องค์ประกอบหลัก", 3)
    add_table(doc,
        ["องค์ประกอบ", "หน้าที่", "Analogy"],
        [
            ["Agents", "ดำเนินงานด้วย LLM + tools", "OS processes"],
            ["Skills", "โมดูล capability ที่นำกลับมาใช้ได้", "OS libraries / drivers"],
            ["Memory", "เก็บ state ถาวร", "File system + RAM"],
            ["Heartbeat", "ตรวจสอบพื้นหลังเป็นระยะ", "OS cron / system daemon"],
            ["Cron", "รันงานตามเวลาที่กำหนด", "OS cron scheduler"],
            ["Message Bus", "การสื่อสาร agent-to-agent", "IPC / message queue"],
        ],
        [4, 7, 5]
    )

    add_heading(doc, "คำอธิบายสถาปัตยกรรม", 3)
    add_body(doc, "สถาปัตยกรรม AI OS ประกอบด้วย 3 ชั้น:")
    add_bullet(doc, "Infrastructure Layer: LLM servers (Ollama), Vector DB (Qdrant), Graph DB (Neo4j), Object storage")
    add_bullet(doc, "Platform Layer: Agent runtime, Skill registry, Memory manager, Event bus, Scheduler")
    add_bullet(doc, "Application Layer: Domain agents (HR, Finance, Sales), User interfaces (Telegram, Web), APIs")

    # Section 6
    add_heading(doc, "ส่วนที่ 6: การประเมินด้วย RAGAS", 2)
    add_body(doc, "RAGAS (Retrieval-Augmented Generation Assessment) ให้ metrics อัตโนมัติในการประเมินคุณภาพระบบ RAG โดยไม่ต้องใช้ label จากมนุษย์")

    add_heading(doc, "4 Metrics หลัก", 3)
    add_table(doc,
        ["Metric", "วัดอะไร", "สูตร", "เป้าหมาย"],
        [
            ["Faithfulness", "คำตอบอ้างอิง context ครบไหม?", "Claims ที่มีหลักฐาน / Claims ทั้งหมด", "> 0.8"],
            ["Answer Relevancy", "คำตอบตอบคำถามไหม?", "Cosine sim(answer, question)", "> 0.8"],
            ["Context Precision", "chunks ที่ดึงมาเกี่ยวข้องไหม?", "Chunks เกี่ยวข้อง / ทั้งหมดที่ดึง", "> 0.7"],
            ["Context Recall", "ดึงข้อเท็จจริงสำคัญครบไหม?", "ข้อเท็จจริงที่ดึง / ทั้งหมด", "> 0.7"],
        ],
        [3.5, 5.5, 4.5, 2.5]
    )

    add_heading(doc, "รัน RAGAS", 3)
    add_code_block(doc,
        "from ragas import evaluate\n"
        "from ragas.metrics import faithfulness, answer_relevancy, context_precision, context_recall\n"
        "from datasets import Dataset\n\n"
        "data = {\n"
        "    'question': ['นโยบายการคืนเงินเป็นอย่างไร?'],\n"
        "    'answer': ['การคืนเงินใช้เวลา 7 วันทำการ'],\n"
        "    'contexts': [['ลูกค้าสามารถขอคืนเงินได้ภายใน 30 วัน...']],\n"
        "    'ground_truth': ['การคืนเงินใช้เวลา 7 วันทำการ']\n"
        "}\n"
        "dataset = Dataset.from_dict(data)\n\n"
        "result = evaluate(\n"
        "    dataset,\n"
        "    metrics=[faithfulness, answer_relevancy, context_precision, context_recall]\n"
        ")\n"
        "print(result)"
    )

    add_heading(doc, "แนวทาง LLM-as-Judge", 3)
    add_body(doc, "ใช้ LLM ประเมินคุณภาพคำตอบเมื่อไม่มี ground truth:")
    add_code_block(doc,
        "judge_prompt = '''\n"
        "ให้คะแนน RAG response นี้ 1-5 ในแต่ละเกณฑ์:\n"
        "- Faithfulness: ทุก claim มีหลักฐานจาก context?\n"
        "- Relevance: ตอบคำถามได้?\n"
        "- Completeness: ครอบคลุมทุกด้านของคำถาม?\n\n"
        "คำถาม: {question}\n"
        "Context: {context}\n"
        "คำตอบ: {answer}\n\n"
        "ส่งกลับ JSON: {{\"faithfulness\": int, \"relevance\": int, \"completeness\": int}}\n"
        "'''"
    )

    add_heading(doc, "A/B Testing Methodology", 3)
    add_bullet(doc, "กำหนด baseline metric (คะแนน RAGAS ของระบบปัจจุบัน)")
    add_bullet(doc, "สร้าง variant B (เปลี่ยน chunking, โมเดล, prompt ฯลฯ)")
    add_bullet(doc, "รัน evaluation บน test set เดียวกันทั้ง A และ B")
    add_bullet(doc, "ความมีนัยสำคัญทางสถิติ: ใช้ t-test หรือ bootstrap (n > 30 ตัวอย่าง)")
    add_bullet(doc, "เกณฑ์ตัดสิน: ปรับปรุง > 5% บน metric หลัก")

    # Section 7
    add_heading(doc, "ส่วนที่ 7: โปรเจกต์สุดท้าย (Capstone)", 2)
    add_heading(doc, "สร้าง Enterprise AI Assistant", 3)
    add_body(doc, "เป้าหมาย: สร้าง AI assistant พร้อม production สำหรับองค์กร ที่รวมแนวคิดทั้งหมดจากการอบรม")
    add_heading(doc, "ข้อกำหนด", 3)
    add_bullet(doc, "Document ingestion: รองรับ PDF, DOCX พร้อมประมวลผลภาษาไทย")
    add_bullet(doc, "Hybrid search: Dense + BM25 กับ RRF fusion")
    add_bullet(doc, "Knowledge graph: ดึง entity และเก็บใน Neo4j")
    add_bullet(doc, "Multi-tool agent: อย่างน้อย 3 tools (RAG, ค้นหา, คำนวณ)")
    add_bullet(doc, "Evaluation: คะแนน RAGAS บน test set 50 คำถาม")
    add_bullet(doc, "UI: Telegram bot หรือ web interface อย่างง่าย")

    add_heading(doc, "เกณฑ์การประเมิน", 3)
    add_table(doc,
        ["เกณฑ์", "น้ำหนัก", "คะแนนขั้นต่ำ"],
        [
            ["Faithfulness (RAGAS)", "25%", "0.75"],
            ["Answer Relevancy (RAGAS)", "25%", "0.75"],
            ["Context Precision (RAGAS)", "20%", "0.70"],
            ["คุณภาพโค้ดและเอกสาร", "15%", "ผ่าน"],
            ["Demo และการนำเสนอ", "15%", "ผ่าน"],
        ],
        [6, 3, 4]
    )

    # Section 8
    add_heading(doc, "ส่วนที่ 8: แบบฝึกหัดและเอกสารอ้างอิง", 2)
    add_heading(doc, "แบบฝึกหัดที่ 1: Agent พร้อม Tools", 3)
    add_body(doc, "งาน: สร้าง agent ที่มี 3 tools: web search, calculator, และ RAG จัดการ tool chaining")
    add_bullet(doc, "Hint: ใช้ while loop จนกว่าจะไม่มี tool_calls ใน response")

    add_heading(doc, "แบบฝึกหัดที่ 2: RAGAS Evaluation", 3)
    add_body(doc, "งาน: สร้างชุด Q&A 20 คู่จากระบบ RAG ของคุณ ประเมินด้วย RAGAS หาจุดอ่อน")
    add_bullet(doc, "Hint: มุ่งเน้น Faithfulness — ระบบ RAG ส่วนใหญ่ล้มเหลวตรงนี้ก่อน")

    add_heading(doc, "แบบฝึกหัดที่ 3: Multi-Agent Orchestration", 3)
    add_body(doc, "งาน: สร้างระบบ 2 agents: Research agent + Summary agent จัดการแบบ sequential")
    add_bullet(doc, "Hint: ส่ง research_result เป็น context ให้ summary agent")

    add_heading(doc, "เอกสารอ้างอิง", 3)
    add_bullet(doc, "Ollama Tools API: https://ollama.ai/blog/tool-support")
    add_bullet(doc, "RAGAS: https://docs.ragas.io")
    add_bullet(doc, "LangGraph (Multi-Agent): https://langchain-ai.github.io/langgraph")
    add_bullet(doc, "OpenClaw (AI OS): https://openclaw.ai")
    add_bullet(doc, "ReAct Paper: https://arxiv.org/abs/2210.03629")

    path = os.path.join(OUTPUT_DIR, "day3-handout-th.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print("Creating AI Workshop Handouts...")
    create_day1_en()
    create_day1_th()
    create_day2_en()
    create_day2_th()
    create_day3_en()
    create_day3_th()
    print("\nAll 6 DOCX files created successfully!")
