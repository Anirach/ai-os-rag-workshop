#!/usr/bin/env node
'use strict';

const pptxgen = require('/Users/pplus/.nvm/versions/node/v22.22.0/lib/node_modules/pptxgenjs');
const path = require('path');

const OUT_DIR = path.join(__dirname, 'slides');

// ─── THEME ────────────────────────────────────────────────────────────────────
const NAVY  = '1B2A4A';
const BLUE  = '2E5090';
const TEAL  = '0D7377';
const WHITE = 'FFFFFF';
const LGRAY = 'F0F2F5';
const DGRAY = '444444';
const MGRAY = 'AAAAAA';

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function newPptx() {
  const prs = new pptxgen();
  prs.layout = 'LAYOUT_WIDE'; // 16:9
  prs.defineLayout({ name: 'LAYOUT_WIDE', width: 13.33, height: 7.5 });
  return prs;
}

/** Title slide */
function addTitleSlide(prs, title, subtitle, tag) {
  const slide = prs.addSlide();
  // Full background
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: NAVY } });
  // Accent bar
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 5.8, w: 13.33, h: 0.08, fill: { color: TEAL } });
  // Tag (day label)
  if (tag) {
    slide.addText(tag, { x: 0.5, y: 0.4, w: 12, h: 0.5, fontSize: 14, color: TEAL, bold: false, fontFace: 'Arial' });
  }
  // Main title
  slide.addText(title, {
    x: 0.8, y: 1.8, w: 11.7, h: 2.2,
    fontSize: 40, bold: true, color: WHITE, fontFace: 'Arial',
    align: 'center', valign: 'middle', wrap: true
  });
  // Subtitle
  slide.addText(subtitle, {
    x: 0.8, y: 4.1, w: 11.7, h: 1,
    fontSize: 20, color: MGRAY, fontFace: 'Arial', align: 'center'
  });
  // Company
  slide.addText('X-Company', {
    x: 0.8, y: 6.6, w: 11.7, h: 0.5,
    fontSize: 13, color: TEAL, fontFace: 'Arial', align: 'center'
  });
  return slide;
}

/** Standard content slide with title bar */
function addContentSlide(prs, title, bullets) {
  const slide = prs.addSlide();
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.1, fill: { color: NAVY } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 1.1, w: 0.08, h: 6.4, fill: { color: TEAL } });
  slide.addText(title, { x: 0.3, y: 0.1, w: 12.7, h: 0.9, fontSize: 24, bold: true, color: WHITE, fontFace: 'Arial', valign: 'middle' });

  const bulletObjs = bullets.map(b => {
    if (typeof b === 'string') return { text: b, options: { fontSize: 18, color: DGRAY, bullet: { type: 'bullet' }, paraSpaceAfter: 6 } };
    return b;
  });
  slide.addText(bulletObjs, { x: 0.5, y: 1.3, w: 12.5, h: 5.9, fontFace: 'Arial', valign: 'top' });
  return slide;
}

/** Slide with title bar + custom body callback */
function addDiagramSlide(prs, title, bodyFn) {
  const slide = prs.addSlide();
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.1, fill: { color: NAVY } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 1.1, w: 0.08, h: 6.4, fill: { color: TEAL } });
  slide.addText(title, { x: 0.3, y: 0.1, w: 12.7, h: 0.9, fontSize: 24, bold: true, color: WHITE, fontFace: 'Arial', valign: 'middle' });
  bodyFn(slide, prs);
  return slide;
}

/** Box helper */
function box(slide, prs, x, y, w, h, label, color, textColor) {
  color = color || BLUE;
  textColor = textColor || WHITE;
  slide.addShape(prs.ShapeType.roundRect, { x, y, w, h, fill: { color }, rectRadius: 0.1, line: { color: WHITE, width: 0 } });
  slide.addText(label, { x, y, w, h, fontSize: 14, bold: true, color: textColor, fontFace: 'Arial', align: 'center', valign: 'middle', wrap: true });
}

/** Arrow right */
function arrowRight(slide, prs, x, y) {
  slide.addShape(prs.ShapeType.rightArrow, { x, y, w: 0.55, h: 0.35, fill: { color: TEAL } });
}

/** Arrow down */
function arrowDown(slide, prs, x, y) {
  slide.addShape(prs.ShapeType.downArrow, { x, y, w: 0.35, h: 0.5, fill: { color: TEAL } });
}

/** Code block */
function addCodeSlide(prs, title, codeLines) {
  const slide = prs.addSlide();
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.1, fill: { color: NAVY } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 1.1, w: 0.08, h: 6.4, fill: { color: TEAL } });
  slide.addText(title, { x: 0.3, y: 0.1, w: 12.7, h: 0.9, fontSize: 24, bold: true, color: WHITE, fontFace: 'Arial', valign: 'middle' });
  // Code box
  slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.3, w: 12.5, h: 6, fill: { color: '1E1E2E' } });
  const codeText = codeLines.map(l => ({ text: l, options: { fontSize: 13, color: 'D4D4D4', fontFace: 'Courier New', paraSpaceAfter: 2 } }));
  slide.addText(codeText, { x: 0.7, y: 1.5, w: 12.1, h: 5.7, valign: 'top' });
  return slide;
}

/** Section divider slide */
function addSectionSlide(prs, label, title) {
  const slide = prs.addSlide();
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: BLUE } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 3.5, w: 13.33, h: 0.06, fill: { color: TEAL } });
  slide.addText(label, { x: 0.5, y: 1.8, w: 12.3, h: 0.6, fontSize: 16, color: TEAL, fontFace: 'Arial', align: 'center' });
  slide.addText(title, { x: 0.5, y: 2.5, w: 12.3, h: 1.5, fontSize: 34, bold: true, color: WHITE, fontFace: 'Arial', align: 'center', valign: 'middle' });
  return slide;
}

/** Two-column comparison */
function addCompareSlide(prs, title, leftTitle, leftItems, rightTitle, rightItems, leftColor, rightColor) {
  leftColor = leftColor || BLUE;
  rightColor = rightColor || TEAL;
  const slide = prs.addSlide();
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.1, fill: { color: NAVY } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 1.1, w: 0.08, h: 6.4, fill: { color: TEAL } });
  slide.addText(title, { x: 0.3, y: 0.1, w: 12.7, h: 0.9, fontSize: 24, bold: true, color: WHITE, fontFace: 'Arial', valign: 'middle' });

  // Left column
  slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.3, w: 6.0, h: 0.55, fill: { color: leftColor } });
  slide.addText(leftTitle, { x: 0.5, y: 1.3, w: 6.0, h: 0.55, fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial', align: 'center', valign: 'middle' });
  const leftBullets = leftItems.map(i => ({ text: i, options: { fontSize: 15, color: DGRAY, bullet: { type: 'bullet' }, paraSpaceAfter: 5 } }));
  slide.addText(leftBullets, { x: 0.5, y: 1.95, w: 6.0, h: 5.3, fontFace: 'Arial', valign: 'top' });

  // Right column
  slide.addShape(prs.ShapeType.rect, { x: 6.8, y: 1.3, w: 6.0, h: 0.55, fill: { color: rightColor } });
  slide.addText(rightTitle, { x: 6.8, y: 1.3, w: 6.0, h: 0.55, fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial', align: 'center', valign: 'middle' });
  const rightBullets = rightItems.map(i => ({ text: i, options: { fontSize: 15, color: DGRAY, bullet: { type: 'bullet' }, paraSpaceAfter: 5 } }));
  slide.addText(rightBullets, { x: 6.8, y: 1.95, w: 6.0, h: 5.3, fontFace: 'Arial', valign: 'top' });

  return slide;
}

// ─── DAY 1 CONTENT ────────────────────────────────────────────────────────────

const DAY1 = {
  en: {
    title: 'Day 1: Data Engineering & RAG Foundation',
    workshopName: 'AI Workshop: Build Enterprise AI with RAG & Knowledge Graphs',
    agenda: [
      'Introduction to Retrieval-Augmented Generation (RAG)',
      'Document Processing & Thai NLP Challenges',
      'Chunking Strategies & Embedding Models',
      'Vector Databases & Hybrid Search',
      'Hands-on: Build a Thai RAG Pipeline',
    ],
    ragWhat: { title: 'What is RAG?', sub: 'Retrieval-Augmented Generation' },
    ragVsFt: {
      left: ['RAG', 'No retraining needed', 'Knowledge updates instantly', 'Source citations possible', 'Lower compute cost', 'Best for: Q&A, Search, Docs'],
      right: ['Fine-tuning', 'Requires model retraining', 'Knowledge baked into weights', 'No automatic citations', 'Higher GPU cost', 'Best for: Style, Format, Tone'],
    },
    docPipeline: { title: 'Document Processing Pipeline' },
    thaiNLP: {
      title: 'Thai NLP Challenges',
      bullets: [
        'No spaces between words → word segmentation required',
        'Multiple valid tokenizations for same sentence',
        'Complex spelling rules and tonal markers',
        'Mixed Thai-English text is common',
        'Limited open-source resources vs. English',
        'Solution: PyThaiNLP + custom dictionaries',
      ],
    },
    pyThaiNLP: {
      title: 'PyThaiNLP Demo',
      code: [
        'from pythainlp.tokenize import word_tokenize',
        '',
        '# Thai text without spaces',
        'text = "ฉันชอบกินข้าวผัดกุ้ง"',
        '',
        '# Tokenize with different engines',
        'tokens_newmm = word_tokenize(text, engine="newmm")',
        '# → ["ฉัน", "ชอบ", "กิน", "ข้าวผัด", "กุ้ง"]',
        '',
        'tokens_attacut = word_tokenize(text, engine="attacut")',
        '# → ["ฉัน", "ชอบ", "กิน", "ข้าว", "ผัด", "กุ้ง"]',
        '',
        '# Sentence segmentation',
        'from pythainlp.tokenize import sent_tokenize',
        'sentences = sent_tokenize(paragraph)',
      ],
    },
    chunking: {
      title: 'Chunking Strategies',
      boxes: [
        { label: 'Fixed Size\n\n256–512 tokens\nSimple & fast\nMay split context', color: BLUE },
        { label: 'Recursive\n\nSplit by paragraph\n→ sentence → token\nPreserves structure', color: TEAL },
        { label: 'Semantic\n\nEmbedding-based\nGroup similar content\nBest quality', color: '7B3F9E' },
      ],
    },
    embeddingWhat: {
      title: 'What are Embeddings?',
      bullets: [
        'Text converted to dense numerical vectors',
        'Semantically similar texts → similar vectors',
        'Vector space: each dimension = learned feature',
        'Typical dimensions: 768, 1024, 1536',
        'Enable: similarity search, clustering, classification',
        'Example: "cat" ≈ "kitten" >> "car" in vector space',
      ],
    },
    embeddingModels: {
      left: ['bge-m3', 'Multilingual (100+ languages)', 'Dimension: 1024', 'Context: 8192 tokens', 'Open-source (BAAI)', 'Excellent Thai support', 'Run locally with Ollama'],
      right: ['nomic-embed-text', 'Primarily English', 'Dimension: 768', 'Context: 8192 tokens', 'Open-source (Nomic AI)', 'Fast & lightweight', 'Good for English docs'],
    },
    vectorDB: {
      title: 'Vector Database Concepts',
      bullets: [
        '① Store: Save embeddings with metadata',
        '② Index: Build ANN index (HNSW, IVF)',
        '③ Search: Find k-nearest neighbors by cosine similarity',
        '',
        'Approximate Nearest Neighbor (ANN) is fast but approximate',
        'Trade-off: Speed ↔ Recall accuracy',
        'Popular: Qdrant, Weaviate, Pinecone, Milvus, pgvector',
      ],
    },
    qdrant: {
      title: 'Qdrant Architecture',
      bullets: [
        'Collections → like database tables (one per domain)',
        'Points → individual records (id + vector + payload)',
        'Payload → JSON metadata (filterable)',
        'Named Vectors → multiple embeddings per point',
        'Sparse Vectors → BM25/SPLADE for keyword search',
        'Quantization → compress vectors, reduce memory',
        'REST & gRPC APIs, Docker-ready, WASM-based filtering',
      ],
    },
    hybridSearch: {
      title: 'Hybrid Search: Dense + Sparse + Reranking',
    },
    bm25: {
      left: ['BM25 (Sparse)', 'Keyword-based matching', 'Exact term recall', 'Fast, interpretable', 'Good for technical terms', 'No semantic understanding', 'TF-IDF variant'],
      right: ['Dense Search', 'Semantic similarity', 'Understands synonyms', 'Slower (ANN index)', 'Misses rare keywords', 'Powerful contextual match', 'Embedding-based'],
    },
    ragPipeline: { title: 'Basic RAG Pipeline' },
    ragCode: {
      title: 'RAG Pipeline — Code Walkthrough',
      code: [
        '# 1. Load & chunk documents',
        'chunks = load_and_chunk_documents("docs/", chunk_size=512)',
        '',
        '# 2. Embed chunks',
        'embeddings = embed_model.encode(chunks)  # shape: (N, 1024)',
        '',
        '# 3. Store in Qdrant',
        'qdrant.upsert(collection="thai-rag",',
        '              points=[PointStruct(id=i, vector=emb, payload={"text": c})',
        '                      for i, (c, emb) in enumerate(zip(chunks, embeddings))])',
        '',
        '# 4. Query',
        'query_vec = embed_model.encode(user_query)',
        'results = qdrant.search(collection="thai-rag", query_vector=query_vec, limit=5)',
        '',
        '# 5. Generate answer with LLM',
        'context = "\\n".join([r.payload["text"] for r in results])',
        'answer = llm.chat(f"Context:\\n{context}\\n\\nQuestion: {user_query}")',
      ],
    },
    ragProblems: {
      title: 'Common RAG Problems',
      bullets: [
        '⚠ Hallucination — LLM adds info not in context → Use citation enforcement',
        '⚠ Context window overflow — Too much retrieved text → Tune top-k & chunk size',
        '⚠ Low relevance — Wrong chunks retrieved → Improve embedding + reranking',
        '⚠ Lost-in-the-middle — LLM ignores middle context → Reorder results',
        '⚠ Stale knowledge — New docs not indexed → Incremental upserts',
        '⚠ Language mismatch — Query in Thai, doc in English → Cross-lingual embed',
      ],
    },
    exerciseOverview: {
      title: 'Hands-on Exercise Overview',
      bullets: [
        '🛠  Environment: Python 3.10+, Docker, Ollama, Qdrant',
        '📄  Dataset: Thai company policy documents (PDF)',
        '🔧  Tools: PyThaiNLP, bge-m3 via Ollama, Qdrant, FastAPI',
        '',
        'Steps:',
        '  1. Load PDFs with pdfplumber',
        '  2. Tokenize Thai text with PyThaiNLP',
        '  3. Chunk with RecursiveTextSplitter',
        '  4. Embed with bge-m3',
        '  5. Store in Qdrant',
        '  6. Build simple FastAPI chat endpoint',
      ],
    },
    exercise: {
      title: 'Exercise: Build Thai RAG Pipeline',
      code: [
        '# Clone the workshop repo',
        'git clone https://github.com/x-company/ai-workshop',
        'cd ai-workshop/day1-rag',
        '',
        '# Start services',
        'docker compose up -d  # Qdrant + Ollama',
        'ollama pull bge-m3',
        '',
        '# Run exercise notebook',
        'jupyter lab notebooks/01_thai_rag.ipynb',
        '',
        '# Or run the script directly',
        'python exercise/build_rag.py --docs data/thai_docs/',
        '',
        '# Test your pipeline',
        'python exercise/test_rag.py --query "นโยบายวันลาพักร้อนคือ?"',
      ],
    },
    summary: {
      title: 'Day 1 Summary',
      bullets: [
        '✅  RAG = Retrieve relevant context → Generate accurate answers',
        '✅  Thai NLP requires explicit tokenization (PyThaiNLP)',
        '✅  Chunk size & strategy significantly affect retrieval quality',
        '✅  Embeddings map text to semantic vector space',
        '✅  Qdrant supports hybrid search (dense + sparse)',
        '✅  Basic RAG pipeline: Load → Chunk → Embed → Store → Retrieve → Generate',
        '',
        '📌  Tomorrow: Knowledge Graphs + GraphRAG for complex reasoning',
      ],
    },
  },

  th: {
    title: 'Day 1: Data Engineering & RAG Foundation',
    workshopName: 'AI Workshop: สร้าง Enterprise AI ด้วย RAG & Knowledge Graphs',
    agenda: [
      'แนะนำ Retrieval-Augmented Generation (RAG)',
      'การประมวลผลเอกสารและความท้าทายของ NLP ภาษาไทย',
      'กลยุทธ์การแบ่ง Chunk และโมเดล Embedding',
      'Vector Database และ Hybrid Search',
      'ลงมือปฏิบัติ: สร้าง Thai RAG Pipeline',
    ],
    ragWhat: { title: 'RAG คืออะไร?', sub: 'Retrieval-Augmented Generation' },
    ragVsFt: {
      left: ['RAG', 'ไม่ต้อง Retrain โมเดล', 'อัปเดตความรู้ได้ทันที', 'อ้างอิงแหล่งที่มาได้', 'ต้นทุนการคำนวณต่ำกว่า', 'เหมาะกับ: Q&A, Search, เอกสาร'],
      right: ['Fine-tuning', 'ต้อง Retrain โมเดล', 'ความรู้ถูกฝังใน Weights', 'ไม่มี Citation อัตโนมัติ', 'ต้องใช้ GPU สูง', 'เหมาะกับ: สไตล์, รูปแบบ, น้ำเสียง'],
    },
    docPipeline: { title: 'Document Processing Pipeline' },
    thaiNLP: {
      title: 'ความท้าทายของ Thai NLP',
      bullets: [
        'ไม่มีช่องว่างระหว่างคำ → ต้องตัดคำก่อน',
        'ประโยคเดียวกันอาจตัดคำได้หลายแบบ',
        'กฎการสะกดและวรรณยุกต์ที่ซับซ้อน',
        'ข้อความผสมไทย-อังกฤษพบบ่อย',
        'ทรัพยากร Open-source น้อยกว่าภาษาอังกฤษ',
        'แนวทาง: PyThaiNLP + Custom Dictionary',
      ],
    },
    pyThaiNLP: {
      title: 'PyThaiNLP Demo',
      code: [
        'from pythainlp.tokenize import word_tokenize',
        '',
        '# ข้อความภาษาไทยไม่มีช่องว่าง',
        'text = "ฉันชอบกินข้าวผัดกุ้ง"',
        '',
        '# ตัดคำด้วย Engine ต่างๆ',
        'tokens_newmm = word_tokenize(text, engine="newmm")',
        '# → ["ฉัน", "ชอบ", "กิน", "ข้าวผัด", "กุ้ง"]',
        '',
        'tokens_attacut = word_tokenize(text, engine="attacut")',
        '# → ["ฉัน", "ชอบ", "กิน", "ข้าว", "ผัด", "กุ้ง"]',
        '',
        '# แบ่งประโยค',
        'from pythainlp.tokenize import sent_tokenize',
        'sentences = sent_tokenize(paragraph)',
      ],
    },
    chunking: {
      title: 'กลยุทธ์การแบ่ง Chunk',
      boxes: [
        { label: 'Fixed Size\n\n256–512 tokens\nเรียบง่าย รวดเร็ว\nอาจตัดกลางบริบท', color: BLUE },
        { label: 'Recursive\n\nแบ่งตาม paragraph\n→ ประโยค → token\nรักษาโครงสร้าง', color: TEAL },
        { label: 'Semantic\n\nใช้ Embedding\nจัดกลุ่มเนื้อหาใกล้เคียง\nคุณภาพดีที่สุด', color: '7B3F9E' },
      ],
    },
    embeddingWhat: {
      title: 'Embedding คืออะไร?',
      bullets: [
        'แปลงข้อความเป็น Vector ตัวเลขที่หนาแน่น',
        'ข้อความที่มีความหมายใกล้เคียง → Vector ใกล้กัน',
        'Vector Space: แต่ละมิติ = คุณลักษณะที่เรียนรู้',
        'จำนวนมิติทั่วไป: 768, 1024, 1536',
        'ใช้ได้กับ: Similarity Search, Clustering, Classification',
        'ตัวอย่าง: "แมว" ≈ "ลูกแมว" >> "รถ" ใน Vector Space',
      ],
    },
    embeddingModels: {
      left: ['bge-m3', 'รองรับหลายภาษา (100+ ภาษา)', 'มิติ: 1024', 'Context: 8192 tokens', 'Open-source (BAAI)', 'รองรับภาษาไทยได้ดีเยี่ยม', 'รันบนเครื่องด้วย Ollama'],
      right: ['nomic-embed-text', 'เน้นภาษาอังกฤษ', 'มิติ: 768', 'Context: 8192 tokens', 'Open-source (Nomic AI)', 'เร็วและเบา', 'เหมาะกับเอกสารภาษาอังกฤษ'],
    },
    vectorDB: {
      title: 'แนวคิด Vector Database',
      bullets: [
        '① Store: บันทึก Embedding พร้อม Metadata',
        '② Index: สร้าง ANN Index (HNSW, IVF)',
        '③ Search: ค้นหา k-nearest neighbor ด้วย Cosine Similarity',
        '',
        'Approximate Nearest Neighbor (ANN) เร็วแต่ประมาณค่า',
        'Trade-off: ความเร็ว ↔ ความแม่นยำ',
        'ที่นิยม: Qdrant, Weaviate, Pinecone, Milvus, pgvector',
      ],
    },
    qdrant: {
      title: 'สถาปัตยกรรมของ Qdrant',
      bullets: [
        'Collections → เหมือน Table ในฐานข้อมูล (แยกตาม Domain)',
        'Points → แต่ละ Record (id + vector + payload)',
        'Payload → JSON Metadata (กรองได้)',
        'Named Vectors → หลาย Embedding ต่อ Point',
        'Sparse Vectors → BM25/SPLADE สำหรับ Keyword Search',
        'Quantization → บีบอัด Vector ลดการใช้หน่วยความจำ',
        'REST & gRPC API พร้อม Docker',
      ],
    },
    hybridSearch: {
      title: 'Hybrid Search: Dense + Sparse + Reranking',
    },
    bm25: {
      left: ['BM25 (Sparse)', 'จับคู่ตาม Keyword', 'ค้นหาคำตรงๆ ได้ดี', 'เร็ว ตีความได้', 'ดีสำหรับคำศัพท์เทคนิค', 'ไม่เข้าใจความหมาย', 'พัฒนาจาก TF-IDF'],
      right: ['Dense Search', 'Semantic Similarity', 'เข้าใจ Synonym', 'ช้ากว่า (ANN Index)', 'พลาด Keyword หายาก', 'จับบริบทได้ดี', 'ใช้ Embedding เป็นฐาน'],
    },
    ragPipeline: { title: 'Basic RAG Pipeline' },
    ragCode: {
      title: 'RAG Pipeline — Code Walkthrough',
      code: [
        '# 1. โหลดและแบ่ง Chunk เอกสาร',
        'chunks = load_and_chunk_documents("docs/", chunk_size=512)',
        '',
        '# 2. สร้าง Embedding',
        'embeddings = embed_model.encode(chunks)  # shape: (N, 1024)',
        '',
        '# 3. บันทึกใน Qdrant',
        'qdrant.upsert(collection="thai-rag",',
        '              points=[PointStruct(id=i, vector=emb, payload={"text": c})',
        '                      for i, (c, emb) in enumerate(zip(chunks, embeddings))])',
        '',
        '# 4. Query',
        'query_vec = embed_model.encode(user_query)',
        'results = qdrant.search(collection="thai-rag", query_vector=query_vec, limit=5)',
        '',
        '# 5. สร้างคำตอบด้วย LLM',
        'context = "\\n".join([r.payload["text"] for r in results])',
        'answer = llm.chat(f"บริบท:\\n{context}\\n\\nคำถาม: {user_query}")',
      ],
    },
    ragProblems: {
      title: 'ปัญหาที่พบบ่อยใน RAG',
      bullets: [
        '⚠ Hallucination — LLM เพิ่มข้อมูลนอก Context → บังคับอ้างอิง Citation',
        '⚠ Context Window เกิน — ดึงข้อความมากเกินไป → ปรับ top-k และ chunk size',
        '⚠ ดึงข้อมูลผิด — Chunk ไม่ตรงกับคำถาม → ปรับ Embedding + Reranking',
        '⚠ Lost-in-the-middle — LLM ละเลยบริบทกลาง → จัดเรียง Result ใหม่',
        '⚠ ข้อมูลล้าสมัย — เอกสารใหม่ไม่ถูก Index → ทำ Incremental Upsert',
        '⚠ ภาษาไม่ตรง — Query ภาษาไทย เอกสารภาษาอังกฤษ → Cross-lingual Embed',
      ],
    },
    exerciseOverview: {
      title: 'ภาพรวมของ Hands-on Exercise',
      bullets: [
        '🛠  สภาพแวดล้อม: Python 3.10+, Docker, Ollama, Qdrant',
        '📄  Dataset: เอกสารนโยบายบริษัท (PDF ภาษาไทย)',
        '🔧  เครื่องมือ: PyThaiNLP, bge-m3 via Ollama, Qdrant, FastAPI',
        '',
        'ขั้นตอน:',
        '  1. โหลด PDF ด้วย pdfplumber',
        '  2. ตัดคำภาษาไทยด้วย PyThaiNLP',
        '  3. แบ่ง Chunk ด้วย RecursiveTextSplitter',
        '  4. สร้าง Embedding ด้วย bge-m3',
        '  5. บันทึกใน Qdrant',
        '  6. สร้าง FastAPI Chat Endpoint',
      ],
    },
    exercise: {
      title: 'Exercise: สร้าง Thai RAG Pipeline',
      code: [
        '# Clone Workshop Repo',
        'git clone https://github.com/x-company/ai-workshop',
        'cd ai-workshop/day1-rag',
        '',
        '# เริ่ม Services',
        'docker compose up -d  # Qdrant + Ollama',
        'ollama pull bge-m3',
        '',
        '# รัน Exercise Notebook',
        'jupyter lab notebooks/01_thai_rag.ipynb',
        '',
        '# หรือรัน Script โดยตรง',
        'python exercise/build_rag.py --docs data/thai_docs/',
        '',
        '# ทดสอบ Pipeline',
        'python exercise/test_rag.py --query "นโยบายวันลาพักร้อนคือ?"',
      ],
    },
    summary: {
      title: 'สรุป Day 1',
      bullets: [
        '✅  RAG = ดึงบริบทที่เกี่ยวข้อง → สร้างคำตอบที่แม่นยำ',
        '✅  Thai NLP ต้องมีการตัดคำก่อน (PyThaiNLP)',
        '✅  Chunk Size และกลยุทธ์ส่งผลต่อคุณภาพการค้นหามาก',
        '✅  Embedding แปลงข้อความเป็น Semantic Vector Space',
        '✅  Qdrant รองรับ Hybrid Search (Dense + Sparse)',
        '✅  Basic RAG: Load → Chunk → Embed → Store → Retrieve → Generate',
        '',
        '📌  พรุ่งนี้: Knowledge Graphs + GraphRAG สำหรับการวิเคราะห์ซับซ้อน',
      ],
    },
  },
};

// ─── DAY 2 CONTENT ────────────────────────────────────────────────────────────

const DAY2 = {
  en: {
    title: 'Day 2: Knowledge Graph & Advanced RAG',
    workshopName: 'AI Workshop: Build Enterprise AI with RAG & Knowledge Graphs',
    agenda: [
      'Knowledge Graph fundamentals & structure',
      'Entity & Relation Extraction',
      'Neo4j & Cypher Query Language',
      'GraphRAG: Combining Vector + Graph',
      'Hands-on: Build Knowledge Graph from Thai docs',
    ],
    kgWhat: {
      title: 'What is a Knowledge Graph?',
      bullets: [
        'Nodes (Entities): People, Organizations, Concepts, Events',
        'Edges (Relations): has, works_at, located_in, related_to',
        'Properties: attributes on nodes and edges',
        'Enables multi-hop reasoning: A → B → C',
        'Structured representation of world knowledge',
        'Powers: Google Knowledge Panel, Wikidata, LinkedIn Graph',
      ],
    },
    kgVsVec: {
      left: ['Knowledge Graph', 'Explicit relationships', 'Traversal queries (multi-hop)', 'Interpretable reasoning', 'Good for structured facts', 'Complex schema needed', 'Harder to build'],
      right: ['Vector DB', 'Implicit similarity', 'Approximate nearest neighbor', 'Black-box retrieval', 'Good for unstructured text', 'Schema-free', 'Easy to build'],
    },
    entityExtract: {
      title: 'Entity Extraction',
      bullets: [
        'NER (Named Entity Recognition): rule-based or ML model',
        'LLM-based extraction: prompt LLM to extract structured JSON',
        'Thai NER: SpaCy + PyThaiNLP NER model',
        '',
        'Example prompt:',
        '  "Extract all entities from this text.',
        '   Return JSON: [{entity, type, description}]"',
        '',
        'Tools: spaCy, Stanza, Hugging Face NER, GPT-4, Claude',
      ],
    },
    entityTypes: {
      title: 'Entity Types',
      bullets: [
        '👤 Person  — employee, executive, customer name',
        '🏢 Organization — company, department, team',
        '📍 Location — city, country, office address',
        '💡 Concept — product, technology, service name',
        '📅 Date/Time — event date, deadline',
        '💰 Amount — financial figures, quantities',
        '📋 Event — meeting, incident, contract signing',
      ],
    },
    relationExtract: {
      title: 'Relation Extraction',
      bullets: [
        'Goal: identify semantic relationship between entity pairs',
        'Pattern-based: regex templates for common relations',
        'ML-based: BERT fine-tuned on relation datasets',
        'LLM-based: zero-shot via structured prompt',
        '',
        'Example relation triple:',
        '  (Elon Musk) —[CEO_OF]→ (Tesla)',
        '  (Tesla) —[LOCATED_IN]→ (Austin, TX)',
        '  (Elon Musk) —[FOUNDED]→ (SpaceX)',
      ],
    },
    graphPipeline: {
      title: 'Graph Construction Pipeline',
      bullets: [
        '① Load documents (PDF, Word, HTML)',
        '② Extract text & tokenize',
        '③ Run NER → collect entity list',
        '④ Run relation extraction → triples',
        '⑤ Deduplicate & merge entities (co-reference)',
        '⑥ Load into Neo4j: CREATE nodes, MERGE relations',
        '⑦ Build vector index on node descriptions',
      ],
    },
    neo4j: {
      title: 'Neo4j Introduction',
      bullets: [
        'Leading graph database — native graph storage',
        'ACID transactions, clustering support',
        'Cypher: declarative query language (like SQL for graphs)',
        'Neo4j Desktop / Docker / AuraDB (cloud)',
        'Supports vector indexes (Neo4j 5.x)',
        'Visualize graphs in Neo4j Browser',
        'LangChain & LlamaIndex have Neo4j integrations',
      ],
    },
    cypher: {
      title: 'Cypher Query Language Basics',
      code: [
        '// Create nodes',
        'CREATE (p:Person {name: "Somchai", role: "Engineer"})',
        'CREATE (c:Company {name: "X-Company"})',
        '',
        '// Create relationship',
        'MATCH (p:Person {name: "Somchai"}), (c:Company {name: "X-Company"})',
        'CREATE (p)-[:WORKS_AT {since: 2020}]->(c)',
        '',
        '// Query: Find who works at X-Company',
        'MATCH (p:Person)-[:WORKS_AT]->(c:Company {name: "X-Company"})',
        'RETURN p.name, p.role',
        '',
        '// Multi-hop: Find colleagues',
        'MATCH (p1:Person)-[:WORKS_AT]->(c)<-[:WORKS_AT]-(p2:Person)',
        'WHERE p1.name = "Somchai"',
        'RETURN p2.name AS colleague',
      ],
    },
    neo4jViz: {
      title: 'Neo4j Visualization',
      bullets: [
        'Neo4j Browser: built-in interactive graph explorer',
        'MATCH (n) RETURN n LIMIT 50 → visual graph',
        'Bloom: business-friendly no-code graph explorer',
        'Node color = label type, edge label = relation type',
        'Export to PNG, SVG, CSV, JSON',
        'Customize with APOC library (Advanced Procedures)',
        'Integration: Gephi, D3.js, vis.js for custom viz',
      ],
    },
    graphRAGWhat: {
      title: 'What is GraphRAG?',
      bullets: [
        'Combines Vector Search + Knowledge Graph traversal',
        'Vector: "find semantically similar chunks"',
        'Graph: "follow relationships to related entities"',
        'Best of both: semantic + structured reasoning',
        'Use case: "Who are Somchai\'s managers, and what projects did they approve?"',
        'Microsoft research: GraphRAG for large document corpora',
        'X-Company implementation: GraphRAG-rs (Rust-based)',
      ],
    },
    graphRAGArch: { title: 'GraphRAG Architecture' },
    graphRAGRs: {
      title: 'GraphRAG-rs: Rust-based Production GraphRAG',
      bullets: [
        'Written in Rust → memory-safe, blazing fast',
        'Integrates Qdrant (vector) + Neo4j (graph)',
        'Supports bge-m3 multilingual embeddings',
        'Community detection: Leiden algorithm',
        'Global summarization: hierarchical community reports',
        'REST API compatible with OpenAI chat format',
        'Used by X-Company for internal document intelligence',
      ],
    },
    hybridRetrieval: { title: 'Hybrid Retrieval: Vector → Graph → Rerank' },
    multiHop: {
      title: 'Multi-hop Reasoning Example',
      bullets: [
        'Question: "What safety incidents occurred at sites managed by projects Somchai supervised?"',
        '',
        'Step 1 — Vector: Find chunks about "safety incidents"',
        'Step 2 — Entity: Extract "Somchai" from query',
        'Step 3 — Graph hop 1: Somchai → SUPERVISED → Projects',
        'Step 4 — Graph hop 2: Projects → AT_SITE → Sites',
        'Step 5 — Graph hop 3: Sites → HAS_INCIDENT → Incidents',
        'Step 6 — Combine: vector chunks + graph facts → LLM',
        'Result: Accurate, cited, traceable answer',
      ],
    },
    community: {
      title: 'Community Detection in Knowledge Graphs',
      bullets: [
        'Groups of highly connected entities = community',
        'Algorithms: Leiden, Louvain, Label Propagation',
        'Each community gets a summary by LLM',
        'Enables global Q&A: "What are the main themes in our docs?"',
        'Microsoft GraphRAG: C0 (global) → C1 → C2 communities',
        'Useful for: org chart analysis, topic clustering, risk areas',
        'X-Company use: identify key departments in policy docs',
      ],
    },
    advancedRetrieval: {
      title: 'Advanced Retrieval Strategies',
      bullets: [
        '🔍 HyDE: generate hypothetical answer → embed → search',
        '🔄 Query Expansion: LLM rewrites query multiple ways',
        '📊 Contextual Compression: extract only relevant sentences',
        '🌳 Parent-Child Chunking: retrieve parent for context',
        '🔗 RAG Fusion: multiple queries + reciprocal rank fusion',
        '🎯 Self-RAG: LLM decides when to retrieve',
        '⚡ Re-ranking: cross-encoder scores retrieved chunks',
      ],
    },
    exercise2: {
      title: 'Exercise: Build Knowledge Graph from Thai Docs',
      code: [
        'cd ai-workshop/day2-kg',
        '',
        '# 1. Extract entities from Thai docs',
        'python extract_entities.py --input data/thai_docs/ --output entities.json',
        '',
        '# 2. Extract relations',
        'python extract_relations.py --entities entities.json --output triples.json',
        '',
        '# 3. Load into Neo4j',
        'python load_neo4j.py --triples triples.json',
        '',
        '# 4. Verify in Neo4j Browser',
        '# Open http://localhost:7474',
        '# MATCH (n) RETURN n LIMIT 50',
        '',
        '# 5. Run GraphRAG query',
        'python graphrag_query.py --query "ใครบริหารโครงการไหนบ้าง?"',
      ],
    },
    summary2: {
      title: 'Day 2 Summary',
      bullets: [
        '✅  Knowledge Graph = Nodes + Edges + Properties',
        '✅  Entity & Relation Extraction → structured knowledge',
        '✅  Neo4j + Cypher for graph storage and multi-hop queries',
        '✅  GraphRAG = Vector (semantic) + Graph (structured)',
        '✅  GraphRAG-rs: production-grade Rust implementation',
        '✅  Community detection for global document understanding',
        '',
        '📌  Tomorrow: AI Agents, AI Operating System & Evaluation',
      ],
    },
  },

  th: {
    title: 'Day 2: Knowledge Graph & Advanced RAG',
    workshopName: 'AI Workshop: สร้าง Enterprise AI ด้วย RAG & Knowledge Graphs',
    agenda: [
      'พื้นฐาน Knowledge Graph และโครงสร้าง',
      'การ Extract Entity และ Relation',
      'Neo4j และ Cypher Query Language',
      'GraphRAG: รวม Vector + Graph เข้าด้วยกัน',
      'ลงมือปฏิบัติ: สร้าง Knowledge Graph จากเอกสารไทย',
    ],
    kgWhat: {
      title: 'Knowledge Graph คืออะไร?',
      bullets: [
        'Nodes (Entity): บุคคล, องค์กร, แนวคิด, เหตุการณ์',
        'Edges (ความสัมพันธ์): has, works_at, located_in, related_to',
        'Properties: คุณลักษณะบน Node และ Edge',
        'รองรับการวิเคราะห์แบบ Multi-hop: A → B → C',
        'แสดงความรู้โลกแบบมีโครงสร้าง',
        'ใช้ใน: Google Knowledge Panel, Wikidata, LinkedIn Graph',
      ],
    },
    kgVsVec: {
      left: ['Knowledge Graph', 'ความสัมพันธ์ชัดเจน', 'Query แบบ Multi-hop', 'การวิเคราะห์ตีความได้', 'เหมาะกับข้อมูลมีโครงสร้าง', 'ต้องออกแบบ Schema', 'สร้างยากกว่า'],
      right: ['Vector DB', 'ความคล้ายแบบ Implicit', 'Approximate Nearest Neighbor', 'การดึงข้อมูลแบบ Black-box', 'เหมาะกับข้อความไม่มีโครงสร้าง', 'ไม่ต้องมี Schema', 'สร้างง่าย'],
    },
    entityExtract: {
      title: 'การ Extract Entity',
      bullets: [
        'NER (Named Entity Recognition): แบบ Rule-based หรือ ML',
        'LLM-based: ให้ LLM Extract JSON มีโครงสร้าง',
        'Thai NER: SpaCy + PyThaiNLP NER Model',
        '',
        'ตัวอย่าง Prompt:',
        '  "Extract ทุก Entity จากข้อความนี้',
        '   ส่งกลับเป็น JSON: [{entity, type, description}]"',
        '',
        'เครื่องมือ: spaCy, Stanza, Hugging Face NER, GPT-4, Claude',
      ],
    },
    entityTypes: {
      title: 'ประเภทของ Entity',
      bullets: [
        '👤 บุคคล (Person) — พนักงาน, ผู้บริหาร, ลูกค้า',
        '🏢 องค์กร (Organization) — บริษัท, แผนก, ทีม',
        '📍 สถานที่ (Location) — เมือง, ประเทศ, ที่อยู่สำนักงาน',
        '💡 แนวคิด (Concept) — ผลิตภัณฑ์, เทคโนโลยี, บริการ',
        '📅 วันเวลา (Date/Time) — วันที่เหตุการณ์, Deadline',
        '💰 จำนวน (Amount) — ตัวเลขทางการเงิน, ปริมาณ',
        '📋 เหตุการณ์ (Event) — ประชุม, อุบัติเหตุ, เซ็นสัญญา',
      ],
    },
    relationExtract: {
      title: 'การ Extract Relation',
      bullets: [
        'เป้าหมาย: ระบุความสัมพันธ์ระหว่างคู่ Entity',
        'Pattern-based: Regex Template สำหรับ Relation ทั่วไป',
        'ML-based: BERT Fine-tuned บน Relation Dataset',
        'LLM-based: Zero-shot ผ่าน Structured Prompt',
        '',
        'ตัวอย่าง Relation Triple:',
        '  (สมชาย) —[ทำงานที่]→ (X-Company)',
        '  (X-Company) —[ตั้งอยู่ที่]→ (กรุงเทพฯ)',
        '  (สมชาย) —[ดูแลโครงการ]→ (Project Alpha)',
      ],
    },
    graphPipeline: {
      title: 'กระบวนการสร้าง Knowledge Graph',
      bullets: [
        '① โหลดเอกสาร (PDF, Word, HTML)',
        '② Extract ข้อความและตัดคำ',
        '③ รัน NER → รวบรวม Entity List',
        '④ รัน Relation Extraction → ได้ Triple',
        '⑤ Deduplicate และรวม Entity (Co-reference)',
        '⑥ โหลดเข้า Neo4j: CREATE nodes, MERGE relations',
        '⑦ สร้าง Vector Index บน Node Description',
      ],
    },
    neo4j: {
      title: 'แนะนำ Neo4j',
      bullets: [
        'ฐานข้อมูล Graph ชั้นนำ — จัดเก็บแบบ Native Graph',
        'รองรับ ACID Transaction และ Clustering',
        'Cypher: Query Language แบบ Declarative (คล้าย SQL)',
        'Neo4j Desktop / Docker / AuraDB (Cloud)',
        'รองรับ Vector Index (Neo4j 5.x)',
        'Visualize Graph ใน Neo4j Browser',
        'LangChain & LlamaIndex มี Neo4j Integration',
      ],
    },
    cypher: {
      title: 'พื้นฐาน Cypher Query Language',
      code: [
        '// สร้าง Node',
        'CREATE (p:Person {name: "สมชาย", role: "วิศวกร"})',
        'CREATE (c:Company {name: "X-Company"})',
        '',
        '// สร้าง Relationship',
        'MATCH (p:Person {name: "สมชาย"}), (c:Company {name: "X-Company"})',
        'CREATE (p)-[:WORKS_AT {since: 2020}]->(c)',
        '',
        '// Query: หาคนที่ทำงานใน X-Company',
        'MATCH (p:Person)-[:WORKS_AT]->(c:Company {name: "X-Company"})',
        'RETURN p.name, p.role',
        '',
        '// Multi-hop: หาเพื่อนร่วมงาน',
        'MATCH (p1:Person)-[:WORKS_AT]->(c)<-[:WORKS_AT]-(p2:Person)',
        'WHERE p1.name = "สมชาย"',
        'RETURN p2.name AS colleague',
      ],
    },
    neo4jViz: {
      title: 'การ Visualize ด้วย Neo4j',
      bullets: [
        'Neo4j Browser: Graph Explorer แบบ Interactive',
        'MATCH (n) RETURN n LIMIT 50 → แสดง Graph ทันที',
        'Bloom: สำหรับผู้ใช้ธุรกิจ ไม่ต้องเขียน Cypher',
        'สีของ Node = ประเภท Label, Label บน Edge = ประเภท Relation',
        'Export เป็น PNG, SVG, CSV, JSON',
        'ปรับแต่งด้วย APOC Library (Advanced Procedures)',
        'Integration: Gephi, D3.js, vis.js สำหรับ Custom Viz',
      ],
    },
    graphRAGWhat: {
      title: 'GraphRAG คืออะไร?',
      bullets: [
        'รวม Vector Search + Knowledge Graph Traversal',
        'Vector: "ค้นหา Chunk ที่มีความหมายใกล้เคียง"',
        'Graph: "ติดตาม Relationship ไปยัง Entity ที่เกี่ยวข้อง"',
        'ดีที่สุดทั้งคู่: Semantic + Structured Reasoning',
        'Use case: "ใครเป็น Manager ของสมชาย และโครงการใดที่อนุมัติ?"',
        'Microsoft Research: GraphRAG สำหรับ Document Corpus ขนาดใหญ่',
        'X-Company: GraphRAG-rs (Rust-based)',
      ],
    },
    graphRAGArch: { title: 'สถาปัตยกรรม GraphRAG' },
    graphRAGRs: {
      title: 'GraphRAG-rs: Production GraphRAG ด้วย Rust',
      bullets: [
        'เขียนด้วย Rust → Memory-safe, ประสิทธิภาพสูง',
        'รวม Qdrant (Vector) + Neo4j (Graph)',
        'รองรับ bge-m3 Multilingual Embedding',
        'Community Detection: Leiden Algorithm',
        'Global Summarization: Hierarchical Community Reports',
        'REST API เข้ากันได้กับ OpenAI Chat Format',
        'X-Company ใช้สำหรับ Internal Document Intelligence',
      ],
    },
    hybridRetrieval: { title: 'Hybrid Retrieval: Vector → Graph → Rerank' },
    multiHop: {
      title: 'ตัวอย่าง Multi-hop Reasoning',
      bullets: [
        'คำถาม: "มีอุบัติเหตุอะไรบ้างในไซต์งานที่สมชายดูแล?"',
        '',
        'Step 1 — Vector: ค้นหา Chunk เกี่ยวกับ "อุบัติเหตุ"',
        'Step 2 — Entity: Extract "สมชาย" จากคำถาม',
        'Step 3 — Graph hop 1: สมชาย → ดูแล → โครงการ',
        'Step 4 — Graph hop 2: โครงการ → อยู่ที่ → ไซต์งาน',
        'Step 5 — Graph hop 3: ไซต์งาน → มีอุบัติเหตุ → เหตุการณ์',
        'Step 6 — รวมกัน: Vector Chunk + Graph Facts → LLM',
        'ผลลัพธ์: คำตอบที่แม่นยำ อ้างอิงได้ ตรวจสอบได้',
      ],
    },
    community: {
      title: 'Community Detection ใน Knowledge Graph',
      bullets: [
        'กลุ่ม Entity ที่เชื่อมต่อกันมาก = Community',
        'Algorithms: Leiden, Louvain, Label Propagation',
        'แต่ละ Community ได้รับ Summary จาก LLM',
        'Q&A ระดับ Global: "หัวข้อหลักในเอกสารของเราคืออะไร?"',
        'Microsoft GraphRAG: C0 (Global) → C1 → C2 Communities',
        'ใช้ได้กับ: วิเคราะห์ Org Chart, Cluster หัวข้อ, พื้นที่เสี่ยง',
        'X-Company: ระบุแผนกสำคัญในเอกสารนโยบาย',
      ],
    },
    advancedRetrieval: {
      title: 'Advanced Retrieval Strategies',
      bullets: [
        '🔍 HyDE: สร้างคำตอบสมมุติ → Embed → ค้นหา',
        '🔄 Query Expansion: LLM เขียนคำถามใหม่หลายแบบ',
        '📊 Contextual Compression: ดึงเฉพาะประโยคที่เกี่ยวข้อง',
        '🌳 Parent-Child Chunking: ดึง Parent เพื่อเพิ่ม Context',
        '🔗 RAG Fusion: Query หลายแบบ + Reciprocal Rank Fusion',
        '🎯 Self-RAG: LLM ตัดสินว่าต้องดึงข้อมูลหรือไม่',
        '⚡ Re-ranking: Cross-encoder ให้คะแนน Chunk ที่ดึงมา',
      ],
    },
    exercise2: {
      title: 'Exercise: สร้าง Knowledge Graph จากเอกสารไทย',
      code: [
        'cd ai-workshop/day2-kg',
        '',
        '# 1. Extract Entity จากเอกสารไทย',
        'python extract_entities.py --input data/thai_docs/ --output entities.json',
        '',
        '# 2. Extract Relation',
        'python extract_relations.py --entities entities.json --output triples.json',
        '',
        '# 3. โหลดเข้า Neo4j',
        'python load_neo4j.py --triples triples.json',
        '',
        '# 4. ตรวจสอบใน Neo4j Browser',
        '# เปิด http://localhost:7474',
        '# MATCH (n) RETURN n LIMIT 50',
        '',
        '# 5. รัน GraphRAG Query',
        'python graphrag_query.py --query "ใครบริหารโครงการไหนบ้าง?"',
      ],
    },
    summary2: {
      title: 'สรุป Day 2',
      bullets: [
        '✅  Knowledge Graph = Nodes + Edges + Properties',
        '✅  Extract Entity & Relation → ความรู้มีโครงสร้าง',
        '✅  Neo4j + Cypher สำหรับจัดเก็บและ Multi-hop Query',
        '✅  GraphRAG = Vector (Semantic) + Graph (Structured)',
        '✅  GraphRAG-rs: Implementation ระดับ Production ด้วย Rust',
        '✅  Community Detection สำหรับความเข้าใจเอกสารระดับ Global',
        '',
        '📌  พรุ่งนี้: AI Agents, AI Operating System และการประเมินผล',
      ],
    },
  },
};

// ─── DAY 3 CONTENT ────────────────────────────────────────────────────────────

const DAY3 = {
  en: {
    title: 'Day 3: AI Agents & AI Operating System',
    workshopName: 'AI Workshop: Build Enterprise AI with RAG & Knowledge Graphs',
    agenda: [
      'AI Agent fundamentals & architecture',
      'Multi-agent systems & communication patterns',
      'AI Operating System (AI OS) concepts',
      'Evaluation: RAGAS, LLM-as-Judge, A/B Testing',
      'Capstone: Build Enterprise AI Assistant',
    ],
    agentWhat: {
      title: 'What is an AI Agent?',
      bullets: [
        'An autonomous system that perceives, reasons, and acts',
        '① Perceive: receive input (text, data, sensor)',
        '② Think: LLM reasons about goal and context',
        '③ Act: call tools, write files, send messages',
        '',
        'Key difference from chatbot: agents can take actions!',
        'Agent loop: Perceive → Think → Act → Observe → Repeat',
        'Examples: ReAct, AutoGPT, Claude Computer Use, OpenClaw',
      ],
    },
    agentArch: {
      title: 'Agent Architecture',
    },
    toolUse: {
      title: 'Tool-Use / Function Calling',
      code: [
        '# Define tool schema (OpenAI format)',
        'tools = [{"type": "function",',
        '          "function": {',
        '            "name": "search_documents",',
        '            "description": "Search company knowledge base",',
        '            "parameters": {',
        '              "type": "object",',
        '              "properties": {',
        '                "query": {"type": "string", "description": "Search query"},',
        '                "limit": {"type": "integer", "default": 5}',
        '              }',
        '            }}}]',
        '',
        '# LLM decides which tool to call',
        'response = client.chat.completions.create(',
        '  model="gpt-4o", messages=messages, tools=tools)',
      ],
    },
    ragAgent: {
      title: 'RAG Agent',
      bullets: [
        'Agent + Knowledge Base = RAG Agent',
        'Agent decides WHEN to retrieve (not always)',
        'Can retrieve from multiple sources in one turn',
        'Self-RAG: LLM evaluates if retrieval is needed',
        'Adaptive RAG: routes to different retrieval strategies',
        'Tools: search_docs, search_graph, search_web, calculate',
        'Memory: conversation history + retrieved context',
      ],
    },
    multiAgent: {
      title: 'Multi-Agent Systems',
      bullets: [
        'Orchestrator agent coordinates specialist agents',
        'Specialist agents: HR, Finance, Engineering, Sales',
        'Benefits: separation of concerns, parallel execution',
        'Orchestrator decides: which agent handles this query?',
        'Result aggregation: merge outputs from multiple agents',
        'Error handling: retry failed sub-agent tasks',
        'Example: OpenClaw multi-persona routing',
      ],
    },
    commPatterns: {
      title: 'Agent Communication Patterns',
    },
    aiOSWhat: {
      title: 'What is an AI Operating System?',
      bullets: [
        'Infrastructure layer for AI agents — like an OS for apps',
        'Manages: agents lifecycle, memory, skills, scheduling',
        'Provides: tool registry, auth, logging, monitoring',
        'Multi-channel: Telegram, Discord, Email, REST API',
        'Proactive: heartbeat scheduler for background tasks',
        'Persistent: survives restarts with memory continuity',
        'Example: OpenClaw (used in this workshop)',
      ],
    },
    aiOSComponents: {
      title: 'AI OS Components',
      bullets: [
        '🤖 Agents — LLM-powered reasoning cores',
        '🔧 Skills — pluggable tool modules (SKILL.md)',
        '🧠 Memory — short-term (context) + long-term (files)',
        '💓 Heartbeat — proactive scheduler (cron-like)',
        '📡 Channels — Telegram, Discord, HTTP integrations',
        '🗂  Workspace — persistent file system for agents',
        '🔐 Auth — per-user permissions and access control',
      ],
    },
    openClaw: {
      title: 'OpenClaw Architecture Overview',
      bullets: [
        'Node.js daemon running on macOS/Linux/Raspberry Pi',
        'Multi-session: each user gets isolated agent session',
        'Skill system: install/update skills from ClawHub',
        'Heartbeat: configurable interval, silent background checks',
        'Memory: SOUL.md (identity) + MEMORY.md + daily logs',
        'Sub-agents: spawn parallel workers for heavy tasks',
        'Security: skill policies, tool whitelisting, data isolation',
      ],
    },
    memory: {
      title: 'Memory Systems',
      bullets: [
        '📝 Short-term Memory: active conversation context window',
        '💾 Long-term Memory: MEMORY.md — curated key facts',
        '📅 Episodic Memory: daily logs (memory/YYYY-MM-DD.md)',
        '🗺  Semantic Memory: knowledge base (RAG / Knowledge Graph)',
        '⚙️  Procedural Memory: skills & tool usage patterns',
        '',
        'Agents write to memory files → survive session restarts',
        'Heartbeat: review daily logs → update long-term memory',
      ],
    },
    heartbeat: {
      title: 'Heartbeat & Proactive Monitoring',
      bullets: [
        'Heartbeat = scheduled background agent activation',
        'Configurable interval (default: every 30 min)',
        'HEARTBEAT.md: defines what to check each run',
        'Checks: email, calendar, system health, data alerts',
        'Proactive: agent reaches out when something needs attention',
        'Silent mode: HEARTBEAT_OK when nothing needs action',
        'Cron jobs: for exact-time scheduled tasks',
      ],
    },
    skills: {
      title: 'Skills & Tool Integration',
      bullets: [
        'Skills = reusable tool bundles (SKILL.md + scripts)',
        'Each skill: description, commands, usage examples',
        'Install: openclaw skill install <name>',
        'Built-in: weather, email, github, gog, pptx, pdf, whisper',
        'Custom: create skill-creator for new tools',
        'ClawHub: public skill registry (like npm for agents)',
        'Tool policy: allowlist/denylist per skill',
      ],
    },
    ragas: {
      title: 'RAGAS Evaluation Framework',
      bullets: [
        'RAGAS = RAG Assessment framework (Python)',
        'Automatic evaluation using LLM-as-Judge',
        'No human labels needed (reference-free metrics)',
        'Evaluates: answer quality + retrieval quality',
        'Metrics: Faithfulness, Answer Relevance, Context Precision',
        'Also: Context Recall, Answer Correctness, Harmlessness',
        'CI/CD integration: evaluate every RAG pipeline change',
      ],
    },
    ragasMetrics: {
      title: 'RAGAS Metrics',
      bullets: [
        '🎯 Faithfulness: answer is grounded in retrieved context',
        '   Score 0–1: 1 = fully supported by context',
        '',
        '🔍 Answer Relevance: answer addresses the question',
        '   Score 0–1: 1 = directly answers the question',
        '',
        '📋 Context Precision: retrieved chunks are relevant',
        '   Score 0–1: 1 = all chunks are useful',
        '',
        '📊 Context Recall: all relevant info was retrieved',
        '   Score 0–1: 1 = no important info was missed',
      ],
    },
    llmJudge: {
      title: 'LLM-as-Judge Evaluation',
      code: [
        'from ragas import evaluate',
        'from ragas.metrics import faithfulness, answer_relevancy',
        'from datasets import Dataset',
        '',
        '# Prepare evaluation data',
        'data = {"question": questions,',
        '        "answer": generated_answers,',
        '        "contexts": retrieved_contexts,',
        '        "ground_truth": reference_answers}',
        '',
        'dataset = Dataset.from_dict(data)',
        '',
        '# Run RAGAS evaluation',
        'result = evaluate(dataset,',
        '                  metrics=[faithfulness, answer_relevancy],',
        '                  llm=ChatOpenAI(model="gpt-4o"))',
        '',
        'print(result)  # → {faithfulness: 0.87, answer_relevancy: 0.91}',
      ],
    },
    abTesting: {
      title: 'A/B Testing for RAG',
      bullets: [
        'Compare two RAG configurations systematically',
        'Variables to test: chunk size, embedding model, top-k, reranker',
        'Metric: RAGAS scores + latency + user satisfaction',
        'Traffic split: 50% → RAG-A, 50% → RAG-B',
        'Statistical significance: need enough samples (n>100)',
        'Shadow mode: run both, log results, no user impact',
        'Gradual rollout: 10% → 25% → 50% → 100% on winner',
      ],
    },
    capstone: {
      title: 'Capstone: Enterprise AI Assistant',
      bullets: [
        '🏗  Project: Build a complete AI assistant for X-Company',
        '',
        'Components:',
        '  • RAG pipeline: Thai HR & policy documents',
        '  • Knowledge Graph: org chart + project relationships',
        '  • Multi-agent: HR agent + Finance agent + Ops agent',
        '  • Interface: Telegram bot via OpenClaw',
        '  • Evaluation: RAGAS dashboard',
        '',
        'Deliverable: Working demo + RAGAS score report',
      ],
    },
    summary3: {
      title: 'Workshop Summary & Next Steps',
      bullets: [
        '✅  Day 1: RAG Pipeline — the foundation of enterprise AI',
        '✅  Day 2: Knowledge Graph — structured reasoning at scale',
        '✅  Day 3: AI Agents + AI OS — production deployment',
        '',
        '🚀  Next Steps:',
        '  • Deploy your RAG pipeline to production',
        '  • Build domain-specific knowledge graph',
        '  • Integrate with your business systems',
        '  • Set up RAGAS CI/CD for continuous evaluation',
        '  • Join X-Company AI Community for support',
        '',
        '🙏  Thank you! Questions?',
      ],
    },
  },

  th: {
    title: 'Day 3: AI Agents & AI Operating System',
    workshopName: 'AI Workshop: สร้าง Enterprise AI ด้วย RAG & Knowledge Graphs',
    agenda: [
      'พื้นฐาน AI Agent และสถาปัตยกรรม',
      'Multi-Agent Systems และรูปแบบการสื่อสาร',
      'AI Operating System (AI OS)',
      'การประเมินผล: RAGAS, LLM-as-Judge, A/B Testing',
      'Capstone: สร้าง Enterprise AI Assistant',
    ],
    agentWhat: {
      title: 'AI Agent คืออะไร?',
      bullets: [
        'ระบบอัตโนมัติที่รับรู้ วิเคราะห์ และลงมือทำ',
        '① Perceive: รับ Input (ข้อความ, ข้อมูล, Sensor)',
        '② Think: LLM วิเคราะห์เป้าหมายและบริบท',
        '③ Act: เรียก Tool, เขียนไฟล์, ส่งข้อความ',
        '',
        'ความต่างจาก Chatbot: Agent ลงมือทำจริงได้!',
        'Agent Loop: Perceive → Think → Act → Observe → วนซ้ำ',
        'ตัวอย่าง: ReAct, AutoGPT, Claude Computer Use, OpenClaw',
      ],
    },
    agentArch: {
      title: 'สถาปัตยกรรมของ Agent',
    },
    toolUse: {
      title: 'Tool-Use / Function Calling',
      code: [
        '# กำหนด Tool Schema (OpenAI Format)',
        'tools = [{"type": "function",',
        '          "function": {',
        '            "name": "search_documents",',
        '            "description": "ค้นหาใน Knowledge Base",',
        '            "parameters": {',
        '              "type": "object",',
        '              "properties": {',
        '                "query": {"type": "string", "description": "คำค้นหา"},',
        '                "limit": {"type": "integer", "default": 5}',
        '              }',
        '            }}}]',
        '',
        '# LLM ตัดสินใจเรียก Tool ใด',
        'response = client.chat.completions.create(',
        '  model="gpt-4o", messages=messages, tools=tools)',
      ],
    },
    ragAgent: {
      title: 'RAG Agent',
      bullets: [
        'Agent + Knowledge Base = RAG Agent',
        'Agent ตัดสินใจว่า WHEN ต้องดึงข้อมูล (ไม่ใช่ทุกครั้ง)',
        'ดึงข้อมูลจากหลายแหล่งในคำถามเดียว',
        'Self-RAG: LLM ประเมินว่าต้องการการค้นหาหรือไม่',
        'Adaptive RAG: เลือก Retrieval Strategy ที่เหมาะสม',
        'Tools: search_docs, search_graph, search_web, calculate',
        'Memory: ประวัติบทสนทนา + Context ที่ดึงมา',
      ],
    },
    multiAgent: {
      title: 'Multi-Agent Systems',
      bullets: [
        'Orchestrator Agent ประสาน Specialist Agent',
        'Specialist: HR, การเงิน, วิศวกรรม, Sales',
        'ข้อดี: แยกความรับผิดชอบ, ทำงานพร้อมกันได้',
        'Orchestrator ตัดสิน: Agent ใดจัดการคำถามนี้?',
        'รวมผลลัพธ์: รวม Output จากหลาย Agent',
        'จัดการข้อผิดพลาด: ลอง Sub-agent ที่ล้มเหลวอีกครั้ง',
        'ตัวอย่าง: OpenClaw Multi-persona Routing',
      ],
    },
    commPatterns: {
      title: 'รูปแบบการสื่อสารของ Agent',
    },
    aiOSWhat: {
      title: 'AI Operating System คืออะไร?',
      bullets: [
        'Layer โครงสร้างพื้นฐานสำหรับ AI Agent — เหมือน OS สำหรับ App',
        'จัดการ: Lifecycle ของ Agent, Memory, Skills, Scheduling',
        'ให้บริการ: Tool Registry, Auth, Logging, Monitoring',
        'Multi-channel: Telegram, Discord, Email, REST API',
        'Proactive: Heartbeat Scheduler สำหรับงานเบื้องหลัง',
        'Persistent: ทนต่อการ Restart ด้วย Memory Continuity',
        'ตัวอย่าง: OpenClaw (ใช้ใน Workshop นี้)',
      ],
    },
    aiOSComponents: {
      title: 'องค์ประกอบของ AI OS',
      bullets: [
        '🤖 Agents — LLM-powered Reasoning Core',
        '🔧 Skills — Tool Module แบบ Pluggable (SKILL.md)',
        '🧠 Memory — Short-term (Context) + Long-term (Files)',
        '💓 Heartbeat — Proactive Scheduler (คล้าย Cron)',
        '📡 Channels — Telegram, Discord, HTTP Integration',
        '🗂  Workspace — File System ถาวรสำหรับ Agent',
        '🔐 Auth — สิทธิ์และ Access Control ต่อผู้ใช้',
      ],
    },
    openClaw: {
      title: 'ภาพรวมสถาปัตยกรรม OpenClaw',
      bullets: [
        'Node.js Daemon บน macOS/Linux/Raspberry Pi',
        'Multi-session: ผู้ใช้แต่ละคนมี Session Agent แยกกัน',
        'Skill System: ติดตั้ง/อัปเดต Skill จาก ClawHub',
        'Heartbeat: Interval ปรับได้, ตรวจสอบเบื้องหลัง',
        'Memory: SOUL.md (ตัวตน) + MEMORY.md + Daily Logs',
        'Sub-agents: Spawn Worker คู่ขนานสำหรับงานหนัก',
        'Security: Skill Policy, Tool Whitelist, Data Isolation',
      ],
    },
    memory: {
      title: 'Memory Systems',
      bullets: [
        '📝 Short-term Memory: Context Window ของบทสนทนาปัจจุบัน',
        '💾 Long-term Memory: MEMORY.md — ข้อเท็จจริงสำคัญที่คัดกรอง',
        '📅 Episodic Memory: บันทึกรายวัน (memory/YYYY-MM-DD.md)',
        '🗺  Semantic Memory: Knowledge Base (RAG / Knowledge Graph)',
        '⚙️  Procedural Memory: Skills และรูปแบบการใช้ Tool',
        '',
        'Agent เขียนลงไฟล์ Memory → ทนต่อการ Restart',
        'Heartbeat: ทบทวน Daily Log → อัปเดต Long-term Memory',
      ],
    },
    heartbeat: {
      title: 'Heartbeat & Proactive Monitoring',
      bullets: [
        'Heartbeat = การกระตุ้น Agent เบื้องหลังตามกำหนดเวลา',
        'Interval ปรับได้ (ค่าเริ่มต้น: ทุก 30 นาที)',
        'HEARTBEAT.md: กำหนดสิ่งที่ต้องตรวจแต่ละรอบ',
        'ตรวจสอบ: Email, ปฏิทิน, สุขภาพระบบ, Data Alert',
        'Proactive: Agent แจ้งเตือนเมื่อมีสิ่งที่ต้องดูแล',
        'Silent Mode: HEARTBEAT_OK เมื่อไม่มีอะไรต้องทำ',
        'Cron Jobs: สำหรับงานที่ต้องการเวลาแน่นอน',
      ],
    },
    skills: {
      title: 'Skills & Tool Integration',
      bullets: [
        'Skills = Bundle เครื่องมือที่ใช้ซ้ำได้ (SKILL.md + Scripts)',
        'แต่ละ Skill: คำอธิบาย, คำสั่ง, ตัวอย่างการใช้',
        'ติดตั้ง: openclaw skill install <ชื่อ>',
        'Built-in: weather, email, github, gog, pptx, pdf, whisper',
        'Custom: สร้างด้วย skill-creator สำหรับ Tool ใหม่',
        'ClawHub: Skill Registry สาธารณะ (เหมือน npm สำหรับ Agent)',
        'Tool Policy: Allowlist/Denylist ต่อ Skill',
      ],
    },
    ragas: {
      title: 'RAGAS Evaluation Framework',
      bullets: [
        'RAGAS = RAG Assessment Framework (Python)',
        'ประเมินอัตโนมัติด้วย LLM-as-Judge',
        'ไม่ต้องมี Label จากมนุษย์ (Reference-free Metrics)',
        'ประเมิน: คุณภาพคำตอบ + คุณภาพการดึงข้อมูล',
        'Metrics: Faithfulness, Answer Relevance, Context Precision',
        'เพิ่มเติม: Context Recall, Answer Correctness, Harmlessness',
        'CI/CD Integration: ประเมินทุกครั้งที่เปลี่ยน Pipeline',
      ],
    },
    ragasMetrics: {
      title: 'RAGAS Metrics',
      bullets: [
        '🎯 Faithfulness: คำตอบอ้างอิงจาก Context ที่ดึงมา',
        '   Score 0–1: 1 = รองรับโดย Context ทั้งหมด',
        '',
        '🔍 Answer Relevance: คำตอบตอบคำถามได้ตรง',
        '   Score 0–1: 1 = ตอบตรงประเด็น',
        '',
        '📋 Context Precision: Chunk ที่ดึงมามีความเกี่ยวข้อง',
        '   Score 0–1: 1 = ทุก Chunk มีประโยชน์',
        '',
        '📊 Context Recall: ดึงข้อมูลที่เกี่ยวข้องครบถ้วน',
        '   Score 0–1: 1 = ไม่มีข้อมูลสำคัญตกหล่น',
      ],
    },
    llmJudge: {
      title: 'LLM-as-Judge Evaluation',
      code: [
        'from ragas import evaluate',
        'from ragas.metrics import faithfulness, answer_relevancy',
        'from datasets import Dataset',
        '',
        '# เตรียมข้อมูลสำหรับการประเมิน',
        'data = {"question": questions,',
        '        "answer": generated_answers,',
        '        "contexts": retrieved_contexts,',
        '        "ground_truth": reference_answers}',
        '',
        'dataset = Dataset.from_dict(data)',
        '',
        '# รัน RAGAS Evaluation',
        'result = evaluate(dataset,',
        '                  metrics=[faithfulness, answer_relevancy],',
        '                  llm=ChatOpenAI(model="gpt-4o"))',
        '',
        'print(result)  # → {faithfulness: 0.87, answer_relevancy: 0.91}',
      ],
    },
    abTesting: {
      title: 'A/B Testing สำหรับ RAG',
      bullets: [
        'เปรียบเทียบ RAG สองรูปแบบอย่างเป็นระบบ',
        'ตัวแปรที่ทดสอบ: Chunk Size, Embedding Model, top-k, Reranker',
        'Metric: RAGAS Score + Latency + ความพึงพอใจผู้ใช้',
        'แบ่ง Traffic: 50% → RAG-A, 50% → RAG-B',
        'ความมีนัยสำคัญทางสถิติ: ต้องมีตัวอย่างพอ (n>100)',
        'Shadow Mode: รันทั้งสอง บันทึกผล ไม่กระทบผู้ใช้',
        'Gradual Rollout: 10% → 25% → 50% → 100% กับ Winner',
      ],
    },
    capstone: {
      title: 'Capstone: Enterprise AI Assistant',
      bullets: [
        '🏗  โปรเจกต์: สร้าง AI Assistant ครบวงจรสำหรับ X-Company',
        '',
        'องค์ประกอบ:',
        '  • RAG Pipeline: เอกสาร HR & นโยบายภาษาไทย',
        '  • Knowledge Graph: Org Chart + ความสัมพันธ์โครงการ',
        '  • Multi-agent: HR Agent + Finance Agent + Ops Agent',
        '  • Interface: Telegram Bot ผ่าน OpenClaw',
        '  • Evaluation: RAGAS Dashboard',
        '',
        'Deliverable: Demo ที่ใช้งานได้ + รายงาน RAGAS Score',
      ],
    },
    summary3: {
      title: 'สรุปทั้งหมดและก้าวต่อไป',
      bullets: [
        '✅  Day 1: RAG Pipeline — รากฐานของ Enterprise AI',
        '✅  Day 2: Knowledge Graph — วิเคราะห์มีโครงสร้างในระดับ Scale',
        '✅  Day 3: AI Agents + AI OS — Deploy ใน Production',
        '',
        '🚀  ก้าวต่อไป:',
        '  • Deploy RAG Pipeline ไปสู่ Production',
        '  • สร้าง Knowledge Graph เฉพาะ Domain',
        '  • Integrate กับระบบธุรกิจของคุณ',
        '  • ตั้ง RAGAS CI/CD สำหรับประเมินผลต่อเนื่อง',
        '  • เข้าร่วม X-Company AI Community',
        '',
        '🙏  ขอบคุณ! มีคำถามไหมครับ?',
      ],
    },
  },
};

// ─── BUILD DECK FUNCTIONS ─────────────────────────────────────────────────────

function buildDay1(lang) {
  const d = DAY1[lang];
  const prs = newPptx();

  // 1. Title
  addTitleSlide(prs, d.title, d.workshopName, lang === 'th' ? 'Workshop Day 1' : 'Workshop Day 1');

  // 2. Agenda
  addContentSlide(prs, lang === 'th' ? 'วัตถุประสงค์การเรียนรู้' : 'Agenda & Learning Objectives',
    d.agenda.map((a, i) => `${i + 1}.  ${a}`));

  // 3. What is RAG
  addDiagramSlide(prs, d.ragWhat.title, (slide, prs) => {
    slide.addText(d.ragWhat.sub, { x: 0.5, y: 1.2, w: 12, h: 0.5, fontSize: 16, color: TEAL, fontFace: 'Arial', align: 'center' });
    const bw = 2.8, bh = 1.2, by = 2.5;
    box(slide, prs, 0.5, by, bw, bh, lang === 'th' ? '① คำถาม\n(Query)' : '① Query\n(User Question)', NAVY);
    arrowRight(slide, prs, 3.45, by + 0.4);
    box(slide, prs, 4.1, by, bw, bh, lang === 'th' ? '② ค้นหา\n(Retrieve)' : '② Retrieve\n(Vector Search)', BLUE);
    arrowRight(slide, prs, 7.05, by + 0.4);
    box(slide, prs, 7.7, by, bw, bh, lang === 'th' ? '③ สร้างคำตอบ\n(Generate)' : '③ Generate\n(LLM Answer)', TEAL);
    const bullets = lang === 'th'
      ? ['ดึงเอกสารที่เกี่ยวข้องจาก Knowledge Base', 'เพิ่ม Context ให้ LLM → คำตอบแม่นยำขึ้น', 'อ้างอิงแหล่งที่มาได้ ไม่ Hallucinate']
      : ['Retrieve relevant docs from Knowledge Base', 'Augment LLM prompt with context → accurate answer', 'Citable, grounded, reduces hallucination'];
    slide.addText(bullets.map(b => ({ text: b, options: { fontSize: 15, color: DGRAY, bullet: { type: 'bullet' }, paraSpaceAfter: 8 } })),
      { x: 0.5, y: 4.1, w: 12.5, h: 3, fontFace: 'Arial' });
  });

  // 4. RAG vs Fine-tuning
  addCompareSlide(prs,
    lang === 'th' ? 'RAG vs Fine-tuning' : 'RAG vs Fine-tuning',
    d.ragVsFt.left[0], d.ragVsFt.left.slice(1),
    d.ragVsFt.right[0], d.ragVsFt.right.slice(1));

  // 5. Document Pipeline
  addDiagramSlide(prs, d.docPipeline.title, (slide, prs) => {
    const labels = lang === 'th'
      ? ['PDF\nไฟล์', 'Extract\nข้อความ', 'Clean\nทำความสะอาด', 'Tokenize\nตัดคำ', 'Chunk\nแบ่งส่วน', 'Embed\nสร้าง Vector']
      : ['PDF\nFile', 'Extract\nText', 'Clean\n& Normalize', 'Tokenize\n(Thai NLP)', 'Chunk\nSplit', 'Embed\nVector'];
    const colors = [NAVY, BLUE, BLUE, TEAL, BLUE, TEAL];
    const bw = 1.7, bh = 1.1, by = 2.6, startX = 0.35;
    labels.forEach((lbl, i) => {
      box(slide, prs, startX + i * 2.1, by, bw, bh, lbl, colors[i]);
      if (i < labels.length - 1) arrowRight(slide, prs, startX + i * 2.1 + bw + 0.05, by + 0.38);
    });
    const note = lang === 'th' ? 'ข้อความแต่ละ Chunk จะถูกแปลงเป็น Vector Embedding สำหรับจัดเก็บใน Vector Database'
      : 'Each text chunk is converted to a vector embedding and stored in the vector database';
    slide.addText(note, { x: 0.5, y: 4.1, w: 12.5, h: 0.7, fontSize: 14, color: DGRAY, fontFace: 'Arial', align: 'center', italic: true });
  });

  // 6. Thai NLP
  addContentSlide(prs, d.thaiNLP.title, d.thaiNLP.bullets);

  // 7. PyThaiNLP
  addCodeSlide(prs, d.pyThaiNLP.title, d.pyThaiNLP.code);

  // 8. Chunking
  addDiagramSlide(prs, d.chunking.title, (slide, prs) => {
    d.chunking.boxes.forEach((b, i) => {
      box(slide, prs, 0.5 + i * 4.2, 1.6, 3.8, 4.5, b.label, b.color);
    });
    const note = lang === 'th' ? 'เลือก Strategy ตาม Use case: Fixed (เร็ว) → Recursive (สมดุล) → Semantic (คุณภาพสูง)'
      : 'Choose strategy by use case: Fixed (fast) → Recursive (balanced) → Semantic (highest quality)';
    slide.addText(note, { x: 0.5, y: 6.3, w: 12.5, h: 0.8, fontSize: 13, color: DGRAY, fontFace: 'Arial', align: 'center', italic: true });
  });

  // 9. What are Embeddings
  addContentSlide(prs, d.embeddingWhat.title, d.embeddingWhat.bullets);

  // 10. Embedding Models
  addCompareSlide(prs, lang === 'th' ? 'โมเดล Embedding เปรียบเทียบ' : 'Embedding Models Comparison',
    d.embeddingModels.left[0], d.embeddingModels.left.slice(1),
    d.embeddingModels.right[0], d.embeddingModels.right.slice(1));

  // 11. Vector DB
  addContentSlide(prs, d.vectorDB.title, d.vectorDB.bullets);

  // 12. Qdrant
  addContentSlide(prs, d.qdrant.title, d.qdrant.bullets);

  // 13. Hybrid Search diagram
  addDiagramSlide(prs, d.hybridSearch.title, (slide, prs) => {
    box(slide, prs, 0.4, 1.5, 2.5, 0.9, lang === 'th' ? 'คำถามผู้ใช้' : 'User Query', NAVY);
    arrowDown(slide, prs, 1.47, 2.5);
    box(slide, prs, 0.4, 3.1, 2.5, 0.9, lang === 'th' ? 'Dense\n(Embedding)' : 'Dense Search\n(Embedding)', BLUE);
    box(slide, prs, 3.3, 3.1, 2.5, 0.9, lang === 'th' ? 'Sparse\n(BM25)' : 'Sparse Search\n(BM25)', BLUE);
    arrowRight(slide, prs, 2.95, 1.9);
    arrowDown(slide, prs, 4.42, 2.5);
    box(slide, prs, 6.3, 3.1, 2.5, 0.9, lang === 'th' ? 'รวม Result\n(Fusion)' : 'Fusion\n(RRF)', TEAL);
    arrowRight(slide, prs, 5.85, 3.55);
    box(slide, prs, 9.1, 3.1, 2.5, 0.9, lang === 'th' ? 'Reranker\n(Cross-encoder)' : 'Reranker\n(Cross-encoder)', '7B3F9E');
    arrowRight(slide, prs, 11.65, 3.55);
    box(slide, prs, 10.8, 5.1, 2.1, 0.9, lang === 'th' ? 'ผลลัพธ์\nสุดท้าย' : 'Final\nResults', NAVY);
    arrowDown(slide, prs, 11.82, 4.05);
  });

  // 14. BM25 vs Dense
  addCompareSlide(prs, lang === 'th' ? 'BM25 vs Dense Search' : 'BM25 vs Dense Search',
    d.bm25.left[0], d.bm25.left.slice(1),
    d.bm25.right[0], d.bm25.right.slice(1), BLUE, TEAL);

  // 15. Basic RAG Pipeline diagram
  addDiagramSlide(prs, d.ragPipeline.title, (slide, prs) => {
    const steps = lang === 'th'
      ? ['โหลดเอกสาร', 'ตัดคำ/Chunk', 'สร้าง Embedding', 'จัดเก็บ Qdrant', '← Query ผู้ใช้', 'Embed Query', 'ค้นหา Vector', 'เพิ่ม Context', 'LLM Generate', 'คำตอบ']
      : ['Load Docs', 'Chunk', 'Embed', 'Store Qdrant', '← User Query', 'Embed Query', 'Vector Search', 'Add Context', 'LLM Generate', 'Answer'];
    const colors = [NAVY, BLUE, BLUE, TEAL, NAVY, BLUE, BLUE, TEAL, BLUE, NAVY];
    const bw = 1.1, bh = 0.75;
    // Indexing row
    for (let i = 0; i < 4; i++) {
      box(slide, prs, 0.3 + i * 1.55, 1.5, bw, bh, steps[i], colors[i]);
      if (i < 3) arrowRight(slide, prs, 0.3 + i * 1.55 + bw + 0.02, 1.82);
    }
    slide.addText(lang === 'th' ? '[ Indexing Phase ]' : '[ Indexing Phase ]', { x: 0.3, y: 1.2, w: 5.5, h: 0.3, fontSize: 11, color: TEAL, fontFace: 'Arial' });
    // Query row
    for (let i = 0; i < 6; i++) {
      box(slide, prs, 0.3 + i * 2.1, 3.2, bw + 0.8, bh, steps[4 + i], colors[4 + i]);
      if (i < 5) arrowRight(slide, prs, 0.3 + i * 2.1 + bw + 0.82, 3.52);
    }
    slide.addText(lang === 'th' ? '[ Query Phase ]' : '[ Query Phase ]', { x: 0.3, y: 2.9, w: 12, h: 0.3, fontSize: 11, color: TEAL, fontFace: 'Arial' });
  });

  // 16. Code
  addCodeSlide(prs, d.ragCode.title, d.ragCode.code);

  // 17. RAG Problems
  addContentSlide(prs, d.ragProblems.title, d.ragProblems.bullets);

  // 18. Exercise overview
  addContentSlide(prs, d.exerciseOverview.title, d.exerciseOverview.bullets);

  // 19. Exercise code
  addCodeSlide(prs, d.exercise.title, d.exercise.code);

  // 20. Summary
  addSectionSlide(prs, lang === 'th' ? 'สรุป Day 1' : 'Day 1 Summary', '');
  addContentSlide(prs, d.summary.title, d.summary.bullets);

  return prs;
}

function buildDay2(lang) {
  const d = DAY2[lang];
  const prs = newPptx();

  addTitleSlide(prs, d.title, d.workshopName, 'Workshop Day 2');

  addContentSlide(prs, lang === 'th' ? 'วาระและวัตถุประสงค์' : 'Agenda & Objectives', d.agenda.map((a, i) => `${i + 1}.  ${a}`));

  // KG diagram
  addDiagramSlide(prs, d.kgWhat.title, (slide, prs) => {
    // Node circles + edges
    box(slide, prs, 1.0, 2.5, 1.6, 1.0, lang === 'th' ? 'สมชาย\n(Person)' : 'Somchai\n(Person)', TEAL, WHITE);
    box(slide, prs, 4.5, 1.5, 1.8, 1.0, lang === 'th' ? 'X-Company\n(Org)' : 'X-Company\n(Org)', BLUE, WHITE);
    box(slide, prs, 4.5, 3.5, 1.8, 1.0, lang === 'th' ? 'กรุงเทพฯ\n(Location)' : 'Bangkok\n(Location)', NAVY, WHITE);
    box(slide, prs, 8.0, 2.5, 1.8, 1.0, lang === 'th' ? 'โครงการ A\n(Project)' : 'Project A\n(Concept)', '7B3F9E', WHITE);
    // Arrow labels
    slide.addText(lang === 'th' ? 'ทำงานที่' : 'WORKS_AT', { x: 2.7, y: 1.9, w: 1.7, h: 0.4, fontSize: 11, color: TEAL, fontFace: 'Arial', align: 'center' });
    slide.addText(lang === 'th' ? 'ตั้งอยู่ที่' : 'LOCATED_IN', { x: 6.4, y: 1.9, w: 1.5, h: 0.4, fontSize: 11, color: TEAL, fontFace: 'Arial', align: 'center' });
    slide.addText(lang === 'th' ? 'ดูแล' : 'MANAGES', { x: 6.4, y: 3.2, w: 1.5, h: 0.4, fontSize: 11, color: TEAL, fontFace: 'Arial', align: 'center' });
    const bullets = d.kgWhat.bullets.slice(3);
    slide.addText(bullets.map(b => ({ text: b, options: { fontSize: 14, color: DGRAY, bullet: { type: 'bullet' }, paraSpaceAfter: 6 } })),
      { x: 0.5, y: 5.0, w: 12.5, h: 2.2, fontFace: 'Arial' });
  });

  addCompareSlide(prs, lang === 'th' ? 'Knowledge Graph vs Vector DB' : 'Knowledge Graph vs Vector DB',
    d.kgVsVec.left[0], d.kgVsVec.left.slice(1),
    d.kgVsVec.right[0], d.kgVsVec.right.slice(1));

  addContentSlide(prs, d.entityExtract.title, d.entityExtract.bullets);
  addContentSlide(prs, d.entityTypes.title, d.entityTypes.bullets);
  addContentSlide(prs, d.relationExtract.title, d.relationExtract.bullets);
  addContentSlide(prs, d.graphPipeline.title, d.graphPipeline.bullets);
  addContentSlide(prs, d.neo4j.title, d.neo4j.bullets);
  addCodeSlide(prs, d.cypher.title, d.cypher.code);
  addContentSlide(prs, d.neo4jViz.title, d.neo4jViz.bullets);
  addContentSlide(prs, d.graphRAGWhat.title, d.graphRAGWhat.bullets);

  // GraphRAG Architecture
  addDiagramSlide(prs, d.graphRAGArch.title, (slide, prs) => {
    box(slide, prs, 0.4, 1.5, 2.2, 0.9, lang === 'th' ? 'คำถามผู้ใช้' : 'User Query', NAVY);
    arrowRight(slide, prs, 2.7, 1.9);
    box(slide, prs, 3.3, 1.5, 2.5, 0.9, lang === 'th' ? 'Query\nAnalyzer' : 'Query\nAnalyzer', BLUE);
    arrowDown(slide, prs, 4.35, 2.5);
    box(slide, prs, 1.3, 3.3, 2.5, 0.9, lang === 'th' ? 'Vector\nSearch' : 'Vector\nSearch', BLUE);
    box(slide, prs, 5.3, 3.3, 2.5, 0.9, lang === 'th' ? 'Graph\nTraversal' : 'Graph\nTraversal', TEAL);
    arrowRight(slide, prs, 3.85, 3.75);
    arrowRight(slide, prs, 7.85, 3.75);
    box(slide, prs, 8.4, 3.3, 2.2, 0.9, lang === 'th' ? 'Reranker' : 'Reranker', '7B3F9E');
    arrowRight(slide, prs, 10.65, 3.75);
    box(slide, prs, 11.2, 3.3, 1.8, 0.9, 'LLM', NAVY);
    arrowDown(slide, prs, 11.92, 4.25);
    box(slide, prs, 10.7, 5.1, 2.5, 0.9, lang === 'th' ? 'คำตอบ' : 'Answer', TEAL);
  });

  addContentSlide(prs, d.graphRAGRs.title, d.graphRAGRs.bullets);

  // Hybrid Retrieval
  addDiagramSlide(prs, d.hybridRetrieval.title, (slide, prs) => {
    box(slide, prs, 0.3, 1.5, 2.3, 0.9, lang === 'th' ? 'Query' : 'Query', NAVY);
    arrowRight(slide, prs, 2.7, 1.9);
    box(slide, prs, 3.3, 1.5, 2.5, 0.9, lang === 'th' ? 'Vector\nSearch' : 'Vector\nSearch', BLUE);
    arrowRight(slide, prs, 5.9, 1.9);
    box(slide, prs, 6.5, 1.5, 2.5, 0.9, lang === 'th' ? 'Graph\nExpansion' : 'Graph\nExpansion', TEAL);
    arrowRight(slide, prs, 9.1, 1.9);
    box(slide, prs, 9.7, 1.5, 2.5, 0.9, lang === 'th' ? 'Reranking' : 'Reranking', '7B3F9E');
    slide.addText(lang === 'th' ? '① ค้นหา Vector ก่อน (Top-20 Candidate)' : '① Vector search first (top-20 candidates)', { x: 0.3, y: 2.8, w: 12.5, h: 0.5, fontSize: 15, color: DGRAY, fontFace: 'Arial' });
    slide.addText(lang === 'th' ? '② Expand ด้วย Knowledge Graph (follow edges)' : '② Expand via Knowledge Graph (follow edges)', { x: 0.3, y: 3.4, w: 12.5, h: 0.5, fontSize: 15, color: DGRAY, fontFace: 'Arial' });
    slide.addText(lang === 'th' ? '③ Rerank ด้วย Cross-encoder → Top-5 Final' : '③ Rerank with cross-encoder → top-5 final', { x: 0.3, y: 4.0, w: 12.5, h: 0.5, fontSize: 15, color: DGRAY, fontFace: 'Arial' });
    slide.addText(lang === 'th' ? '④ ส่ง Context ให้ LLM สร้างคำตอบ' : '④ Send context to LLM for generation', { x: 0.3, y: 4.6, w: 12.5, h: 0.5, fontSize: 15, color: DGRAY, fontFace: 'Arial' });
  });

  addContentSlide(prs, d.multiHop.title, d.multiHop.bullets);
  addContentSlide(prs, d.community.title, d.community.bullets);
  addContentSlide(prs, d.advancedRetrieval.title, d.advancedRetrieval.bullets);
  addCodeSlide(prs, d.exercise2.title, d.exercise2.code);
  addSectionSlide(prs, lang === 'th' ? 'สรุป Day 2' : 'Day 2 Summary', '');
  addContentSlide(prs, d.summary2.title, d.summary2.bullets);

  return prs;
}

function buildDay3(lang) {
  const d = DAY3[lang];
  const prs = newPptx();

  addTitleSlide(prs, d.title, d.workshopName, 'Workshop Day 3');

  addContentSlide(prs, lang === 'th' ? 'วาระและวัตถุประสงค์' : 'Agenda & Objectives', d.agenda.map((a, i) => `${i + 1}.  ${a}`));

  addContentSlide(prs, d.agentWhat.title, d.agentWhat.bullets);

  // Agent Architecture
  addDiagramSlide(prs, d.agentArch.title, (slide, prs) => {
    box(slide, prs, 4.5, 1.5, 4.3, 1.0, 'LLM\n(Reasoning Core)', NAVY);
    box(slide, prs, 0.3, 3.3, 2.5, 0.9, lang === 'th' ? '🔧 Tools' : '🔧 Tools', BLUE);
    box(slide, prs, 3.4, 3.3, 2.5, 0.9, lang === 'th' ? '🧠 Memory' : '🧠 Memory', TEAL);
    box(slide, prs, 6.5, 3.3, 2.5, 0.9, lang === 'th' ? '📋 Planning' : '📋 Planning', BLUE);
    box(slide, prs, 9.6, 3.3, 2.5, 0.9, lang === 'th' ? '📡 Channel' : '📡 Channel', NAVY);
    slide.addText('↕', { x: 1.5, y: 2.6, w: 0.5, h: 0.7, fontSize: 22, color: TEAL, fontFace: 'Arial', align: 'center' });
    slide.addText('↕', { x: 4.6, y: 2.6, w: 0.5, h: 0.7, fontSize: 22, color: TEAL, fontFace: 'Arial', align: 'center' });
    slide.addText('↕', { x: 7.7, y: 2.6, w: 0.5, h: 0.7, fontSize: 22, color: TEAL, fontFace: 'Arial', align: 'center' });
    slide.addText('↕', { x: 10.8, y: 2.6, w: 0.5, h: 0.7, fontSize: 22, color: TEAL, fontFace: 'Arial', align: 'center' });
    const note = lang === 'th' ? 'LLM เป็นศูนย์กลาง ประสาน Tool, Memory, Planning และ Channel เข้าด้วยกัน'
      : 'LLM is the reasoning core, coordinating tools, memory, planning, and communication channels';
    slide.addText(note, { x: 0.5, y: 4.6, w: 12.5, h: 0.7, fontSize: 14, color: DGRAY, fontFace: 'Arial', align: 'center', italic: true });
  });

  addCodeSlide(prs, d.toolUse.title, d.toolUse.code);
  addContentSlide(prs, d.ragAgent.title, d.ragAgent.bullets);
  addContentSlide(prs, d.multiAgent.title, d.multiAgent.bullets);

  // Communication Patterns
  addDiagramSlide(prs, d.commPatterns.title, (slide, prs) => {
    // Sequential
    slide.addText(lang === 'th' ? 'Sequential (ลำดับ)' : 'Sequential', { x: 0.3, y: 1.3, w: 3.5, h: 0.4, fontSize: 14, bold: true, color: NAVY, fontFace: 'Arial' });
    box(slide, prs, 0.3, 1.8, 0.9, 0.6, 'A1', BLUE); arrowRight(slide, prs, 1.25, 2.0);
    box(slide, prs, 1.85, 1.8, 0.9, 0.6, 'A2', BLUE); arrowRight(slide, prs, 2.8, 2.0);
    box(slide, prs, 3.4, 1.8, 0.9, 0.6, 'A3', TEAL);
    // Parallel
    slide.addText(lang === 'th' ? 'Parallel (คู่ขนาน)' : 'Parallel', { x: 5.0, y: 1.3, w: 3.5, h: 0.4, fontSize: 14, bold: true, color: NAVY, fontFace: 'Arial' });
    box(slide, prs, 5.0, 1.8, 1.0, 0.6, 'Orch', NAVY);
    box(slide, prs, 6.7, 1.5, 0.9, 0.55, 'A1', BLUE);
    box(slide, prs, 6.7, 2.1, 0.9, 0.55, 'A2', BLUE);
    box(slide, prs, 6.7, 2.7, 0.9, 0.55, 'A3', BLUE);
    arrowRight(slide, prs, 6.05, 2.0);
    // Routing
    slide.addText(lang === 'th' ? 'Routing (เลือกเส้นทาง)' : 'Routing', { x: 9.0, y: 1.3, w: 4.0, h: 0.4, fontSize: 14, bold: true, color: NAVY, fontFace: 'Arial' });
    box(slide, prs, 9.0, 1.8, 1.0, 0.6, 'Router', TEAL);
    box(slide, prs, 10.7, 1.5, 1.0, 0.55, 'HR', BLUE);
    box(slide, prs, 10.7, 2.1, 1.0, 0.55, 'Fin', BLUE);
    box(slide, prs, 10.7, 2.7, 1.0, 0.55, 'Ops', BLUE);
    arrowRight(slide, prs, 10.05, 2.0);

    const bullets = lang === 'th'
      ? ['Sequential: งานต้องทำตามลำดับ (output A1 → input A2)', 'Parallel: งานอิสระทำพร้อมกัน → รวมผลลัพธ์', 'Routing: Orchestrator เลือก Agent ที่เหมาะสม']
      : ['Sequential: tasks depend on each other (A1 output → A2 input)', 'Parallel: independent tasks run simultaneously → merge results', 'Routing: orchestrator picks the right specialist agent'];
    slide.addText(bullets.map(b => ({ text: b, options: { fontSize: 15, color: DGRAY, bullet: { type: 'bullet' }, paraSpaceAfter: 8 } })),
      { x: 0.5, y: 3.8, w: 12.5, h: 3.3, fontFace: 'Arial' });
  });

  addContentSlide(prs, d.aiOSWhat.title, d.aiOSWhat.bullets);
  addContentSlide(prs, d.aiOSComponents.title, d.aiOSComponents.bullets);
  addContentSlide(prs, d.openClaw.title, d.openClaw.bullets);
  addContentSlide(prs, d.memory.title, d.memory.bullets);
  addContentSlide(prs, d.heartbeat.title, d.heartbeat.bullets);
  addContentSlide(prs, d.skills.title, d.skills.bullets);
  addContentSlide(prs, d.ragas.title, d.ragas.bullets);
  addContentSlide(prs, d.ragasMetrics.title, d.ragasMetrics.bullets);
  addCodeSlide(prs, d.llmJudge.title, d.llmJudge.code);
  addContentSlide(prs, d.abTesting.title, d.abTesting.bullets);
  addContentSlide(prs, d.capstone.title, d.capstone.bullets);
  addSectionSlide(prs, lang === 'th' ? 'สรุปและก้าวต่อไป' : 'Workshop Wrap-up', '');
  addContentSlide(prs, d.summary3.title, d.summary3.bullets);

  return prs;
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────

async function main() {
  console.log('🎨 Generating AI Workshop slide decks...\n');

  const tasks = [
    { fn: buildDay1, lang: 'th', file: 'day1-slides-th.pptx', label: 'Day 1 Thai' },
    { fn: buildDay1, lang: 'en', file: 'day1-slides-en.pptx', label: 'Day 1 English' },
    { fn: buildDay2, lang: 'th', file: 'day2-slides-th.pptx', label: 'Day 2 Thai' },
    { fn: buildDay2, lang: 'en', file: 'day2-slides-en.pptx', label: 'Day 2 English' },
    { fn: buildDay3, lang: 'th', file: 'day3-slides-th.pptx', label: 'Day 3 Thai' },
    { fn: buildDay3, lang: 'en', file: 'day3-slides-en.pptx', label: 'Day 3 English' },
  ];

  for (const t of tasks) {
    process.stdout.write(`  Building ${t.label}... `);
    const prs = t.fn(t.lang);
    const outPath = path.join(OUT_DIR, t.file);
    await prs.writeFile({ fileName: outPath });
    console.log(`✅  ${t.file}`);
  }

  console.log('\n✨ All 6 slide decks generated!');
  console.log(`📁  Output: ${OUT_DIR}`);
}

main().catch(err => { console.error(err); process.exit(1); });
