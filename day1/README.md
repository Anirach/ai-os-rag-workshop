# 📚 Day 1: Data Engineering & RAG Foundation

> วิศวกรรมข้อมูลและพื้นฐาน RAG

---

## 🎯 เป้าหมายของวันนี้

หลังจากเรียนวันนี้จบ คุณจะสามารถ:
1. ✅ อ่านและ process เอกสาร PDF/DOCX ภาษาไทยได้
2. ✅ ตัดคำภาษาไทยและสร้าง Chunks อย่างมีประสิทธิภาพ
3. ✅ สร้าง Embeddings ด้วย bge-m3 ผ่าน Ollama
4. ✅ จัดเก็บและค้นหาใน Qdrant Vector Database
5. ✅ สร้าง Basic RAG Pipeline แบบ end-to-end

---

## 📋 เนื้อหา

| Section | หัวข้อ | เวลาโดยประมาณ |
|---------|--------|---------------|
| 1.1 | Document Processing | 30 นาที |
| 1.2 | Thai NLP & Tokenization | 30 นาที |
| 1.3 | Chunking Strategies | 30 นาที |
| 1.4 | Embedding Generation | 30 นาที |
| 1.5 | Qdrant Vector Database | 45 นาที |
| 1.6 | Hybrid Search | 30 นาที |
| 1.7 | Basic RAG Pipeline | 45 นาที |
| 💪 | Exercises | 60 นาที |

**รวม: ~5.5 ชั่วโมง**

---

## 🏗️ Architecture วันนี้

```
PDF/DOCX Documents
       ↓
  Text Extraction (PyMuPDF)
       ↓
  Thai Tokenization (PyThaiNLP)
       ↓
  Chunking (LangChain)
       ↓
  Embedding (Ollama: bge-m3)
       ↓
  Vector Store (Qdrant)
       ↓
  [Query Time]
  User Query → Embed → Search → Context → LLM → Answer
```

---

## 📂 ไฟล์ในโฟลเดอร์นี้

```
day1/
├── README.md                     # ไฟล์นี้
├── day1_data_engineering.ipynb   # Notebook หลัก (เนื้อหา)
├── day1_exercises.ipynb          # แบบฝึกหัด
└── sample_data/
    └── README.md                 # วิธีเตรียม sample data
```

---

## 🔧 Requirements สำหรับวันนี้

```bash
# Services ที่ต้องรัน:
# 1. Ollama (พร้อม bge-m3 model)
ollama serve
ollama pull bge-m3
ollama pull qwen3:8b

# 2. Qdrant
docker run -d --name qdrant -p 6333:6333 qdrant/qdrant:latest

# ทดสอบ
curl http://localhost:11434/api/tags    # Ollama
curl http://localhost:6333/health      # Qdrant
```

---

## 💡 Key Concepts

### RAG คืออะไร?
**RAG (Retrieval-Augmented Generation)** คือเทคนิคที่ทำให้ LLM ตอบคำถามจากเอกสารของเราได้ โดยไม่ต้อง fine-tune model ใหม่

```
แบบเดิม: Question → LLM → Answer (จาก training data เท่านั้น)
RAG:      Question → Search Docs → Context + Question → LLM → Answer
```

### ทำไมต้อง Local?
- 🔒 **Privacy**: เอกสาร HR, Finance ไม่ควรส่งออก Cloud
- 💰 **Cost**: ไม่มี API call costs
- 🚀 **Speed**: Latency ต่ำกว่าเพราะอยู่ใน LAN
- 🔧 **Control**: ปรับแต่งได้ทุกอย่าง

---

## 📝 Notes จากระบบจริง (PPLUS)

ระบบ GraphRAG-rs ที่ PPLUS ใช้จริง:
- **1,728 chunks** จาก 40 documents
- **bge-m3** embeddings (multilingual, รองรับไทย)
- **Qdrant** เป็น vector database
- Search latency ประมาณ **50-200ms**

วันนี้เราจะ build pipeline ใกล้เคียงกันนี้ตั้งแต่ต้น! 🚀
