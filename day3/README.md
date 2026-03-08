# 🤖 Day 3: AI Agents & AI Operating System

> AI Agent และระบบ AI Operating System

---

## 🎯 เป้าหมายของวันนี้

หลังจากเรียนวันนี้จบ คุณจะสามารถ:
1. ✅ เข้าใจ AI Agent Architecture (Intent, Tools, Memory, Planning)
2. ✅ สร้าง Tool-Use Agent ที่เรียก function ได้
3. ✅ Build RAG Agent ที่ตอบจาก Knowledge Base
4. ✅ ออกแบบ Multi-Agent System
5. ✅ ประเมิน RAG ด้วย RAGAS metrics

---

## 📋 เนื้อหา

| Section | หัวข้อ | เวลาโดยประมาณ |
|---------|--------|---------------|
| 3.1 | AI Agent Architecture | 30 นาที |
| 3.2 | Tool-Use Agents | 45 นาที |
| 3.3 | RAG Agent | 30 นาที |
| 3.4 | Multi-Agent Systems | 45 นาที |
| 3.5 | AI Operating System (OpenClaw) | 30 นาที |
| 3.6 | Evaluation with RAGAS | 45 นาที |
| 3.7 | Capstone Project | 60 นาที |
| 💪 | Exercises | 60 นาที |

**รวม: ~5.5 ชั่วโมง**

---

## 🏗️ Architecture วันนี้

```
User Message
     ↓
  AI OS Layer (OpenClaw)
     ↓ intent detection
  Router Agent
  ├── HR Agent      → GraphRAG (HR docs) → Neo4j
  ├── Finance Agent → GraphRAG (Finance docs)
  └── IT Agent      → GraphRAG (Tech docs)
     ↓
  Response Generation (Ollama LLM)
     ↓
  User Answer
```

---

## 🏢 Production Reference: X-Company AI OS

OpenClaw ที่ X-Company ทำงานอย่างไร:
- **5 modes**: HR, Finance, Operations, Engineering, Sales
- **Auto-routing**: detect mode จาก keywords
- **GraphRAG-rs**: ค้นหาใน 40 documents
- **Neo4j**: entity relationships สำหรับ multi-hop queries
- **Telegram**: interface หลักสำหรับ 50+ employees

---

## 💡 Key Concepts

### AI Agent คืออะไร?

```
LLM-only: Question → LLM → Answer
Agent:    Question → Intent → Plan → Tools → Observation → LLM → Answer
                              ↑                                    ↓
                              ←←←←←←← Loop ←←←←←←←←←←←←←←←←←←←
```

### ReAct Pattern (Reason + Act)

```python
# Loop until answer found:
# 1. THINK: "ฉันต้องการข้อมูลเกี่ยวกับ..."
# 2. ACT: search_tool("คำค้นหา")
# 3. OBSERVE: "ผลลัพธ์จาก tool..."
# 4. THINK: "จากข้อมูลนี้ฉันสามารถตอบว่า..."
# 5. ANSWER: "คำตอบสุดท้าย"
```
