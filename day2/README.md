# 🧠 Day 2: Knowledge Graph & Advanced RAG

> Knowledge Graph และ RAG ขั้นสูง

---

## 🎯 เป้าหมายของวันนี้

หลังจากเรียนวันนี้จบ คุณจะสามารถ:
1. ✅ เข้าใจแนวคิด Knowledge Graph (Nodes, Edges, Properties)
2. ✅ Extract entities จากเอกสารภาษาไทยด้วย NER + LLM
3. ✅ สร้าง Knowledge Graph ใน Neo4j
4. ✅ เขียน Cypher queries สำหรับ graph traversal
5. ✅ Implement GraphRAG: Vector Search + Graph Context

---

## 📋 เนื้อหา

| Section | หัวข้อ | เวลาโดยประมาณ |
|---------|--------|---------------|
| 2.1 | Knowledge Graph Basics | 30 นาที |
| 2.2 | Entity Extraction (NER + LLM) | 45 นาที |
| 2.3 | Graph Construction | 30 นาที |
| 2.4 | Neo4j Setup & Cypher | 60 นาที |
| 2.5 | GraphRAG Concepts | 30 นาที |
| 2.6 | GraphRAG-rs Integration | 45 นาที |
| 2.7 | Advanced Retrieval | 30 นาที |
| 💪 | Exercises | 60 นาที |

**รวม: ~5.5 ชั่วโมง**

---

## 🏗️ Architecture วันนี้

```
Documents (Day 1 → Chunks)
         ↓
   Entity Extraction
   (LLM: identify persons, orgs, policies, dates)
         ↓
   Graph Construction
   (Node: Entity | Edge: Relationship)
         ↓
   Neo4j Graph Database
         ↓
   [Query Time]
   Question → Vector Search → Graph Expand → Merged Context → LLM → Answer
```

---

## 🔧 Requirements สำหรับวันนี้

```bash
# Neo4j (เพิ่มจาก Day 1)
docker run -d \
  --name neo4j \
  -p 7474:7474 -p 7687:7687 \
  -e NEO4J_AUTH=neo4j/workshop2026 \
  neo4j:latest

# Browser UI: http://localhost:7474
# Bolt: bolt://localhost:7687

# Python
pip install neo4j
```

---

## 💡 Key Concepts

### Knowledge Graph คืออะไร?

```
Vector Search:  "ประกัน" → chunks ที่มีคำว่าประกัน
Knowledge Graph: "ประกัน" → Entity(ประกัน MTL) → [covers] → Entity(OPD) 
                                                → [provider] → Entity(MTL)
                                                → [for] → Entity(พนักงาน)
```

### GraphRAG vs RAG ธรรมดา

| ลักษณะ | RAG | GraphRAG |
|--------|-----|----------|
| Context | Flat chunks | + Graph relationships |
| Multi-hop | ❌ | ✅ |
| Entity linking | ❌ | ✅ |
| Explainability | Low | High |

### Production Reference (PPLUS)

GraphRAG-rs ของ PPLUS:
- **Graph**: Employee → Benefit → Policy relationships
- **Multi-hop**: "ประกัน" → covers → "OPD" → requires → "ใบรับรองแพทย์"
- **Neo4j**: 200+ nodes, 500+ relationships จากเอกสาร HR + Safety
