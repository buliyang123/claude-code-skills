---
name: document-ai-search
description: AI-powered semantic document search. Use when user needs to find documents by meaning rather than exact keywords. Triggers on queries like "search documents for X", "find files about Y", "look through folder for Z", "帮我搜索哪些文档提到了X". Supports PDF, DOCX, DOC, XLSX, XLS, TXT, and images (OCR) with semantic relevance scoring and automatic summary generation.
---

# Document AI Semantic Search

Performs intelligent semantic search across document folders using AI understanding rather than keyword matching.

## Overview

Unlike traditional keyword search, this skill uses semantic understanding to determine document relevance. Claude analyzes the actual meaning of document content to find matches, even when the exact keywords don't appear.

## When to Use

Trigger this skill when users ask to:
- "Search my documents for contracts about termination"
- "Find files mentioning budget planning"
- "Look through this folder for anything about API integrations"
- "帮我搜索哪些文档提到了可行性报告"
- "分析这个文件夹，找找有没有关于历史信息化项目的内容"

**Key Indicator**: User wants to find documents by **meaning** or **topic**, not exact keyword matching.

## Supported File Formats

| Format | Extension | Notes |
|--------|-----------|-------|
| PDF | `.pdf` | Text-based PDFs (image-only PDFs return no content) |
| Word | `.docx` | Office 2007+ format |
| Word | `.doc` | Legacy Word format (requires docx2txt) |
| Excel | `.xlsx` | Office 2007+ format (reads all sheets) |
| Excel | `.xls` | Legacy Excel format (97-2003) |
| Text | `.txt` | UTF-8, GBK, GB2312, Latin-1 encodings |
| Markdown | `.md` | Read as plain text |
| Images | `.jpg`, `.jpeg`, `.png` | OCR text extraction (requires Tesseract-OCR) |

## Usage

### Basic Usage

```
用户: "帮我搜索 D:\Projects 文件夹里有没有提到'可行性报告'或'历史信息化项目'的文档"
```

Claude will:
1. Scan the folder for supported files
2. Extract text content from each document
3. Use semantic understanding to find relevant documents
4. Generate a Markdown report with results

### Command Line Usage

```bash
# Basic search
python scripts/search_documents.py <folder_path> <query>

# With output file
python scripts/search_documents.py ./documents "feasibility report" -o results.md

# Custom batch size (for large folders)
python scripts/search_documents.py ~/Documents "API integration" --batch-size 10 --max-docs 100

# Quiet mode (less output)
python scripts/search_documents.py /data/files "budget" -q
```

### Arguments

| Argument | Description |
|----------|-------------|
| `folder` | Folder path to search (required) |
| `query` | Search keywords or question (required) |
| `-o, --output` | Output markdown file (default: `search_results.md`) |
| `--batch-size` | Documents per batch (default: 5) |
| `--max-docs` | Maximum documents to process (default: 50) |
| `-q, --quiet` | Suppress progress output |

## Workflow

```
1. Scan folder for supported files
   ↓
2. Extract text content from each document
   [Progress: 25%] Extracting: 12/50 - contract.pdf
   ↓
3. Analyze documents in batches for semantic relevance
   [Progress: 50%] Analyzing batch 3/10 (15/50 docs)
   ↓
4. Filter results by relevance threshold (30%)
   ↓
5. Generate Markdown report
   [Progress: 100%] Found 8 relevant documents
```

## Output Format

Generated Markdown report includes:

### Metadata
- Search query
- Folder path
- Timestamp
- Statistics (scanned, matched, skipped)

### Matched Documents
Each matched document shows:
- File name and path
- Relevance score (0-100)
- Content summary
- Relevant excerpts

### Statistics Table
- Total files scanned
- Successfully read
- Relevant matches
- Files skipped (with errors)

## How It Works

### Semantic Understanding vs Keyword Matching

**Traditional Keyword Search:**
- Finds only documents containing exact words
- Misses documents with different wording
- Returns many false positives

**AI Semantic Search:**
- Understands meaning and context
- Finds documents with related concepts
- Uses OR logic (any relevant concept = match)
- Provides relevance scoring

**Example:**
```
Query: "termination clause"

Keyword search finds: "termination clause"
AI search also finds:
- "contract ending provisions"
- "early termination rights"
- "exit conditions"
```

### Relevance Scoring

| Score | Meaning |
|-------|---------|
| 90-100 | Direct, comprehensive answer |
| 70-89 | Highly relevant, partial answer |
| 50-69 | Somewhat related, tangential content |
| 30-49 | Minimal relevance |
| 0-29 | Not relevant |

Documents below 30% are filtered out.

## Error Handling

| Error Type | Handling |
|------------|----------|
| Encrypted PDF | Skipped, logged in report |
| Corrupted file | Skipped, error message shown |
| Unsupported format | Warning, file skipped |
| Empty content | Silently skipped |
| Encoding issues (TXT) | Tries multiple encodings |

## Dependencies

Install required packages:

```bash
pip install pypdf openpyxl xlrd docx2txt pytesseract Pillow
```

Or use the requirements.txt:

```bash
pip install -r requirements.txt
```

### OCR Support (Optional)

For image text extraction (.jpg, .png, .jpeg), you need to install Tesseract-OCR separately:

**Windows:**
1. Download from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install to default location: `C:\Program Files\Tesseract-OCR\`
3. Download Chinese language pack (`chi_sim.traineddata`) to `tessdata` folder

**Linux/Mac:**
```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr tesseract-ocr-chi-sim

# macOS with Homebrew
brew install tesseract tesseract-lang
```

## Examples

### Example 1: Legal Document Search

```
用户: "搜索 D:\Legal 文件夹里关于违约条款的文档"
```

**Sample Output:**
```markdown
# Document Search Results

**Query:** `违约条款`
**Folder:** `D:\Legal`
**Scanned:** 47 documents | **Matched:** 8 documents

---

## Matched Documents

### 1. 服务合同 - TechCorp
**File:** `contracts/techcorp_service.docx`
**Relevance:** 92/100
**Summary:** 该服务合同包含详细的违约条款，包括违约通知要求（30天通知期）、补救期和违约金条款。

**Relevant Excerpts:**
> "任何一方如有重大违约，另一方可在发出书面通知30天后终止本协议"

> "违约方应向守约方支付违约金，金额为合同总价的20%"
```

### Example 2: Project Documentation Search

```
用户: "帮我找找有没有关于'历史信息化项目'的文档"
```

**Sample Output:**
```markdown
# Document Search Results

**Query:** `历史信息化项目`
**Folder:** `D:\Projects`
**Scanned:** 23 documents | **Matched:** 3 documents

---

## Matched Documents

### 1. 项目建设方案
**File:** `2024项目建设方案.docx`
**Relevance:** 88/100
**Summary:** 文档包含历史信息化项目回顾，分析了过往项目的经验教训，以及本期项目的改进措施。

**Relevant Excerpts:**
> "基于历史信息化项目的建设经验，本期项目重点解决以下问题..."

> "既往项目存在的主要问题：系统间集成度不足、数据孤岛现象严重"
```

## Tips for Best Results

1. **Use natural language queries**: Describe what you're looking for in your own words
2. **Be specific about context**: Include domain-specific terms for better relevance
3. **Use Chinese or English**: Queries work in both languages
4. **Check the excerpts**: Review relevant excerpts to verify semantic match
5. **Adjust batch size**: For large folders, increase `--batch-size` for faster processing

## Limitations

- **Scanned/image-only PDFs**: Return no content (need OCR)
- **Password-protected files**: Cannot be opened
- **Very large files**: Content truncated at 100,000 characters
- **Processing time**: AI analysis takes time per document

## See Also

- `file_reader.py` - Document extraction logic
- `search_documents.py` - Main search orchestration
