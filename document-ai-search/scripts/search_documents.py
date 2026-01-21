#!/usr/bin/env python3
"""
Document AI Semantic Search - Main Script

Searches documents using AI semantic understanding (not keyword matching).
Supports: PDF, DOCX, DOC, XLSX, XLS, TXT, MD, JPG, JPEG, PNG

Usage:
    python search_documents.py <folder_path> <query> [options]

See SKILL.md for full documentation.
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Any, Tuple

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from file_reader import DocumentReader, ReadError


class DocumentSearcher:
    """Main search orchestrator for AI-powered document search."""

    # Minimum relevance score to include in results
    RELEVANCE_THRESHOLD = 30

    def __init__(
        self,
        folder: Path,
        query: str,
        batch_size: int = 5,
        max_docs: int = 50,
        verbose: bool = True
    ):
        self.folder = Path(folder)
        self.query = query
        self.batch_size = batch_size
        self.max_docs = max_docs
        self.verbose = verbose
        self.reader = DocumentReader()

        # Statistics
        self.stats = {
            'total_found': 0,
            'successfully_read': 0,
            'skipped': 0,
            'matched': 0,
            'errors': []
        }

    def log(self, message: str):
        """Print message if verbose mode is on."""
        if self.verbose:
            print(message)

    def discover_files(self) -> List[Path]:
        """Find supported document files in folder."""
        extensions = {'.pdf', '.docx', '.doc', '.xlsx', '.xls', '.txt', '.md', '.jpg', '.jpeg', '.png'}
        files = []

        self.log(f"\n🔍 Scanning folder: {self.folder}")

        for ext in extensions:
            found = list(self.folder.rglob(f'*{ext}'))
            files.extend(found)

        # Sort and limit
        files = sorted(files)
        if self.max_docs and len(files) > self.max_docs:
            files = files[:self.max_docs]

        self.stats['total_found'] = len(files)
        self.log(f"   Found {len(files)} supported files")

        return files

    def extract_content(self, file: Path) -> tuple[str, bool]:
        """
        Extract text from document.

        Returns:
            (content, success_flag)
        """
        try:
            content = self.reader.read(file)
            return content, True
        except ReadError as e:
            self.stats['skipped'] += 1
            self.stats['errors'].append({
                'file': str(file.relative_to(self.folder)),
                'error': str(e)
            })
            return None, False

    def _query_matches_path(self, file: Path) -> Tuple[bool, List[str], int]:
        """
        Check if query matches file path (filename or folder names).

        Args:
            file: Path to the file

        Returns:
            (is_match, matched_terms, relevance_score)
        """
        path_str = str(file).lower()
        filename = file.name.lower()

        # Split query into terms (handle Chinese and English separators)
        terms = re.split(r'[\s,，、;；]', self.query)
        terms = [t.strip().lower() for t in terms if t.strip()]

        matched = []
        for term in terms:
            if term in path_str:
                matched.append(term)

        if not matched:
            return False, [], 0

        # Calculate relevance based on match location
        relevance = 0
        for term in matched:
            # Filename match: higher score
            if term in filename:
                relevance += 40
            # Parent folder name match: medium score
            else:
                # Check if term is in any parent folder name
                for parent in file.parents:
                    if term in parent.name.lower():
                        relevance += 30
                        break
                else:
                    # Path contains term but not in folder name (partial path match)
                    relevance += 20

        return True, matched, min(relevance, 100)

    def analyze_with_claude(self, documents: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Analyze documents for semantic relevance using Claude.

        This method should be called by Claude (the AI) when running as a skill.
        When run as a standalone script, it returns placeholder results.

        Args:
            documents: List of dicts with 'file' and 'content' keys

        Returns:
            List of analysis results with relevance, summary, excerpts
        """
        # Prepare prompt for Claude
        docs_info = []
        for doc in documents:
            # Truncate content for prompt
            content_preview = doc['content'][:3000]
            docs_info.append({
                'file': str(doc['file'].relative_to(self.folder)),
                'content_preview': content_preview
            })

        prompt = f"""分析以下文档与查询的语义相关性：

查询：{self.query}

对每个文档评估：
1. 相关性评分 (0-100)：
   - 90-100: 直接且全面的回答
   - 70-89: 高度相关，部分回答
   - 50-69: 有些相关，边缘内容
   - 30-49: 最小相关性
   - 0-29: 不相关

2. 内容摘要 (2-3 句话)：文档讨论了什么

3. 相关摘录：2-3 个最相关的段落

**重要 - 使用 OR 逻辑**：
- 如果文档的任何部分与查询的任何概念相关，即标记为相关
- 使用语义理解，不是关键词匹配
- 考虑同义词、相关概念、上下文含义

文档列表：
{json.dumps(docs_info, ensure_ascii=False, indent=2)}

请以 JSON 格式回复：
{{
  "analyses": [
    {{
      "file": "文件路径",
      "relevance": 85,
      "summary": "文档讨论了...",
      "excerpts": ["摘录1", "摘录2"]
    }}
  ]
}}"""

        # When running as skill, Claude will analyze this prompt
        # For now, return placeholder
        print("\n" + "="*60)
        print("📝 Claude Analysis Prompt:")
        print("="*60)
        print(prompt)
        print("="*60)
        print("\n⚠️  Note: This script is designed to run as a Claude Code skill.")
        print("    Claude will analyze the documents and return structured results.")
        print("="*60 + "\n")

        # Placeholder results
        return []

    def search(self) -> List[Dict[str, Any]]:
        """
        Execute search with progress tracking.

        Returns:
            List of matched documents with analysis results
        """
        # Step 1: Discover files
        files = self.discover_files()
        if not files:
            self.log("   No supported files found!")
            return []

        # Step 1.5: Path matching (check filenames and folder names)
        self.log(f"\n🔍 Checking file/folder names for: '{self.query}'")
        path_match_results = {}
        for file in files:
            is_match, matched_terms, relevance = self._query_matches_path(file)
            if is_match:
                path_match_results[str(file)] = {
                    'file': file,
                    'path_relevance': relevance,
                    'matched_terms': matched_terms
                }

        path_match_count = len(path_match_results)
        if path_match_count > 0:
            self.log(f"   ✅ Found {path_match_count} files with matching names")

        # Step 2: Extract content from all files
        self.log(f"\n📄 Extracting content from {len(files)} files...")

        documents = []
        for i, file in enumerate(files, 1):
            progress = f"   [{i}/{len(files)}] {file.name}"
            self.log(progress)

            content, success = self.extract_content(file)
            if success and content.strip():
                documents.append({
                    'file': file,
                    'content': content
                })
                self.stats['successfully_read'] += 1

        self.log(f"   ✅ Successfully read: {self.stats['successfully_read']} files")
        if self.stats['skipped'] > 0:
            self.log(f"   ⚠️  Skipped: {self.stats['skipped']} files")

        if not documents:
            self.log("\n   No content could be extracted!")

        # Step 3: Process in batches (only if we have content)
        all_analyses = []
        if documents:
            self.log(f"\n🤖 Analyzing document content for: '{self.query}'")
            self.log(f"   Batch size: {self.batch_size} | Batches: {(len(documents) + self.batch_size - 1) // self.batch_size}")

            batches = [
                documents[i:i + self.batch_size]
                for i in range(0, len(documents), self.batch_size)
            ]

            for batch_num, batch in enumerate(batches, 1):
                self.log(f"   🔄 Batch {batch_num}/{len(batches)} ({len(batch)} docs)...")

                # This will be handled by Claude when running as skill
                analyses = self.analyze_with_claude(batch)
                all_analyses.extend(analyses)

        # Step 4: Merge path matches with content matches
        # Track which files had content analysis
        content_analyzed_files = {a.get('file', ''): a for a in all_analyses}

        final_results = []

        # Process path matches
        for file_key, path_info in path_match_results.items():
            if file_key in content_analyzed_files:
                # File has both path match AND content match - combine scores
                content_analysis = content_analyzed_files[file_key]
                content_relevance = content_analysis.get('relevance', 0)
                path_relevance = path_info['path_relevance']

                # Combined score (weighted average, path match gives boost)
                combined_relevance = min(100, content_relevance + (path_relevance * 0.3))

                final_results.append({
                    'file': path_info['file'],
                    'relevance': int(combined_relevance),
                    'summary': content_analysis.get('summary', ''),
                    'excerpts': content_analysis.get('excerpts', []),
                    'match_sources': ['path', 'content'],
                    'matched_terms': path_info['matched_terms'],
                    'path_relevance': path_relevance,
                    'content_relevance': content_relevance
                })
            else:
                # File only has path match (no content or content read failed)
                final_results.append({
                    'file': path_info['file'],
                    'relevance': path_info['path_relevance'],
                    'summary': f"文件/文件夹名包含: {', '.join(path_info['matched_terms'])}",
                    'excerpts': [f"路径匹配: {path_info['file'].relative_to(self.folder)}"],
                    'match_sources': ['path'],
                    'matched_terms': path_info['matched_terms']
                })

        # Add content-only matches (no path match)
        for analysis in all_analyses:
            file_key = str(analysis.get('file', ''))
            if file_key not in path_match_results:
                analysis['match_sources'] = ['content']
                final_results.append(analysis)

        # Filter by relevance threshold
        matched = [
            r for r in final_results
            if r.get('relevance', 0) >= self.RELEVANCE_THRESHOLD
        ]
        matched.sort(key=lambda x: x.get('relevance', 0), reverse=True)

        self.stats['matched'] = len(matched)
        self.log(f"\n   📊 Results: {len(matched)} relevant documents found")

        return matched

    def generate_report(self, results: List[Dict[str, Any]], output_path: Path):
        """Generate Markdown report."""
        report_lines = [
            "# Document Search Results",
            "",
            f"**Query:** `{self.query}`",
            f"**Folder:** `{self.folder}`",
            f"**Scanned:** {self.stats['total_found']} documents | **Matched:** {self.stats['matched']} documents | **Skipped:** {self.stats['skipped']}",
            f"**Generated:** {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')} UTC",
            "",
            "---",
            ""
        ]

        # Matched documents
        if results:
            report_lines.append("## Matched Documents\n")

            for i, result in enumerate(results, 1):
                relevance = result.get('relevance', 0)
                file_path = result.get('file', 'Unknown')
                summary = result.get('summary', 'No summary available.')
                excerpts = result.get('excerpts', [])
                match_sources = result.get('match_sources', ['content'])
                matched_terms = result.get('matched_terms', [])

                # Get just the filename for heading
                file_name = Path(file_path).name
                if file_name != file_path:
                    file_info = f"{file_name}\n**File:** `{file_path}`"
                else:
                    file_info = f"`{file_path}`"

                # Add match source badge
                source_badge = " 🔍" if 'path' in match_sources else ""

                report_lines.append(f"### {i}. {file_name}{source_badge}")
                report_lines.append(f"**Relevance:** {relevance}/100")

                # Show match sources if path match
                if 'path' in match_sources:
                    source_labels = []
                    if 'path' in match_sources:
                        source_labels.append("文件/文件夹名")
                    if 'content' in match_sources:
                        source_labels.append("内容")
                    report_lines.append(f"**Match Sources:** {', '.join(source_labels)}")

                    # Show matched terms
                    if matched_terms:
                        report_lines.append(f"**Matched Terms:** {', '.join(matched_terms)}")

                    # Show breakdown if both matches
                    if 'content' in match_sources:
                        path_rel = result.get('path_relevance', 0)
                        content_rel = result.get('content_relevance', 0)
                        report_lines.append(f"**Score Breakdown:** 路径匹配 {path_rel} + 内容匹配 {content_rel}")

                report_lines.append(f"**Summary:** {summary}")

                if excerpts:
                    report_lines.append("\n**Relevant Excerpts:**")
                    for excerpt in excerpts:
                        report_lines.append(f"> {excerpt}")

                report_lines.append("\n---\n")

        # Statistics
        report_lines.extend([
            "## Statistics\n",
            "| Metric | Count |",
            "|--------|-------|",
            f"| Total files scanned | {self.stats['total_found']} |",
            f"| Successfully read | {self.stats['successfully_read']} |",
            f"| Relevant matches (≥{self.RELEVANCE_THRESHOLD}%) | {self.stats['matched']} |",
            f"| Files skipped (errors) | {self.stats['skipped']} |",
            ""
        ])

        # Skipped files
        if self.stats['errors']:
            report_lines.append("## Skipped Files\n")
            for error in self.stats['errors']:
                report_lines.append(f"- `{error['file']}` - {error['error']}")
            report_lines.append("")

        # Footer
        report_lines.extend([
            "---",
            "",
            "*Generated by Document AI Search skill*"
        ])

        # Write report
        report_content = "\n".join(report_lines)
        output_path.write_text(report_content, encoding='utf-8')

        self.log(f"\n✅ Report saved to: {output_path}")


def main():
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="AI-powered semantic document search",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s ./documents "feasibility report"
  %(prog)s ~/Documents "API integration" -o results.md
  %(prog)s /data/files "budget forecast" --batch-size 10 --max-docs 100
        """
    )

    parser.add_argument(
        "folder",
        type=Path,
        help="Folder to search"
    )

    parser.add_argument(
        "query",
        help="Search query or keywords"
    )

    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=Path("search_results.md"),
        help="Output markdown file (default: search_results.md)"
    )

    parser.add_argument(
        "--batch-size",
        type=int,
        default=5,
        metavar="N",
        help="Documents per batch (default: 5)"
    )

    parser.add_argument(
        "--max-docs",
        type=int,
        default=50,
        metavar="N",
        help="Maximum documents to process (default: 50)"
    )

    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress progress output"
    )

    args = parser.parse_args()

    # Validate folder
    if not args.folder.exists():
        print(f"❌ Error: Folder not found: {args.folder}", file=sys.stderr)
        sys.exit(1)

    if not args.folder.is_dir():
        print(f"❌ Error: Not a folder: {args.folder}", file=sys.stderr)
        sys.exit(1)

    # Create searcher
    searcher = DocumentSearcher(
        folder=args.folder,
        query=args.query,
        batch_size=args.batch_size,
        max_docs=args.max_docs,
        verbose=not args.quiet
    )

    # Execute search
    results = searcher.search()

    # Generate report
    searcher.generate_report(results, args.output)

    # Exit with appropriate code
    if results:
        print(f"\n🎉 Found {len(results)} relevant documents!")
        sys.exit(0)
    else:
        print(f"\n📭 No relevant documents found.")
        sys.exit(0)


if __name__ == "__main__":
    main()
