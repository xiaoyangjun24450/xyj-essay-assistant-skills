#!/usr/bin/env python3
"""
Academic Reference Searcher for docx-essay-writer skill.

Searches OpenAlex via PyAlex, formats results in GB/T 7714-2015 style.
Completely self-contained — does not import any other skill modules.

CRITICAL RULE: Never fabricate references. Only present data returned by the API.
"""

import argparse
import sys
from datetime import datetime
from typing import Callable, Dict, List, Optional

try:
    from pyalex import Works, config
except ImportError:
    print("Error: pyalex not installed. Run: pip install pyalex", file=sys.stderr)
    sys.exit(1)


class ReferenceSearcher:
    """Search for academic references using OpenAlex / PyAlex."""

    def __init__(self, api_key: Optional[str] = None,
                 formatter: Optional[Callable] = None):
        if api_key:
            config.api_key = api_key
        self.formatter = formatter or self._default_format

    # ------------------------------------------------------------------
    # Single-topic search
    # ------------------------------------------------------------------

    def search(
        self,
        query: str,
        min_year: Optional[int] = None,
        max_year: Optional[int] = None,
        min_results: int = 5,
        language: Optional[str] = None,
        min_citations: Optional[int] = None,
        open_access: Optional[bool] = None,
    ) -> List[str]:
        """Search papers and return formatted reference strings."""
        filters: Dict = {}

        if min_year and max_year:
            filters["publication_year"] = f"{min_year}-{max_year}"
        elif min_year:
            filters["publication_year"] = f">{min_year - 1}"
        elif max_year:
            filters["publication_year"] = f"<{max_year + 1}"

        if language:
            filters["language"] = language
        if open_access is not None:
            filters["is_oa"] = open_access

        work_query = Works().search(query)
        if filters:
            work_query = work_query.filter(**filters)
        work_query = work_query.sort(cited_by_count="desc")

        try:
            results = work_query.get()
        except Exception as e:
            if "401" in str(e) or "Unauthorized" in str(e):
                print(
                    "Warning: OpenAlex API requires authentication.\n"
                    "  1. Get a free key from https://openalex.org/settings/api\n"
                    "  2. Pass --api-key or set pyalex.config.api_key",
                    file=sys.stderr,
                )
                return []
            raise

        if min_citations:
            results = [w for w in results
                       if w.get("cited_by_count", 0) >= min_citations]

        references: List[str] = []
        for i, work in enumerate(results[: min_results * 2], 1):
            if len(references) >= min_results:
                break
            ref = self.formatter(work, i)
            if ref:
                references.append(ref)

        return references

    # ------------------------------------------------------------------
    # Multi-topic search with deduplication
    # ------------------------------------------------------------------

    def search_multi(
        self,
        topics: List[str],
        min_results: int = 5,
        **kwargs,
    ) -> List[str]:
        """Search multiple topics and merge results (deduplicated)."""
        all_refs: List[str] = []
        seen_titles: set = set()

        per_topic = max(2, min_results // len(topics) + 1)

        for topic in topics:
            refs = self.search(query=topic, min_results=per_topic, **kwargs)
            for ref in refs:
                title_key = ref.split(". ")[1] if ". " in ref else ref
                if title_key not in seen_titles:
                    seen_titles.add(title_key)
                    all_refs.append(ref)

        # Renumber
        renumbered: List[str] = []
        for idx, ref in enumerate(all_refs[: min_results], 1):
            renumbered.append(f"[{idx}]" + ref.split("]", 1)[-1])
        return renumbered

    # ------------------------------------------------------------------
    # Default GB/T 7714-2015 formatter
    # ------------------------------------------------------------------

    @staticmethod
    def _default_format(work: dict, index: int = 0) -> str:
        """Format an OpenAlex work dict into GB/T 7714-2015 reference string."""
        try:
            authorships = work.get("authorships", [])
            author_names: List[str] = []
            for auth in authorships[:3]:
                name = auth.get("author", {}).get("display_name", "")
                if name:
                    author_names.append(name)
            if len(authorships) > 3:
                author_names.append("et al")
            authors_str = ", ".join(author_names) if author_names else "Unknown Author"

            title = work.get("display_title") or work.get("title", "Untitled")

            host = work.get("host_venue") or {}
            if not isinstance(host, dict):
                host = {}
            if not host:
                loc = work.get("primary_location") or {}
                host = loc.get("source") or {}
            journal = host.get("display_name") or host.get("name", "") if isinstance(host, dict) else ""

            year = work.get("publication_year", "n.d.")

            biblio = work.get("biblio", {}) or {}
            volume = biblio.get("volume", "")
            issue = biblio.get("issue", "")
            first_page = biblio.get("first_page", "")
            last_page = biblio.get("last_page", "")

            pages = ""
            if first_page and last_page:
                pages = f": {first_page}-{last_page}"
            elif first_page:
                pages = f": {first_page}"

            vol_issue = ""
            if volume and issue:
                vol_issue = f", {volume}({issue})"
            elif volume:
                vol_issue = f", {volume}"

            if journal:
                return f"[{index}] {authors_str}. {title}[J]. {journal}, {year}{vol_issue}{pages}."
            return f"[{index}] {authors_str}. {title}[C]. {year}."
        except Exception as exc:
            print(f"Warning: format error: {exc}", file=sys.stderr)
            return ""


# ------------------------------------------------------------------
# CLI
# ------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Search academic references via OpenAlex"
    )
    parser.add_argument("--query", "-q", required=True, help="Search query")
    parser.add_argument("--min-year", type=int, help="Min publication year")
    parser.add_argument("--max-year", type=int, default=datetime.now().year,
                        help="Max publication year")
    parser.add_argument("--min-results", type=int, default=5,
                        help="Min results to return (default 5)")
    parser.add_argument("--language", help="Language code, e.g. 'en'")
    parser.add_argument("--min-citations", type=int,
                        help="Min citation count")
    parser.add_argument("--open-access", action="store_true",
                        help="Only open access papers")
    parser.add_argument("--api-key", help="OpenAlex API key")
    parser.add_argument("--output", "-o", help="Output file path")
    args = parser.parse_args()

    searcher = ReferenceSearcher(api_key=args.api_key)

    print(f"Searching: {args.query}", file=sys.stderr)
    references = searcher.search(
        query=args.query,
        min_year=args.min_year,
        max_year=args.max_year,
        min_results=args.min_results,
        language=args.language,
        min_citations=args.min_citations,
        open_access=args.open_access,
    )

    output = "\n".join(references)
    print(output)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output + "\n")
        print(f"\nSaved to: {args.output}", file=sys.stderr)

    return 0 if references else 1


if __name__ == "__main__":
    sys.exit(main())
