#!/usr/bin/env python3
"""
PyAlex Academic Reference Searcher
Search for academic papers and generate formatted references.
"""

import argparse
import sys
from typing import List, Optional, Callable
from datetime import datetime

try:
    from pyalex import Works, config
except ImportError:
    print("Error: pyalex not installed. Run: pip install pyalex")
    sys.exit(1)


class ReferenceSearcher:
    """Search for academic references using OpenAlex/PyAlex."""
    
    def __init__(self, api_key: Optional[str] = None, formatter: Optional[Callable] = None):
        """
        Initialize the reference searcher.
        
        Args:
            api_key: OpenAlex API key (optional but recommended)
            formatter: Custom formatting function for references
        """
        if api_key:
            config.api_key = api_key
        self.formatter = formatter or self._default_format
    
    def search(
        self,
        query: str,
        min_year: Optional[int] = None,
        max_year: Optional[int] = None,
        min_results: int = 5,
        language: Optional[str] = None,
        min_citations: Optional[int] = None,
        open_access: Optional[bool] = None
    ) -> List[str]:
        """
        Search for academic papers and return formatted references.
        
        Args:
            query: Search query string
            min_year: Minimum publication year (inclusive)
            max_year: Maximum publication year (inclusive)
            min_results: Minimum number of results to return
            language: Language code (e.g., "en" for English)
            min_citations: Minimum citation count
            open_access: Filter for open access papers only
            
        Returns:
            List of formatted reference strings
        """
        # Build filter dictionary
        filters = {}
        
        # Handle year filter
        if min_year and max_year:
            filters["publication_year"] = f"{min_year}-{max_year}"
        elif min_year:
            filters["publication_year"] = f">{min_year-1}"
        elif max_year:
            filters["publication_year"] = f"<{max_year+1}"
            
        if language:
            filters["language"] = language
            
        if open_access is not None:
            filters["is_oa"] = open_access
        
        # Build query
        work_query = Works().search(query)
        
        # Apply filters
        if filters:
            work_query = work_query.filter(**filters)
            
        # Sort by citation count for relevance
        work_query = work_query.sort(cited_by_count="desc")
        
        # Get results
        try:
            results = work_query.get()
        except Exception as e:
            if "401" in str(e) or "Unauthorized" in str(e):
                print("Warning: OpenAlex API requires authentication. Please set an API key:")
                print("  1. Get a free API key from https://openalex.org/settings/api")
                print("  2. Use --api-key parameter or set pyalex.config.api_key")
                return []
            raise
        
        # Filter by minimum citations if specified
        if min_citations:
            results = [w for w in results if w.get("cited_by_count", 0) >= min_citations]
        
        # Format references
        references = []
        for i, work in enumerate(results[:min_results * 2], 1):  # Get more to ensure quality
            if len(references) >= min_results:
                break
            ref = self.formatter(work, i)
            if ref:
                references.append(ref)
        
        return references
    
    def search_multi(
        self,
        topics: List[str],
        min_year: Optional[int] = None,
        max_year: Optional[int] = None,
        min_results: int = 5,
        **kwargs
    ) -> List[str]:
        """
        Search multiple topics and merge results.
        
        Args:
            topics: List of search queries
            min_year: Minimum publication year
            max_year: Maximum publication year
            min_results: Minimum total results
            **kwargs: Additional filters
            
        Returns:
            List of formatted reference strings (deduplicated)
        """
        all_refs = []
        seen_ids = set()
        
        for topic in topics:
            results = self.search(
                query=topic,
                min_year=min_year,
                max_year=max_year,
                min_results=max(2, min_results // len(topics)),
                **kwargs
            )
            
            for ref in results:
                # Simple deduplication based on work ID in the reference
                work_id = ref.split('.')[0] if '.' in ref else ref
                if work_id not in seen_ids:
                    seen_ids.add(work_id)
                    all_refs.append(ref)
        
        return all_refs[:min_results]
    
    def _default_format(self, work: dict, index: int = 0) -> str:
        """
        Format a work into GB/T 7714-2015 style reference.
        
        Args:
            work: OpenAlex work dictionary
            index: Reference number
            
        Returns:
            Formatted reference string
        """
        try:
            # Extract authors
            authorships = work.get("authorships", [])
            author_names = []
            
            for auth in authorships[:3]:
                author = auth.get("author", {})
                name = author.get("display_name", "")
                if name:
                    # Convert "First Last" to "First Last" format for Chinese papers
                    author_names.append(name)
            
            if len(authorships) > 3:
                author_names.append("et al")
            
            authors_str = ", ".join(author_names) if author_names else "Unknown Author"
            
            # Extract title
            title = work.get("display_title") or work.get("title", "Untitled")
            
            # Extract journal/source
            host_venue = work.get("host_venue") or work.get("primary_location", {}).get("source", {})
            if isinstance(host_venue, dict):
                journal = host_venue.get("display_name") or host_venue.get("name", "")
            else:
                journal = ""
            
            # Extract year
            year = work.get("publication_year", "n.d.")
            
            # Extract volume/issue/pages
            biblio = work.get("biblio", {})
            volume = biblio.get("volume", "")
            issue = biblio.get("issue", "")
            first_page = biblio.get("first_page", "")
            last_page = biblio.get("last_page", "")
            
            # Build pages string
            pages = ""
            if first_page and last_page:
                pages = f": {first_page}-{last_page}"
            elif first_page:
                pages = f": {first_page}"
            
            # Build vol/issue string
            vol_issue = ""
            if volume and issue:
                vol_issue = f", {volume}({issue})"
            elif volume:
                vol_issue = f", {volume}"
            
            # Build reference string in GB/T 7714-2015 format
            if journal:
                # Journal article: [N] Author. Title[J]. Journal, Year, Vol(Issue): Pages.
                ref = f"[{index}] {authors_str}. {title}[J]. {journal}, {year}{vol_issue}{pages}."
            else:
                # Other types
                ref = f"[{index}] {authors_str}. {title}[C]. {year}."
            
            return ref
            
        except Exception as e:
            print(f"Warning: Failed to format work: {e}")
            return ""


def main():
    """Command line interface."""
    parser = argparse.ArgumentParser(
        description="Search for academic references using OpenAlex"
    )
    parser.add_argument(
        "--query", "-q",
        required=True,
        help="Search query for papers"
    )
    parser.add_argument(
        "--min-year",
        type=int,
        help="Minimum publication year"
    )
    parser.add_argument(
        "--max-year",
        type=int,
        default=datetime.now().year,
        help="Maximum publication year (default: current year)"
    )
    parser.add_argument(
        "--min-results",
        type=int,
        default=5,
        help="Minimum number of results (default: 5)"
    )
    parser.add_argument(
        "--language",
        help="Language filter (e.g., 'en' for English)"
    )
    parser.add_argument(
        "--min-citations",
        type=int,
        help="Minimum citation count"
    )
    parser.add_argument(
        "--open-access",
        action="store_true",
        help="Only include open access papers"
    )
    parser.add_argument(
        "--api-key",
        help="OpenAlex API key"
    )
    parser.add_argument(
        "--output", "-o",
        help="Output file path"
    )
    
    args = parser.parse_args()
    
    # Create searcher
    searcher = ReferenceSearcher(api_key=args.api_key)
    
    # Search
    print(f"Searching for: {args.query}")
    if args.min_year:
        print(f"Year range: {args.min_year}-{args.max_year}")
    if args.language:
        print(f"Language: {args.language}")
    print("-" * 50)
    
    references = searcher.search(
        query=args.query,
        min_year=args.min_year,
        max_year=args.max_year,
        min_results=args.min_results,
        language=args.language,
        min_citations=args.min_citations,
        open_access=args.open_access
    )
    
    # Output results
    output = "\n".join(references)
    print(output)
    
    # Save to file if specified
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output + "\n")
        print(f"\nSaved to: {args.output}")
    
    return 0 if references else 1


if __name__ == "__main__":
    sys.exit(main())
