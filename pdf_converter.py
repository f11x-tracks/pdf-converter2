#!/usr/bin/env python3
"""
PDF to Excel Converter
Converts PDF files containing tables to Excel format using multiple extraction methods.
"""

import pandas as pd
import pdfplumber
import tabula
from pathlib import Path
import sys
import logging
from typing import List, Optional
import argparse

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFToExcelConverter:
    def __init__(self, pdf_path: str, output_path: Optional[str] = None):
        """
        Initialize the PDF to Excel converter.
        
        Args:
            pdf_path (str): Path to the input PDF file
            output_path (str, optional): Path for the output Excel file
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        if output_path:
            self.output_path = Path(output_path)
        else:
            self.output_path = self.pdf_path.with_suffix('.xlsx')
    
    def extract_tables_with_pdfplumber(self) -> List[pd.DataFrame]:
        """
        Extract tables using pdfplumber library.
        
        Returns:
            List[pd.DataFrame]: List of extracted tables as DataFrames
        """
        tables = []
        logger.info("Extracting tables using pdfplumber...")
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    logger.info(f"Processing page {page_num}/{len(pdf.pages)}")
                    
                    # Extract tables from the page
                    page_tables = page.extract_tables()
                    
                    for table_num, table in enumerate(page_tables, 1):
                        if table and len(table) > 1:  # Ensure table has data
                            try:
                                # Convert to DataFrame
                                df = pd.DataFrame(table[1:], columns=table[0])
                                df = self._clean_dataframe(df)
                                
                                if not df.empty:
                                    df.name = f"Page_{page_num}_Table_{table_num}"
                                    tables.append(df)
                                    logger.info(f"Extracted table from page {page_num}, table {table_num}: {df.shape}")
                            except Exception as e:
                                logger.warning(f"Error processing table on page {page_num}: {e}")
                    
                    # Also try to extract text and look for structured data
                    if not page_tables:
                        text = page.extract_text()
                        if text and self._looks_like_tabular_data(text):
                            try:
                                df = self._extract_table_from_text(text, page_num)
                                if df is not None and not df.empty:
                                    df.name = f"Page_{page_num}_Text_Table"
                                    tables.append(df)
                                    logger.info(f"Extracted text table from page {page_num}: {df.shape}")
                            except Exception as e:
                                logger.warning(f"Error extracting text table from page {page_num}: {e}")
        
        except Exception as e:
            logger.error(f"Error extracting tables with pdfplumber: {e}")
        
        return tables
    
    def extract_tables_with_tabula(self) -> List[pd.DataFrame]:
        """
        Extract tables using tabula-py library.
        
        Returns:
            List[pd.DataFrame]: List of extracted tables as DataFrames
        """
        tables = []
        logger.info("Extracting tables using tabula-py...")
        
        try:
            # Extract all tables from all pages
            dfs = tabula.read_pdf(str(self.pdf_path), pages='all', multiple_tables=True)
            
            for i, df in enumerate(dfs, 1):
                if not df.empty:
                    df = self._clean_dataframe(df)
                    if not df.empty:
                        df.name = f"Tabula_Table_{i}"
                        tables.append(df)
                        logger.info(f"Extracted table {i} with tabula: {df.shape}")
        
        except Exception as e:
            logger.error(f"Error extracting tables with tabula: {e}")
        
        return tables
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and preprocess the extracted DataFrame.
        
        Args:
            df (pd.DataFrame): Raw extracted DataFrame
            
        Returns:
            pd.DataFrame: Cleaned DataFrame
        """
        # Remove completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Remove rows where all values are None or empty strings
        df = df[~df.apply(lambda row: all(pd.isna(val) or val == '' for val in row), axis=1)]
        
        # Strip whitespace from string columns
        for col in df.columns:
            try:
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].replace('nan', '')
            except Exception:
                # Handle cases where dtype check might fail
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace('nan', '')
        
        return df
    
    def _looks_like_tabular_data(self, text: str) -> bool:
        """
        Check if text contains patterns that suggest tabular data.
        Enhanced to detect various table formats and patterns.
        
        Args:
            text (str): Text to analyze
            
        Returns:
            bool: True if text appears to contain tabular data
        """
        lines = text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        if len(lines) < 3:
            return False
        
        # Enhanced detection patterns
        table_indicators = 0
        
        # Check for consistent column patterns
        separator_count = 0
        for line in lines[:10]:  # Check first 10 lines
            if len(line.split()) > 2:  # Multiple columns
                separator_count += 1
        
        if separator_count > 2:
            table_indicators += 1
        
        # Look for common table separators
        separator_chars = ['|', '\t', '  ', ',']
        for char in separator_chars:
            if sum(1 for line in lines[:5] if char in line) >= 3:
                table_indicators += 1
                break
        
        # Check for numeric patterns (common in tables)
        numeric_lines = 0
        for line in lines[:10]:
            if any(char.isdigit() for char in line):
                numeric_lines += 1
        
        if numeric_lines >= 3:
            table_indicators += 1
        
        # Look for header-like patterns (all caps, underscores, etc.)
        if lines:
            first_line = lines[0]
            if (first_line.isupper() or '_' in first_line or 
                any(word in first_line.lower() for word in ['name', 'date', 'id', 'code', 'amount', 'total'])):
                table_indicators += 1
        
        return table_indicators >= 2
    
    def extract_mixed_content(self) -> tuple[List[pd.DataFrame], List[dict]]:
        """
        Extract both tables and text content from PDF.
        
        Returns:
            tuple: (tables_list, text_content_list)
                - tables_list: List of DataFrames containing table data
                - text_content_list: List of dicts with page text content
        """
        tables = []
        text_content = []
        logger.info("Extracting mixed content (tables and text)...")
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    logger.info(f"Processing page {page_num}/{len(pdf.pages)} for mixed content")
                    
                    # Extract tables first
                    page_tables = page.extract_tables()
                    table_count = 0
                    
                    if page_tables:
                        for i, table in enumerate(page_tables):
                            try:
                                if table and len(table) > 1:
                                    df = pd.DataFrame(table[1:], columns=table[0])
                                    df = self._clean_dataframe(df)
                                    if df is not None and not df.empty:
                                        df.name = f"Page_{page_num}_Table_{i+1}"
                                        tables.append(df)
                                        table_count += 1
                                        logger.info(f"Extracted table from page {page_num}, table {i+1}: {df.shape}")
                            except Exception as e:
                                logger.warning(f"Error processing table on page {page_num}: {e}")
                    
                    # Extract all text from page
                    try:
                        full_text = page.extract_text()
                        if full_text:
                            # Identify non-tabular text sections
                            text_sections = self._extract_text_sections(full_text, page_num)
                            if text_sections:
                                text_content.append({
                                    'page': page_num,
                                    'sections': text_sections,
                                    'has_tables': table_count > 0
                                })
                    except Exception as e:
                        logger.warning(f"Error extracting text from page {page_num}: {e}")
                        
        except Exception as e:
            logger.error(f"Error processing PDF: {e}")
        
        return tables, text_content
    
    def _extract_text_sections(self, text: str, page_num: int) -> List[dict]:
        """
        Extract and categorize text sections from a page.
        
        Args:
            text (str): Full page text
            page_num (int): Page number
            
        Returns:
            List[dict]: List of text sections with metadata
        """
        sections = []
        paragraphs = text.split('\n\n')
        
        for i, paragraph in enumerate(paragraphs):
            paragraph = paragraph.strip()
            if not paragraph:
                continue
                
            # Categorize the text section
            section_type = self._classify_text_section(paragraph)
            
            sections.append({
                'section_id': i + 1,
                'type': section_type,
                'content': paragraph,
                'word_count': len(paragraph.split()),
                'line_count': len(paragraph.split('\n'))
            })
        
        return sections
    
    def _classify_text_section(self, text: str) -> str:
        """
        Classify a text section by type.
        
        Args:
            text (str): Text to classify
            
        Returns:
            str: Section type (header, paragraph, list, etc.)
        """
        text_lower = text.lower()
        lines = text.split('\n')
        
        # Check for headers (short, potentially all caps, special formatting)
        if len(text) < 100 and (text.isupper() or text.istitle()):
            return "header"
        
        # Check for lists
        if any(line.strip().startswith(('•', '-', '*', '1.', '2.', 'a)', 'i)')) for line in lines):
            return "list"
        
        # Check for table-like structure
        if self._looks_like_tabular_data(text):
            return "table_text"
        
        # Check for metadata/footer
        if any(word in text_lower for word in ['page', 'copyright', '©', 'confidential', 'proprietary']):
            return "metadata"
        
        # Default to paragraph
        return "paragraph"
    
    def _extract_table_from_text(self, text: str, page_num: int) -> Optional[pd.DataFrame]:
        """
        Try to extract tabular data from plain text.
        
        Args:
            text (str): Text to parse
            page_num (int): Page number for reference
            
        Returns:
            Optional[pd.DataFrame]: Extracted DataFrame or None
        """
        lines = text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        if len(lines) < 2:
            return None
        
        # Try to identify columns by looking for consistent spacing
        table_data = []
        for line in lines:
            # Split by multiple spaces or tabs
            parts = [part.strip() for part in line.split() if part.strip()]
            if len(parts) > 1:
                table_data.append(parts)
        
        if len(table_data) < 2:
            return None
        
        # Use first row as headers, rest as data
        try:
            max_cols = max(len(row) for row in table_data)
            
            # Pad rows to have the same number of columns
            for row in table_data:
                while len(row) < max_cols:
                    row.append('')
            
            df = pd.DataFrame(table_data[1:], columns=table_data[0])
            return self._clean_dataframe(df)
        except Exception:
            return None
    
    def save_to_excel(self, tables: List[pd.DataFrame], single_sheet: bool = False) -> None:
        """
        Save extracted tables to Excel file.
        
        Args:
            tables (List[pd.DataFrame]): List of tables to save
            single_sheet (bool): If True, combine all tables into one sheet
        """
        if not tables:
            logger.warning("No tables found to save")
            return
        
        logger.info(f"Saving {len(tables)} tables to {self.output_path}")
        
        if single_sheet:
            self._save_to_single_sheet(tables)
        else:
            self._save_to_multiple_sheets(tables)
        
        logger.info(f"Excel file saved successfully: {self.output_path}")
    
    def save_mixed_content_to_excel(self, tables: List[pd.DataFrame], text_content: List[dict], 
                                   mixed_format: str = "separate_sheets") -> None:
        """
        Save mixed content (tables and text) to Excel with different organization options.
        
        Args:
            tables (List[pd.DataFrame]): Extracted tables
            text_content (List[dict]): Extracted text content
            mixed_format (str): How to organize content ("separate_sheets", "combined_sheets", "text_summary")
        """
        if not tables and not text_content:
            logger.warning("No content found to save")
            return
        
        logger.info(f"Saving {len(tables)} tables and {len(text_content)} pages of text content")
        
        with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
            
            if mixed_format == "separate_sheets":
                # Save tables to separate sheets
                for i, table in enumerate(tables):
                    sheet_name = getattr(table, 'name', f'Table_{i+1}')
                    table.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
                # Create text summary sheets
                self._save_text_summary_sheets(writer, text_content)
                
            elif mixed_format == "combined_sheets":
                # Combine content by page
                self._save_combined_page_sheets(writer, tables, text_content)
                
            elif mixed_format == "text_summary":
                # Save tables normally, add comprehensive text summary
                for i, table in enumerate(tables):
                    sheet_name = getattr(table, 'name', f'Table_{i+1}')
                    table.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
                # Create comprehensive text analysis
                self._save_comprehensive_text_analysis(writer, text_content)
        
        logger.info(f"Mixed content saved successfully: {self.output_path}")
    
    def _save_text_summary_sheets(self, writer, text_content: List[dict]):
        """Save text content organized by type to separate sheets."""
        
        # Organize text by type
        text_by_type = {}
        all_text_data = []
        
        for page_data in text_content:
            page_num = page_data['page']
            for section in page_data['sections']:
                section_type = section['type']
                
                if section_type not in text_by_type:
                    text_by_type[section_type] = []
                
                text_by_type[section_type].append({
                    'Page': page_num,
                    'Section': section['section_id'],
                    'Content': section['content'][:500] + '...' if len(section['content']) > 500 else section['content'],
                    'Word_Count': section['word_count'],
                    'Full_Content': section['content']
                })
                
                all_text_data.append({
                    'Page': page_num,
                    'Section': section['section_id'],
                    'Type': section_type,
                    'Content': section['content'][:200] + '...' if len(section['content']) > 200 else section['content'],
                    'Word_Count': section['word_count']
                })
        
        # Save overview sheet
        if all_text_data:
            df_overview = pd.DataFrame(all_text_data)
            df_overview.to_excel(writer, sheet_name='Text_Overview', index=False)
        
        # Save sheets by content type
        for content_type, content_list in text_by_type.items():
            if content_list:
                df_type = pd.DataFrame(content_list)
                # Remove Full_Content for display sheets (too long)
                if 'Full_Content' in df_type.columns:
                    df_type = df_type.drop('Full_Content', axis=1)
                sheet_name = f"Text_{content_type.title()}"[:31]
                df_type.to_excel(writer, sheet_name=sheet_name, index=False)
    
    def _save_combined_page_sheets(self, writer, tables: List[pd.DataFrame], text_content: List[dict]):
        """Save content organized by page (tables + text per page)."""
        
        # Group tables by page
        tables_by_page = {}
        for table in tables:
            if hasattr(table, 'name') and 'Page_' in table.name:
                try:
                    page_num = int(table.name.split('_')[1])
                    if page_num not in tables_by_page:
                        tables_by_page[page_num] = []
                    tables_by_page[page_num].append(table)
                except (ValueError, IndexError):
                    pass
        
        # Create combined sheets for each page that has content
        all_pages = set()
        all_pages.update(tables_by_page.keys())
        all_pages.update(page_data['page'] for page_data in text_content)
        
        for page_num in sorted(all_pages):
            sheet_data = []
            
            # Add tables from this page
            if page_num in tables_by_page:
                for table in tables_by_page[page_num]:
                    sheet_data.append(['=== TABLE ==='])
                    sheet_data.extend(table.values.tolist())
                    sheet_data.append([''])  # Empty row
            
            # Add text content from this page
            page_text = next((p for p in text_content if p['page'] == page_num), None)
            if page_text:
                sheet_data.append(['=== TEXT CONTENT ==='])
                for section in page_text['sections']:
                    sheet_data.append([f"[{section['type'].upper()}]"])
                    # Split long text into multiple rows
                    content_lines = section['content'].split('\n')
                    for line in content_lines[:20]:  # Limit to first 20 lines
                        if line.strip():
                            sheet_data.append([line.strip()])
                    sheet_data.append([''])  # Empty row
            
            if sheet_data:
                df_page = pd.DataFrame(sheet_data)
                sheet_name = f"Page_{page_num}"
                df_page.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    
    def _save_comprehensive_text_analysis(self, writer, text_content: List[dict]):
        """Save comprehensive text analysis and statistics."""
        
        # Text statistics
        stats_data = []
        for page_data in text_content:
            page_num = page_data['page']
            total_words = sum(section['word_count'] for section in page_data['sections'])
            section_types = [section['type'] for section in page_data['sections']]
            type_counts = {t: section_types.count(t) for t in set(section_types)}
            
            stats_data.append({
                'Page': page_num,
                'Total_Words': total_words,
                'Total_Sections': len(page_data['sections']),
                'Has_Tables': page_data['has_tables'],
                'Section_Types': ', '.join(f"{k}:{v}" for k, v in type_counts.items())
            })
        
        if stats_data:
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='Text_Statistics', index=False)
    
    def _save_to_single_sheet(self, tables: List[pd.DataFrame]) -> None:
        """
        Save all tables to a single Excel sheet with separators.
        
        Args:
            tables (List[pd.DataFrame]): List of tables to save
        """
        combined_rows = []
        
        for i, df in enumerate(tables):
            table_name = getattr(df, 'name', f'Table_{i+1}')
            
            # Add table header
            combined_rows.append([f"=== {table_name} ==="])
            
            # Add the table data
            if not df.empty:
                # Add column headers
                combined_rows.append(list(df.columns))
                
                # Add data rows
                for _, row in df.iterrows():
                    combined_rows.append(list(row))
            
            # Add empty separator row (except after last table)
            if i < len(tables) - 1:
                combined_rows.append([""])
        
        # Create DataFrame from all rows
        if combined_rows:
            # Find the maximum number of columns needed
            max_cols = max(len(row) for row in combined_rows)
            
            # Pad all rows to have the same number of columns
            for row in combined_rows:
                while len(row) < max_cols:
                    row.append('')
            
            # Create column names
            columns = [f'Column_{i+1}' for i in range(max_cols)]
            
            final_df = pd.DataFrame(combined_rows, columns=columns)
            
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name='All_Tables', index=False)
                logger.info(f"Saved all {len(tables)} tables to single sheet: All_Tables")
    
    def _save_to_multiple_sheets(self, tables: List[pd.DataFrame]) -> None:
        """
        Save tables to separate Excel sheets.
        
        Args:
            tables (List[pd.DataFrame]): List of tables to save
        """
        with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
            for i, df in enumerate(tables):
                sheet_name = getattr(df, 'name', f'Table_{i+1}')
                # Excel sheet names have a 31 character limit
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                logger.info(f"Saved table to sheet: {sheet_name}")
    
    def convert(self, method: str = 'pdfplumber', single_sheet: bool = False) -> None:
        """
        Convert PDF to Excel using specified method.
        
        Args:
            method (str): Extraction method ('pdfplumber', 'tabula', or 'both')
            single_sheet (bool): If True, combine all tables into one sheet
        """
        all_tables = []
        
        if method in ['pdfplumber', 'both']:
            pdfplumber_tables = self.extract_tables_with_pdfplumber()
            all_tables.extend(pdfplumber_tables)
        
        # Skip tabula if Java is not available
        if method in ['tabula', 'both']:
            try:
                tabula_tables = self.extract_tables_with_tabula()
                all_tables.extend(tabula_tables)
            except Exception as e:
                logger.warning(f"Tabula extraction failed (Java may not be installed): {e}")
        
        if all_tables:
            self.save_to_excel(all_tables, single_sheet=single_sheet)
            print(f"✅ Successfully converted '{self.pdf_path}' to '{self.output_path}'")
            print(f"📊 Extracted {len(all_tables)} tables")
            if single_sheet:
                print("📋 All tables saved to a single sheet")
            else:
                print("📋 Tables saved to separate sheets")
        else:
            print(f"❌ No tables found in '{self.pdf_path}'")
            print("💡 This PDF might not contain structured tables, or the content might be in image format.")
            logger.warning("No tables were extracted from the PDF")

def main():
    """Main function to run the PDF to Excel converter."""
    parser = argparse.ArgumentParser(description='Convert PDF tables to Excel format with text extraction')
    parser.add_argument('pdf_file', help='Path to the input PDF file')
    parser.add_argument('-o', '--output', help='Output Excel file path (optional)')
    parser.add_argument('-m', '--method', choices=['pdfplumber', 'tabula', 'both'], 
                        default='both', help='Extraction method to use')
    parser.add_argument('-s', '--single-sheet', action='store_true',
                        help='Combine all tables into a single Excel sheet')
    parser.add_argument('--mixed-content', action='store_true',
                        help='Extract both tables and text content (not just tables)')
    parser.add_argument('--text-format', choices=['separate_sheets', 'combined_sheets', 'text_summary'], 
                        default='separate_sheets',
                        help='How to organize mixed content in Excel')
    
    args = parser.parse_args()
    
    try:
        converter = PDFToExcelConverter(args.pdf_file, args.output)
        
        if args.mixed_content:
            # Extract mixed content (tables + text)
            tables, text_content = converter.extract_mixed_content()
            converter.save_mixed_content_to_excel(tables, text_content, args.text_format)
            
            print(f"✅ Successfully extracted mixed content from '{args.pdf_file}'")
            print(f"📊 Found {len(tables)} tables")
            print(f"📄 Found text content from {len(text_content)} pages")
            print(f"💾 Saved to: {converter.output_path}")
        else:
            # Standard table-only extraction
            converter.convert(args.method, args.single_sheet)
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # If no command line arguments, convert the ProZ-FP.pdf file
    if len(sys.argv) == 1:
        print("Converting ProZ-FP.pdf to Excel...")
        print("Content extraction options:")
        print("1. Tables only - Separate sheets for each table (default)")
        print("2. Tables only - All tables in a single sheet")
        print("3. Mixed content - Tables + text with separate organization")
        print("4. Mixed content - Tables + text combined by page")
        print("5. Mixed content - Tables + comprehensive text analysis")
        choice = input("Enter your choice (1-5): ").strip()
        
        try:
            converter = PDFToExcelConverter("ProZ-FP.pdf")
            
            if choice in ["3", "4", "5"]:
                # Mixed content extraction
                tables, text_content = converter.extract_mixed_content()
                
                if choice == "3":
                    text_format = "separate_sheets"
                elif choice == "4":
                    text_format = "combined_sheets"
                else:  # choice == "5"
                    text_format = "text_summary"
                
                # Update output filename for mixed content
                converter.output_path = converter.output_path.with_stem(
                    converter.output_path.stem + f"_mixed_{text_format}"
                )
                
                converter.save_mixed_content_to_excel(tables, text_content, text_format)
                
                print(f"✅ Successfully extracted mixed content")
                print(f"📊 Found {len(tables)} tables")
                print(f"📄 Found text content from {len(text_content)} pages")
                print(f"💾 Saved to: {converter.output_path}")
            else:
                # Standard table extraction
                single_sheet = choice == "2"
                converter.convert(single_sheet=single_sheet)
                
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)
    else:
        main()