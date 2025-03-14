"""

Made by Yigit KOSEALI

"""

import os
import argparse
from pathlib import Path
import sys
import time
import re
import subprocess
import tempfile
from contextlib import contextmanager


def check_dependencies():
    """Check if required dependencies are installed"""
    missing_deps = []
    
    # Check for markitdown
    try:
        import markitdown
    except ImportError:
        missing_deps.append(("markitdown", "pip install markitdown"))
    
    # Check for direct file handling dependencies
    try:
        import PyPDF2
    except ImportError:
        missing_deps.append(("PyPDF2", "pip install PyPDF2"))
    
    try:
        import docx
    except ImportError:
        missing_deps.append(("python-docx", "pip install python-docx"))
    
    try:
        import pandas
        import openpyxl
    except ImportError:
        missing_deps.append(("pandas & openpyxl", "pip install pandas openpyxl"))
    
    try:
        import lxml
    except ImportError:
        missing_deps.append(("lxml", "pip install lxml"))
    
    # Display results
    if missing_deps:
        print("\nSome optional dependencies are missing:")
        for dep, install_cmd in missing_deps:
            print(f"  - {dep}: {install_cmd}")
        print("\nThese may improve conversion quality for certain file types.")
    
    return len(missing_deps) == 0


def markdown_to_text(markdown_string):
    """Convert markdown to plain text by removing formatting"""
    # Remove headers
    text = re.sub(r'#{1,6}\s+', '', markdown_string)
    
    # Remove emphasis markers (* and _)
    text = re.sub(r'\*\*?(.*?)\*\*?', r'\1', text)
    text = re.sub(r'__?(.*?)__?', r'\1', text)
    
    # Remove blockquotes
    text = re.sub(r'^\s*>\s+', '', text, flags=re.MULTILINE)
    
    # Remove inline code
    text = re.sub(r'`([^`]*)`', r'\1', text)
    
    # Remove code blocks
    text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
    
    # Remove links but keep the text
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    
    # Remove images
    text = re.sub(r'!\[.*?\]\(.*?\)', '', text)
    
    # Remove horizontal rules
    text = re.sub(r'^\s*[-*_]{3,}\s*$', '', text, flags=re.MULTILINE)
    
    # Remove HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    
    # Clean up multiple blank lines
    text = re.sub(r'\n\s*\n', '\n\n', text)
    
    return text.strip()


def extract_text_from_doc(file_path):
    """Extract text from a DOC file using multiple fallback methods"""
    
    # Method 1: Try using textutil on macOS (most reliable for macOS)
    if sys.platform == 'darwin':  # macOS
        try:
            temp_txt = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
            temp_txt.close()
            
            result = subprocess.run([
                'textutil', '-convert', 'txt', '-output', 
                temp_txt.name, str(file_path)
            ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            if result.returncode == 0:
                with open(temp_txt.name, 'r', errors='ignore') as f:
                    text = f.read()
                os.unlink(temp_txt.name)
                return text
        except Exception as e:
            print(f"  macOS textutil failed: {str(e)}")
    
    # Method 2: Try using textract if available
    try:
        import textract
        text = textract.process(file_path).decode('utf-8')
        return text
    except ImportError:
        print("  textract not installed, trying alternative methods...")
    except Exception as e:
        print(f"  textract extraction failed: {str(e)}")
    
    # Method 3: Try using antiword if available
    try:
        result = subprocess.run(['antiword', str(file_path)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            return result.stdout.decode('utf-8')
    except FileNotFoundError:
        print("  antiword not found, trying next method...")
    except Exception as e:
        print(f"  antiword extraction failed: {str(e)}")
    
    # Method 4: Try using catdoc if available
    try:
        result = subprocess.run(['catdoc', str(file_path)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            return result.stdout.decode('utf-8')
    except FileNotFoundError:
        print("  catdoc not found, trying next method...")
    except Exception as e:
        print(f"  catdoc extraction failed: {str(e)}")
    
    # Method 5: Try using libreoffice to convert to txt
    try:
        temp_dir = tempfile.mkdtemp()
        result = subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'txt:Text', 
            '--outdir', temp_dir, str(file_path)
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        if result.returncode == 0:
            txt_filename = os.path.basename(str(file_path)).replace('.doc', '.txt')
            txt_path = os.path.join(temp_dir, txt_filename)
            with open(txt_path, 'r', errors='ignore') as f:
                text = f.read()
            return text
    except FileNotFoundError:
        print("  libreoffice not found...")
    except Exception as e:
        print(f"  libreoffice conversion failed: {str(e)}")
    
    # If we get here, all methods failed
    raise Exception("Failed to extract text from DOC file. Try installing textract, antiword, catdoc, or LibreOffice.")


def extract_text_from_docx(file_path):
    """Extract text from a DOCX file directly"""
    try:
        from docx import Document
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        raise Exception(f"Failed to extract text from DOCX: {str(e)}")


def extract_text_from_pdf(file_path):
    """Extract text from a PDF file directly"""
    try:
        import PyPDF2
        text = []
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text.append(page.extract_text() or "")
        return '\n\n'.join(text)
    except Exception as e:
        raise Exception(f"Failed to extract text from PDF: {str(e)}")


@contextmanager
def open_excel_file(file_path):
    """Safely open an Excel file with proper resource management"""
    import pandas as pd
    ext = file_path.suffix.lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    excel_file = None
    
    try:
        excel_file = pd.ExcelFile(file_path, engine=engine)
        yield excel_file
    finally:
        if excel_file is not None:
            excel_file.close()


def extract_text_from_excel(file_path):
    """Extract text from Excel files directly with proper resource management"""
    try:
        import pandas as pd
        
        all_sheets = []
        
        with open_excel_file(file_path) as excel_file:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                all_sheets.append(f"Sheet: {sheet_name}\n")
                all_sheets.append(df.to_string(index=False))
                all_sheets.append("\n\n")
        
        return "".join(all_sheets)
    except Exception as e:
        raise Exception(f"Failed to extract text from Excel: {str(e)}")


def extract_text_from_csv(file_path):
    """Extract text from a CSV file directly"""
    try:
        import pandas as pd
        df = pd.read_csv(file_path, low_memory=False)
        return df.to_string(index=False)
    except Exception as e:
        # Fall back to basic CSV reading
        import csv
        rows = []
        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as file:
                csv_reader = csv.reader(file)
                for row in csv_reader:
                    rows.append('\t'.join(row))
            return '\n'.join(rows)
        except Exception as csv_e:
            raise Exception(f"Failed to extract text from CSV: {str(e)}, {str(csv_e)}")


def extract_text_from_xml(file_path):
    """Extract text from an XML file directly with robust error handling"""
    # First try using lxml's etree
    try:
        import lxml.etree as ET
        
        # Function to extract text from elements
        def get_element_text(element):
            result = []
            if element.text and element.text.strip():
                result.append(element.text.strip())
            
            for child in element:
                # Recursively get text from children
                child_text = get_element_text(child)
                if child_text:
                    result.append(child_text)
                
                # Get tail text if any
                if child.tail and child.tail.strip():
                    result.append(child.tail.strip())
                    
            # Join results if we have multiple pieces
            if isinstance(result, list):
                return " ".join(result) if result else ""
            return result
        
        # Try to parse the XML
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Add a simple representation of the XML structure
            result = f"XML Document: {file_path.name}\n\n"
            
            # Try to get structure info
            try:
                # Generate a structured representation
                def print_element(element, depth=0):
                    try:
                        text = ""
                        element_name = element.tag
                        if '}' in element_name:
                            element_name = element_name.split('}', 1)[1]  # Remove namespace
                        
                        indent = "  " * depth
                        text += f"{indent}{element_name}"
                        
                        # Add attributes if any
                        if element.attrib:
                            attrs = ", ".join(f"{k}='{v}'" for k, v in element.attrib.items())
                            text += f" ({attrs})"
                        
                        # Add element text if any
                        if element.text and element.text.strip():
                            element_text = element.text.strip()
                            if len(element_text) > 50:  # Truncate very long text
                                element_text = element_text[:47] + "..."
                            text += f": {element_text}"
                        
                        text += "\n"
                        
                        # Process children
                        for child in element:
                            text += print_element(child, depth + 1)
                        
                        return text
                    except Exception as e:
                        return f"{indent}Error processing element: {str(e)}\n"
                
                # Add structure representation
                result += "Structure:\n"
                result += print_element(root)
            except Exception as structure_e:
                result += f"Error generating structure: {str(structure_e)}\n"
            
            # Collect all text
            try:
                # Add full text content
                result += "\nContent:\n"
                
                # Alternative method to extract all text
                all_text = []
                for elem in root.iter():
                    if elem.text and elem.text.strip():
                        all_text.append(elem.text.strip())
                    if elem.tail and elem.tail.strip():
                        all_text.append(elem.tail.strip())
                
                result += "\n".join(all_text)
            except Exception as content_e:
                result += f"Error extracting text content: {str(content_e)}\n"
            
            return result
            
        except ET.XMLSyntaxError:
            # If XML parsing failed, try as HTML
            parser = ET.HTMLParser()
            tree = ET.parse(file_path, parser)
            root = tree.getroot()
            
            # Extract text from HTML-like XML
            text_parts = []
            for element in root.iter():
                if element.text and element.text.strip():
                    text_parts.append(element.text.strip())
            
            return "\n".join(text_parts)
            
    except Exception as lxml_e:
        # Fall back to Python's built-in XML parser
        try:
            import xml.etree.ElementTree as ET
            
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Simple extraction of all text content
            result = f"XML Document: {file_path.name}\n\n"
            result += "Content:\n"
            
            # Collect all text
            text_parts = []
            for elem in root.iter():
                if elem.text and elem.text.strip():
                    text_parts.append(elem.text.strip())
                
            result += "\n".join(text_parts)
            return result
            
        except Exception as builtin_e:
            # Last resort: try to read as plain text
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                # Remove angle brackets and keep content
                import re
                # Replace XML/HTML tags with newlines
                text = re.sub(r'<[^>]*>', '\n', content)
                # Clean up whitespace
                text = re.sub(r'\s+', ' ', text)
                # Split into lines and filter empty lines
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                
                return "\n".join(lines)
            except Exception as raw_e:
                raise Exception(f"All XML parsing methods failed: {str(lxml_e)}, {str(builtin_e)}, {str(raw_e)}")


def extract_text_from_youtube(url):
    """
    Extract text from a YouTube URL by downloading the transcript.
    """
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
        
        # Extract video ID from URL
        video_id = None
        patterns = [
            r'(?:youtube\.com\/watch\?v=|youtu\.be\/)([a-zA-Z0-9_-]{11})',
            r'youtube\.com\/shorts\/([a-zA-Z0-9_-]{11})'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                video_id = match.group(1)
                break
        
        if not video_id:
            raise ValueError(f"Could not extract video ID from URL: {url}")
        
        # Get transcript
        transcript_list = YouTubeTranscriptApi.get_transcript(video_id)
        
        # Merge transcript parts
        transcript_text = ""
        for item in transcript_list:
            transcript_text += item['text'] + " "
        
        return transcript_text.strip()
    
    except Exception as e:
        raise Exception(f"Failed to extract transcript from YouTube: {str(e)}\n"
                        f"Make sure to install youtube-transcript-api with:\n"
                        f"pip install youtube-transcript-api")


def is_youtube_url_file(file_path):
    """Check if a text file contains YouTube URLs"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Simple check for YouTube URLs
        youtube_patterns = ['youtube.com/watch', 'youtu.be/', 'youtube.com/shorts/']
        return any(pattern in content for pattern in youtube_patterns)
    except:
        return False


def process_youtube_urls(file_path):
    """Process a file containing YouTube URLs"""
    results = []
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            youtube_patterns = ['youtube.com/watch', 'youtu.be/', 'youtube.com/shorts/']
            if any(pattern in line for pattern in youtube_patterns):
                try:
                    transcript = extract_text_from_youtube(line)
                    results.append(f"URL: {line}\n\nTRANSCRIPT:\n{transcript}\n\n")
                except Exception as e:
                    results.append(f"URL: {line}\n\nError: Could not extract transcript - {str(e)}\n\n")
            else:
                results.append(line + "\n")
    
    except Exception as e:
        raise Exception(f"Failed to process YouTube URLs from file: {str(e)}")
    
    return "".join(results)


def process_file(md, file_path, output_dir, supported_extensions, force=False):
    """Process a single file with special handling for problematic formats"""
    file_ext = file_path.suffix.lower()
    if file_ext not in supported_extensions:
        print(f"Skipping unsupported file: {file_path}")
        return False
    
    # Create output file path with .txt extension
    output_file = output_dir / f"{file_path.stem}.txt"
    
    # Check if output file already exists
    if output_file.exists() and not force:
        response = input(f"File {output_file} already exists. Overwrite? (y/n): ")
        if response.lower() not in ['y', 'yes']:
            print(f"Skipping {file_path}")
            return False
    
    print(f"Converting: {file_path}")
    
    # Try direct extraction for formats that MarkItDown struggles with
    text_content = None
    
    # Special handling for DOC files - MarkItDown often fails with these
    if file_ext == '.doc':
        try:
            text_content = extract_text_from_doc(file_path)
            print(f"  (Used direct DOC extraction)")
        except Exception as e:
            print(f"  Direct DOC extraction failed: {str(e)}")
            # For DOC files, don't fall back to MarkItDown as it usually fails
            print(f"✗ Error converting {file_path}: Direct extraction failed and MarkItDown cannot process .doc files reliably")
            return False
    
    # For other formats, try direct extraction where we have it
    elif file_ext == '.pdf':
        try:
            text_content = extract_text_from_pdf(file_path)
            print(f"  (Used direct PDF extraction)")
        except Exception as e:
            print(f"  Direct PDF extraction failed: {str(e)}")
    
    elif file_ext == '.docx':
        # For DOCX files, try MarkItDown first, then fall back to direct extraction
        try:
            # Try using MarkItDown
            result = md.convert(str(file_path))
            text_content = markdown_to_text(result.text_content)
            print(f"  (Used MarkItDown conversion for DOCX)")
        except Exception as e:
            print(f"  MarkItDown conversion failed: {str(e)}")
            # Fall back to direct extraction
            try:
                text_content = extract_text_from_docx(file_path)
                print(f"  (Used direct DOCX extraction as fallback)")
            except Exception as e2:
                print(f"  Direct DOCX extraction also failed: {str(e2)}")
    
    elif file_ext in ['.xlsx', '.xls']:
        try:
            text_content = extract_text_from_excel(file_path)
            print(f"  (Used direct Excel extraction)")
        except Exception as e:
            print(f"  Direct Excel extraction failed: {str(e)}")
    
    elif file_ext == '.csv':
        try:
            text_content = extract_text_from_csv(file_path)
            print(f"  (Used direct CSV extraction)")
        except Exception as e:
            print(f"  Direct CSV extraction failed: {str(e)}")
    
    elif file_ext == '.xml':
        try:
            text_content = extract_text_from_xml(file_path)
            print(f"  (Used direct XML extraction)")
        except Exception as e:
            print(f"  Direct XML extraction failed: {str(e)}")
    
    # YouTube URL handling (for files that contain YouTube URLs)
    elif file_ext == '.txt' and is_youtube_url_file(file_path):
        try:
            text_content = process_youtube_urls(file_path)
            print(f"  (Processed YouTube URLs from file)")
        except Exception as e:
            print(f"  YouTube URL processing failed: {str(e)}")
    
    # For all other formats or if direct extraction failed
    if text_content is None and file_ext not in ['.doc', '.xml', '.docx']:  # Skip MarkItDown for problematic formats and DOCX (already handled)
        try:
            # Try using MarkItDown
            result = md.convert(str(file_path))
            text_content = markdown_to_text(result.text_content)
            print(f"  (Used MarkItDown conversion)")
        except Exception as e:
            print(f"  MarkItDown conversion failed: {str(e)}")
            return False
    
    # If we still don't have text content, fail
    if text_content is None:
        print(f"✗ Error converting {file_path}: All extraction methods failed")
        return False
    
    # Write the text content
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(text_content)
        print(f"✓ Successfully converted to: {output_file}")
        return True
    except Exception as e:
        print(f"✗ Error writing output file: {str(e)}")
        return False


def process_directory(md, directory, output_dir, supported_extensions, recursive=False, force=False):
    """Process all files in a directory"""
    print(f"Processing directory: {directory}")
    
    total_files = 0
    successful = 0
    failed = 0
    
    for item in directory.iterdir():
        if item.is_file():
            result = process_file(md, item, output_dir, supported_extensions, force)
            total_files += 1
            if result:
                successful += 1
            else:
                failed += 1
        elif item.is_dir() and recursive:
            # Create corresponding output subdirectory
            relative_path = item.relative_to(directory)
            new_output_dir = output_dir / relative_path
            new_output_dir.mkdir(exist_ok=True)
            
            # Process the subdirectory
            sub_total, sub_successful, sub_failed = process_directory(
                md, item, new_output_dir, supported_extensions, recursive, force
            )
            total_files += sub_total
            successful += sub_successful
            failed += sub_failed
    
    return total_files, successful, failed


def main():
    parser = argparse.ArgumentParser(description="Batch convert files to plain text using Microsoft's MarkItDown")
    parser.add_argument("input", help="Input file or directory path")
    parser.add_argument("-o", "--output", help="Output directory (default: ./text_output)")
    parser.add_argument("-r", "--recursive", action="store_true", help="Process directories recursively")
    parser.add_argument("-d", "--docintel", action="store_true", help="Use Azure Document Intelligence")
    parser.add_argument("-e", "--endpoint", help="Azure Document Intelligence endpoint")
    parser.add_argument("--llm", action="store_true", help="Use LLM for image descriptions")
    parser.add_argument("--model", default="gpt-4o", help="LLM model to use (default: gpt-4o)")
    parser.add_argument("--api-key", help="OpenAI API key (otherwise uses OPENAI_API_KEY environment variable)")
    parser.add_argument("--force", action="store_true", help="Overwrite existing text files without confirmation")
    parser.add_argument("--rtf", action="store_true", help="Include RTF files (experimental support)")
    args = parser.parse_args()

    # Check dependencies
    check_dependencies()
    
    # Import the markitdown package - make sure the script has a different name than markitdown.py!
    try:
        from markitdown import MarkItDown
    except ImportError:
        print("Error: Could not import MarkItDown.")
        print("Make sure you've installed it with 'pip install markitdown'")
        print("Also, make sure there's no file named markitdown.py in your current directory.")
        sys.exit(1)

    # Set up output directory
    output_dir = args.output if args.output else "./text_output"
    os.makedirs(output_dir, exist_ok=True)

    # Initialize MarkItDown with optional parameters
    md_params = {}
    
    if args.docintel:
        if not args.endpoint:
            print("Error: Document Intelligence endpoint (-e) is required when using -d flag")
            sys.exit(1)
        md_params["docintel_endpoint"] = args.endpoint
    
    if args.llm:
        try:
            from openai import OpenAI
            if args.api_key:
                client = OpenAI(api_key=args.api_key)
            else:
                client = OpenAI()  # Uses OPENAI_API_KEY environment variable
            md_params["llm_client"] = client
            md_params["llm_model"] = args.model
        except ImportError:
            print("Error: OpenAI package is required for LLM integration")
            print("Install it using: pip install openai")
            sys.exit(1)
    
    # Create MarkItDown instance
    try:
        md = MarkItDown(**md_params)
    except Exception as e:
        print(f"Error initializing MarkItDown: {str(e)}")
        sys.exit(1)
    
    # Supported file extensions
    supported_extensions = [
        # Document formats
        '.pdf', '.pptx', '.ppt', '.docx', '.doc', '.xlsx', '.xls',
        # Image formats
        '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
        # Audio formats 
        '.mp3', '.wav', '.m4a', '.flac',
        # Web and text formats
        '.html', '.htm', '.csv', '.json', '.xml', '.txt',
        # Archives
        '.zip'
    ]
    
    # Add RTF if requested (experimental)
    if args.rtf:
        supported_extensions.append('.rtf')
    
    # Process files
    total_files = 0
    successful_conversions = 0
    failed_conversions = 0
    
    start_time = time.time()
    input_path = Path(args.input)
    
    if input_path.is_file():
        result = process_file(md, input_path, Path(output_dir), supported_extensions, args.force)
        total_files = 1
        if result:
            successful_conversions = 1
        else:
            failed_conversions = 1
    elif input_path.is_dir():
        total_files, successful_conversions, failed_conversions = process_directory(
            md, input_path, Path(output_dir), supported_extensions, args.recursive, args.force
        )
    else:
        print(f"Error: Input path '{args.input}' does not exist")
        sys.exit(1)
    
    # Print summary
    elapsed_time = time.time() - start_time
    print("\n" + "="*50)
    print("Conversion Summary:")
    print(f"Total files processed: {total_files}")
    print(f"Successful conversions: {successful_conversions}")
    print(f"Failed conversions: {failed_conversions}")
    print(f"Time elapsed: {elapsed_time:.2f} seconds")
    print("="*50)


if __name__ == "__main__":
    main()