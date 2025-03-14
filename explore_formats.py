import os
import sys
import win32com.client
import argparse
from pathlib import Path
import time
import glob
import shutil
from tabulate import tabulate
from PyPDF2 import PdfReader
import mimetypes
import traceback

def detect_file_type(filepath):
    """Attempt to detect the type of a file using multiple methods"""
    try:
        # Check if it's a directory first
        if os.path.isdir(filepath):
            return "Directory"
        
        # Try to detect by extension first
        ext = Path(filepath).suffix.lower()
        mime_type = mimetypes.guess_type(filepath)[0]
        
        # Special handling for known types
        if ext == '.pdf':
            try:
                reader = PdfReader(filepath)
                num_pages = len(reader.pages)
                return f"Valid PDF ({num_pages} pages)"
            except:
                return "Invalid PDF"
        elif ext == '.htm' or ext == '.html':
            return "HTML document"
        elif ext == '.txt':
            return "Text document"
        elif ext == '.rtf':
            return "Rich Text document"
        elif ext == '.pub':
            return "Publisher document"
        elif mime_type:
            return mime_type
        
        # If we can't detect the type, return the extension
        return f"Unknown ({ext})"
    except Exception as e:
        return f"Error detecting type: {str(e)}"

def clean_output_dir(output_dir):
    """Clean all files from the output directory except .gitkeep"""
    for item in os.listdir(output_dir):
        if item != '.gitkeep':
            path = os.path.join(output_dir, item)
            try:
                if os.path.isfile(path):
                    os.unlink(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)
            except Exception as e:
                print(f'Failed to delete {path}: {e}')

def test_format_constant(pub_path, output_dir, format_constant):
    """Test a single format constant and return information about what it produced"""
    publisher = None
    doc = None
    result = {
        'constant': format_constant,
        'files': [],
        'error': None
    }
    
    try:
        # Clean output directory
        clean_output_dir(output_dir)
        
        # Initialize Publisher
        publisher = win32com.client.gencache.EnsureDispatch("Publisher.Application")
        
        # Open document with absolute path
        abs_pub_path = str(Path(pub_path).absolute())
        doc = publisher.Open(abs_pub_path)
        time.sleep(0.5)  # Give document time to open
        
        # Try to save with the format constant using absolute path
        base_name = Path(pub_path).stem
        output_path = str(Path(output_dir).absolute() / base_name)
        print(f"\nDebug - SaveAs parameters:")
        print(f"  Path: {output_path}")
        print(f"  Format constant: {format_constant} (type: {type(format_constant)})")
        doc.SaveAs(output_path, format_constant)
        
        # Check what files were created
        time.sleep(0.5)  # Give filesystem time to settle
        created_files = []
        for item in os.listdir(output_dir):
            if item != '.gitkeep':
                filepath = os.path.join(output_dir, item)
                filetype = detect_file_type(filepath)
                created_files.append({
                    'name': item,
                    'type': filetype
                })
        
        result['files'] = created_files
        
    except Exception as e:
        if hasattr(e, 'excepinfo'):
            result['error'] = f"COM Error: {e.excepinfo}"
        else:
            result['error'] = str(e)
    
    finally:
        try:
            if doc:
                doc.Close()
            if publisher:
                publisher.Quit()
        except:
            pass
    
    return result

def explore_format_constants(pub_path, output_dir, start=1, end=100):
    """
    Test a range of format constants and generate a report
    """
    # Ensure start and end are integers
    start = int(start)
    end = int(end)
    print(f"Start type: {type(start)}, value: {start}")
    print(f"End type: {type(end)}, value: {end}")
    
    results = []
    
    print(f"Testing format constants {start} through {end}...")
    print(f"Input file: {pub_path}")
    print(f"Output directory: {output_dir}")
    print()
    
    for i in range(start, end + 1):
        sys.stdout.write(f"\rTesting format constant {i}...")
        sys.stdout.flush()
        
        result = test_format_constant(pub_path, output_dir, i)
        if result['files'] or result['error'] != "COM Error: None":  # Only keep interesting results
            results.append(result)
    
    print("\n\nResults:")
    
    # Prepare table data
    table_data = []
    for r in results:
        if r['files']:
            files_str = "\n".join([f"{f['name']}: {f['type']}" for f in r['files']])
        else:
            files_str = "No files created"
        
        table_data.append([
            r['constant'],
            files_str if not r['error'] else f"Error: {r['error']}"
        ])
    
    # Print results table
    print(tabulate(table_data, headers=['Constant', 'Output'], tablefmt='grid'))
    
    # Clean up at the end
    clean_output_dir(output_dir)

def main():
    parser = argparse.ArgumentParser(description="Explore Publisher SaveAs format constants")
    parser.add_argument("pub_file", help="Path to the Publisher file to test with")
    parser.add_argument("--output-dir", default="output", 
                      help="Directory to save output files in (default: output)")
    parser.add_argument("--start", type=int, default=1,
                      help="Starting format constant (default: 1)")
    parser.add_argument("--end", type=int, default=100,
                      help="Ending format constant (default: 100)")
    args = parser.parse_args()
    
    try:
        explore_format_constants(args.pub_file, args.output_dir, args.start, args.end)
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()