import os
import sys
import argparse
import time
from pathlib import Path
from convert import convert_pub_to_html, check_output_exists

def count_pub_files(input_path):
    """
    First pass: Count all .pub files in the directory tree
    """
    input_path = Path(input_path).resolve()
    count = 0
    
    print("Scanning directory structure...")
    for root, _, files in os.walk(input_path):
        pub_files = [f for f in files if f.lower().endswith('.pub')]
        count += len(pub_files)
    
    return count

def convert_directory(input_path, output_path):
    """
    Second pass: Convert all .pub files to HTML, with progress tracking
    """
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()
    
    if not input_path.exists():
        raise FileNotFoundError(f"Input directory not found: {input_path}")
    
    # First pass - count files
    total_files = count_pub_files(input_path)
    print(f"\nFound {total_files} Publisher files to convert")
    
    if total_files == 0:
        print("No files to convert")
        return 0, 0
    
    # Second pass - convert files
    files_converted = 0
    files_skipped = 0
    start_time = time.time()
    
    for root, dirs, files in os.walk(input_path):
        current_dir = Path(root)
        print(f"\nEntering directory: {current_dir}")
        
        # Filter for .pub files
        pub_files = [f for f in files if f.lower().endswith('.pub')]
        
        # Only process directory if it contains .pub files
        if pub_files:
            # Get relative path and create output directory
            rel_path = current_dir.relative_to(input_path)
            out_dir = str(output_path / rel_path)  # Convert to string for convert_pub_to_html
            os.makedirs(out_dir, exist_ok=True)
            
            # Process .pub files
            for file in pub_files:
                pub_file = current_dir / file
                base_name = Path(file).stem
                output_base = Path(out_dir) / base_name  # Use Path for check_output_exists
                
                # Check if already converted
                if check_output_exists(str(output_base)):  # Convert to string
                    files_skipped += 1
                    print(f"\nSkipping already converted file ({files_skipped} skipped): {pub_file}")
                    continue
                
                files_converted += 1
                files_processed = files_converted + files_skipped
                
                # Calculate progress and timing
                elapsed_time = time.time() - start_time
                avg_time_per_file = elapsed_time / files_converted if files_converted > 0 else 0
                est_remaining_time = avg_time_per_file * (total_files - files_processed)
                
                # Format times as HH:MM:SS
                elapsed_str = time.strftime('%H:%M:%S', time.gmtime(elapsed_time))
                remaining_str = time.strftime('%H:%M:%S', time.gmtime(est_remaining_time))
                
                print(f"\nConverting file {files_processed} of {total_files}: {pub_file}")
                print(f"Time elapsed: {elapsed_str}, estimated time remaining: {remaining_str}")
                print(f"Progress: {files_converted} converted, {files_skipped} skipped")
                
                try:
                    html_path = convert_pub_to_html(str(pub_file), out_dir)
                    print(f"Converted to: {html_path}")
                except Exception as e:
                    print(f"Error converting {pub_file}: {str(e)}", file=sys.stderr)
    
    # Final timing and stats
    total_time = time.time() - start_time
    total_time_str = time.strftime('%H:%M:%S', time.gmtime(total_time))
    print(f"\nConversion complete:")
    print(f"Total files: {total_files}")
    print(f"Converted: {files_converted}")
    print(f"Skipped: {files_skipped}")
    print(f"Total time: {total_time_str}")
    
    return files_converted, files_skipped

def main():
    parser = argparse.ArgumentParser(
        description="Recursively convert Publisher (.pub) files to HTML"
    )
    parser.add_argument("input_path", 
                      help="Root directory to search for .pub files")
    parser.add_argument("output_path",
                      help="Root directory to output HTML files to")
    args = parser.parse_args()
    
    try:
        print(f"Starting conversion from {args.input_path} to {args.output_path}")
        converted, skipped = convert_directory(args.input_path, args.output_path)
        print(f"\nConversion complete. Converted {converted} file(s), skipped {skipped} file(s).")
        
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()