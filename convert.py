import os
import sys
import win32com.client
import argparse
from pathlib import Path
import time
import glob
import shutil
import psutil
import subprocess

def kill_publisher_processes():
    """Kill any running Publisher processes"""
    killed = False
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if 'MSPUB.EXE' in proc.info['name'].upper():
                print(f"Killing Publisher process {proc.info['pid']}...")
                proc.kill()
                killed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    
    if killed:
        # Wait for processes to fully terminate
        time.sleep(2)
        # Double check they're gone
        for proc in psutil.process_iter(['name']):
            try:
                if 'MSPUB.EXE' in proc.info['name'].upper():
                    print("Warning: Publisher process still running")
                    return False
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    return True

def validate_html(html_path):
    """
    Validate that the HTML file and its supporting files were created
    Args:
        html_path: Path to the HTML file (without extension)
    Raises:
        RuntimeError: If the HTML files are missing or incomplete
    """
    # Check for main HTML file
    html_file = html_path + '.htm'
    if not os.path.exists(html_file):
        raise RuntimeError(f"HTML file was not created at {html_file}")
    
    # Check for supporting files directory
    files_dir = html_path + '_files'
    if not os.path.exists(files_dir) or not os.path.isdir(files_dir):
        raise RuntimeError(f"Supporting files directory not found at {files_dir}")
    
    # Check if the directory contains any files
    if not os.listdir(files_dir):
        raise RuntimeError(f"Supporting files directory is empty: {files_dir}")
    
    print(f"HTML validation successful:")
    print(f"- Main file: {html_file}")
    print(f"- Supporting files: {files_dir}")

def check_output_files(base_path):
    """
    Check for any files created with the base filename
    Args:
        base_path: Base path without extension
    Returns:
        List of files created
    """
    # Remove extension if present
    base_path = str(Path(base_path).with_suffix(''))
    # Check for any files starting with the base name
    pattern = base_path + '*'
    return glob.glob(pattern)

def check_output_exists(output_base):
    """
    Check if both the HTML file and supporting files directory already exist
    Args:
        output_base: Base path without extension
    Returns:
        bool: True if both html file and supporting directory exist
    """
    html_file = output_base + '.htm'
    files_dir = output_base + '_files'
    return os.path.exists(html_file) and os.path.isdir(files_dir) and os.listdir(files_dir)

def convert_pub_to_html(pub_path, output_dir="output", format_constant=7, max_retries=3):
    """
    Convert a Publisher file to HTML using COM automation
    Args:
        pub_path: Path to the .pub file
        output_dir: Directory to save the HTML in (default: output)
        format_constant: Publisher SaveAs format constant (default: 7 for HTML)
        max_retries: Maximum number of retry attempts after killing Publisher
    Returns:
        Path to the generated HTML file
    """
    # Convert to absolute paths
    pub_path = str(Path(pub_path).resolve())
    output_dir = str(Path(output_dir).resolve())
    
    print(f"Input file path: {pub_path}")
    print(f"Output directory: {output_dir}")
    
    if not os.path.exists(pub_path):
        raise FileNotFoundError(f"Publisher file not found: {pub_path}")
    
    if not pub_path.lower().endswith('.pub'):
        raise ValueError("File must have .pub extension")

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Create output path with same base name
    base_name = os.path.splitext(os.path.basename(pub_path))[0]
    output_base = os.path.join(output_dir, base_name)
    print(f"Output base path: {output_base}")
    
    # Check if file is already converted
    if check_output_exists(output_base):
        print(f"File already converted, skipping: {output_base}.htm")
        return output_base + '.htm'
    
    last_error = None
    for attempt in range(max_retries):
        publisher = None
        doc = None
        try:
            if attempt > 0:
                print(f"\nRetry attempt {attempt + 1} of {max_retries}...")
                # Kill any hanging Publisher processes before retry
                if not kill_publisher_processes():
                    print("Warning: Could not kill all Publisher processes")
                    time.sleep(5)  # Give more time for processes to die
            
            # Initialize Publisher application with security settings
            print("Initializing Publisher...")
            publisher = win32com.client.gencache.EnsureDispatch("Publisher.Application")
            
            # Set security settings to disable macros and active content
            publisher.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            
            # Open the document
            print("Opening document...")
            abs_pub_path = str(Path(pub_path).absolute())
            doc = publisher.Open(abs_pub_path, False, False)  # ReadOnly=False, OpenAndRepair=False
            time.sleep(1)  # Give document time to open
            
            # Save as HTML
            print(f"Saving as HTML...")
            abs_output_base = str(Path(output_base).absolute())
            doc.SaveAs(abs_output_base, format_constant)
            
            # Check what files were created
            print("Checking created files...")
            created_files = check_output_files(output_base)
            if created_files:
                print("Created files:")
                for f in created_files:
                    print(f"- {f}")
            else:
                print("No files were created")
            
            # Validate the HTML output
            print("Validating HTML...")
            validate_html(output_base)
            
            # If we get here, conversion was successful
            return output_base + '.htm'
        
        except Exception as e:
            last_error = e
            print(f"Error on attempt {attempt + 1}: {str(e)}", file=sys.stderr)
            if hasattr(e, 'excepinfo'):
                print(f"Exception info: {e.excepinfo}", file=sys.stderr)
            
            # Clean up failed attempt
            try:
                if doc:
                    doc.Close()
                if publisher:
                    publisher.Quit()
            except:
                pass
            
            # Only continue retrying for certain errors
            if hasattr(e, 'excepinfo'):
                error_code = e.excepinfo[5] if len(e.excepinfo) > 5 else None
                # -2147221457 is the "modal dialog" error
                if error_code != -2147221457:
                    break  # Don't retry for other error types
            else:
                break  # Don't retry for non-COM errors
        
        finally:
            # Clean up in finally block
            print("Cleaning up...")
            try:
                if doc:
                    doc.Close()
                if publisher:
                    publisher.Quit()
            except Exception as e:
                print(f"Warning: Cleanup failed: {str(e)}", file=sys.stderr)
    
    # If we get here, all retries failed
    raise RuntimeError(f"Failed to convert file after {max_retries} attempts: {str(last_error)}")

def main():
    parser = argparse.ArgumentParser(description="Convert Publisher (.pub) file to HTML")
    parser.add_argument("pub_file", help="Path to the Publisher file to convert")
    parser.add_argument("--output-dir", default="output", 
                      help="Directory to save HTML in (default: output)")
    parser.add_argument("--format-constant", type=int, default=7, 
                      help="Publisher SaveAs format constant to use (default: 7 for HTML)")
    args = parser.parse_args()
    
    try:
        html_path = convert_pub_to_html(args.pub_file, args.output_dir, args.format_constant)
        print(f"Successfully converted to: {html_path}")
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()