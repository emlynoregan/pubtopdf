import winreg
import win32com.client
import os
from pathlib import Path

def find_publisher_typelib():
    try:
        # First try HKCR
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'CLSID') as clsid_key:
            # Look for Publisher.Application
            for i in range(winreg.QueryInfoKey(clsid_key)[0]):
                try:
                    clsid = winreg.EnumKey(clsid_key, i)
                    with winreg.OpenKey(clsid_key, f"{clsid}\\ProgID") as prog_key:
                        if winreg.QueryValue(prog_key, None) == "Publisher.Application":
                            print(f"Found Publisher CLSID: {clsid}")
                            # Try to get TypeLib
                            with winreg.OpenKey(clsid_key, f"{clsid}\\TypeLib") as typelib_key:
                                typelib_id = winreg.QueryValue(typelib_key, None)
                                print(f"TypeLib ID: {typelib_id}")
                                # Try to get path
                                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"TypeLib\\{typelib_id}\\1.0\\0\\win32") as lib_key:
                                    path = winreg.QueryValue(lib_key, None)
                                    print(f"TypeLib Path: {path}")
                                    return path
                except:
                    continue
    except Exception as e:
        print(f"Error searching registry: {e}")
        return None

def inspect_type_library(path):
    try:
        publisher = win32com.client.Dispatch('Publisher.Application')
        print("\nTrying to enumerate Publisher methods:")
        for attr in dir(publisher):
            if 'SaveAs' in attr:
                print(f"Found method: {attr}")
                # Try to get method signature
                try:
                    method = getattr(publisher, attr)
                    print(f"Method info: {method.__doc__}")
                except:
                    pass
    finally:
        try:
            publisher.Quit()
        except:
            pass

if __name__ == "__main__":
    print("Searching for Publisher type library...")
    typelib_path = find_publisher_typelib()
    if typelib_path:
        inspect_type_library(typelib_path)