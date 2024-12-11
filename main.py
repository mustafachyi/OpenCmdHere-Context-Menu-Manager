import winreg as reg
import os
import ctypes
import sys

# Constants for registry key paths and command details
KEY_PATH = r'Directory\Background\shell\OpenCmdHere'
COMMAND_KEY_PATH = r'Directory\Background\shell\OpenCmdHere\command'
ICON_PATH = r'C:\Windows\System32\cmd.exe'
COMMAND = r'cmd.exe /k cd %V'

def is_admin():
    """
    Check if the script is running with administrative privileges.
    Returns True if running as admin, False otherwise.
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_open_cmd_here():
    """
    Check if the 'Open Command Window Here' context menu entry exists.
    Returns True if the entry exists, False otherwise.
    """
    try:
        reg.OpenKey(reg.HKEY_CLASSES_ROOT, KEY_PATH)
        return True
    except FileNotFoundError:
        return False

def create_registry_key(path, name, value):
    """
    Create a registry key with the specified path, name, and value.
    """
    key = reg.CreateKey(reg.HKEY_CLASSES_ROOT, path)
    reg.SetValue(key, '', reg.REG_SZ, name)
    reg.SetValueEx(key, 'Icon', 0, reg.REG_SZ, ICON_PATH)
    command_key = reg.CreateKey(reg.HKEY_CLASSES_ROOT, COMMAND_KEY_PATH)
    reg.SetValue(command_key, '', reg.REG_SZ, COMMAND)

def delete_registry_key(path):
    """
    Delete the registry key at the specified path.
    """
    reg.DeleteKey(reg.HKEY_CLASSES_ROOT, path + r'\command')
    reg.DeleteKey(reg.HKEY_CLASSES_ROOT, path)

def add_open_cmd_here():
    """
    Add the 'Open Command Window Here' context menu entry.
    """
    try:
        create_registry_key(KEY_PATH, 'Open Command Window Here', COMMAND)
        print("\n\033[32m[+] Context menu entry 'Open Command Window Here' added successfully.\033[0m")
    except Exception as e:
        print(f"\n\033[31m[-] Failed to add context menu entry: {e}\033[0m")

def remove_open_cmd_here():
    """
    Remove the 'Open Command Window Here' context menu entry.
    """
    try:
        delete_registry_key(KEY_PATH)
        print("\n\033[32m[+] Context menu entry 'Open Command Window Here' removed successfully.\033[0m")
    except Exception as e:
        print(f"\n\033[31m[-] Failed to remove context menu entry: {e}\033[0m")

def main():
    """
    Main function to manage the 'Open Command Window Here' context menu entry.
    """
    if not is_admin():
        print("\033[31m[-] This script requires administrative privileges to run.\033[0m")
        print("\033[33mPlease run the script as an administrator.\033[0m")
        input("\033[36mPress Enter to exit...\033[0m")
        sys.exit(1)

    # ASCII Art for "Context menu manager"
    print("\033[36m")
    print(r"""
       _____            _            _                                                                              
      / ____|          | |          | |                                                                             
     | |     ___  _ __ | |_ _____  _| |_   _ __ ___   ___ _ __  _   _   _ __ ___   __ _ _ __   __ _  __ _  ___ _ __ 
     | |    / _ \| '_ \| __/ _ \ \/ / __| | '_ ` _ \ / _ \ '_ \| | | | | '_ ` _ \ / _` | '_ \ / _` |/ _` |/ _ \ '__|
     | |___| (_) | | | | ||  __/>  <| |_  | | | | | |  __/ | | | |_| | | | | | | | (_| | | | | (_| | (_| |  __/ |   
      \_____\___/|_| |_|\__\___/_/\_\\__| |_| |_| |_|\___|_| |_|\__,_| |_| |_| |_|\__,_|_| |_|\__,_|\__, |\___|_|   
                                                                                                     __/ |          
                                                                                                    |___/           
    """)
    print("\033[0m")
    
    if check_open_cmd_here():
        print("\033[33mThe 'Open Command Window Here' context menu entry already exists.\033[0m\n")
        print("\033[34mOptions:\033[0m")
        print("  \033[31m1.\033[0m Remove the 'Open Command Window Here' context menu entry")
        print("  \033[32m2.\033[0m Exit without making changes\n")
        choice = input("\033[36mPlease enter your choice: \033[0m").strip()
        if choice == '1':
            remove_open_cmd_here()
        elif choice == '2':
            print("\n\033[33mExiting without making changes.\033[0m")
        else:
            print("\n\033[31m[-] Invalid choice. Exiting without making changes.\033[0m")
    else:
        print("\033[33mThe 'Open Command Window Here' context menu entry does not exist.\033[0m\n")
        print("\033[34mOptions:\033[0m")
        print("  \033[32m1.\033[0m Add the 'Open Command Window Here' context menu entry")
        print("  \033[32m2.\033[0m Exit without making changes\n")
        choice = input("\033[36mPlease enter your choice: \033[0m").strip()
        if choice == '1':
            add_open_cmd_here()
        elif choice == '2':
            print("\n\033[33mExiting without making changes.\033[0m")
        else:
            print("\n\033[31m[-] Invalid choice. Exiting without making changes.\033[0m")

if __name__ == "__main__":
    main()