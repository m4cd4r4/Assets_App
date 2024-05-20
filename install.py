# install.py
import os
import subprocess
import sys
import venv

# Define the directory for the virtual environment
venv_dir = os.path.join(os.getcwd(), 'venv')

def create_virtual_env():
    """Create a virtual environment."""
    if not os.path.exists(venv_dir):
        venv.create(venv_dir, with_pip=True)
        print("Virtual environment created.")
    else:
        print("Virtual environment already exists.")

def install_dependencies():
    """Install the dependencies from requirements.txt."""
    python_executable = os.path.join(venv_dir, 'Scripts' if sys.platform == 'win32' else 'bin', 'python')
    pip_executable = os.path.join(venv_dir, 'Scripts' if sys.platform == 'win32' else 'bin', 'pip')
    subprocess.check_call([python_executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    subprocess.check_call([pip_executable, 'install', '-r', 'requirements.txt'])
    print("Dependencies installed.")

def create_shortcut():
    """Create a shortcut to run the application."""
    script_path = os.path.join(os.getcwd(), 'your_script.py')
    if sys.platform == "win32":
        shortcut_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'YourApp.lnk')
        with open(shortcut_path, 'w') as shortcut:
            shortcut.write(f"""[InternetShortcut]
URL=file:///{script_path}
IconIndex=0
IconFile={os.path.join(os.getcwd(), 'icon.ico')}
""")
        print("Shortcut created on Desktop.")
    elif sys.platform == "darwin" or sys.platform.startswith('linux'):
        desktop_entry = f"""
[Desktop Entry]
Name=YourApp
Exec={python_executable} {script_path}
Type=Application
Icon={os.path.join(os.getcwd(), 'icon.png')}
Terminal=false
"""
        shortcut_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'YourApp.desktop')
        with open(shortcut_path, 'w') as shortcut:
            shortcut.write(desktop_entry)
        os.chmod(shortcut_path, 0o755)
        print("Shortcut created on Desktop.")

def main():
    create_virtual_env()
    install_dependencies()
    create_shortcut()
    print("Installation complete. Please use the shortcut on your Desktop to run the application.")

if __name__ == "__main__":
    main()
