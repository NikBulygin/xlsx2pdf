import os
import shutil
from PyInstaller.__main__ import run

if __name__ == "__main__":
    options = [
        'main.py',
        '--onefile',
        '--name=report_generator',
        # '--icon-icon.ico'
    ]

    run(options)

    output_dir = os.path.join('dist')
    config_file = os.path.join('config.json')
    if os.path.exists(config_file):
        shutil.copy(config_file, output_dir)
    os.mkdir('./dist/out')
    os.mkdir('./dist/templates')