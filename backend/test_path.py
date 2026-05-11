from pathlib import Path
import os

print(f"CWD: {os.getcwd()}")
print(f"__file__: {__file__}")
base_dir = Path(__file__).parent.parent
print(f"BASE_DIR: {base_dir.absolute()}")
db_file = base_dir / "kb.db"
print(f"DB_FILE: {db_file.absolute()}")
print(f"DB exists: {db_file.exists()}")
if db_file.exists():
    print(f"DB size: {db_file.stat().st_size}")
