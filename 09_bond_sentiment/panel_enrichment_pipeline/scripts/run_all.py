"""
Запуск парсеров по очереди.
Закомментируй ненужные строки в SCRIPTS.
"""

import subprocess
import sys
from datetime import datetime

SCRIPTS = [
    # "expertra_parser.py",
    # "expertra_issuer_parser.py",
    "acra_parser.py",
    "acra_issuer_parser.py",
]

for script in SCRIPTS:
    print(f"\n{'='*60}")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Запуск: {script}")
    print(f"{'='*60}\n")

    result = subprocess.run([sys.executable, script])

    if result.returncode != 0:
        print(f"\n⚠️  {script} завершился с ошибкой (код {result.returncode})")
        print("Продолжаю следующий...")

print(f"\n{'='*60}")
print(f"[{datetime.now().strftime('%H:%M:%S')}] Все скрипты завершены")
print(f"{'='*60}")
