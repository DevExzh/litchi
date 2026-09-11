"""Continue frozen captures serially after both candidate builds succeed."""
import subprocess,sys
from run import HERE
commands = [
    ['capture.py', 'after', 'preflight'],
    ['capture.py', 'after', 'r1'],
    ['capture.py', 'after', 'r2'],
    ['capture.py', 'before', 'r2'],
    ['capture.py', 'after', 'allocator-r1'],
    ['capture.py', 'after', 'allocator-r2'],
    ['capture.py', 'after', 'profile'],
    ['check.py', 'fmt'],
    ['check.py', 'clippy'],
    ['check.py', 'rustdoc'],
    ['check.py', 'boundaries'],
    ['check.py', 'claims'],
]
for script, *args in commands:
    subprocess.run([sys.executable, '-B', str(HERE/script), *args], check=True)
