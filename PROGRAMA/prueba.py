import psutil
import platform
from datetime import datetime

def get_size(bytes, suffix="B"):
    factor = 1024
    for unit in ["","K","M","G","T","P"]: 
        if bytes < factor:
            return f"{bytes:.2f}{unit}{suffix}"
        bytes /= factor

uname = platform.uname()
# uname.node = Nombre del equipo
print(f"system:{psutil.cpu_count(logical=True)}")