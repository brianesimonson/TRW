"""Scoped file handling, replacing clipboard waits and shared temporary paths."""
import os
from pathlib import Path
import tempfile
import time


def wait_for_file(path,timeout=2.0,interval=0.2):
    deadline=time.monotonic()+timeout
    while not Path(path).is_file():
        if time.monotonic()>=deadline:
            raise FileNotFoundError(f'File did not appear within {timeout}s: {path}')
        time.sleep(min(interval,max(0,deadline-time.monotonic())))
    return Path(path)


def save_document(document,destination):
    destination=Path(destination)
    destination.parent.mkdir(parents=True,exist_ok=True)
    if destination.exists():
        raise FileExistsError(f'Refusing to overwrite {destination}')
    with tempfile.NamedTemporaryFile(dir=destination.parent,suffix='.docx',delete=False) as handle:
        temporary=Path(handle.name)
    try:
        document.save(temporary)
        os.chmod(temporary,0o644)
        os.link(temporary,destination)
    finally:
        temporary.unlink(missing_ok=True)


def save_bytes(data,destination):
    destination=Path(destination)
    destination.parent.mkdir(parents=True,exist_ok=True)
    if destination.exists():
        raise FileExistsError(f'Refusing to overwrite {destination}')
    with tempfile.NamedTemporaryFile(dir=destination.parent,delete=False) as handle:
        temporary=Path(handle.name)
        handle.write(data)
    try:
        os.chmod(temporary,0o644);os.link(temporary,destination)
    finally:
        temporary.unlink(missing_ok=True)
