"""Export via the shared artifact workbook engine, keeping it out of UI details."""
import json
import shutil
import subprocess
import tempfile
from pathlib import Path
from classifier import ROOT


def excel_bytes(frame) -> bytes:
    node = shutil.which('node')
    if not node:
        raise RuntimeError('Ekspor Excel belum tersedia di server. Hubungi pengelola aplikasi.')
    folder = ROOT / '.cache' / 'exports'
    folder.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(dir=folder) as temp:
        source = Path(temp)/'data.json'
        dest = Path(temp)/'hasil.xlsx'
        frame.to_json(source,orient='records',force_ascii=False,date_format='iso')
        proc=subprocess.run([node,str(ROOT/'export_xlsx.mjs'),str(source),str(dest)],
                            capture_output=True,text=True,timeout=180,cwd=ROOT)
        # Some native render bindings exit nonzero after a successful artifact save.
        if not dest.exists():
            raise RuntimeError('Ekspor Excel gagal. Hasil kategorisasi tetap tersimpan; coba kembali.')
        return dest.read_bytes()
