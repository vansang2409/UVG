"""Use the packaged DOCX renderer with Microsoft Word as Windows PDF engine."""
import importlib.util
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
SKILL = Path(r'C:\Users\thuong ngo\.codex\plugins\cache\openai-primary-runtime\documents\26.1026.10807\skills\documents')
POPLER = Path(r'C:\Program Files\WindowsApps\OpenAI.CodexPrimaryRuntime.v26-1026-487-0_26.1026.487.0_x64__3k8sg7r9htsxt\dependencies\native\poppler\Library\bin')
os.environ['PATH'] = str(POPLER) + os.pathsep + os.environ.get('PATH', '')
spec = importlib.util.spec_from_file_location('packaged_render_docx', SKILL / 'render_docx.py')
renderer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(renderer)

def convert_with_word(doc_path, user_profile, convert_tmp_dir, stem, verbose=False):
    output = str(Path(convert_tmp_dir) / (stem + '.pdf'))
    result = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ROOT / 'export_word_pdf.ps1'), '-InputDocx', str(Path(doc_path).resolve()), '-OutputPdf', output], capture_output=True, text=True, timeout=60)
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    return output, result.stdout

renderer.convert_to_pdf = convert_with_word
renderer.main()
