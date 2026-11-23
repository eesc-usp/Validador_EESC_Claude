import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

EXEMPLOS = RAIZ / "exemplos"
TEMPLATE = RAIZ / "assets" / "modelos" / "TEMPLATE_Dissertacao_Tese_ABNT_EESC.docx"
