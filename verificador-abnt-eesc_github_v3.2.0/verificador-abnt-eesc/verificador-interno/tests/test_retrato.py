"""Nada muda sem querer: os resultados de todos os documentos públicos batem com o retrato guardado."""
import json

import pytest
import retrato


@pytest.mark.skipif(not retrato.ESPERADO.exists(), reason="retrato ainda não gerado")
def test_resultados_iguais_ao_retrato():
    esperado = json.loads(retrato.ESPERADO.read_text("utf-8"))
    atual = json.loads(json.dumps(retrato.calcular(), ensure_ascii=False, default=str))
    difs = retrato.diferencas(esperado, atual)
    assert not difs, (
        f"{len(difs)} resultado(s) mudaram:\n  " + "\n  ".join(difs[:30]) +
        "\n\nSe a mudança foi intencional (regra nova ou corrigida), atualize o retrato e confira a diferença:\n"
        "    python tests/retrato.py --atualizar")
