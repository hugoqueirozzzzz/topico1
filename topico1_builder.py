"""
Motor de geração do "Tópico 1" a partir de uma lição completa da EBD em Foco.

Regra (validada contra os exemplos reais Sansão / Paulo, jovens / adultos):

1. Todos os slides ANTES do primeiro bumper de seção ("I. ...", "I – ...")
   são mantidos (capa, texto áureo/principal, resumo, objetivos, texto
   bíblico, introdução).
2. O bumper do Tópico I e TODO o seu conteúdo (I.1, I.2, I.3 ...) são
   mantidos até o próximo bumper de nível superior.
3. A partir do segundo bumper em diante (II., III., IV. ...), mantém-se
   APENAS o slide do bumper em si — todo o conteúdo daquele tópico é
   descartado.
4. Tudo depois do último bumper (conclusão, revisão, encerramento
   original) é descartado.
5. Um pacote fixo de slides promocionais é anexado ao final (varia
   conforme "jovens" ou "adultos").
6. Na capa (slide 0), a caixa de texto do nome do professor é substituída
   por "SLIDE DE AMOSTRA / APENAS TÓPICO I".
"""
import copy
import re
import subprocess
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu

from slide_copy import append_slide, delete_slide

BASE_DIR = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"

PROMO_TEMPLATES = {
    "jovens": ASSETS_DIR / "promo_jovens.pptx",
    "adultos": ASSETS_DIR / "promo_adultos.pptx",
}

# Numeral romano no início do texto do slide, sem ser seguido de dígito
# (o que indicaria uma subseção tipo "I.1", "II.2" etc.)
_BUMPER_RE = re.compile(
    r"^(M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3}))"
    r"\s*[\.\)\u2013\u2014-]\s*(?!\d)\S",
)

TITLE_SLOT_TEXTS = {"nome do professor"}
SAMPLE_LABEL = "SLIDE DE AMOSTRA\nAPENAS TÓPICO I"
SAMPLE_LABEL_HEIGHT = Emu(830997)


def _slide_text(slide) -> str:
    parts = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            parts.append(shape.text_frame.text.strip())
    return " ".join(parts)


def _is_bumper(slide) -> bool:
    text = _slide_text(slide)
    if not text or len(text) > 150:
        return False
    if len(slide.shapes) > 3:
        return False
    return bool(_BUMPER_RE.match(text))


def _stamp_sample_label(slide) -> bool:
    """Troca a caixa de texto 'Nome do Professor' pelo aviso de amostra,
    preservando exatamente a formatação original (fonte, tamanho, cor,
    negrito, alinhamento) — apenas o texto muda, quebrado em duas linhas."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape.text_frame.text.strip().lower() not in TITLE_SLOT_TEXTS:
            continue

        txBody = shape.text_frame._txBody
        paragraphs = txBody.findall(qn("a:p"))
        if not paragraphs:
            return False

        first_p = paragraphs[0]
        runs = first_p.findall(qn("a:r"))
        template_rPr = runs[0].find(qn("a:rPr")) if runs else None
        template_pPr = first_p.find(qn("a:pPr"))

        for p in paragraphs:
            txBody.remove(p)

        for line in SAMPLE_LABEL.split("\n"):
            new_p = OxmlElement("a:p")
            if template_pPr is not None:
                new_p.append(copy.deepcopy(template_pPr))
            new_r = OxmlElement("a:r")
            if template_rPr is not None:
                new_r.append(copy.deepcopy(template_rPr))
            new_t = OxmlElement("a:t")
            new_t.text = line
            new_r.append(new_t)
            new_p.append(new_r)
            txBody.append(new_p)

        shape.height = SAMPLE_LABEL_HEIGHT
        return True
    return False


def build_topico1(source_path: str, tipo: str, output_pptx_path: str) -> dict:
    """Gera o pptx do Tópico 1 a partir do arquivo completo da lição.

    Retorna um dicionário com informações de diagnóstico (útil para logs /
    conferência), incluindo o índice dos slides mantidos/descartados.
    """
    if tipo not in PROMO_TEMPLATES:
        raise ValueError(f"tipo inválido: {tipo!r} (use 'jovens' ou 'adultos')")

    promo_path = PROMO_TEMPLATES[tipo]
    if not promo_path.exists():
        raise FileNotFoundError(f"Template promocional não encontrado: {promo_path}")

    src = Presentation(source_path)
    n = len(src.slides)

    bumper_indices = [i for i in range(n) if _is_bumper(src.slides[i])]
    if not bumper_indices:
        raise ValueError(
            "Não encontrei nenhum slide de seção (ex: 'I. ...') no arquivo "
            "enviado. Confirme se é o arquivo da lição completa."
        )

    first_bumper = bumper_indices[0]
    other_bumpers = bumper_indices[1:]

    keep = set(range(0, first_bumper))  # tudo antes do Tópico I
    # Tópico I inteiro: do primeiro bumper até (mas sem incluir) o segundo bumper
    topic1_end = other_bumpers[0] if other_bumpers else n
    keep.update(range(first_bumper, topic1_end))
    # Demais bumpers: só o próprio slide de título
    keep.update(other_bumpers)

    kept_sorted = sorted(keep)
    dropped = sorted(set(range(n)) - keep)

    # Monta o deck de saída a partir de uma cópia do arquivo original
    # (preserva master/layouts/tema perfeitamente) removendo o que não
    # deve entrar, na ordem inversa para não invalidar os índices.
    out = Presentation(source_path)
    for i in range(n - 1, -1, -1):
        if i not in keep:
            delete_slide(out, i)

    stamped = _stamp_sample_label(out.slides[0])

    # Anexa o pacote promocional fixo
    promo_prs = Presentation(str(promo_path))
    for promo_slide in promo_prs.slides:
        append_slide(out, promo_slide)

    Path(output_pptx_path).parent.mkdir(parents=True, exist_ok=True)
    out.save(output_pptx_path)

    return {
        "total_slides_origem": n,
        "bumpers_encontrados": bumper_indices,
        "slides_mantidos": kept_sorted,
        "slides_descartados": dropped,
        "capa_carimbada": stamped,
        "slides_promocionais_anexados": len(promo_prs.slides),
        "total_slides_saida": len(out.slides),
    }


def convert_pptx_to_pdf(pptx_path: str, out_dir: str, timeout: int = 120) -> str:
    """Converte um .pptx em .pdf usando o LibreOffice headless.
    Retorna o caminho do PDF gerado."""
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        "soffice", "--headless", "--norestore",
        "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    pdf_path = out_dir / (pptx_path.stem + ".pdf")
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            f"Falha ao converter para PDF: {result.stdout}\n{result.stderr}"
        )
    return str(pdf_path)
