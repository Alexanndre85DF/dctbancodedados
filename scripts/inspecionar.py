# -*- coding: utf-8 -*-
import sqlite3
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
root = Path(__file__).resolve().parent.parent
conn = sqlite3.connect(root / "saida" / "dct_tocantins.sqlite")
conn.row_factory = sqlite3.Row


def show(title, sql):
    print("\n===", title, "===")
    rows = conn.execute(sql).fetchall()
    for r in rows:
        print(dict(r))
    print("n=", len(rows))


show(
    "CI sem enunciado",
    """
    SELECT h.codigo, h.enunciado, al.ano, al.bimestre, al.unidade_tematica, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='CI' AND length(coalesce(h.enunciado,''))<20
    LIMIT 20
    """,
)
show(
    "CI bom",
    """
    SELECT h.codigo, substr(h.enunciado,1,100) enc, al.ano, al.bimestre,
           al.unidade_tematica, substr(al.objetos_conhecimento,1,60) obj, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='CI' AND length(coalesce(h.enunciado,''))>=40
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "CI codigo estranho",
    """
    SELECT h.codigo, count(*) n
    FROM habilidade h
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='CI' AND instr(h.codigo,'CI')=0
    GROUP BY h.codigo
    """,
)
show(
    "LP amostra",
    """
    SELECT h.codigo, substr(h.enunciado,1,100) enc, al.ano, al.bimestre,
           al.unidade_tematica, al.eixo, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='LP'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "GE amostra",
    """
    SELECT h.codigo, substr(h.enunciado,1,100) enc, al.ano, al.bimestre,
           al.unidade_tematica, substr(al.sugestao_pedagogica,1,70) sug, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='GE'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "MA amostra",
    """
    SELECT h.codigo, substr(h.enunciado,1,90) enc, al.ano, al.bimestre,
           al.unidade_tematica, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='MA'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "codigo vs componente",
    """
    SELECT c.codigo as comp, substr(h.codigo,5,2) as sigla, count(*) n
    FROM habilidade h
    JOIN componente c ON c.id=h.componente_id
    GROUP BY c.codigo, substr(h.codigo,5,2)
    ORDER BY c.codigo, n DESC
    """,
)
show(
    "CI amostra limpa",
    """
    SELECT h.codigo, substr(h.enunciado,1,110) enc, al.ano, al.bimestre,
           al.unidade_tematica, al.tema_transversal, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='CI'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 10
    """,
)
show(
    "GE amostra limpa",
    """
    SELECT h.codigo, substr(h.enunciado,1,110) enc, al.ano, al.bimestre,
           al.unidade_tematica, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='GE'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "MA amostra limpa",
    """
    SELECT h.codigo, substr(h.enunciado,1,90) enc, al.ano, al.bimestre,
           al.unidade_tematica, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='MA'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "LP amostra limpa",
    """
    SELECT h.codigo, substr(h.enunciado,1,90) enc, al.ano, al.bimestre,
           al.unidade_tematica, al.eixo, al.pagina
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE c.codigo='LP'
    ORDER BY al.ano, al.bimestre, h.codigo
    LIMIT 8
    """,
)
show(
    "sem enunciado restante",
    """
    SELECT c.codigo, h.codigo, al.ano, al.bimestre, al.pagina, h.enunciado
    FROM alocacao al
    JOIN habilidade h ON h.id=al.habilidade_id
    JOIN componente c ON c.id=h.componente_id
    WHERE length(coalesce(h.enunciado,''))<20
    """,
)
