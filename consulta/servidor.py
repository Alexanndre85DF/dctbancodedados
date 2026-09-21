# -*- coding: utf-8 -*-
"""Página de consulta do DCT Tocantins — lê o SQLite e serve a interface."""

from __future__ import annotations

import json
import os
import socket
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = ROOT / "saida" / "dct_tocantins.sqlite"
STATIC_DIR = (Path(__file__).resolve().parent / "static").resolve()
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8765"))

MIME = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".ico": "image/x-icon",
    ".svg": "image/svg+xml",
}


def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def json_bytes(payload, status=200):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    return status, "application/json; charset=utf-8", body


def resumo():
    conn = db()
    try:
        totais = conn.execute(
            """
            SELECT
                (SELECT COUNT(*) FROM habilidade) AS habilidades,
                (SELECT COUNT(*) FROM alocacao) AS alocacoes,
                (SELECT COUNT(*) FROM articulacao) AS articulacoes
            """
        ).fetchone()
        por_comp = conn.execute(
            """
            SELECT c.nome, c.codigo, COUNT(DISTINCT h.id) AS habilidades, COUNT(al.id) AS alocacoes
            FROM componente c
            LEFT JOIN habilidade h ON h.componente_id = c.id
            LEFT JOIN alocacao al ON al.habilidade_id = h.id
            GROUP BY c.id
            ORDER BY c.codigo
            """
        ).fetchall()
        return {
            "habilidades": totais["habilidades"],
            "alocacoes": totais["alocacoes"],
            "articulacoes": totais["articulacoes"],
            "componentes": [dict(r) for r in por_comp],
        }
    finally:
        conn.close()


def consulta(params: dict):
    componente = (params.get("componente") or [""])[0].strip()
    ano = (params.get("ano") or [""])[0].strip()
    bimestre = (params.get("bimestre") or [""])[0].strip()
    origem = (params.get("origem") or [""])[0].strip()
    q = (params.get("q") or [""])[0].strip()
    try:
        limite = min(int((params.get("limite") or ["80"])[0]), 200)
    except ValueError:
        limite = 80

    where = ["1=1"]
    args: list = []
    if componente:
        where.append("componente = ?")
        args.append(componente)
    if ano:
        where.append("ano = ?")
        args.append(int(ano))
    if bimestre:
        where.append("bimestre = ?")
        args.append(int(bimestre))
    if origem in {"BNCC", "DCT_TO"}:
        where.append("origem = ?")
        args.append(origem)
    if q:
        like = f"%{q}%"
        where.append(
            "(codigo LIKE ? OR enunciado LIKE ? OR unidade_tematica LIKE ? OR objetos_conhecimento LIKE ?)"
        )
        args.extend([like, like, like, like])

    sql_where = " AND ".join(where)
    conn = db()
    try:
        total = conn.execute(
            f"SELECT COUNT(*) FROM v_consulta WHERE {sql_where}", args
        ).fetchone()[0]
        rows = conn.execute(
            f"""
            SELECT * FROM v_consulta
            WHERE {sql_where}
            ORDER BY ano, bimestre, codigo
            LIMIT ?
            """,
            [*args, limite],
        ).fetchall()
        itens = [dict(r) for r in rows]
        if itens:
            codes = list({item["codigo"] for item in itens})
            placeholders = ",".join("?" * len(codes))
            arts = conn.execute(
                f"SELECT de_codigo, para_codigo FROM articulacao WHERE de_codigo IN ({placeholders})",
                codes,
            ).fetchall()
            mapa: dict[str, list[str]] = {}
            for de, para in arts:
                mapa.setdefault(de, []).append(para)
            for item in itens:
                item["articulacoes"] = mapa.get(item["codigo"], [])
        return {"total": total, "mostrando": len(itens), "itens": itens}
    finally:
        conn.close()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[consulta] {self.address_string()} {fmt % args}")

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        params = parse_qs(parsed.query)

        if path in {"/api/resumo"}:
            status, ctype, body = json_bytes(resumo())
            extra = {}
        elif path in {"/api/consulta"}:
            status, ctype, body = json_bytes(consulta(params))
            extra = {}
        elif path.startswith("/download/"):
            status, ctype, body, extra = self._download(path)
        else:
            status, ctype, body = self._static(path)
            extra = {}

        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in extra.items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def _download(self, path: str):
        arquivos = {
            "/download/dct_tocantins.sqlite": (
                DB_PATH,
                "application/vnd.sqlite3",
                "dct_tocantins.sqlite",
            ),
            "/download/dct_tocantins.sql": (
                ROOT / "saida" / "dct_tocantins.sql",
                "application/sql; charset=utf-8",
                "dct_tocantins.sql",
            ),
            "/download/dct_tocantins.xlsx": (
                ROOT / "saida" / "dct_tocantins.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "dct_tocantins.xlsx",
            ),
        }
        info = arquivos.get(path)
        if not info:
            return 404, "text/plain; charset=utf-8", b"Not found", {}
        arquivo, ctype, nome = info
        if not arquivo.is_file():
            return 404, "text/plain; charset=utf-8", b"Not found", {}
        extra = {"Content-Disposition": f'attachment; filename="{nome}"'}
        return 200, ctype, arquivo.read_bytes(), extra

    def _static(self, path: str):
        if path in {"/", "/index.html"}:
            rel = "index.html"
        else:
            rel = path.lstrip("/").replace("\\", "/")
        target = (STATIC_DIR / rel).resolve()
        if STATIC_DIR not in target.parents and target != STATIC_DIR:
            return 403, "text/plain; charset=utf-8", b"Forbidden"
        if not target.is_file():
            return 404, "text/plain; charset=utf-8", b"Not found"
        return 200, MIME.get(target.suffix, "application/octet-stream"), target.read_bytes()


def ips_da_rede() -> list[str]:
    encontrados: list[str] = []
    try:
        nome = socket.gethostname()
        for info in socket.getaddrinfo(nome, None, socket.AF_INET):
            ip = info[4][0]
            if ip in encontrados:
                continue
            if ip.startswith("192.168.") or ip.startswith("10."):
                encontrados.append(ip)
            elif ip.startswith("172."):
                segundo = int(ip.split(".")[1])
                if 16 <= segundo <= 31:
                    encontrados.append(ip)
    except OSError:
        pass
    return encontrados


def main():
    if not DB_PATH.exists():
        raise SystemExit(f"Banco não encontrado: {DB_PATH}\nRode antes: python scripts/montar_dct.py")
    httpd = ThreadingHTTPServer((HOST, PORT), Handler)
    print("Consulta DCT Tocantins")
    print(f"Ouvindo em {HOST}:{PORT}")
    print(f"Neste computador:  http://127.0.0.1:{PORT}/")
    for ip in ips_da_rede():
        print(f"Na rede (outros PCs): http://{ip}:{PORT}/")
    print("Deixe esta janela aberta. Ctrl+C para encerrar.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrado.")
        httpd.server_close()


if __name__ == "__main__":
    main()
