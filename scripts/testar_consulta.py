# -*- coding: utf-8 -*-
import json
import sys
import urllib.parse
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
BASE = "http://127.0.0.1:8765"


def get(path: str):
    with urllib.request.urlopen(BASE + path) as response:
        return response.status, response.headers.get_content_type(), response.read()


status, ctype, body = get("/")
print(
    "HOME",
    status,
    ctype,
    "bytes",
    len(body),
    "form",
    b'id="filtros"' in body,
    "css",
    b"styles.css" in body,
    "js",
    b"app.js" in body,
)

status, ctype, body = get("/styles.css")
print("CSS", status, len(body))
status, ctype, body = get("/app.js")
print("JS", status, len(body))

cases = [
    "/api/consulta?componente=Geografia&ano=1&bimestre=2",
    "/api/consulta?origem=DCT_TO&limite=5",
    "/api/consulta?" + urllib.parse.urlencode({"q": "EF04GE13TO"}),
    "/api/consulta?"
    + urllib.parse.urlencode({"componente": "Matemática", "ano": "3", "bimestre": "1"}),
    "/api/consulta?"
    + urllib.parse.urlencode({"componente": "Língua Portuguesa", "ano": "1", "bimestre": "1"}),
    "/api/consulta?componente=Arte&ano=9",
    "/api/consulta?componente=Ciências&ano=6&bimestre=4",
]
for path in cases:
    status, ctype, body = get(path)
    data = json.loads(body)
    first = data["itens"][0]["codigo"] if data["itens"] else "-"
    enc = (data["itens"][0].get("enunciado") or "")[:50] if data["itens"] else ""
    print(f"{data['total']:4} {data['mostrando']:3} {first:14} {enc}")
