const form = document.getElementById("filtros");
const lista = document.getElementById("lista");
const count = document.getElementById("count");
const stats = document.getElementById("stats");

function paramsFromForm() {
  const data = new FormData(form);
  const q = new URLSearchParams();
  for (const [key, value] of data.entries()) {
    if (String(value).trim()) q.set(key, String(value).trim());
  }
  return q;
}

function meta(item) {
  const bits = [];
  if (item.componente) bits.push(item.componente);
  if (item.ano) bits.push(`${item.ano}º ano`);
  if (item.bimestre) bits.push(`${item.bimestre}º bimestre`);
  if (item.unidade_tematica) bits.push(item.unidade_tematica);
  else if (item.eixo) bits.push(item.eixo);
  return bits.join(" · ");
}

function bloco(titulo, texto) {
  if (!texto) return "";
  const safe = escapeHtml(texto);
  return `<div class="bloco"><h3>${titulo}</h3><p>${safe}</p></div>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function card(item) {
  const to = item.origem === "DCT_TO";
  const arts = (item.articulacoes || [])
    .map(
      (code) =>
        `<button type="button" class="art" data-code="${escapeHtml(code)}">${escapeHtml(code)}</button>`
    )
    .join("");
  return `
    <article class="card${to ? " to" : ""}">
      <div class="card-head">
        <span class="codigo">${escapeHtml(item.codigo)}</span>
        <span class="pill${to ? " to" : ""}">${to ? "Acrescentada no TO" : "Código da BNCC"}</span>
      </div>
      <p class="meta">${escapeHtml(meta(item))}</p>
      <p class="enunciado">${escapeHtml(item.enunciado || "Enunciado não extraído neste recorte.")}</p>
      <details>
        <summary>Objetos, sugestão e fonte</summary>
        ${bloco("Unidade / campo", item.unidade_tematica)}
        ${bloco("Eixo", item.eixo)}
        ${bloco("Tema transversal", item.tema_transversal)}
        ${bloco("Objetos de conhecimento", item.objetos_conhecimento)}
        ${bloco("Sugestão pedagógica", item.sugestao_pedagogica)}
        ${arts ? `<div class="bloco"><h3>Articula com</h3><div class="arts">${arts}</div></div>` : ""}
        ${bloco("Fonte", [item.fonte_pdf, item.pagina ? `p. ${item.pagina}` : ""].filter(Boolean).join(" · "))}
      </details>
    </article>
  `;
}

async function carregarResumo() {
  const data = await fetch("/api/resumo").then((r) => r.json());
  stats.hidden = false;
  stats.innerHTML = `
    <div><dt>Habilidades</dt><dd>${data.habilidades}</dd></div>
    <div><dt>Alocações</dt><dd>${data.alocacoes}</dd></div>
    <div><dt>Articulações</dt><dd>${data.articulacoes}</dd></div>
  `;
}

async function buscar() {
  const q = paramsFromForm();
  count.textContent = "Carregando…";
  const data = await fetch(`/api/consulta?${q}`).then((r) => r.json());
  if (!data.itens.length) {
    lista.innerHTML = `<p class="vazio">Nenhuma habilidade neste recorte. Tente outro ano, bimestre ou código.</p>`;
    count.textContent = "0 habilidades";
    return;
  }
  const extra =
    data.total > data.mostrando
      ? ` Mostrando ${data.mostrando} de ${data.total}. Aperte a busca para afunilar.`
      : "";
  count.textContent = `${data.total} habilidade(s) neste recorte.${extra}`;
  lista.innerHTML = data.itens.map(card).join("");
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  buscar();
});
for (const id of ["componente", "ano", "bimestre", "origem"]) {
  document.getElementById(id).addEventListener("change", buscar);
}
lista.addEventListener("click", (event) => {
  const btn = event.target.closest("button.art");
  if (!btn) return;
  document.getElementById("q").value = btn.dataset.code;
  document.getElementById("componente").value = "";
  document.getElementById("ano").value = "";
  document.getElementById("bimestre").value = "";
  buscar();
});

document.getElementById("componente").value = "Ciências";
document.getElementById("ano").value = "6";
carregarResumo().then(buscar);
