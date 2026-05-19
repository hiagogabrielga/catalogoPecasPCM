"use strict";

// ===============================
// ESTADO GLOBAL
// ===============================
let listaCompleta = [];
let debounceTimer;
let atualizacaoDisponivel = false;

const IMAGEM_PADRAO = "placeholder.png";
const CACHE_KEY = "catalogoPecasPCM";
const CACHE_VERSION_KEY = "catalogoPecasPcmVersao";
const VERSAO_ATUAL = "2.0.0";

// ===============================
// UTIL
// ===============================
function removerAcentos(texto) {
  return texto
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

async function verificarAtualizacao() {
  try {
    console.log("Verificando atualização...");

    const [materiais, estoque] = await Promise.all([
      lerExcel("dados-catalago-pcm-interno/nomes_usuais.xlsx"),
      lerExcel("dados-catalago-pcm-interno/relatorio.xlsx"),
    ]);

    const estoqueMap = new Map(
      estoque.map((item) => [
        String(item.material).replace(/\s+/g, ""),
        item,
      ])
    );

    const novaLista = materiais.map((mat) => {
      const cod = String(mat.material).replace(/\s+/g, "");
      const est = estoqueMap.get(cod);

      return {
        codigo: cod,
        quantidade: est?.quantidade ?? 0,
      };
    });

    const atual = JSON.stringify(
      listaCompleta.map((i) => ({
        codigo: i.codigo,
        quantidade: i.quantidade,
      }))
    );

    const novo = JSON.stringify(novaLista);

    if (atual !== novo && !atualizacaoDisponivel) {
      console.log("🔄 Mudança detectada no estoque");

      atualizacaoDisponivel = true;

      mostrarModalAtualizacao();
    } else {
      console.log("Sem alterações");
    }

  } catch (erro) {
    console.error("Erro ao verificar atualização:", erro);
  }
}

// ===============================
// TOAST
// ===============================
function mostrarToast(texto) {
  const toast = document.getElementById("toast");

  toast.textContent = texto;
  toast.classList.add("mostrar");

  setTimeout(() => {
    toast.classList.remove("mostrar");
  }, 2000);
}

// ===============================
// COPIAR CÓDIGO
// ===============================
function copiarCodigo(codigo) {
  navigator.clipboard.writeText(codigo);
  mostrarToast(`Código ${codigo} copiado`);
}

// ===============================
// MODAL UPDATE
// ===============================
function mostrarModalAtualizacao() {
  document.getElementById("updateModal").classList.remove("hidden");
}

function atualizarPagina() {
  localStorage.removeItem(CACHE_KEY);
  localStorage.removeItem(CACHE_VERSION_KEY);
  location.reload();
}

// ===============================
// EXCEL
// ===============================
async function lerExcel(caminho) {
  const response = await fetch(`${caminho}?v=${Date.now()}`, {
    cache: "no-store",
  });

  const arrayBuffer = await response.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
  });

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  return XLSX.utils.sheet_to_json(sheet);
}

// ===============================
// IMAGEM
// ===============================
async function imagemExiste(url) {
  try {
    const response = await fetch(url, { method: "HEAD" });
    if (response.ok) return true;
  } catch (e) {}

  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve(true);
    img.onerror = () => resolve(false);
    img.src = url;
  });
}

async function encontrarImagem(codigo) {
  const extensoes = ["png", "jpg", "jpeg", "gif"];

  for (let ext of extensoes) {
    const caminho = `dados-catalago-pcm-interno/imagens/${codigo}.${ext}`;
    if (await imagemExiste(caminho)) return caminho;
  }

  return IMAGEM_PADRAO;
}

// ===============================
// CARD
// ===============================
function criarCard(item) {
  const card = document.createElement("div");

  const estoqueOk =
    Number(String(item.quantidade).replace(",", ".")) > 0;

  card.className = `card ${estoqueOk ? "estoque-ok" : "estoque-zero"}`;

  card.innerHTML = `
    <div class="status-bar"></div>

    <img src="${item.imagem}" alt="${item.nome}">

    <div class="codigo-linha">
      <span class="codigo">Código: ${item.codigo}</span>
      <button class="copiar-btn" onclick="copiarCodigo('${item.codigo}')">📋</button>
    </div>

    <div class="nome">${item.nome}</div>

    <div class="estoque">
      <span class="quantidade">${item.quantidade}</span>
      <span class="unidade">${item.unidade}</span>
    </div>
  `;

  return card;
}

// ===============================
// AGRUPAMENTO
// ===============================
function agruparPorCategoria(lista) {
  return lista.reduce((acc, item) => {
    const cat = item.categoria || "Sem categoria";

    if (!acc[cat]) acc[cat] = [];

    acc[cat].push(item);

    return acc;
  }, {});
}

// ===============================
// RENDER
// ===============================
function renderizar(lista) {
  const container = document.getElementById("cards");
  container.innerHTML = "";

  const agrupado = agruparPorCategoria(lista);

  Object.entries(agrupado).forEach(([categoria, itens]) => {
    const categoriaDiv = document.createElement("div");
    categoriaDiv.className = "categoria";

    const header = document.createElement("div");
    header.className = "categoria-header";

    const titulo = document.createElement("h2");
    titulo.textContent = categoria;

    const botao = document.createElement("button");
    botao.textContent = "⯆";

    const conteudo = document.createElement("div");
    conteudo.className = "categoria-conteudo";

    itens.forEach((item) => {
      conteudo.appendChild(criarCard(item));
    });

    botao.onclick = () => {
      const fechado = conteudo.classList.toggle("fechado");
      botao.textContent = fechado ? "⯈" : "⯆";
    };

    header.appendChild(titulo);
    header.appendChild(botao);

    categoriaDiv.appendChild(header);
    categoriaDiv.appendChild(conteudo);

    container.appendChild(categoriaDiv);
  });
}

// ===============================
// FILTRO
// ===============================
function filtrarDebounced() {
  clearTimeout(debounceTimer);

  debounceTimer = setTimeout(() => {
    const termo = removerAcentos(
      document.getElementById("search").value
    );

    const filtrados = listaCompleta.filter((item) => {
      return (
        removerAcentos(item.codigo).includes(termo) ||
        removerAcentos(item.nome).includes(termo)
      );
    });

    renderizar(filtrados);
  }, 200);
}

// ===============================
// CARREGAMENTO PRINCIPAL
// ===============================
async function carregarDados() {
  try {
    const versaoSalva = localStorage.getItem(CACHE_VERSION_KEY);
    const dadosSalvos = localStorage.getItem(CACHE_KEY);

    if (dadosSalvos && versaoSalva === VERSAO_ATUAL) {
      listaCompleta = JSON.parse(dadosSalvos);
      renderizar(listaCompleta);
      return;
    }

    const [materiais, estoque] = await Promise.all([
      lerExcel("dados-catalago-pcm-interno/nomes_usuais.xlsx"),
      lerExcel("dados-catalago-pcm-interno/relatorio.xlsx"),
    ]);

    const estoqueMap = new Map(
      estoque.map((item) => [
        String(item.material).replace(/\s+/g, ""),
        item,
      ])
    );

    const candidatos = await Promise.all(
      materiais.map(async (mat) => {
        const cod = String(mat.material).replace(/\s+/g, "");
        const est = estoqueMap.get(cod);

        return {
          codigo: cod,
          nome: mat.nome,
          categoria: mat.categoria,
          imagem: await encontrarImagem(cod),
          unidade: est?.unidade ?? "UN",
          quantidade: est?.quantidade ?? 0,
          deposito: est?.deposito ?? "PC01",
        };
      })
    );

    listaCompleta = candidatos.filter(Boolean);
    renderizar(listaCompleta);

    localStorage.setItem(CACHE_KEY, JSON.stringify(listaCompleta));
    localStorage.setItem(CACHE_VERSION_KEY, VERSAO_ATUAL);

  } catch (err) {
    console.error("Erro:", err);
  }
}

// ===============================
// INIT
// ===============================
window.onload = async () => {
  await carregarDados();

  setInterval(verificarAtualizacao, 5 * 60 * 1000);
};