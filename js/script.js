const btnProcessar = document.getElementById("btnProcessar");
const tabelaResultado = document.getElementById("resultadoTabela");

let bancoDados = {};

// CAMINHO DO EXCEL
const excelPath = "dados-catalago-pcm-interno/relatorio.xlsx";

// CARREGAR EXCEL AUTOMATICAMENTE
async function carregarExcel() {
  try {
    const response = await fetch(excelPath);

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
    });

    const primeiraAba = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[primeiraAba];

    const json = XLSX.utils.sheet_to_json(worksheet);

    json.forEach((item) => {
      const codigo = String(item.Código || item.material || "").trim();

      const local = String(item.Local || item.deposito || "").trim();

      if (codigo) {
        bancoDados[codigo] = local;
      }
    });

    console.log("Planilha carregada");
    console.log(bancoDados);
  } catch (erro) {
    console.error("Erro ao carregar Excel:", erro);
  }
}

function escolherPatim(linha) {
  return new Promise((resolve) => {
    const modal = document.getElementById("modalPatim");
    const texto = document.getElementById("modalTexto");

    const btnCavalo = document.getElementById("btnCavalo");
    const btnCarreta = document.getElementById("btnCarreta");

    texto.innerText = linha;

    modal.classList.remove("hidden");

    btnCavalo.onclick = () => {
      modal.classList.add("hidden");
      resolve("cavalo");
    };

    btnCarreta.onclick = () => {
      modal.classList.add("hidden");
      resolve("carreta");
    };
  });
}

function escolherOpcaoNivel2(tipo, linha) {
  return new Promise((resolve) => {
    const modal = document.getElementById("modalOpcao2");
    const texto = document.getElementById("modalTexto2");

    const btnA = document.getElementById("btnOpcaoA");
    const btnB = document.getElementById("btnOpcaoB");
    const btnC = document.getElementById("btnOpcaoC");
    const btnD = document.getElementById("btnOpcaoD");

    texto.innerText = linha;

    modal.classList.remove("hidden");

    // reset
    btnA.style.display = "none";
    btnB.style.display = "none";
    btnC.style.display = "none";
    btnD.style.display = "none";

    // =====================================
    // CAVALO -> A e B
    // =====================================
    if (tipo === "cavalo") {
      btnA.style.display = "inline-block";
      btnB.style.display = "inline-block";

      btnA.onclick = () => {
        modal.classList.add("hidden");
        resolve("scania-dianteira");
      };

      btnB.onclick = () => {
        modal.classList.add("hidden");
        resolve("scania-tracao");
      };
    }

    // =====================================
    // CARRETA -> C e D
    // =====================================
    else {
      btnC.style.display = "inline-block";
      btnD.style.display = "inline-block";

      btnC.onclick = () => {
        modal.classList.add("hidden");
        resolve("librelato");
      };

      btnD.onclick = () => {
        modal.classList.add("hidden");
        resolve("randon");
      };
    }
  });
}

function extrairCodigo(texto) {
  texto = texto.replace(/\t/g, " ").replace(/\s+/g, " ").trim();

  texto = texto.replace(/^\d+\s*\.\s*/, "");

  if (texto.includes(" - ")) {
    return texto.split(" - ")[0].trim();
  }

  const match = texto.match(/\d{5,}/);

  return match ? match[0] : "";
}

function extrairQuantidade(texto) {
  texto = texto.replace(/\t/g, " ").replace(/\s+/g, " ").trim();
  texto = texto.replace("(", "").replace(")", "")
  let matchInicio = texto.match(/^(\d+)\s*\.\s*/);

  if (matchInicio) {
    return parseInt(matchInicio[1]);
  }

  let parteDepoisTraco = "";

  if (texto.includes(" - ")) {
    parteDepoisTraco = texto.split(" - ")[1].trim();
  } else {
    parteDepoisTraco = texto;
  }

  let matchDepoisTraco = parteDepoisTraco.match(/^(\d+)\s+/);

  if (matchDepoisTraco) {
    return parseInt(matchDepoisTraco[1]);
  }

  let matchUn = parteDepoisTraco.match(/(\d+)\s*un$/i);

  if (matchUn) {
    return parseInt(matchUn[1]);
  }

  let matchUnd = parteDepoisTraco.match(/(\d+)\s*und$/i);

  if (matchUnd) {
    return parseInt(matchUnd[1]);
  }


  let matchFinal = parteDepoisTraco.match(/(\d+)$/);

  if (matchFinal) {
    return parseInt(matchFinal[1]);
  }

  return 0;
}

function criarLinhaTabela(codigo, quantidade, local, encontrado) {
  const tr = document.createElement("tr");

  tr.innerHTML = `
        <td>${codigo}</td>
        <td> </td>
        <td>${quantidade}</td>
        <td> </td>
        <td>L</td>
        <td>${local}</td>
        <td><input type="checkbox" name="" id=""></td>
    `;

  tabelaResultado.appendChild(tr);
}

// PROCESSAR DADOS
async function processarDados() {
  tabelaResultado.innerHTML = "";

  const entrada = document.getElementById("entrada").value;

  const linhas = entrada.split("\n");

  for (let linha of linhas) {
    if (!linha.trim()) continue;

    const texto = linha.toLowerCase();

    const quantidade = extrairQuantidade(linha);

    if ((texto.includes("patim") && texto.includes("rebitado")) || texto.includes("patinho")) {
      const tipo = await escolherPatim(linha);
      const opcao = await escolherOpcaoNivel2(tipo, linha);
      let quantidade_jg_lona = Math.ceil(Number(quantidade) / 4);
      if (tipo === "cavalo") {
        if (opcao === "scania-dianteira") {
          criarLinhaTabela("1010020541", quantidade_jg_lona, "PC01", true);
        }

        if (opcao === "scania-tracao") {
          criarLinhaTabela("1010020879", quantidade_jg_lona, "PC01", true);
        }
        criarLinhaTabela("1010029427", quantidade_jg_lona * 130, "PC01", true);
      } else {
        if (opcao === "librelato") {
          criarLinhaTabela("1040077380", quantidade_jg_lona, "PC01", true);
        }

        if (opcao === "randon") {
          criarLinhaTabela("409250", quantidade_jg_lona, "PC01", true);
        }
        criarLinhaTabela("1010029427", quantidade_jg_lona * 130, "PC01", true);
        criarLinhaTabela("1010021543", quantidade, "PC01", true);
        criarLinhaTabela("1010084315", quantidade, "PC01", true);
      }
      continue;
    }

    const codigo = extrairCodigo(linha);

    let local = bancoDados[codigo];

    let encontrado = true;

    if (!local) {
      local = "PC01";
      encontrado = false;
    }

    if (local === "PC02") {
      local = "PC01";
    }

    criarLinhaTabela(codigo, quantidade, local, encontrado);
  }
}
// EVENTO
btnProcessar.addEventListener("click", processarDados);

// INICIAR
carregarExcel();
