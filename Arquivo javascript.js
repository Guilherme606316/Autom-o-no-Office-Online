Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    // Passo 1: Preencher cabecalho automático se faltar
    await Word.run(async (context) => {
      const body = context.document.body;
      const hoje = new Date().toLocaleDateString("pt-BR");
      const searchText = `Empresa construção - |Pintura Demarcação Viária| Forro Gípseos| Drywall| - ${hoje}`;
      const results = body.search(searchText, { matchCase: false });
      results.load("items");
      await context.sync();

      if (results.items.length === 0) {
        body.insertParagraph(
          `CONSTRUTORA PORTO - |Pintura| Demarcação Viária| Forro Gípseos| Drywall| Quarta - ${hoje}`,
          Word.InsertLocation.start
        );
        await context.sync();
      }
    }).catch(console.error);

    // Passo 2: Atribuir handlers dos botões
    document.getElementById("btnCadastrar").onclick = showCadastroForm;
    document.getElementById("btnAtualizar").onclick = startUpdateFlow;

    // Passo 3: Navegação por setas nos botões fixos
    const fixedInputs = [
      document.getElementById("btnCadastrar"),
      document.getElementById("btnAtualizar")
    ];
    enableArrowNavigation(fixedInputs);
  }
});

// Exibe mensagem de status
function showStatus(msg, isError = false) {
  const st = document.getElementById("status");
  st.textContent = msg;
  st.style.color = isError ? "red" : "green";
}

// Restaura UI principal
function resetUI() {
  const form = document.getElementById("inlineForm");
  form.innerHTML = "";
  document.getElementById("btnCadastrar").style.display = "";
  document.getElementById("btnAtualizar").style.display = "";
}

// Habilita setas para navegar no array de elementos
function enableArrowNavigation(elements) {
  elements.forEach((el, idx) => {
    el.addEventListener("keydown", (e) => {
      let target;
      if (e.key === "ArrowDown") target = elements[idx + 1];
      if (e.key === "ArrowUp") target = elements[idx - 1];
      if (target) {
        e.preventDefault();
        target.focus();
      }
    });
  });
}

// Monta formulário de cadastro inline
function showCadastroForm() {
  document.getElementById("btnAtualizar").style.display = "none";
  const f = document.getElementById("inlineForm");
  f.innerHTML = `
    <div class="form-group">
      <label for="cad_os">OS (5 dígitos)</label>
      <input type="text" id="cad_os" maxlength="5" />
      <button id="skip_os" type="button">Recusar</button>
    </div>
    <div class="form-group">
      <label for="cad_tipo">Tipo</label>
      <datalist id="cad_tipos">
        <option value="">Selecione</option>
        <option>Execução</option>
        <option>Vistoria</option>
        <option>Continuar</option>
        <option>Execução/vistoria</option>
        <option>Vistoria/execução</option>
      </datalist>
      <input type="text" id="cad_tipo" list="cad_tipos" />
      <button id="skip_tipo" type="button">Recusar</button>
    </div>
    <div class="form-group">
      <label for="cad_desc">Descrição</label>
      <input type="text" id="cad_desc" />
      <button id="skip_desc" type="button">Recusar</button>
    </div>
    <button id="cad_submit" type="button">Enviar</button>
  `;

  // Botões recusar
  document.getElementById("skip_os").onclick = () => document.getElementById("cad_os").value = "Sem OS";
  document.getElementById("skip_tipo").onclick = () => document.getElementById("cad_tipo").value = "Sem Tipo";
  document.getElementById("skip_desc").onclick = () => document.getElementById("cad_desc").value = "Sem Descrição";

  // Navegação por setas
  enableArrowNavigation([
    document.getElementById("cad_os"),
    document.getElementById("skip_os"),
    document.getElementById("cad_tipo"),
    document.getElementById("skip_tipo"),
    document.getElementById("cad_desc"),
    document.getElementById("skip_desc"),
    document.getElementById("cad_submit")
  ]);

  // Handler de envio
  document.getElementById("cad_submit").onclick = async () => {
    const os = document.getElementById("cad_os").value.trim();
    const tipo = document.getElementById("cad_tipo").value.trim();
    const desc = document.getElementById("cad_desc").value.trim();

    if (!/^[0-9]{5}$/.test(os) && os !== "Sem OS") {
      return showStatus("OS inválida", true);
    }
    if (!tipo && tipo !== "Sem Tipo") {
      return showStatus("Tipo obrigatório", true);
    }
    if (!desc && desc !== "Sem Descrição") {
      return showStatus("Descrição obrigatória", true);
    }

    await Word.run(async ctx => {
      const txt = `#${os} – ${tipo} – ${desc}`;
      ctx.document.body.insertParagraph(txt, Word.InsertLocation.end);
      await ctx.sync();
    }).catch(console.error);

    showStatus("Cadastrado!");
    resetUI();
  };
}

// Inicia fluxo de update inline
async function startUpdateFlow() {
  document.getElementById("btnCadastrar").style.display = "none";
  const placeholders = ["Sem OS", "Sem Tipo", "Sem Descrição"];

  for (const ph of placeholders) {
    await Word.run(async ctx => {
      const results = ctx.document.body.search(ph, { matchCase: false });
      results.load("items");
      await ctx.sync();
      if (results.items.length === 0) return;

      for (const item of results.items) {
        const ranges = item.getTextRanges(["\n"], false);
        ranges.load("items");
        await ctx.sync();
        const text = ranges.items[0] ?.text.trim() || "";

        // Monta inline replacement form
        const f = document.getElementById("inlineForm");
        f.innerHTML = `
          <div class="form-group">
            <label>Substituir ${ph} em:</label>
            <div>${text}</div>
            <input type="text" id="upd_val" />
            <button id="upd_ok"   type="button">OK</button>
            <button id="upd_skip" type="button">Pular</button>
          </div>
        `;

        const ok = document.getElementById("upd_ok");
        const skip = document.getElementById("upd_skip");
        const val = document.getElementById("upd_val");
        val.focus();

        // Navegação
        enableArrowNavigation([val, ok, skip]);

        const answer = await new Promise(res => {
          ok.onclick = () => res(val.value.trim());
          skip.onclick = () => res(null);
        });

        if (answer !== null && answer !== "") {
          item.insertText(answer, Word.InsertLocation.replace);
          await ctx.sync();
        }
      }
    }).catch(console.error);
  }

  showStatus("Todas as pendências foram atualizadas.");
  resetUI();
}
