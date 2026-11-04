/*
 * vml-apresentacao-pptx.js — v1.0
 * Preenche um template .pptx marcado com tags e baixa o arquivo, dentro de um artefato HTML do WPP Open.
 * Não conhece layout: lê as tags do próprio template.
 *
 * Dependência, carregada antes deste arquivo:
 *   https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js
 *
 * Tags no template (texto digitado de uma vez, sem trocar formatação no meio):
 *   {{nome|45}}               texto com até 45 caracteres
 *   {{nome|Baixo/Médio/Alto}} uma das opções
 * O texto do parágrafo herda a formatação do primeiro trecho. Texto sem tag não muda.
 *
 * Contrato de dados (D):
 * { template: "ford-01", nome_arquivo: "proposta-ford", valores: { titulo: "...", c1_nome: "..." } }
 * template pode ser um id registrado em TEMPLATES (arquivo em templates/, ao lado deste script) ou uma URL.
 * Sem template acessível, o artefato pede o .pptx ao usuário.
 *
 * Navegador: VMLApresentacao.montar(D, { preview, botao, entrada, avisos })
 */
(function (root) {
  var TEMPLATES = {
    "ford-01": { arquivo: "templates/ford-01.pptx", nome: "Ford, comparação de três opções" }
  };
  var A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  var P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  var TAG = /\{\{(\w+)\|([^}]+)\}\}/g;
  var EMU = 914400;

  var base = (function () {
    var s = document.currentScript && document.currentScript.src;
    return s ? s.slice(0, s.lastIndexOf("/") + 1) : "";
  })();

  function slides(zip) {
    return Object.keys(zip.files)
      .filter(function (n) { return /^ppt\/slides\/slide\d+\.xml$/.test(n); })
      .sort(function (a, b) { return parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]); });
  }

  function textoDo(p) {
    var ts = p.getElementsByTagNameNS(A, "t"), s = "";
    for (var i = 0; i < ts.length; i++) s += ts[i].textContent;
    return s;
  }

  function carregar(fonte) {
    if (fonte instanceof Blob) return fonte.arrayBuffer().then(JSZip.loadAsync);
    var url = TEMPLATES[fonte] ? base + TEMPLATES[fonte].arquivo : fonte;
    return fetch(url).then(function (r) {
      if (!r.ok) throw new Error("template não encontrado: " + fonte);
      return r.arrayBuffer();
    }).then(JSZip.loadAsync);
  }

  function lerDocs(zip) {
    return Promise.all(slides(zip).map(function (nome) {
      return zip.file(nome).async("string").then(function (xml) {
        return { nome: nome, doc: new DOMParser().parseFromString(xml, "application/xml") };
      });
    }));
  }

  function esquema(docs) {
    var campos = {};
    docs.forEach(function (d) {
      var ps = d.doc.getElementsByTagNameNS(A, "p");
      for (var i = 0; i < ps.length; i++) {
        textoDo(ps[i]).replace(TAG, function (_, nome, lim) {
          campos[nome] = /^\d+$/.test(lim)
            ? { tipo: "texto", max_caracteres: +lim }
            : { tipo: "opcao", opcoes: lim.split("/") };
        });
      }
    });
    return campos;
  }

  function validar(campos, valores) {
    var erros = [], avisos = [];
    Object.keys(campos).forEach(function (nome) {
      var r = campos[nome], v = String(valores[nome] || "").trim();
      if (!v) avisos.push(nome + ": vazio");
      else if (r.tipo === "texto" && v.length > r.max_caracteres) erros.push(nome + ": " + v.length + " caracteres, máximo " + r.max_caracteres);
      else if (r.tipo === "opcao" && r.opcoes.indexOf(v) < 0) erros.push(nome + ": '" + v + "' fora de " + r.opcoes.join("/"));
    });
    Object.keys(valores).forEach(function (nome) {
      if (!campos[nome]) avisos.push(nome + ": não existe no template, ignorado");
    });
    return { erros: erros, avisos: avisos };
  }

  function preencherDocs(docs, valores) {
    docs.forEach(function (d) {
      var ps = d.doc.getElementsByTagNameNS(A, "p");
      for (var i = 0; i < ps.length; i++) {
        var texto = textoDo(ps[i]);
        if (!texto.match(TAG)) continue;
        var ts = ps[i].getElementsByTagNameNS(A, "t");
        ts[0].textContent = texto.replace(TAG, function (_, nome) { return String(valores[nome] || "").trim(); });
        for (var j = 1; j < ts.length; j++) ts[j].textContent = "";
      }
    });
  }

  // Prévia aproximada: cada caixa de texto na posição e no corpo do template, sem imagens nem fontes da marca.
  function preview(docs, el, larguraSlide) {
    el.innerHTML = "";
    docs.forEach(function (d) {
      var w = el.clientWidth || 960, k = w / larguraSlide;
      var tela = document.createElement("div");
      tela.style.cssText = "position:relative;width:" + w + "px;height:" + Math.round(w * 9 / 16) + "px;background:#fff;border:1px solid #ccc;overflow:hidden;margin-bottom:12px;font-family:Arial,sans-serif";
      var sps = d.doc.getElementsByTagNameNS(P, "sp");
      for (var i = 0; i < sps.length; i++) {
        var off = sps[i].getElementsByTagNameNS(A, "off")[0], ext = sps[i].getElementsByTagNameNS(A, "ext")[0];
        var texto = textoDo(sps[i]);
        if (!off || !texto.trim()) continue;
        var rpr = sps[i].getElementsByTagNameNS(A, "rPr")[0];
        var cor = sps[i].querySelector("rPr solidFill srgbClr, defRPr solidFill srgbClr");
        var cx = document.createElement("div");
        cx.textContent = texto;
        cx.style.cssText = "position:absolute;white-space:pre-wrap;line-height:1.15;left:" + off.getAttribute("x") * k + "px;top:" + off.getAttribute("y") * k +
          "px;width:" + ext.getAttribute("cx") * k + "px;font-size:" + ((rpr && rpr.getAttribute("sz")) || 1800) / 100 * EMU / 72 * k +
          "px;font-weight:" + (rpr && rpr.getAttribute("b") === "1" ? "700" : "400") + ";color:#" + (cor ? cor.getAttribute("val") : "111");
        tela.appendChild(cx);
      }
      el.appendChild(tela);
    });
  }

  function nomeArquivo(D) {
    return (String(D.nome_arquivo || "apresentacao").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^\w-]+/g, "-").slice(0, 60) || "apresentacao") + ".pptx";
  }

  function gerar(fonte, D) {
    return carregar(fonte).then(function (zip) {
      return Promise.all([lerDocs(zip), zip.file("ppt/presentation.xml").async("string")]).then(function (r) {
        var docs = r[0], larg = +(r[1].match(/<p:sldSz[^>]*cx="(\d+)"/) || [0, 12192000])[1];
        var campos = esquema(docs), check = validar(campos, D.valores || {});
        return {
          campos: campos, erros: check.erros, avisos: check.avisos,
          preview: function (el) { var copia = docs.map(function (d) { return { doc: d.doc.cloneNode(true) }; }); preencherDocs(copia, D.valores || {}); preview(copia, el, larg); },
          baixar: function () {
            preencherDocs(docs, D.valores || {});
            docs.forEach(function (d) { zip.file(d.nome, new XMLSerializer().serializeToString(d.doc)); });
            return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }).then(function (blob) {
              var a = document.createElement("a");
              a.href = URL.createObjectURL(blob); a.download = nomeArquivo(D);
              document.body.appendChild(a); a.click(); a.remove();
            });
          }
        };
      });
    });
  }

  function montar(D, ui) {
    function mostrar(res) {
      var msgs = res.erros.map(function (e) { return "Erro: " + e; }).concat(res.avisos.map(function (a) { return "Aviso: " + a; }));
      if (ui.avisos) ui.avisos.textContent = msgs.join("\n");
      if (ui.preview) res.preview(ui.preview);
      if (ui.botao) { ui.botao.disabled = res.erros.length > 0; ui.botao.onclick = res.baixar; }
      return res;
    }
    function pedirArquivo(motivo) {
      if (ui.avisos) ui.avisos.textContent = motivo + " Selecione o arquivo .pptx do template.";
      if (ui.entrada) {
        ui.entrada.style.display = "";
        ui.entrada.onchange = function () { if (ui.entrada.files[0]) gerar(ui.entrada.files[0], D).then(mostrar); };
      }
    }
    if (ui.botao) ui.botao.disabled = true;
    if (!D.template) return pedirArquivo("Template não informado.");
    return gerar(D.template, D).then(mostrar, function () { pedirArquivo("Não foi possível carregar o template " + D.template + "."); });
  }

  root.VMLApresentacao = { version: "1.0", templates: TEMPLATES, gerar: gerar, montar: montar, validar: validar };
})(this);
