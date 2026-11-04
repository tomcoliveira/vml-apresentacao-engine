/*
 * vml-apresentacao-pptx.js \u2014 v1.6
 * Preenche um template .pptx marcado com tags e baixa o arquivo, dentro de um artefato HTML do WPP Open.
 * N\u00e3o conhece layout: l\u00ea as tags do pr\u00f3prio template.
 *
 * Depend\u00eancia, carregada antes deste arquivo:
 *   https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js
 *
 * Tags no template (texto digitado de uma vez, sem trocar formata\u00e7\u00e3o no meio):
 *   {{nome|45}}               texto com at\u00e9 45 caracteres
 *   {{nome|Baixo/M\u00e9dio/Alto}} uma das op\u00e7\u00f5es
 * O texto do par\u00e1grafo herda a formata\u00e7\u00e3o do primeiro trecho. Texto sem tag n\u00e3o muda.
 *
 * Contrato de dados (D):
 * { template: "ford-01", nome_arquivo: "proposta-ford", valores: { titulo: "...", c1_nome: "..." } }
 * template pode ser um id registrado em TEMPLATES (arquivo em templates/, ao lado deste script) ou uma URL.
 * Sem template acess\u00edvel, o artefato pede o .pptx ao usu\u00e1rio.
 *
 * Navegador: VMLApresentacao.montar(D, { preview, botao, entrada, avisos })
 */
(function (root) {
  var TEMPLATES = {
    "ford-01": { arquivo: "templates/ford-01.js", nome: "Ford, compara\u00e7\u00e3o de tr\u00eas op\u00e7\u00f5es" }
  };
  // O jsDelivr bloqueia .pptx; o template publicado vai em base64 dentro de templates/<id>.js, gerado a partir do .pptx de mesmo nome.
  var A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  var P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  var TAG = /\{\{(\w+)\|([^}]+)\}\}/g;
  var EMU = 914400;

  var base = (function () {
    var s = document.currentScript && document.currentScript.src;
    return s ? s.slice(0, s.lastIndexOf("/") + 1) : "https://cdn.jsdelivr.net/gh/tomcoliveira/vml-apresentacao-engine@v1.6/";
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

  // Template registrado vem por <script> (templates/<id>.js define o base64); o Open pode bloquear fetch, mas carrega scripts.
  function porScript(id) {
    root.VMLTemplates = root.VMLTemplates || {};
    if (root.VMLTemplates[id]) return Promise.resolve(root.VMLTemplates[id]);
    return new Promise(function (ok, falha) {
      var s = document.createElement("script");
      s.src = base + "templates/" + id + ".js";
      s.onload = function () { root.VMLTemplates[id] ? ok(root.VMLTemplates[id]) : falha(new Error("template vazio: " + id)); };
      s.onerror = function () { falha(new Error("template n\u00e3o carregou: " + id)); };
      (document.head || document.body).appendChild(s);
    });
  }

  function carregar(fonte) {
    if (fonte instanceof Blob) return fonte.arrayBuffer().then(JSZip.loadAsync);
    if (TEMPLATES[fonte]) {
      return porScript(fonte).then(function (b64) { return JSZip.loadAsync(b64, { base64: true }); });
    }
    return fetch(fonte).then(function (r) {
      if (!r.ok) throw new Error("template n\u00e3o encontrado: " + fonte);
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

  // O Open renderiza o prompt como markdown e pode comer o "_" dos nomes; c1_nome e c1nome valem o mesmo.
  function chave(nome) { return String(nome).toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function valorDe(valores, nome) {
    var k = chave(nome);
    for (var n in valores) if (chave(n) === k) return String(valores[n] || "").trim();
    return "";
  }

  function validar(campos, valores) {
    var erros = [], avisos = [], conhecidas = Object.keys(campos).map(chave);
    Object.keys(campos).forEach(function (nome) {
      var r = campos[nome], v = valorDe(valores, nome);
      if (!v) avisos.push(nome + ": vazio");
      else if (r.tipo === "texto" && v.length > r.max_caracteres) erros.push(nome + ": " + v.length + " caracteres, m\u00e1ximo " + r.max_caracteres);
      else if (r.tipo === "opcao" && r.opcoes.indexOf(v) < 0) erros.push(nome + ": '" + v + "' fora de " + r.opcoes.join("/"));
    });
    Object.keys(valores).forEach(function (nome) {
      if (conhecidas.indexOf(chave(nome)) < 0) erros.push(nome + ": n\u00e3o existe no template");
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
        ts[0].textContent = texto.replace(TAG, function (_, nome) { return valorDe(valores, nome); });
        for (var j = 1; j < ts.length; j++) ts[j].textContent = "";
      }
    });
  }

  // Pr\u00e9via aproximada: cada caixa de texto na posi\u00e7\u00e3o e no corpo do template, sem imagens nem fontes da marca.
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
    return (String(D.nome_arquivo || D.nomearquivo || "apresentacao").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\w-]+/g, "-").slice(0, 60) || "apresentacao") + ".pptx";
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
            var MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
            return Promise.all([zip.generateAsync({ type: "blob", mimeType: MIME }), zip.generateAsync({ type: "base64" })]).then(function (r) {
              var nome = nomeArquivo(D);
              try {
                var a = document.createElement("a");
                a.href = URL.createObjectURL(r[0]); a.download = nome;
                document.body.appendChild(a); a.click(); a.remove();
              } catch (e) {}
              // Plano B: alguns navegadores ignoram o download autom\u00e1tico dentro do artefato do Open.
              var caixa = document.getElementById("vml-links");
              if (!caixa) { caixa = document.createElement("div"); caixa.id = "vml-links"; document.body.insertBefore(caixa, document.body.firstChild); }
              caixa.style.cssText = "font-size:13px;margin:0 0 12px";
              caixa.innerHTML = "";
              caixa.appendChild(document.createTextNode("Se o download n\u00e3o come\u00e7ou: "));
              var l = document.createElement("a");
              l.href = "data:" + MIME + ";base64," + r[1]; l.download = nome; l.target = "_blank";
              l.textContent = "Baixar o arquivo .pptx";
              caixa.appendChild(l);
              caixa.appendChild(document.createTextNode(" (ou clique com o bot\u00e3o direito e escolha Salvar link como)."));
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
    if (!D.template) return pedirArquivo("Template n\u00e3o informado.");
    return gerar(D.template, D).then(mostrar, function () { pedirArquivo("N\u00e3o foi poss\u00edvel carregar o template " + D.template + "."); });
  }

  // Artefato m\u00ednimo: monta a interface e carrega o JSZip sozinho, para o agente s\u00f3 escrever D.
  function artefato(D) {
    if (root.__vmlRodou) return;
    root.__vmlRodou = true;
    function iniciar() {
      var corpo = document.body;
      var aviso = document.getElementById("vml-carregando");
      if (aviso) aviso.remove();
      corpo.style.cssText += ";font-family:Arial,sans-serif;margin:16px";
      var avisos = document.createElement("div");
      avisos.style.cssText = "white-space:pre-wrap;font-size:13px;color:#a00;margin-bottom:8px";
      var entrada = document.createElement("input");
      entrada.type = "file"; entrada.accept = ".pptx"; entrada.style.display = "none";
      var botao = document.createElement("button");
      botao.textContent = "Baixar PowerPoint";
      botao.style.cssText = "font-size:14px;padding:8px 16px;margin-bottom:12px";
      var vb = document.createElement("div");
      vb.style.cssText = "width:100%;max-width:960px";
      var links = document.createElement("div");
      links.id = "vml-links";
      [avisos, entrada, botao, links, vb].forEach(function (e) { corpo.appendChild(e); });
      montar(D, { preview: vb, botao: botao, entrada: entrada, avisos: avisos });
    }
    function comZip() {
      if (root.JSZip) return iniciar();
      var s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js";
      s.onload = iniciar;
      s.onerror = function () { document.body.textContent = "N\u00e3o foi poss\u00edvel carregar o JSZip do jsDelivr."; };
      (document.head || document.body).appendChild(s);
    }
    if (document.body) comZip(); else document.addEventListener("DOMContentLoaded", comZip);
  }

  root.VMLApresentacao = { version: "1.6", templates: TEMPLATES, gerar: gerar, montar: montar, validar: validar, artefato: artefato };
  // O Open pode rodar o script do artefato antes deste carregar; o artefato deixa os dados em VMLD e o engine roda sozinho.
  if (root.VMLD) artefato(root.VMLD);
})(typeof window !== "undefined" ? window : this);
