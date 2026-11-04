# vml-apresentacao-engine

Preenche um template .pptx marcado com tags e baixa o arquivo, dentro de um artefato HTML do WPP Open. O agente monta o objeto `D` e o artefato carrega este engine, mostra a prévia e baixa o PowerPoint. Mesmo padrão do vml-documento-engine.

## Uso no artefato

```html
<div id="avisos"></div>
<input type="file" id="tpl" accept=".pptx" style="display:none">
<button id="baixar">Baixar PowerPoint</button>
<div id="vb"></div>
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/tomcoliveira/vml-apresentacao-engine@v1.4/vml-apresentacao-pptx.js"></script>
<script>
const D = { template: "ford-01", nome_arquivo: "exemplo", valores: { titulo: "..." } };
VMLApresentacao.montar(D, { preview: vb, botao: baixar, entrada: tpl, avisos: avisos });
</script>
```

## Tags

Cada texto editável do template recebe uma tag, digitada de uma vez sem trocar a formatação no meio:
- `{{nome|45}}`: texto com até 45 caracteres.
- `{{nome|Baixo/Médio/Alto}}`: uma das opções.

O parágrafo herda a formatação do primeiro trecho. Texto sem tag não muda. Campo acima do limite ou fora das opções aparece como erro e bloqueia o download.

## Templates

Ficam em `templates/` e são registrados em `TEMPLATES`. O jsDelivr bloqueia .pptx (responde 403), então cada template é publicado também em base64, em `<id>.b64.txt`, que é o arquivo que o engine carrega. Depois de editar o .pptx: `base64 -i templates/<id>.pptx -b 76 > templates/<id>.b64.txt`. Sem template acessível, o artefato pede o .pptx ao usuário. É o caminho para template que não pode ser público.

- `ford-01`: Ford, comparação de três opções (26 campos).

## Teste

`teste.html` usa o engine local. Servir a pasta (`python3 -m http.server`) e abrir `teste.html`.

## Versões

Cada tag é um endereço fixo no jsDelivr. Mudança de engine ou de template = nova tag + troca da versão no prompt do agente.
