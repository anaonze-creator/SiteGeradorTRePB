# Gerador de TR e PB

Site estático (sem build, sem backend) que gera o Termo de Referência de Bens
em `.docx` a partir de um formulário guiado.

## Como abrir

O site **precisa ser servido por HTTP** — abrir `index.html` por duplo clique
(`file://`) não funciona: os navegadores bloqueiam módulos ES (`type="module"`)
e o `fetch` do modelo `templates/Template_TR_Bens.docx` nessa origem.

Na raiz do projeto:

```bash
python -m http.server 8000
```

Depois abra <http://localhost:8000>.

Esta é uma limitação conhecida, a ser resolvida junto com a escolha de
hospedagem definitiva pela equipe de TI.

## Regerar o modelo Word

O `templates/Template_TR_Bens.docx` é gerado offline por
`templates/build_template_tr_bens.py` (python-docx). Rodar de dentro de
`templates/` sempre que o texto oficial mudar.

## Testes

```bash
node --test tests/*.test.mjs
```
