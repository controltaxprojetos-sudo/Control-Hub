# Control Hub — guia para quem mexe no código

Painel web (HTML/CSS/JS puro, sem dependências) com os atalhos das ferramentas do escritório.
Várias contas do Claude trabalham neste repositório: siga o fluxo abaixo para não sobrescrever o trabalho de outra pessoa.

## Onde o app está publicado

- Artefato (versão usada pela equipe): https://claude.ai/artifact/SGDHdEfbJi2qBjfNRo3vXS
- Capacidades declaradas: `db` (lista compartilhada), `user` (saber se o visitante pode editar), `downloads` (exportar JSON).

## Estrutura

```
index.html                 estrutura da página, modal de cadastro e diálogo de confirmação
css/styles.css             estilos (tokens de tema claro/escuro, grade/lista, responsivo)
js/app.js                  estado, renderização, persistência, import/export
scripts/build-artifact.mjs gera dist/control-hub.html (arquivo único publicado como artefato)
```

## Dois modos de execução (js/app.js)

- **Artefato** (`window.claude.use` existe): a lista vive no banco compartilhado do artefato, coleção `tools`,
  um documento por ferramenta (id do documento = id da ferramenta; o corpo não repete o id).
  Todos os que têm acesso de Colaborador (Contributor) ou mais editam; Leitores veem em modo leitura.
  A checagem de status fica escondida, porque o artefato bloqueia requisições a outros sites.
- **Local** (abrir `index.html` ou servir a pasta): a lista fica no `localStorage`, começando por `DEFAULT_TOOLS`.

Restrições do artefato que o código já respeita e devem continuar respeitadas:
sem `confirm()`/`alert()`/`prompt()` (use `ask()`), sem `window.open` (use `<a href target="_blank">`),
sem `<a download>` (use a capacidade `downloads`), sem `fetch` para outros domínios.

## Fluxo para alterar e publicar

1. Trabalhe numa branch e abra PR para `main`; não publique a partir de código que não está no GitHub.
2. Teste localmente: `python3 -m http.server 8080` e abra http://localhost:8080.
3. Gere o arquivo do artefato: `node scripts/build-artifact.mjs` (a pasta `dist/` não vai para o git).
4. Publique **atualizando o artefato existente**: Artifact tool com `url: https://claude.ai/artifact/SGDHdEfbJi2qBjfNRo3vXS`
   e `file_path: dist/control-hub.html`. Leia o artefato antes (`action: "read"`). Sem `url`, cria-se um artefato novo e
   a equipe fica com dois links.
   Omitir `capabilities` mantém as atuais; só passe o objeto se for mudar as capacidades (ele substitui o conjunto inteiro).
5. Para mudar dados (e não código), edite pela própria página ou pelo ArtifactData na coleção `tools`.
   Não coloque ferramentas fixas no código do modo artefato.

Para publicar no mesmo link, a conta precisa de acesso de **Editor** ao artefato (menu Compartilhar).
