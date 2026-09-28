# Control Hub

Painel web central para acessar, organizar e monitorar as várias ferramentas usadas no dia a dia.

## Funcionalidades

- **Cartões de ferramentas** com ícone, cor, descrição, categoria, tags e notas.
- **Categorias** geradas automaticamente na barra lateral, além das visões *Todas*, *Favoritas*, *Recentes* e *Com problema*.
- **Busca instantânea** (atalhos `/` ou `Ctrl+K`; `Enter` abre o primeiro resultado).
- **Verificação de status** (online / fora do ar) por ferramenta ou de todas de uma vez.
- **Favoritas**, contagem de uso e ordenação (manual, nome, mais usadas, recentes).
- **Arrastar e soltar** para reorganizar os cartões na ordem manual.
- **Grade ou lista**, **tema claro/escuro** e layout responsivo (celular).
- **Exportar / importar** a configuração em JSON para compartilhar com a equipe ou fazer backup.
- Atalho `N` para cadastrar uma nova ferramenta.

## Versão publicada

A equipe usa o app como artefato do Claude, com a lista de ferramentas compartilhada entre as contas:
https://claude.ai/artifact/SGDHdEfbJi2qBjfNRo3vXS

Para dar acesso a outra conta, use o menu **Compartilhar** do artefato: *Colaborador* edita a lista,
*Leitor* só visualiza e *Editor* também pode publicar novas versões do código. Veja `CLAUDE.md` para o
fluxo de alteração e publicação.

## Como usar localmente

Não há dependências nem etapa de build — é HTML, CSS e JavaScript puro.

```bash
# opção 1: abrir direto no navegador
open index.html

# opção 2: servir localmente
python3 -m http.server 8080
# acesse http://localhost:8080
```

Também pode ser publicado em qualquer hospedagem estática (GitHub Pages, Netlify, servidor interno etc.).

## Dados

As ferramentas ficam salvas no `localStorage` do navegador. Use **Exportar** para gerar um arquivo
`control-hub-AAAA-MM-DD.json` e **Importar** para carregá-lo em outro navegador (substituindo ou mesclando).

Não guarde senhas no campo de notas — use um cofre de senhas.

No artefato, a lista fica no banco compartilhado do próprio artefato e toda alteração aparece na hora para
todos que estão com a página aberta.

## Sobre a verificação de status

Disponível só no modo local: o artefato não permite requisições a outros sites.


O status é verificado pelo próprio navegador com uma requisição `no-cors`: se o servidor responde, a
ferramenta aparece como **Online**; se houver erro de rede ou demorar mais de 8 s, como **Fora do ar**.
É uma checagem de alcance (DNS/rede), não de saúde da aplicação — um site que responde com erro 500 ainda
aparece como online.

## Estrutura

```
index.html       # estrutura da página e modal de cadastro
css/styles.css   # estilos (temas claro/escuro, grade/lista, responsivo)
js/app.js        # estado, renderização, persistência, status, import/export
scripts/build-artifact.mjs  # gera dist/control-hub.html para publicar como artefato
```
