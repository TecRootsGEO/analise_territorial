---
description: 'Guia rápido de instalação do Graphify (skill que transforma uma pasta de código/docs em um grafo de conhecimento consultável) para uso com Claude Code e GitHub Copilot, via uv.'
---

# Instalação do Graphify

[Graphify](https://github.com/Graphify-Labs/graphify) transforma uma pasta (código, schemas SQL, docs, PDFs, imagens, vídeos) em um grafo de conhecimento consultável. Código é parseado localmente via tree-sitter (nada sai da máquina); conteúdo não-código usa o modelo do assistente de IA ou uma API key configurada.

## 1. Instalar o CLI via uv

```bash
uv tool install graphifyy
```

> O pacote no PyPI é `graphifyy` (com dois "y"), mas o comando instalado é `graphify`.

## 2. Configurar para Claude Code

```bash
graphify install
```

Use dentro do Claude Code:

```bash
/graphify .
```

Para que o Claude Code sempre consulte o grafo em perguntas sobre o código (gera `CLAUDE.md` + hook `PreToolUse`):

```bash
graphify claude install
```

## 3. Configurar para GitHub Copilot

```bash
graphify install --platform copilot
```

Para uso sempre ativo no Copilot CLI:

```bash
graphify copilot install
```

## Comandos básicos de uso

```bash
/graphify .                          # constrói o grafo de conhecimento da pasta atual (dentro do assistente)
/graphify ./docs --update            # re-extrai apenas os arquivos alterados
/graphify add <url>                  # adiciona um paper ou vídeo (ex.: YouTube) ao grafo

graphify query "pergunta em linguagem natural"   # retorna o subgrafo que responde à pergunta
graphify path "A" "B"                            # traça a conexão mais curta entre dois conceitos
graphify explain "Nome"                          # detalha um nó: origem, comunidade e conexões

graphify extract ./docs              # extração headless, útil em CI (requer API key configurada)
graphify hook install                # cria git hook para reconstruir o grafo a cada commit
graphify prs                         # dashboard de PRs: estado de CI, revisão e mapeamento de worktree
```

## Desinstalar

```bash
graphify claude uninstall
graphify copilot uninstall
```
