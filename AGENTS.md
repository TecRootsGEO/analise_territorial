# Regras Globais de Desenvolvimento de IA

Essas regras definem como a IA deve se comportar em todos os projetos, ambientes de trabalho e gerações de código. Elas são obrigatórias e sempre ativas.

---

## 1. Idioma e Comunicação

- **Idioma obrigatório:** Responder, planejar, comentar o código e redigir documentação sempre em **Português do Brasil** (mantendo termos técnicos padrão em inglês quando apropriado).
- **Objetividade:** Ser direto, claro e focado na solução técnica sem rodeios.

---

## 2. Princípios Fundamentais de Engenharia

- Priorizar simplicidade em vez de complexidade.
- Preferir manutenibilidade em vez de engenhosidade.
- Evitar abstrações desnecessárias ou superengenharia.
- Reutilizar padrões existentes antes de introduzir novos.
- Sempre assumir requisitos de qualidade de nível de produção.
- Nunca quebrar funcionalidades existentes, a menos que solicitado explicitamente.

---

## 3. Padrões de Qualidade de Código

- Escrever código limpo, legível e previsível.
- Usar nomes significativos e descritivos para variáveis e funções.
- Manter funções pequenas e focadas em uma única responsabilidade.
- Evitar duplicação (princípio DRY), mas não super-abstrair.
- Preferir lógica explícita em vez de comportamento mágico oculto.
- Remover código morto ou não utilizado assim que identificado.
- Preferir operações vetorizadas (Pandas/NumPy/GeoPandas) em vez de loops manuais em dados tabulares ou espaciais.

---

## 4. Arquitetura, Pipelines e Backend

- Seguir os padrões de arquitetura existentes no projeto.
- **Separação de responsabilidades:** Em aplicações web, separar rotas, serviços e acesso a dados. Em pipelines de dados e dashboards, separar claramente **ingestão/leitura**, **processamento/regras de negócio**, **exportação** e **camada de visualização**.
- Validar todos os dados de entrada antes de processá-los.
- Use o servidor MCP `context7` (se estiver disponível e ativo no ambiente) para consultar documentação atualizada de bibliotecas antes de gerar código.

---

## 5. Regras de UI, Dashboards e Frontend

- Sempre seguir os sistemas de design e temas existentes no projeto.
- Reutilizar componentes de UI existentes e manter espaçamento, layout e alinhamento consistentes.
- Priorizar a legibilidade e fluidez em telas **desktop** para dashboards analíticos, mapas e tabelas complexas, garantindo que o layout não quebre em diferentes resoluções de monitor.
- Garantir estruturas de resposta e estados visuais previsíveis (ex: indicadores de carregamento para operações pesadas).

---

## 6. Regras de Segurança e Proteção de Contexto

- Nunca confiar na entrada do usuário (sempre validar e higienizar).
- Proteger contra ataques de injeção (SQL, NoSQL, etc.).
- Nunca expor dados sensíveis, chaves ou credenciais no código, logs ou respostas.
- Usar variáveis de ambiente (`.env`) para todos os segredos e credenciais.
- **Proteção contra arquivos pesados:** Nunca carregar arquivos binários ou datasets volumosos inteiros (`.tif`, `.gpkg`, `.shp`, `.parquet`, `.csv`, `.xlsx`) diretamente no contexto do chat. Ler apenas metadados, esquemas ou pequenas amostras (`.head()`).

---

## 7. Regras de Desempenho e Dados

- Evitar computações e re-renderizações desnecessárias (utilizar cache em dashboards quando aplicável).
- Otimizar consultas ao banco de dados e evitar carregamento integral de grandes conjuntos de dados na memória sem necessidade.
- Garantir fluxo de dados previsível e manter o estado mínimo e localizado.

---

## 8. Tratamento de Erros e Confiabilidade

- Sempre tratar erros explicitamente (especialmente falhas de I/O, arquivos ausentes, projeções incompatíveis ou dados nulos).
- Nunca permitir exceções silenciosas ou não tratadas em código de produção.
- Fornecer mensagens de erro claras e acionáveis sem expor informações sensíveis.

---

## 9. Padrões de Documentação

- **Comentários de código:** Usar comentários apenas para explicar **por que** uma decisão foi tomada (regras de negócio, casos extremos, ajustes de projeção), nunca **o que** o código faz de forma óbvia.
- **Docstrings de funções:** Nos módulos principais de produção (`src/`), documentar claramente: Propósito, Entradas, Saídas e Efeitos colaterais. (Dispensável em scripts temporários de exploração).
- **Arquitetura e APIs:** Documentar fluxo de dados, endpoints (se houver) e trade-offs em recursos complexos.
- **README.md:** Todo projeto deve manter atualizado: configuração do ambiente, variáveis de ambiente, pilha tecnológica, estrutura de pastas e instruções de execução.

---

## 10. Regras de Comportamento da IA (STRICT)

- Sempre seguir as convenções de projeto existentes primeiro.
- Não introduzir nova arquitetura ou novas dependências sem instrução explícita.
- Se várias soluções existirem, escolher a mais simples e confiável.
- Nunca presumir requisitos ausentes — perguntar ao usuário ou escolher o padrão mais seguro.
- Combinar o estilo de código existente exatamente.
- Não refatorar código não relacionado durante a implementação de recursos.
- Preservar a lógica existente, a menos que a mudança seja solicitada explicitamente.

---

## 11. Limitações Rigorosas de Controle de Versão

- A IA **NÃO** deve executar nenhuma operação Git.
- A IA **NÃO** deve sugerir commits, pushes, merges ou gerenciamento de ramificações.
- A IA **NÃO** deve modificar o estado de controle de versão de nenhuma maneira.
- Todas as ações de controle de versão são tratadas manualmente apenas pelo desenvolvedor.

---

## 12. Regra de Ouro

Se houver incerteza:

- → Escolher a solução mais segura, padrão e mantível.
- → Nunca otimizar a engenhosidade em detrimento da clareza.
