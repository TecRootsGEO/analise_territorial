# Guia de Estilos e Design System: Plataforma Geo e Inovação

Este documento serve como referência de design para a construção de novas páginas e componentes dentro da Plataforma Geo, garantindo consistência visual, hierarquia e harmonia em todo o sistema.

---

## 1. Paleta de Cores

As cores da plataforma foram definidas seguindo a identidade corporativa, com funções específicas para cada tom:

### Cores Principais

- **Azul Institucional** (`#262DA7`): Cor principal da identidade.
  - **Uso**: Header (Topbar), Footer, Títulos (`h2`, `h3`), Navegação e Elementos estruturais.
  - **Variável CSS**: `--azul-inst`
- **Amarelo Destaque** (`#FAFD00`): Cor de chamada de ação. Deve ser usada com parcimônia.
  - **Uso**: Botões primários ("Entrar", "Upload e Buscar", "Começar agora"), Alertas importantes.
  - **Variável CSS**: `--amarelo-destaque`

### Cores Interativas

- **Azul Elétrico** (`#455AFF`): Cor de interação e destaque dinâmico.
  - **Uso**: Links, Botões secundários (Ghost buttons), Seleções, Tabs, Estados ativos (hover/focus).
  - **Variável CSS**: `--azul-eletrico`

### Cores de Status e Contexto

- **Verde Institucional** (`#01781E`):
  - **Uso**: Mensagens de Sucesso, Status "Ativo", Indicadores positivos, Sustentabilidade e Dados relacionados à agricultura.
  - **Variável CSS**: `--verde-inst`
- **Azul Céu** (`#74D5F8`):
  - **Uso**: Alertas de informação, Dados geoespaciais, Cards informativos e Elementos relacionados a mapas.
  - **Variável CSS**: `--azul-ceu`

### Cores de Apoio e Temáticas

- **Turquesa** (`#6DDECF`):
  - **Uso**: Ícones, Elementos decorativos (bordas de painéis), Mapas e Áreas de apoio.
  - **Variável CSS**: `--turquesa`
- **Roxo** (`#735CC4`):
  - **Uso**: Funcionalidades de Inteligência Artificial, Análises avançadas, Tecnologia e Categorias específicas de gráficos.
  - **Variável CSS**: `--roxo`

### Cores Base (Neutras)

- **Texto Principal** (`#333333`): `--ink`
- **Texto Secundário/Muted** (`#666666`): `--muted`
- **Fundo de Painéis/Cards** (`#ffffff`): `--paper`
- **Fundo da Página**: `#f4f6f8`

---

## 2. Tipografia

O projeto utiliza a família de fontes proprietária **BancoDoBrasilTextos**, localizada no diretório `/fonts`.

- **Regular** (`font-weight: 400`): Textos corridos, parágrafos, labels de formulários e descrições genéricas. (Arquivo: `BancoDoBrasilTextos-Regular.ttf`)
- **Medium** (`font-weight: 500`): Botões, destaques intermediários e subtítulos menores. (Arquivo: `BancoDoBrasilTextos-Medium.ttf`)
- **Bold** (`font-weight: 700`): Títulos principais, dados em destaque (ex: números grandes de estatísticas) e texto de botões primários. (Arquivo: `BancoDoBrasilTextos-Bold.ttf`)

---

## 3. Elementos Estruturais (Layout)

A estrutura base das páginas HTML deve seguir o encapsulamento padrão `.shell` que centraliza o conteúdo.

### Grid e Espaçamentos

- **Grid de 2 Colunas** (`.container-2col`): Divide a tela em duas partes iguais (1fr 1fr), com distanciamento (`gap`) de 24px. Ideal para separar Configuração/Filtros dos Resultados. Em mobile, empilha automaticamente (1fr).
- **Painéis / Cards** (`.panel`):
  - Fundo branco (`--paper`), bordas arredondadas de 16px (`--radius-lg`), *padding* interno de 24px e sombra leve.
  - Usado para agrupar conteúdos logicamente.
- **Full-Width** (`.full-width`): Classe utilitária para painéis que precisam ocupar 100% da largura dentro de um grid.

---

## 4. Componentes Globais

### Cabeçalho (Topbar)

- Possui fundo `--azul-inst` com textos brancos.
- **Logo**: O logo (`image/geo.png`) possui altura fixa de `96px` e um pequeno raio de borda, acompanhado pelo título "Plataforma Geo e Inovação" (fonte 1.8rem Bold) e o subtítulo descritivo (1.1rem na cor Azul Céu).
- Fica dentro do encapsulamento `.shell` mas destaca-se por seu preenchimento (`padding`) e sombra.

### Botões

- **Primário** (`.button-primary`):
  - Fundo `--amarelo-destaque`, texto `--azul-inst`, `font-weight: 700`.
  - Efeito de *hover*: Elevação (-2px), Sombra e leve escurecimento do amarelo.
- **Secundário / Ghost** (`.button-ghost`):
  - Sem fundo, borda de 2px sólida na cor `--azul-eletrico`, texto `--azul-eletrico`, `font-weight: 500`.
  - Efeito de *hover*: Fundo levemente translúcido na mesma cor elétrica.

### Formulários e Inputs

- Todos os inputs (text, file, date, select) possuem borda neutra (`--line`), fundo branco e `border-radius: 8px`.
- O estado de **Focus** recebe contorno (outline) e box-shadow usando a cor `--azul-eletrico`, provendo excelente feedback visual.

### Alertas (Notificações)

Caixas de alerta (`.alert`) para avisos na tela:

- **Sucesso** (`.alert.success`): Fundo verde translúcido, borda lateral `--verde-inst`.
- **Informação** (`.alert.info`): Fundo azul translúcido, borda lateral `--azul-ceu`.
- **Erro** (`.alert.error`): Tonalidade de vermelho clássico para destacar bloqueios ou falhas.

---

## 5. Boas Práticas e Recomendações

1. **Reutilização do CSS**: Todas essas regras já estão declaradas no arquivo `/web/assets/styles.css`. Sempre prefira utilizar as classes pré-existentes (`.panel`, `.button-primary`, `.form-group`) ao invés de criar novos estilos *inline*.
2. **Ícones e Logos**: Mantenha o padrão de imagens arredondadas levemente (4px a 8px) com fundo transparente ou branco para garantir o contraste no cabeçalho ou elementos institucionais.
3. **Escalabilidade**: Sempre utilize variáveis do CSS (ex: `var(--azul-inst)`) para novas customizações, nunca repita o valor `#HEX` direto nas folhas de estilo de novas páginas. Isso garante que futuras alterações de *branding* ocorram em apenas um lugar.
