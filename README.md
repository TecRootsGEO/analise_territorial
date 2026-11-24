# Análise de sobreposição entre glebas e Camada Base

O script `analise_sobreposicao.py` cruza glebas em **CSV (WKT), GeoPackage (`.gpkg`), GeoJSON (`.geojson`/`.json`) ou GeoParquet (`.parquet`)** com uma ou mais camadas de referência (`Camada Base`) em **GeoPackage, GeoJSON ou GeoParquet** e, opcionalmente, uma ou mais tabelas auxiliares em **CSV ou Parquet** pela coluna-chave detectada (como `REF_BACEN`, `COD_IMOVEL`, `CD_CAR` ou outra informada via `--coluna-ref`). Os resultados são consolidados no Excel com as abas `Sobreposicoes`, `Intersecoes_Detalhe` (1:N gleba × feição) e `Resumo`.

## Instalação

Recomenda-se Python 3.10 ou superior, em um ambiente virtual:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Execução

Exemplo com os arquivos encontrados na pasta `dados` deste projeto:

```powershell
python .\analise_sobreposicao.py `
  --glebas ".\dados\sicor_glebas_wkt_2024.csv" `
  --operacoes ".\dados\SICOR_OPERACAO_BASICA_CONTRAT.csv" `
  --terra-indigena ".\dados\TI_Uru_Eu_Wau_Wau 1.gpkg" `
  --terra-indigena ".\dados\TI_Governador 1.gpkg" `
  --saida ".\resultado\glebas_sobrepostas.xlsx" `
  --saida-gpkg ".\resultado\glebas_sobrepostas.gpkg"
```

Repita `--glebas` para analisar e consolidar vários arquivos (por exemplo `--glebas ".\dados\glebas_1.csv" --glebas ".\dados\glebas_2.parquet"`). Arquivos CSV devem usar o CRS indicado por `--crs-glebas`; arquivos GeoPackage, GeoJSON e GeoParquet usam seu CRS embutido quando presente.

O argumento `--operacoes` é opcional (é possível cruzar apenas glebas × Camada Base). Para acrescentar outras tabelas CSV ou Parquet, repita `--operacoes` com cada caminho:

```powershell
python .\analise_sobreposicao.py `
  --glebas ".\dados\sicor_glebas_wkt_2024.csv" `
  --operacoes ".\dados\operacoes_1.csv" `
  --operacoes ".\dados\operacoes_2.parquet" `
  --terra-indigena ".\dados\TI_Uru_Eu_Wau_Wau 1.gpkg" `
  --area-minima-ha 0.01 `
  --percentual-minimo 0.1 `
  --saida ".\resultado\glebas_sobrepostas.xlsx"
```

O script detecta o delimitador, a coluna geométrica (WKT/WKB) e a coluna-chave comum entre glebas e tabelas auxiliares (priorizando `REF_BACEN`, `ID_GLEBA`, `CD_CAR`, `COD_IMOVEL`, etc., ou aceitando definição explícita via `--coluna-ref` e `--coluna-geometria`). O CRS de origem padrão das coordenadas WKT é `EPSG:4674`. Para áreas, o padrão `--crs-area auto` analisa a extensão longitudinal da Camada Base: até **6° de longitude** seleciona automaticamente a zona **SIRGAS 2000 / UTM** correspondente; acima de **6°**, utiliza projeção equivalente **SIRGAS 2000 / Brazil Albers (`EPSG:10857`)**, evitando distorções em análises multiestado ou nacionais. Os parâmetros opcionais `--area-minima-ha` e `--percentual-minimo` descartam sobreposições residuais (*slivers*).

### Interface gráfica

Para usar a interface web conectada ao processamento Python, instale as dependências e inicie o serviço:

```powershell
python -m pip install -r requirements.txt
python .\web_app.py
```

Abra `http://127.0.0.1:8000` no navegador. A página envia os arquivos selecionados ao serviço local, que executa a mesma função de análise do script, fornece logs de progresso e disponibiliza os resultados interativos e downloads. Assim que arquivos são adicionados, a interface realiza uma pré-inspeção leve (`POST /api/inspect` com fatia inicial dos CSVs) para sugerir colunas no mapeamento avançado e antecipar o plano de consolidação (`UNION` vs. `JOIN`). O serviço aceita múltiplos arquivos de glebas (`.csv`, `.gpkg`, `.geojson`, `.parquet`), tabelas auxiliares opcionais (`.csv`, `.parquet`) e Camada Base (`.gpkg`, `.geojson`, `.parquet`), além de tolerância mínima de área/percentual e mapeamento manual opcional (`--coluna-ref`, `--coluna-geometria`, `--camada-ti`, `--coluna-nome-ti`, `--coluna-valor`). Por padrão, limita o total recebido por análise a 10 GiB (`GEO_MAX_UPLOAD_BYTES`) e processa uma análise por vez.

Na interface web, os quadros 1, 2 (opcional) e 3 ficam lado a lado com listas roláveis. Após a análise, os indicadores de resumo aparecem acima do quadro 5, que exibe um mapa Leaflet (`preferCanvas`) com troca de mapa base (`Satélite ESRI` ou `OpenStreetMap`) e controle de camadas para **Camada Base**, **Glebas com sobreposição** e **Recorte de interseção** (geometria exata da interseção 1:N), além de popups ricos com coordenadas do centroide para cópia rápida.

Na web, apenas o banco `dashboard.sqlite` é gerado de forma síncrona para liberar o mapa e a tabela rapidamente; o **Excel completo** (`resultado_completo.xlsx`) é preparado em segundo plano, enquanto as exportações sob demanda respeitam os filtros ativos, incluindo **Excel filtrado**, **CSV (`;` em UTF-8 BOM)**, **Parquet tabular**, **GeoParquet 1.1 (WKB + metadado `geo`)** e **GeoPackage com duas camadas (`glebas_sobrepostas` e `intersecoes`)**. O mapa acompanha os filtros e exibe até 500 feições por camada sem simplificar geometrias; o foco de linhas da tabela já presentes no mapa ocorre sem recarregar camadas.

A seção tabular oferece duas abas sincronizadas com os mesmos filtros (**Visão por Registro / Gleba** e **Visão 1:N (Gleba × Camada Base)**), seletor de colunas visíveis, busca textual livre, filtro por Camada Base, filtros analíticos (**área mínima**, **percentual mínimo** e **status de correspondência** quando há tabelas auxiliares), filtros SICOR e até 4 filtros categóricos automáticos para tabelas genéricas. Quando existir `CNPJ_IF` e `dados/SICOR_LISTA_IFS.csv` estiver presente, o nome da instituição financeira é enriquecido automaticamente. Na interface web (tabela, popup e resumo), a área sobreposta e o percentual são exibidos com até 4 casas decimais (`< 0,0001` para sobreposições positivas muito pequenas), enquanto os arquivos exportados e o SQLite preservam a precisão integral.

O serviço é vinculado exclusivamente a `127.0.0.1:8000` e grava arquivos temporários sob `.web_jobs`; os arquivos de entrada são removidos ao fim da análise e as saídas ficam disponíveis para download enquanto o serviço estiver ativo. Use a interface web por esse endereço, não abrindo `interface.html` diretamente como arquivo.

A interface gráfica Tkinter também pode ser aberta com:

```powershell
python .\interface_grafica.py
```

Opções comuns:

```powershell
python .\analise_sobreposicao.py --help
```

Repita `--operacoes` para selecionar várias tabelas auxiliares (CSV ou Parquet). Tabelas de mesmo esquema (por exemplo, anos diferentes de operação básica) são empilhadas verticalmente (`UNION`), enquanto tabelas complementares (como complemento da operação, mutuários e propriedades) são consolidadas horizontalmente pela coluna-chave (e `NU_ORDEM` quando compartilhada) na mesma linha (`JOIN`). Colunas com o mesmo nome dentro de um mesmo esquema são alinhadas; `ARQUIVO_OPERACOES` e `ARQUIVO_GLEBAS` identificam a origem dos registros. Cabeçalhos duplicados dentro de um arquivo são reportados como erro.

Use `--saida-gpkg caminho.gpkg` para gravar as camadas `glebas_sobrepostas` (polígonos originais sobrepostos e atributos consolidados) e `intersecoes` (geometria recortada de cada par gleba × Camada Base) no CRS de cálculo de área (`--crs-area`).

A detecção automática de codificação dos CSVs valida os primeiros 4 MiB de cada arquivo (UTF-8 vs. `cp1252`/`latin-1`) e aplica fallback automático em fluxo para `cp1252` caso um caractere incompatível apareça mais adiante. A leitura e agregação das tabelas auxiliares usam streaming e agrupamento nativo em Polars.

## Lógica de processamento

1. Valida arquivos e inspeciona esquemas/colunas sem carregar tabelas inteiras na memória; normaliza cabeçalhos (incluindo `#` do SICOR).
2. Lê a Camada Base (`.gpkg`, `.geojson` ou `.parquet`), corrige polígonos inválidos e reprojeta para o CRS projetado (UTM até 6° ou Albers `EPSG:10857` acima de 6°), construindo índices espaciais reutilizáveis (`STRtree` com geometrias preparadas).
3. Lê glebas em blocos (CSV via Polars streaming com fallback UTF-8 → `cp1252`, Parquet/GeoParquet nativamente via Polars WKT/WKB, ou GPKG/GeoJSON fatiado via `pyogrio`) aplicando estratégia *filter-then-refine*.
4. Calcula interseções exatas e união sem dupla contagem por gleba, aplicando eventual tolerância de slivers (`--area-minima-ha`, `--percentual-minimo`) e preservando simultaneamente o detalhamento 1:N (`intersecoes`).
5. Filtra e agrega tabelas auxiliares nativamente em Polars pela coluna-chave antes da junção final.
6. Gera Excel em modo streaming (`Sobreposicoes`, `Intersecoes_Detalhe` e `Resumo`) e GeoPackage multi-camada na CLI/Tkinter, ou `dashboard.sqlite` com exportações sob demanda na interface web.

## Desempenho e observações

- Polars lê CSVs e Parquets em blocos com projeção de colunas e pushdown de filtro, mantendo baixo consumo de memória.
- O Excel suporta no máximo 1.048.576 linhas por aba; para volumes maiores na web, utilize os downloads em **CSV**, **Parquet** ou **GeoParquet** diretamente dos filtros ativos.
- O GeoParquet 1.1 é gerado nativamente em Polars (codificação WKB e metadado `geo` completo) sem exigir `pyarrow`.
#   a n a l i s e _ t e r r i t o r i a l  
 