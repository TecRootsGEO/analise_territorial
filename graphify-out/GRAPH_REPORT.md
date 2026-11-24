# Graph Report - . (2026-11-20)

## Corpus Check

- Large corpus: 32 files · ~993,305 words. Semantic extraction will be expensive (many Claude tokens). Consider running on a subfolder.

## Summary

- 297 nodes · 771 edges · 13 communities
- Extraction: 75% EXTRACTED · 25% INFERRED · 0% AMBIGUOUS · INFERRED: 189 edges (avg confidence: 0.88)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)

- [Identidade Visual e Plataforma Geo](#community-0---identidade-visual-e-plataforma-geo)
- [Manual e Esquema de Dados SICOR](#community-1---manual-e-esquema-de-dados-sicor)
- [Orquestração Web de Jobs e Saídas](#community-2---orquestração-web-de-jobs-e-saídas)
- [Dashboard SQLite, Filtros e Guias](#community-3---dashboard-sqlite-filtros-e-guias)
- [Motor de Sobreposição e Exportação](#community-4---motor-de-sobreposição-e-exportação)
- [Leitura Streaming Polars e CRS](#community-5---leitura-streaming-polars-e-crs)
- [Interface Gráfica Desktop Tkinter](#community-6---interface-gráfica-desktop-tkinter)
- [Upload Multipart Streaming HTTP](#community-7---upload-multipart-streaming-http)
- [Cruzamento e Normalização de Operações](#community-8---cruzamento-e-normalização-de-operações)
- [Dependências Python Geoespaciais](#community-9---dependências-python-geoespaciais)
- [Painéis Power BI Agropecuários](#community-10---painéis-power-bi-agropecuários)
- [Mapa Leaflet e Tokens CSS](#community-11---mapa-leaflet-e-tokens-css)
- [Saneamento Geométrico de Polígonos](#community-12---saneamento-geométrico-de-polígonos)

## God Nodes (most connected - your core abstractions)

1. `sicor_operacao_basica_estado (Operações Básicas do SICOR)` - 39 edges
2. `Plataforma GeoInovação` - 37 edges
3. `Manual de Dados do SICOR Completo (Banco Central do Brasil)` - 37 edges
4. `run()` - 29 edges
5. `guia_estilo.md (Style Guide & Design System)` - 27 edges
6. `AnalysisWindow` - 19 edges
7. `read_matching_operations()` - 18 edges
8. `SICOR REF_BACEN` - 18 edges
9. `str` - 17 edges
10. `Path` - 16 edges

## Surprising Connections (you probably didn't know these)

- `_create_filtered_dashboard_excel()` --semantically_similar_to--> `write_excel()`  [INFERRED] [semantically similar]
  web_app.py → analise_sobreposicao.py
- `Chunked Polars + Spatial Index Overlap Pipeline` --rationale_for--> `run()`  [INFERRED]
  README.md → analise_sobreposicao.py
- `pyogrio (>=0.9)` --conceptually_related_to--> `glebas_sobrepostas.gpkg (Overlapping Plots GeoPackage)`  [INFERRED]
  requirements.txt → analise_sobreposicao.py
- `openpyxl (>=3.1)` --conceptually_related_to--> `resultado.xlsx / resultado_filtrado.xlsx (Excel Report)`  [INFERRED]
  requirements.txt → analise_sobreposicao.py
- `Operação Básica Contratada (SICOR_OPERACAO_BASICA_CONTRAT)` --shares_data_with--> `SICOR REF_BACEN`  [INFERRED]
  dados/manualDadosSicorCompleto.pdf → analise_sobreposicao.py

## Hyperedges (group relationships)

- **Streaming Spatial Overlap & SICOR Join Pipeline** — analise_sobreposicao_run, analise_sobreposicao_automatic_area_crs, analise_sobreposicao_load_indigenous_lands, analise_sobreposicao_iter_csv_batches, analise_sobreposicao_clean_wkt_chunk, analise_sobreposicao_find_overlaps, analise_sobreposicao_read_matching_operations, camada_base, sicor_ref_bacen, sicor_gt_geometria, sicor_vl_parc_credito [EXTRACTED 1.00]
- **Web Multipart Upload & Asynchronous Analysis Execution Flow** — interface_html, web_app_requesthandler_do_post, web_app_parse_multipart, web_app_multipartreader, web_app_execute_job, analise_sobreposicao_run, analise_sobreposicao_write_dashboard_database, analise_sobreposicao_write_excel, analise_sobreposicao_write_geopackage, analise_sobreposicao_write_preview_geojson [EXTRACTED 1.00]
- **SQLite-Backed Interactive Dashboard, Map & Filtered Export Flow** — dashboard_sqlite, sicor_lista_ifs_csv, web_app_read_financial_institutions, web_app_requesthandler_dashboard, web_app_requesthandler_map, web_app_create_filtered_dashboard_excel, web_app_requesthandler_download, resultado_xlsx, sicor_cnpj_if, sicor_ref_bacen_efetivo [EXTRACTED 1.00]
- **End-to-End Web Analysis, Polling, Map & Filtered Dashboard Flow** — interface_html_start_analysis, web_app_requesthandler_do_post, interface_html_monitor_job, web_app_requesthandler_do_get, analise_sobreposicao_run, dashboard_sqlite, interface_html_render_dashboard, web_app_requesthandler_dashboard, interface_html_render_filtered_map, web_app_requesthandler_map [EXTRACTED 1.00]
- **Plataforma Geo Design System Brand Implementation** — plataforma_geo_inovacao, guia_estilo_md, guia_estilo_palette_tokens, guia_estilo_font_bb_textos, guia_estilo_layout_components, image_geo_png, interface_html [EXTRACTED 1.00]
- **Territorial Overlap & SICOR Credit Enrichment Pipeline** — analise_sobreposicao_py, camada_base, sicor_ref_bacen, sicor_gt_geometria, sicor_vl_parc_credito, sicor_cnpj_if, sicor_lista_ifs_csv, glebas_sobrepostas_gpkg, resultado_xlsx [EXTRACTED 1.00]
- **BCB SICOR Relational Microdata Core Schema** — manualdadossicorcompleto_pdf, sicor_operacao_basica_estado, sicor_complemento_operacao_basica, sicor_glebas_wkt, sicor_mutuarios, sicor_propriedades, sicor_lista_ifs_csv [EXTRACTED 1.00]
- **SICOR Spatial Overlap & REF_BACEN Join Pipeline** — sicor_glebas_wkt, sicor_gt_geometria, camada_base, sicor_ref_bacen, sicor_operacao_basica_estado, analise_sobreposicao_py [INFERRED 0.95]
- **SICOR Complementary & Financial Institution Dashboard Enrichment** — sicor_complemento_operacao_basica, sicor_mutuarios, sicor_lista_ifs_csv, sicor_ref_bacen_efetivo, sicor_cnpj_if, web_app_py [INFERRED 0.95]
- **Suite de Painéis Power BI — Plataforma GEO e Inovação** — image_painel1_png, image_painel2_png, image_painel3_png, plataforma_geo_inovacao, image_powerbi_logo_png, interface_html [INFERRED 0.85]
- **Fluxo Integrado de Valoração Rural (Imóvel + CAR + INCRA/SIGEF + Mapa)** — painel1_apoio_incremento_valoracao, painel1_tabela_imovel_valor, painel1_tabela_car, painel1_tabela_incra_sigef, painel1_mapa_geoespacial_brasil [EXTRACTED 1.00]
- **Fluxo de Monitoramento Remoto ATNI (Produções + Compatibilidade CAR + RL/APP)** — painel2_monitoramento_producao_remota, painel2_kpi_qtd_producoes, painel2_compatibilidade_car, painel2_area_preservacao_rl_app, painel2_carteira_atni_producao [EXTRACTED 1.00]
- **Plataforma Geo Multi-Resolution Favicon & PWA Icon Family** — image_favicon_16_png, image_favicon_32_png, image_apple_touch_icon_png, image_android_chrome_192_png, image_android_chrome_512_png [INFERRED 0.95]
- **Identidade Visual Corporativa Banco do Brasil para Plataforma Geo e Inovação** — image_logo_bb_png, image_geo_png, design_system_paleta_bb_geo, guia_estilo_md, interface_html, plataforma_geo_inovacao [INFERRED 0.95]
- **Ativos Visuais InfoGEO (Badge Geométrico vs Ilustração Rural Duplicada)** — image_infogeo_png, image_infogeo_img_png, image_infogeo_old_png, brand_infogeo [INFERRED 0.85]
- **AgTech Cyber-Tree Background Lineage for Plataforma Geo e Inovação** — image_background_old_png, image_background_png, background_cyber_tree_emblem, background_digital_agriculture_motif, plataforma_geo_inovacao [INFERRED 0.85]
- **Visual Branding Evolution from Legacy InfoGEO to Plataforma Geo e Inovação** — image_infogeo_png, image_infogeo_old_png, image_infogeo_img_png, image_background_old_png, image_background_png, guia_estilo_md, interface_html [INFERRED 0.75]
- **Plataforma Geo Feature Module Card Visual Suite** — image_consulta_png, image_ferramentas_png, image_satelite_png, image_valoracao_png, image_parecer_geo_png, image_busca_car_png, interface_html, guia_estilo_md, plataforma_geo_inovacao [INFERRED 0.85]
- **Cadastral Query, CAR Overlap Inspection, Satellite Monitoring, and Geo Opinion Flow** — consulta_camadas_cadastrais, busca_car_deteccao_sobreposicao, satelite_sensoriamento_multiespectral, parecer_geo_relatorio_conformidade, plataforma_geo_inovacao [INFERRED 0.75]
- **Agro-Computational Analytics, Remote Sensing, and Rural Valuation Suite** — ferramentas_agrocomputacionais, satelite_sensoriamento_multiespectral, valoracao_imovel_rural_georreferenciada, plataforma_geo_inovacao [INFERRED 0.75]

## Communities (13 total, 0 thin omitted)

### Community 0 - "Identidade Visual e Plataforma Geo"

Cohesion: 0.11
Nodes (50): Cyber-Tree Emblem (Half-Circuit, Half-Leaf), Digital Agriculture Circuit-Furrow Motif, UI Light-Canvas Bottom Fade Treatment, Banco do Brasil (Identidade Corporativa), InfoGEO, Microsoft Power BI, CAR Parcel Search & Multi-Polygon Overlap Inspection, Multi-Layer Cadastral Query (CAR, SIGEF, INCRA) (+42 more)

### Community 1 - "Manual e Esquema de Dados SICOR"

Cohesion: 0.08
Nodes (47): SICOR — Sistema de Operações do Crédito Rural e do Proagro (BCB), Modelo Relacional de Cruzamento SICOR (REF_BACEN + NU_ORDEM), Manual de Dados do SICOR Completo (Banco Central do Brasil), CD_CAR (CAR do Proprietário), CD_CNPJ_CPF (CPF/CNPJ do Proprietário Rural), CD_CPF_CNPJ (CPF/CNPJ do Mutuário), CD_DAP (Código da DAP do Mutuário), CD_EMPREENDIMENTO (Código do Empreendimento) (+39 more)

### Community 2 - "Orquestração Web de Jobs e Saídas"

Cohesion: 0.10
Nodes (26): Any, BaseHTTPRequestHandler, glebas_sobrepostas.gpkg (Overlapping Plots GeoPackage), CSS Token --amarelo-destaque (#FAFD00), SimpleNamespace, Excel & GeoPackage Download Toggles (/api/jobs/{jobId}/download/{kind}), Job Polling Loop (monitorJob -> GET /api/jobs/{jobId}), Analysis Submission Handler (#start-button -> POST /api/analyze) (+18 more)

### Community 3 - "Dashboard SQLite, Filtros e Guias"

Cohesion: 0.12
Nodes (29): infer_delimiter(), Detecta UTF-8; para arquivos incompatíveis, escolhe Windows-1252 ou Latin-1., Detecta o delimitador pelo cabeçalho sem ler o arquivo inteiro., resolve_csv_encoding(), Caveman Mode Terse Response & Exception Policy, Connection, dashboard.sqlite (SQLite Dashboard & Map DB), Graphify Knowledge Graph CLI & Claude/Copilot Workflow (+21 more)

### Community 4 - "Motor de Sobreposição e Exportação"

Cohesion: 0.14
Nodes (28): clean_wkt_chunk(), find_overlaps(), main(), Spatial Index & Union Overlap Deduplication Rationale, parse_args(), int, Converte WKT e prepara polígonos válidos; retorna contagens de rejeição., Usa índice espacial para calcular área de interseção sem dupla contagem. (+20 more)

### Community 5 - "Leitura Streaming Polars e CRS"

Cohesion: 0.15
Nodes (22): automatic_area_crs(), encoding_is_valid(), is_utf8_encoding(), iter_csv_batches(), load_indigenous_lands(), polars_compatible_csv(), polars_encoding(), Path (+14 more)

### Community 6 - "Interface Gráfica Desktop Tkinter"

Cohesion: 0.14
Nodes (3): AnalysisWindow, main(), Tk

### Community 7 - "Upload Multipart Streaming HTTP"

Cohesion: 0.35
Nodes (8): BinaryIO, bytes, Incremental MultipartReader for <=10 GiB Uploads, MultipartReader, _parse_content_disposition(),_parse_multipart(), int, Leitor multipart incremental para aceitar uploads grandes sem carregá-los na RAM

### Community 8 - "Cruzamento e Normalização de Operações"

Cohesion: 0.22
Nodes (11): normalize_column_name(), parse_financial_values(), float, Resolve uma coluna sem diferenciar maiúsculas/minúsculas., Padroniza cabeçalhos, removendo BOM, espaços e o # do SICOR., Converte valores monetários em número, aceitando decimal com ponto ou vírgula., Lê operações em blocos e retém somente referências com sobreposição., read_matching_operations() (+3 more)

### Community 9 - "Dependências Python Geoespaciais"

Cohesion: 0.36
Nodes (9): requirements.txt (Python Package Dependencies), geopandas (>=1.0,<2.0), numpy (>=1.24), openpyxl (>=3.1), pandas (>=2.0,<4.0), polars (>=1.30,<2.0), pyogrio (>=0.9), shapely (>=2.0,<3.0) (+1 more)

### Community 10 - "Painéis Power BI Agropecuários"

Cohesion: 0.33
Nodes (10): Painel de Apoio Incremento/Valoração, Mapa Geoespacial de Imóveis no Brasil (TomTom/OSM), Tabela CAR (Cadastro Ambiental Rural), Tabelas de Imóvel, Valoração e Avaliador BB (MCI), Tabela INCRA / Parcela SIGEF Certificada, Tabela de Área de Preservação (RL, APP) por BB_CD_REG_CAR, Carteira ATNI (ATR) e Detalhamento de Produções por MCI, Indicadores de Compatibilidade CAR por Status e SG_UF (+2 more)

### Community 11 - "Mapa Leaflet e Tokens CSS"

Cohesion: 0.31
Nodes (9): CSS Token --azul-ceu (#74D5F8), CSS Token --azul-inst (#262DA7), CSS Token --roxo (#735CC4), CSS Token --turquesa (#6DDECF), CSS Token --verde-inst (#01781E), Corporate Color Palette & Neutral CSS Tokens, ESRI World Imagery Satellite Basemap Layer, Leaflet 1.9.4 Interactive Map (#map) (+1 more)

### Community 12 - "Saneamento Geométrico de Polígonos"

Cohesion: 0.50
Nodes (5): polygonal_geometry(), Retém somente partes poligonais de geometrias corrigidas., Corrige geometria inválida, descartando feições que não sejam polígonos., repair_polygon(), BaseGeometry

## Ambiguous Edges - Review These

- `interface.html` → `infogeo_old.png — Legacy InfoGEO Rural Aerial Landscape`  [AMBIGUOUS]
  image/infogeo_old.png · relation: conceptually_related_to

## Knowledge Gaps

- **20 isolated node(s):** `CRS`, `float`, `LogRecord`, `Connection`, `float` (+15 more)
  These have ≤1 connection - possible missing edges or undocumented components.

## Suggested Questions

Questions this graph is uniquely positioned to answer:

- **What is the exact relationship between `interface.html` and `infogeo_old.png — Legacy InfoGEO Rural Aerial Landscape`?**
  *Edge tagged AMBIGUOUS (relation: conceptually_related_to) - confidence is low.*
- **Why does `Plataforma GeoInovação` connect `Identidade Visual e Plataforma Geo` to `Dependências Python Geoespaciais`, `Dashboard SQLite, Filtros e Guias`, `Interface Gráfica Desktop Tkinter`?**
  *High betweenness centrality (0.231) - this node is a cross-community bridge.*
- **Why does `Manual de Dados do SICOR Completo (Banco Central do Brasil)` connect `Manual e Esquema de Dados SICOR` to `Dashboard SQLite, Filtros e Guias`?**
  *High betweenness centrality (0.116) - this node is a cross-community bridge.*
- **Why does `AnalysisWindow` connect `Interface Gráfica Desktop Tkinter` to `Orquestração Web de Jobs e Saídas`, `Motor de Sobreposição e Exportação`?**
  *High betweenness centrality (0.101) - this node is a cross-community bridge.*
- **Are the 2 inferred relationships involving `sicor_operacao_basica_estado (Operações Básicas do SICOR)` (e.g. with `analise_sobreposicao.py` and `Operação Básica Contratada (SICOR_OPERACAO_BASICA_CONTRAT)`) actually correct?**
  *`sicor_operacao_basica_estado (Operações Básicas do SICOR)` has 2 INFERRED edges - model-reasoned connections that need verification.*
- **Are the 31 inferred relationships involving `Plataforma GeoInovação` (e.g. with `analise_sobreposicao.py` and `Cyber-Tree Emblem (Half-Circuit, Half-Leaf)`) actually correct?**
  *`Plataforma GeoInovação` has 31 INFERRED edges - model-reasoned connections that need verification.*
- **Are the 22 inferred relationships involving `guia_estilo.md (Style Guide & Design System)` (e.g. with `UI Light-Canvas Bottom Fade Treatment` and `Caveman Mode Terse Response & Exception Policy`) actually correct?**
  *`guia_estilo.md (Style Guide & Design System)` has 22 INFERRED edges - model-reasoned connections that need verification.*
