# Plano — Dashboard Profarma para teste

## Objetivo e escopo

Entregar uma versão provisória, navegável e interativa do dashboard Profarma para o usuário testar, usando o dashboard do ZIP e o PBIX somente como referências visuais e funcionais. O relatório de referência contém cinco páginas: Indicadores, Apontamento, Motoristas, Viagens e Cliente. O foco é reproduzir suas áreas, métricas e interações principais — inclusive seleção cruzada em gráficos e tabelas — em uma aplicação web, sem exigir Power BI.

O código executável recebido dentro dos anexos não será iniciado nem copiado como código. Reutilizam-se apenas dados declarativos não sensíveis necessários à estrutura visual (páginas, visuais, dimensões e rótulos). Códigos de clientes e dicionários com nomes legais de transportadoras/rotas foram removidos do artefato; os rótulos relacionados virão da planilha que o usuário carregar localmente. A NOVEMBRO.xlsx e o BINOVO.pbix não serão incorporados ao site ou ao ZIP.

## Implementação e dados

- Criar uma aplicação estática em HTML/CSS/JavaScript, sem framework nem backend. O `index.html` novo carrega a aplicação nova e a configuração declarativa segura do layout; não executa o JavaScript recebido.
- Usar SheetJS via CDN para ler XLSX/CSV no navegador. Os arquivos carregados permanecem no IndexedDB do perfil do navegador e são organizados em pastas mensais; não são enviados pelo código do app. Os arquivos Excel mensais são gerados localmente como um único `.xlsx` por pasta, com as planilhas individuais preservadas para ações separadas. O IndexedDB não é criptografado.
- Manter as cinco páginas do BI, com indicadores, gráficos, tabelas, ordenação de colunas, slicers e seletor de período manual. Implementar seleção cruzada dos pontos dos gráficos e linhas das tabelas, com estado visível e limpeza explícita da seleção.
- Aplicar apenas as regras de preparação atualmente descritas no código, sem listas privadas de códigos de clientes ou nomes de transportadoras/rotas. Isso pode produzir resultados diferentes do Power BI. A diferença de Aderência de novembro não será resolvida sem confirmação da definição DAX e dos filtros do PBIX; não mudar a fórmula por palpite.
- Para Segmento, priorizar uma coluna explícita `SEGMENTO` (Distribuição/Varejo) quando existir; sem ela, manter a classificação provisória já existente. Slicers apresentam valores vazios como “(Em branco)”. Zona Prioridade deriva dos códigos de prioridade tratados pela função de mapeamento no código e inclui as categorias da referência quando presentes nos dados.
- Usar o intervalo total encontrado como padrão; o usuário pode selecionar as datas inicial e final manualmente. Não incluir feriados nem ajuste automático de semanas.
- Em telas menores, manter escala legível com rolagem horizontal; manter navegação centralizada no rodapé e engrenagem isolada no canto superior.

## Decisões de design

**Movimento:** ferramenta operacional de BI corporativo, baseada na linguagem visual do relatório Power BI fornecido.

**Princípios:** (1) preservar a hierarquia e densidade úteis do BI; (2) verde Profarma para orientação e estado ativo, vermelho apenas para desvios; (3) tornar filtros e seleções reversíveis; (4) não ocultar dados ou regras de negócio inventadas.

**Filosofia de cor:** branco e cinzas neutros mantêm a leitura de operação; o verde `#006440` conecta a interface ao relatório original e indica seleção/resultado favorável; vermelho `#D64550` destaca exceções.

**Paradigma de layout:** páginas analíticas em canvas de proporção fixa semelhante ao relatório original, redimensionadas para desktop e roláveis horizontalmente no celular; filtros em coluna lateral conforme o BI e navegação centralizada no rodapé.

**Elementos de assinatura:** cartões numéricos compactos; cabeçalhos de slicers verdes; tabelas densas com totais e realces de dados.

**Interação:** filtros de seleção múltipla, período manual, ordenação por cabeçalho, seleção de gráfico/linha cruzando para os demais visuais e controle de limpar seleção. Feedback deve deixar claro quando há filtro ativo.

**Animação:** transições curtas (até 180 ms) apenas para estados de foco/hover e abertura de painel; sem animações contínuas que prejudiquem a leitura operacional.

**Tipografia:** Segoe UI do relatório de origem, fallback para Arial/system sans-serif. Valores de KPI têm maior peso; rótulos e cabeçalhos permanecem compactos.

**Essência da marca:** painel de operação Profarma para leitura rápida de desempenho de transporte; personalidade objetiva, confiável e analítica.

**Voz:** clara e funcional. Exemplos: “Selecione um ponto do gráfico para filtrar as demais áreas.” “O período selecionado pode ser ajustado manualmente.”

**Wordmark/logo:** reaproveitar apenas o logo estático incorporado no dashboard de origem, sem substituir por marca genérica; não executar código que o carregue.

**Cor de assinatura:** verde Profarma `#006440`.

## Estrutura do projeto

- `index.html` — estrutura da aplicação e carregamento dos recursos novos.
- `app.css` — estilos novos inspirados no relatório de referência.
- `app.js` — renderização, cálculos, filtros, interações, leitura local de arquivos e criação de ZIP.
- `report-spec.json` — configuração declarativa não sensível das páginas, visuais, dimensões e campos.
- `manus-routes.json` — declaração da rota `/` para Preview.
- `dados/arquivos.json` — manifesto vazio, sem dados da planilha privada.
- `profarma-logo.png` — logo estático extraído como imagem, sem executar conteúdo embutido.
- `README.md` — instruções de uso/teste e aviso de privacidade/limites de equivalência.
- `TODO.md` — estado dos resultados acordados.

## Limites de equivalência

A meta é reproduzir fielmente a aparência e as interações principais registradas no PBIX. Power BI usa um motor de modelo/DAX que não está disponível no dashboard estático. A regra DAX de Aderência e a divergência de novembro (96,53% vs. 96,47%) permanecem sem confirmação. Para evitar distribuir dados internos, a versão omite listas privadas de exclusão e dicionários de nomes; por isso, não deve ser apresentada como equivalente ao PBIX nem como resultado reconciliado.
