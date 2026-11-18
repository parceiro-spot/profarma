# Dashboard Profarma — versão provisória

Aplicação estática de teste inspirada nas páginas **Indicadores**, **Apontamento**, **Motoristas**, **Viagens** e **Cliente** do relatório de referência. Inclui filtros, período manual, gráficos e tabelas interativos, ordenação e organização de planilhas por mês.

## Teste no Preview

Abra o Preview do projeto e use **Carregar planilhas** para selecionar um arquivo XLSX, XLS ou CSV. Os dados são lidos no navegador e armazenados no IndexedDB deste perfil até serem excluídos no painel de organização ou até os dados do site serem limpos. O dashboard não envia o arquivo a um servidor da aplicação; carregue planilhas apenas em um dispositivo/perfil autorizado. O IndexedDB local não é criptografado e não deve ser tratado como mecanismo de confidencialidade.

Clique em um ponto de gráfico ou linha de tabela para cruzar a seleção; use **Limpar seleção** para removê-la. O intervalo inicial/final pode ser alterado manualmente. A engrenagem no canto superior abre as pastas mensais, permite visualizar, ativar/desativar ou excluir cada planilha individual e baixar cada mês como um único arquivo `.xlsx` formatado, com a aba `itineraries`.

O filtro **Segmento** usa a coluna `SEGMENTO` da planilha quando disponível; na ausência dela, mantém a classificação provisória existente. Os slicers preservam valores vazios como **(Em branco)**. A lista de **Zona Prioridade** contempla Agendamento, Dist Baixada, Dist Zona Oeste, Dist Zona Sul, DMIL Rua e DMIL Shopping quando as prioridades correspondentes estiverem presentes nos dados carregados.

## Privacidade e limites de equivalência

A planilha e o PBIX de referência não estão no projeto nem no ZIP. Para não embutir dados internos, o código não inclui códigos de clientes nem dicionários de nomes legais de transportadoras/rotas; os rótulos correspondentes vêm da planilha que o usuário carregar localmente. Assim, esta versão é provisória e os resultados podem diferir do Power BI. A Aderência é calculada como `OK / Total de NFs` após os filtros implementados; a medida DAX do PBIX não foi confirmada, e a divergência de novembro (96,53% vs. 96,47%) permanece sem reconciliação.

## Abrir uma cópia local

Sirva esta pasta por HTTP e abra `index.html`. A página carrega `app.js`, `app.css`, `report-spec.json`, o logo em `assets/` e a biblioteca SheetJS por CDN. Abrir diretamente como `file://` pode impedir o carregamento desses recursos.
