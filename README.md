# Comparador

Algoritmo que compara duas tabelas com base nas colunas selecionadas e exibe: <br />
Linhas excluídas da primeira tabela;<br />
Linhas adicionadas à segunda tabela;<br />
Linhas que possuem as colunas selecionadas iguais porem com outros campos diferentes.<br />
<br />
<br />


## ACCESS


Um programa em python que compara as tabelas contidas em arquivos Access (.accdb) ou Excel (.xlsx), exibindo os resultados em tabelas na própria interface.<br />
O programa suporta mudar as colunas usadas para a comparação e a ordem das colunas nas tabelas, também sendo possível utilizar arquivos com múltiplas tabelas.<br />
Não é necessária a instalação de nenhuma das ferramentas da Microsoft para o correto funcionamento do algoritmo, uma vez que é utilizada a ferramenta mdbtools que extrai os dados do arquivo diretamente.<br />
Da maneira que está configurado, o algoritmo utiliza dos executáveis da pasta 'mdbtools' contida nesse repositório, tais arquivos foram compilados para utilização em Windows e não foram testados em outros sistemas operacionais.<br />
O programa é capaz de exportar as tabelas carregadas e o relatório dos resultados para arquivos Excel (.xlsx) utilizando o menu exportar, caso o Excel esteja instalado na máquina o arquivo abre automaticamente. <br />
Para ajuda de como o programa funciona existe um tutorial presente no menu 'ajuda'.<br />
No menu opções existe a possibilidade de mudar a cor das células que possuem valores discrepantes e a localização das ocorrências nas abas das tabelas originais, entretanto essa opção aumenta consideravelmente o tempo de compilação (principalmente em tabelas grandes com muitas ocorrências) devido a iteração nos dataframes.
Para inúmeras ocorrências existe um bug no visualizador do pacote pandastable onde as células são marcadas incorretamente após realizar uma rolagem pela tabela.<br /

### Pacotes necessários:
`$ pip install subprocess`<br />
`$ pip install pandas`<br />
`$ pip install tkinter`<br />
`$ pip install openpyxl`<br />
`$ pip install pandastable`<br />
