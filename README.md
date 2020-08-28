# CODIGO-FONTE DE 2007

# extratificacao-de-dados

Código usado para fazer estatística de pesquisa de opinião publica. Desenvolvido em ASP Clássico para rodar em formato de Intranet, com o objetivo de ler um arquivo com colunas e linhas dinamicas para gerar um arquivo WORD com o resultado.

## Como funcionava?

1. Eram impresso milhares de cedula de pesquisa;
2. Diversos coletores de dados andavam todos os bairros da cidade para registrar as opniões (processo no qual também participei)
3. Alguem digitava todos os votos em uma planilha de excel (Normalmente era eu)
4. Rodava esse sistema, que era um site "intranet", onde digitava uma senha, colocava o arquivo de excel, digitava todos os campos da cedula de votação e o resultado era um arquivo WORD com a extratificação da pesquisa.

## Como instalar

Precisa rodar no Windows XP com IIS 2006
Colocar um arquivo de excel na pasta  "planilha/velha"
Criar um banco MySQL com: UID=root; PWD=luisrevolution2007;
Rodar o IIS e abrir o site localhost;
No browser digitar user "luis" e senha "luisdemarchi123";