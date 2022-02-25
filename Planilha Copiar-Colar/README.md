# Planilha Copiar-Colar
<p align="center">
  <img src="https://img.shields.io/static/v1?label=VBA&message=MsExcel&color=green&style=for-the-badge&logo=microsoftoffice"/>
  <img src="http://img.shields.io/static/v1?label=SIZE&message=19.4 KB&color=blue&style=for-the-badge"/>
  <img src="http://img.shields.io/static/v1?label=STATUS&message=CONCLUIDO&color=GREEN&style=for-the-badge"/>
</p>

 <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Planilha%20Copiar-Colar/Readme-repository/cover.jpg"></p>

### Sumário
🔹 [Descrição do projeto](#descrição-do-projeto)

🔹 [Pré-requisitos](#pré-requisitos)

🔹 [Rodando a aplicação](#rodando-a-aplicação)





## Descrição do projeto 
<p align="justify">
  Planilha em Excel que copia o conteúdo de todas as planilhas de uma determinada pasta e cola na principal.
</p>



## Pré-requisitos
* Ter o Microsoft Office Excel instalado



## Rodando a aplicação
Assim que abrir a planilha, localize os alertas de Edição e Macro e os habilite

  <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Planilha%20Copiar-Colar/Readme-repository/ativacaoexcel.png"></p>
  
Aperte as teclas <b>Alt+F8</b> para abrir o gerenciador de Macros
Selecione a opção Importar_XLS e clique na opção <b>Depurar</b>

 <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Planilha%20Copiar-Colar/Readme-repository/macroimportar.jpg"></p>
 
 Ao abrir o ambiente VBA, o código deverá ser editado conforme as linhas selecionadas
 
 <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Planilha%20Copiar-Colar/Readme-repository/codemacro.jpg" width=415></p>

* <b><u>sPath="C:\Users\Usuario\Desktop\NomedaPasta\"</u></b></br>
Deve ser alterada com o caminho onde as planilhas que serão lidas estão (o caminho deve sempre estar entre aspas e ter uma barra \ no final)

* <b>r = shPadrao.Cells(Rows.Count, <u>"B"</u>).End(xlUp).Row</b></br>
Começa a contar a partir da linha especificada, se a leitura deve começar a partir da primeira, troque o valor por A, se tiver um cabeçalho, coloque B e assim por diante.

* <b>rTemp = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, <u>"B"</u>).End(xlUp).Row</b></br>
Deve seguir o mesmo parametro que o anterior

* <b>ActiveWorkbook.ActiveSheet.Range(<u>"A6:F"</u> & rTemp).Copy shPadrao.Range(<u>"A"</u> & r + 1)</b></br>
O Caractere F deve ser trocado de acordo com o limite da coluna (se as informações das planilhas a serem lidas forem até a coluna J, o valor deverá ser A6:J)

Salve as alterações e execute a planilha

