# Sistema Hoteleiro
<p align="center">
  <img src="https://img.shields.io/static/v1?label=VBA&message=MsExcel&color=green&style=for-the-badge&logo=microsoftoffice"/>
  <img src="http://img.shields.io/static/v1?label=SIZE&message=1.5 MB&color=blue&style=for-the-badge"/>
  <img src="http://img.shields.io/static/v1?label=STATUS&message=CONCLUIDO&color=GREEN&style=for-the-badge"/>
</p>

 <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/cover.jpg"></p>

### Sumário
🔹 [Descrição do projeto](#descrição-do-projeto)

🔹 [Pré-requisitos](#pré-requisitos)

🔹 [Instalação e ativação](#instalação-e-ativação)

🔹 [Rodando a aplicação](#rodando-a-aplicação)

🔹 [Possíveis Erros](#possíveis-erros)

🔹 [Planilhas](#planilhas)

🔹 [Telas](#telas)


## Descrição do projeto 
<p align="justify">
  Planilha em Excel que simula um sistema hoteleiro, realiza cadastros, consultas, cálculos e exclusão de dados.
</p>



## Pré-requisitos
* Ter o Microsoft Office Excel instalado
* Ter o  Microsoft Date and Time Picker Control 6.0 (SP6) instalado [(Confira Aqui)](#instalação-e-ativação)



## Instalação e ativação
Primeiramente faça download do módulo <a href="https://github.com/almeidastor/VBAs/raw/main/Sistema%20Hoteleiro/MSCOMCT2.zip" download>Microsoft Date and Time Picker</a> e extraia o arquivo da pasta zipada.

><h4>Windows 32bits</h4>
* Mova o arquivo mscomct2.ocx para a pasta C:\Windows\System32
* Execute o Prompt de Comando como Admnistrador e digite C:\Windows\System32\regsvr32.exe mscomct2.ocx


><h4>Windows 64bits</h4>
* Mova o arquivo mscomct2.ocx para a pasta C:\Windows\SysWoW64
* Execute o Prompt de Comando como Admnistrador e digite C:\Windows\SysWoW64\regsvr32.exe mscomct2.ocx

Em seguida aparecerá a mensagem de confirmação</br>

<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/instalacaoregistered.png"></p>


Execute o Arquivo <a href="https://github.com/almeidastor/VBAs/raw/main/Sistema%20Hoteleiro/Sistema%20Hoteleiro.xlsm" download>Sistema Hoteleiro.xlsm</a> e  [faça as ativações necessárias](#rodando-a-aplicação)

Habilite a guia do Desenvolvedor (Na aba <b>Arquivos</b> > <b>Opções</b> > <b>Personalizar Faixa de Opções</b> : e no quadro da direita habilite a box do <b> Desenvolvedor</b>) 

Busque a Aba do Desenvolvedor na Página Inicial e selecione no menu a opção <b>Inserir</b> e em seguida escolha a ultima opção <b>Mais controles</b> representado pelos icones das ferramentas (circulado em vermelho)
  
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/desenvolvedoractiv.png"></p>
  
Selecione a opção Microsoft Date and Time Picker Control 6.0 (SP6) e clique em OK
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/desenvolvedorcontroles.jpg"></p>

Assim sua aplicação está pronta para ser rodada (ou confira [Possíveis Erros](#possíveis-erros))
  
  

## Rodando a aplicação
Assim que abrir a planilha, localize os alertas de Edição e Macro e os habilite

  <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/ativacaoexcel.png"></p>
  
  
  
## Possíveis Erros
* Erro de Compilação: É impossivel localizar o projeto ou a biblioteca

>Se ao clicar em algum elemento da planilha, e o alerta Erro de Compilação aparecer, busque a opção de <b>Referências</b> na aba de <b>Ferramentas</b> e então desabilite as caixas marcadas como AUSENTE:


  <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/errodecompiler.jpg"></p>
  
  
  <p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/erroobjetoausente.GIF"></p>
  
  
  ## Planilhas
  * Aba Hotel


  * Aba Banco de Dados


  * Aba Valores

  ## Planilhas
  * Aba Hotel
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/abahotel.jpg"></p>

  * Aba Banco de Dados
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/ababcodados.jpg"></p>

  * Valores
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/abavalores.jpg"></p>


  ## Telas
  * Botão Novo Cadastro
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/btninserir.png"></p>

  * Botão Consulta e Edições   
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/btnconed.png"></p>
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/btnconsulta.png"></p>
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/btnedicoes.png"></p>

* Botão Reajuste
<p align="center"><img src="https://github.com/almeidastor/VBAs/blob/main/Sistema%20Hoteleiro/README-repository/btnreajust.png"></p>
