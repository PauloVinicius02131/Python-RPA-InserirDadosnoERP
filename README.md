# Inserir dados repetitivos no ERP. (Compras)
<div align="center" display="inline-block">
<img src="https://img.shields.io/badge/Python-FFD43B?style=for-the-badge&logo=python&logoColor=blue">
</img>
<img src="https://img.shields.io/badge/Selenium-43B02A?style=for-the-badge&logo=Selenium&logoColor=white">
</img>
<img src="https://img.shields.io/badge/Pandas-2C2D72?style=for-the-badge&logo=pandas&logoColor=white">
</img>
<img src="https://img.shields.io/badge/Numpy-777BB4?style=for-the-badge&logo=numpy&logoColor=white">
</img>
</div>

<br>
   Este projeto surgiu da necessidade de inserir dados repetitivos no ERP Transnet, neste caso são dados relacionados a pedidos de compras ja que surgiu como solução a fim de viabilizar um setor de compras corporativo unificando diversas empresas em um único sistema.

   O objetivo principal é minimizar o custo de horas de 12 funcionários que por semana gastam cerca de 1 hora cada para inserir tais dados.

   O projeto foi desenvolvido em Python 3.6 e utiliza a biblioteca Selenium para automatizar o processo de inserção de dados no ERP.

## Instalação 

    Para instalar o projeto basta clonar o repositório e instalar as dependências.
    
    ```bash
    git clone

## Configuração

    Para configurar o projeto corretamente é necessário :

    - Incluir o endereço URL do sistema no arquivo /lib/funcoes.py na variável URL.

    - Incluir o nome de usuario do sistema no arquivo /lib/funcoes.py na variável USUARIO.

    - Incluir senha de sistema no arquivo /lib/funcoes.py na variável SENHA.

## Uso

 Na planilha entrada/DadosEntrada.xlsx deve ser inserido os dados de acordo com o modelo.

 Para utilizar o projeto basta executar o arquivo project.py

 Após a finalização do processo de inserção de dados o sistema irá gerar um arquivo com a data e hora da execução no formato .csv na área de trabalho do usuário que executou o projeto.

Este arquivo possui as informações de cada peça inserida no sistema. Podendo conter alertas de erros ou informações de sucesso.


## Contribuição
    
        Para contribuir com o projeto basta enviar um pull request.



<img src="https://img.shields.io/badge/Status-Validado-green.svg"></img>


<img src="https://img.shields.io/badge/License-MIT-green.svg"></img>

<img src="https://img.shields.io/badge/Version-1.0-green.svg"></img></img>

<img src="https://img.shields.io/badge/Author-Paulo Vinicius-%2300BFFF.svg"></img>
