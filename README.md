Automação de Impressão com Python para o software Fiery 

Sumário

•Sobre o projeto

•	About the project

•	Linguagens e tecnologias usadas

•	Bibliotecas usadas

•	Passo a passo das soluções

•	Resumo das funcionalidades importantes do código

•	Captura de tela

•	Conclusões

•	Autores

Sobre o projeto

Este repositório contém uma solução para o projeto Automação de Impressão com o uso do software Fiery voltado para as impressoras Xerox C70 e C60. A automação também pode ser adaptada para outras impressoras que utilizam o software Fiery para configuração de impressão. O projeto foi desenvolvido inicialmente no Visual Studio Code, utilizando majoritariamente a biblioteca PyAutoGUI. Essa biblioteca é amplamente utilizada para a automação de comandos do mouse e do teclado, permitindo interações dinâmicas com interfaces gráficas. Um de seus pontos positivos é a facilidade de implementação e a independência de ferramentas externas, tornando-a uma escolha acessível para quem está iniciando em automação com Python.
Desafio: O desafio principal consistiu em criar um processo automatizado para configurar e gerenciar tarefas de impressão através do software Fiery. Embora os desafios encontrados durante o desenvolvimento tenham sido poucos, algumas adequações foram necessárias, principalmente relacionadas a detalhes nas funções e comportamentos específicos do software.
Este projeto pode ser útil para profissionais que desejam otimizar processos de impressão e aprender mais sobre as possibilidades de automação com Python em um ambiente corporativo.



About the project

This repository contains a solution for the Printing Automation project using Fiery software for the Xerox C70 and C60 printers. The automation can also be adapted to other printers that use Fiery software for printing configuration. The project was initially developed in Visual Studio Code, using mainly the PyAutoGUI library. This library is widely used for automating mouse and keyboard commands, allowing dynamic interactions with graphical interfaces. One of its strengths is its ease of implementation and independence from external tools, making it an accessible choice for those starting out in automation with Python.
Challenge: The main challenge was to create an automated process to configure and manage printing jobs using Fiery software. Although there were few challenges encountered during development, some adjustments were necessary, mainly related to details in the specific functions and behaviors of the software.
This project can be useful for professionals who want to optimize printing processes and learn more about the possibilities of automation with Python in a corporate environment.

Linguagens e tecnologias usadas

•	Visual Studio Code
•	Python 3.11.9
•	Markdown

Bibliotecas usadas

•	Win32api
•	Os
•	Time
•	Pyautogui
•	Pymsgbox

Passo a passo das soluções
1. Entrada de dados do usuário
•	O programa solicita ao usuário três entradas:
o	Qual arquivo deseja imprimir (1 a 9, correspondendo a diferentes pastas).
o	Quantidade de cópias para cada arquivo.
o	Máquina de impressão (entre 6 e 10).
•	O programa valida as entradas. Se houver erro, uma exceção é levantada com uma mensagem de erro.
2. Determinação do caminho do arquivo
•	Dependendo da escolha do usuário, o programa define o caminho do diretório onde os arquivos a serem impressos estão localizados.
•	Para arquivos específicos (8 e 9), há um ajuste no valor da variável anexo para alterar a configuração de impressão.
3. Listagem de arquivos
•	A função os.listdir obtém todos os arquivos no diretório escolhido pelo usuário. Esses arquivos serão impressos um por um.
4. Automação da impressão
Para cada arquivo listado:
1.	Abrir o arquivo:
O arquivo é aberto usando win32api.ShellExecute, simulando uma interação manual com o sistema.
2.	Função imprimir:
o	A função imprimir(a, b, c) usa a biblioteca pyautogui para realizar as ações de automação da impressão.
o	Passos realizados pela função:
	Simula o comando Ctrl + P para abrir a interface de impressão.
	Realiza cliques e seleções nos menus de configuração de impressão, ajustando opções de cópias (a), máquina (b), e anexos (c).
	Preenche campos e confirma as opções.
3.	Monitorar o progresso:
o	A função verificar(a) verifica se a impressão foi concluída. Ela tenta localizar uma imagem na tela (ex.: "progresso.png") que indica o status de impressão.
o	Enquanto o status estiver ativo, o programa espera.
4.	Fechar o arquivo:
o	Após concluir a impressão de um arquivo, o programa fecha a janela do arquivo usando o comando Ctrl + W.
5. Fechamento do programa
•	Quando todos os arquivos no diretório escolhido forem processados, o programa fecha o sistema usando o comando Ctrl + Q.
6. Tratamento de erros
•	Caso ocorra qualquer erro durante a execução (ex.: entrada inválida, falha no caminho, etc.), ele será capturado e exibido na tela para o usuário.

Resumo das funcionalidades importantes do código
1.	Automação de impressão: Utiliza pyautogui para simular interações humanas no sistema de impressão.
2.	Gerenciamento de arquivos: Usa os e win32api para acessar arquivos no sistema operacional e abri-los automaticamente.
3.	Validação de entrada: Garante que os dados fornecidos pelo usuário são válidos antes de executar o restante do programa.
4.	Monitoramento de progresso: Verifica o status da impressão antes de prosseguir para o próximo arquivo.

Captura de tela
PNG progresso.
 

Conclusões

A solução desenvolvida para o projeto Automação de Impressão com o uso do software Fiery cumpre seu objetivo de automatizar a configuração e o envio de tarefas de impressão para impressoras Xerox C70 e C60, além de ser adaptável a outros equipamentos que utilizem o mesmo software. A automação foi implementada utilizando majoritariamente a biblioteca PyAutoGUI, que apresenta vantagens, mas também limitações inerentes ao seu funcionamento.
Um ponto interessante do PyAutoGUI é sua simplicidade na simulação de interações humanas com a interface gráfica. Contudo, a biblioteca apresenta algumas limitações:
•	Dependência da resolução da tela: Os comandos como pyautogui.click() utilizam a quantidade de pixels para determinar as posições de clique, o que exige adaptações quando o código é transferido para outro computador.
•	Impossibilidade de uso simultâneo do computador: Como as interações ocorrem diretamente na interface gráfica, não é possível realizar outras atividades enquanto o script está em execução.
•	Tempos de espera arbitrários: É necessário definir intervalos entre os comandos para evitar erros, e esses intervalos podem variar dependendo da velocidade do computador ou da rede, o que torna o script menos robusto.
Apesar das limitações, o código criado demonstra ser funcional e eficiente para realizar automações com o software Fiery, podendo ser adaptado para outros fluxos de trabalho que envolvam sistemas de impressão. Adaptações adicionais seriam necessárias para automatizar tarefas específicas, como login no sistema, reconhecimento automático de arquivos corretos para impressão e a geração de relatórios baseados em dados de impressão.
Autores

Criado por Felipe Ivo da Silva e João Pedro Iannoni Milaré 

