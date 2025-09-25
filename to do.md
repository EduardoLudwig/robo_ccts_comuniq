fazer os ccts comuniq e rede M:\Remuneração e Benefícios\SINDICAIS\Analises sindicatos lojas 2025\09_SETEMBRO\ENVIADAS\thai e da pasta "atualizar no salário certo" 

e checar os documentos desse caminho: M:\Remuneração e Benefícios\SINDICAIS\Analises sindicatos lojas 2025\09_SETEMBRO\ENVIADAS   teoricamente foram feitos na rede e comuniq


**melhorias para o robo do comuniq:**

headless
testar criar uma rotina para todo dia 1 ele upar (talvez isso seja util para o robo da rede)

criar alguma marcação para os pdfs que já foram upados, talvez mudar o nome

ex: 23_TRÊS DE MAIO_SIND0019_2025-2026.pdf --> *_*23_TRÊS DE MAIO_SIND0019_2025-2026.pdf 

para o robo entender o que foi upado ele pode contar quantos arquivos tinham antes de upar e quantos tem depois de salvar (ideia) 
ou contar quantas vezes a palavra remover aparece no campo de upload (antes e depois de upar)

e o que não foi upado tanto por erro ou por sequencia de bugs ele marca como outro código para tentar novamente

fazer alguma medida para impedir as sequencias de stacktrace error (SUPER COMUM NO COMUNIQ)


**isso também deve se aplicar no robo dos ccts da rede**


caso não encontre a loja deve criar um txt explicando que alguem precisa criar a loja.... 

colocar algum icone girando ou uma ampulheta para sinalizar que ele esta trabalhando 

se depois vigência tem algo mas não esta com um _ (underline) só ignorar e formular o nome parando na vigência, as vezes quem analisa escreve ok ou algo do tipo e atrapalha o robo  

ex: 483, 505, 549, 604_SAO JOSE DO CERRITO_FLORIANOPOLIS _SIND0210_2025-2026 (Bruna)

nesse caso formataria no padrao ignorando o que vem depois da vigência

e se caso não tem vigência, só não fazer e marcar de alguma forma para tentar novamente (ele já não faz o que não tem vigência)

no final do processo deve mostrar uma messagebox de alerta com quais documentos deram errado e o porque
e deve pedir para tentar novamente.

obviamente os que estão marcados como feitos não vai ser feito o upload novamente.