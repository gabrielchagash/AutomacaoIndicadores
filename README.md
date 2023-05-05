# AutomacaoIndicadores
 
Neste projeto criei um processo de forma automática para analisar a base de dados e calcular o OnePage de cada loja, bem como seus indicadores, e enviar um email para o gerente de cada loja com o seu OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo.

Utilizei as bibliotecas: Pandas, Pathlib e win32com.client.

Para analisar as bases de dados, utilzei o Pandas. Para arquivar as bases de dados criadas e alteradas, utilizei o Pathlib e para enviar os e-mails utilizei a win32com.client.

A lógica do processo todo, é como se segue:

Importar base de dados.

Definir e criar uma tabela para cada Loja e definir o dia do indicador
Criar um discionário para cada uma das lojas

Salvar cada planilha em excel na pasta de backup em suas devidas pastas, organizada de acordo com as lojas

Defini as metas de faturamento do dia e ano, de quantidade de produto do dia e ano, e do ticket médio do dia e ano

Calculei o indicador para cada loja utilizando o for

Enviei os emails, para isso precisei definir as cores para definir o cenário de cada loja, para melhor análise. Sendo verde para positivo e vermelho para negativo, caso a loja batesse a meta ou não. Para a criação do email utilizei uma integração com HTML para que cada tabela fosse modificada automaticamente, junto com o envio de cada planilha e email para os respectivos gerentes.

Criei um ranking de faturamento das lojas para o dia e ano.

Enviei o ranking também para o diretor com os rankings criado.


O projeto foi feito para automatizar um processo de indicadores, analisando a base de dados, integração de Python com HTML, automatizando envio de emails, criação de planilhas em excel e organização das pastas que foram criadas para cada planilha feita.

