Este código Python realiza a extração e análise de dados de uma consulta SQL para calcular o volume carregado e a quantidade de viagens, com base em um arquivo de consulta SQL. 
Após a leitura dos dados, ele filtra as informações conforme a data de inclusão e realiza cálculos de volume carregado diário e acumulado. O código também gera um gráfico percentual do volume carregado, salva-o como imagem e cria um resumo em formato HTML.
Se for dentro do horário de envio (entre 00h-10h ou 12h-23h), ele envia um e-mail com o resumo diário e acumulado, bem como o gráfico do percentual de volume carregado, para o endereço de e-mail especificado. Caso contrário, o envio do e-mail é ignorado. 
Além disso, ele ajusta a formatação dos valores de volume e apresenta os resultados de forma bem organizada.
