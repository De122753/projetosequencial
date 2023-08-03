# Projeto Sequencial Escola Técnica.
Automação em Python desenvolvida como projeto final para a escola Maria Helena Faria Lima e Cunha onde leciono.

# Tutorial de execução do código

- Primeiro passo
  - Acessar o site https://replit.com/
  - Fazer login na plataforma;
  - Criar um template Python


- Segundo passo
  - Instalar os módulos pandas e openpyxl usando o Shell do ambiente;
  - <code>pip install pandas</code>
  - <code>pip install openpyxl</code>

- Terceiro passo
  - Copiar o código do arquivo [main.py](main.py) que está neste repositório para main.py do ambiente Repl.it.

- Quarto passo
  - Editar valor da variável email_user da linha 23 com seu e-mail do gmail;
  - Editar valor da variável email_pass da linha 24 com a senha que será gerada na verificação de duas etapas do gmail;
    - Para obter a senha da verificação de duas estapas: acesse as configurações de sua conta Google>Segurança>Como você acessa o seu e-mail>Verificação de dua etapas>Ativar verificação>Senhas de App>Outros>Nomear>Gerar;
    - Copiar a senha gerada para a linha 24;
  - Editar, na linha 75, o valor da variável receiver_email. Atualizar com o e-mail do destinatário.
 
- Quinto passo
  - No Repl.it vá em Files>Upload file e faça o upload da tabela [Vendas.xlsx](Vendas.xlsx) que se encontra neste repositório.
 
- Sexto passo
  - No Shell do ambiente Repl.it executar o comando <code>python main.py</code>

# Considerações finais
O objetivo deste código é a manipulação dos dados da tabela Vendas.xlsx, impressão desses dados manipulados no Shell e envio deles por e-mail para algum detinatário. A automação envia um e-mail formatado por meio de tags html com as informações manipuladas.


  
  
