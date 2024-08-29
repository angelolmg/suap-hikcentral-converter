# Conversor de planilha SUAP para HikCentral

## Sobre

O **Conversor de planilha SUAP para HikCentral** é um script simples para converter planilhas na formatação SUAP para a formatação apropriada para upload na plataforma HikCentral Professional.

## QuickStart

### Pré-requisitos

Antes de começar, certifique-se de ter o Python instalado na sua máquina. Este script requer algumas bibliotecas Python, listadas no arquivo `requirements.txt`.

### Instalação

1. **Clone este repositório**:
	```bash
	 git clone https://github.com/seu-usuario/conversor-suap-hikcentral.git	
	 cd conversor-suap-hikcentral
	```

2. **Instale as dependências**:

   Utilize o `requirements.txt` para instalar as bibliotecas necessárias:

   ```bash
   pip install -r requirements.txt
   ```

### Como usar

1. **Coloque o arquivo de entrada** no mesmo diretório do script. O arquivo pode estar nos formatos `.xls`, `.xlsx` ou `.csv`.

2. **Execute o script**:

   ```bash
   python script.py
   ```

   O script irá procurar pelo primeiro arquivo `.xls`, `.xlsx`, ou `.csv` no diretório atual e criar um novo arquivo `.xlsx` com o nome modificado, adicionando `Modificado` ao nome original.

3. **Verifique a saída**:

   O arquivo convertido será salvo no mesmo diretório, com as formatações ajustadas para upload no HikCentral Professional. A estrutura do arquivo de saída incluirá as seguintes colunas:

   - **ID**: Número da matrícula (formatação como "número").
   - **Nome próprio**: Primeiro nome extraído da coluna "Nome".
   - **Apelido**: Sobrenome extraído da coluna "Nome".
   - **Departamento**: Valor fixo "All Departments/xyz".
   - **Hora de início do período de vigência**: Data atual no formato `aaaa/mm/dd hh:mm:ss`.
   - **Hora do fim do período de vigência**: Data de início + 10 anos.
   - **E-mail**: E-mail Pessoal.
   - **Matricula**: Campo deixado em branco.

### Dicas

- Certifique-se de que o arquivo de entrada contém as colunas necessárias conforme esperado pelo script (`#`, `Nome`, `Matrícula`, `Curso`, `Campus`, `Polo`, `Situação`, `E-mail Acadêmico`, `E-mail Pessoal`, `Ano/Periodo Letivo`).
- Se houver mais de um arquivo no diretório, o script usará o primeiro que encontrar, com base na ordem alfabética.

### Contribuindo

Contribuições são bem-vindas! Se você encontrar um problema ou tiver uma sugestão de melhoria, sinta-se à vontade para abrir uma issue ou enviar um pull request.

### Licença

Este projeto está licenciado sob a licença MIT.
