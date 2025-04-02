# Envio Automático de E-mails com Outlook e Assinatura

Este script permite o envio automático de e-mails via Microsoft Outlook com base em um arquivo CSV contendo os destinatários. Ele também anexa uma imagem de assinatura automaticamente.

## 📌 Funcionalidades

- Selecionar um arquivo CSV contendo os e-mails dos destinatários.
- Gerar e-mails personalizados com base nos dados do CSV.
- Anexar automaticamente uma assinatura em formato de imagem.
- Enviar os e-mails via Outlook.

## 🚀 Como Usar

### 1️⃣ **Pré-requisitos**

- Windows com Microsoft Outlook instalado e configurado.
- Python 3.x instalado.
- As seguintes bibliotecas Python instaladas:
  ```sh
  pip install pandas pywin32
  ```

### 2️⃣ **Estrutura do CSV**

O arquivo CSV deve conter as seguintes colunas:

```csv
Email,Vencimento
joao@email.com,2025-04-10
maria@email.com,2025-04-12
```

### 3️⃣ **Executando o Script**

1. Clone ou baixe este repositório.
2. Certifique-se de que a imagem `assinatura.jpeg` está no mesmo diretório do script.
3. Execute o script:
   ```sh
   python envio_boletos.py
   ```
4. Selecione o arquivo CSV quando solicitado.
5. O script enviará os e-mails automaticamente.

## 🛠️ Como Criar um Executável

Caso queira transformar o script em um executável `.exe` para facilitar sua execução:

1. Instale o PyInstaller:
   ```sh
   pip install pyinstaller
   ```
2. Gere o executável:
   ```sh
   pyinstaller --onefile --windowed --add-data "assinatura.jpeg;." envio_boletos.py
   ```
3. O executável estará disponível na pasta `dist`.

## ⚠️ Observações

- O Outlook deve estar aberto e configurado na máquina.
- Caso haja problemas com a assinatura, verifique se o arquivo `assinatura.jpeg` está no mesmo diretório do executável.