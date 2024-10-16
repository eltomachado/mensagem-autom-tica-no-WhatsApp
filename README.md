## 📅 Notificador de Vencimento de Planos

Este projeto em Python envia notificações via WhatsApp para clientes sobre a data de vencimento de seus planos. Utilizando uma planilha Excel, o aplicativo verifica as datas de vencimento e envia mensagens personalizadas aos clientes.

## 🚀 Funcionalidades

- 📊 Leitura de planilhas Excel com informações dos clientes.
- 🖥️ Interface gráfica para entrada de dados, como chave Pix e nome da empresa.
- 💬 Envio automático de mensagens via WhatsApp para clientes próximos ao vencimento de seus planos.

## 🛠️ Tecnologias Utilizadas

- Python
- Pandas
- PyWhatKit
- Tkinter

## 📋 Pré-requisitos

Certifique-se de ter os seguintes itens instalados:

- Python 3.x
- Bibliotecas necessárias:
  ```bash
  pip install pandas pywhatkit openpyxl
  ```

## 📝 Como Usar

1. Clone o repositório:
   ```bash
   git clone https://github.com/eltomachado/mensagem-autom-tica-no-WhatsApp
   ```
2. Navegue até o diretório do projeto:
   ```bash
   cd mensagem-autom-tica-no-WhatsApp
   ```
3. Execute o script:
   ```bash
   python notificador.py
   ```
4. Preencha a chave Pix e o nome da empresa na interface gráfica.
5. Certifique-se de que a planilha `clientes.xlsx` esteja no mesmo diretório do script e siga o formato adequado.

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## 📄 Licença

Este projeto é licenciado sob a MIT License - consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
