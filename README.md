# Relatório de Validação de Dados Financeiros

Este projeto automatiza a extração, validação e notificação de informações financeiras relacionadas a títulos de renda fixa (CRA/CRI), integrando dados de fontes distintas (B3 e sistema interno Galáxia) e destacando divergências.

---

## **Funcionalidades**

### **1. Extração de Dados**
- **Base de Dados MySQL (Azure):**
  - Realiza consultas SQL para obter eventos financeiros programados para títulos de corporações específicas (ex.: VIRGO e ISEC).
- **API do Galáxia:**
  - Coleta informações adicionais sobre os títulos financeiros usando uma API REST.

### **2. Validação de Dados**
- Compara os dados extraídos da B3 com os dados do sistema Galáxia.
- Identifica discrepâncias nos seguintes tipos de eventos:
  - **Pagamento de Juros.**
  - **Amortização.**
  - **Amortização Extraordinária.**

### **3. Geração de Relatórios**
- Apresenta os dados em uma interface gráfica desenvolvida com **PySimpleGUI**:
  - Relatório de eventos extraídos da B3.
  - Relatório de eventos extraídos do Galáxia.
  - Comparativo de valores, com destaques para divergências.

### **4. Notificações Automáticas**
- Envia e-mails ao responsável sempre que forem identificadas discrepâncias entre os valores da B3 e do Galáxia.
- Mensagens personalizadas detalham os erros e solicitam correções.

---

## **Tecnologias Utilizadas**
- **Python 3.9**
- **PySimpleGUI:** Criação da interface gráfica.
- **SQLAlchemy:** Conexão com o banco de dados MySQL.
- **Urllib:** Consumo de API REST.
- **Win32com:** Envio de e-mails via Outlook.
- **JSON:** Manipulação de dados retornados pela API do Galáxia.
- **Datetime:** Manipulação de datas.

---

## **Como Executar**

### **1. Requisitos**
- Python 3.9 ou superior.
- Dependências listadas no arquivo `requirements.txt`:
  ```
  PySimpleGUI
  SQLAlchemy
  PyMySQL
  pywin32
  ```
- Certificado SSL `DigiCertGlobalRootCA.crt.pem` para conexão segura com o banco de dados.

### **2. Configuração do Ambiente**
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Adicione o certificado SSL ao diretório do projeto.
3. Certifique-se de que as credenciais de conexão com o banco de dados estejam corretas no código.

### **3. Execução**
1. Execute o script principal:
   ```bash
   python script.py
   ```
2. Navegue entre as abas da interface para visualizar os relatórios:
   - Relatório B3
   - Relatório Galáxia
   - Comparativo
3. Caso divergências sejam detectadas, o sistema enviará e-mails automaticamente.

---

## **Estrutura do Projeto**
```
.
├── DigiCertGlobalRootCA.crt.pem  # Certificado SSL
├── script.py                     # Script principal
├── requirements.txt              # Dependências do projeto
├── galaxia.ico                   # Ícone do aplicativo
├── img.png                       # Imagem utilizada na interface
└── README.md                     # Documentação
```

---

## **Melhorias Futuras**
- Implementar logs para auditoria e rastreamento de erros.
- Automatizar a atualização de credenciais e certificados.
- Adicionar suporte para outros tipos de títulos financeiros.
- Incluir testes automatizados para validação do funcionamento do sistema.

---

## **Autores**
- Desenvolvido por: [Seu Nome Aqui]  
  Contato: [seu.email@exemplo.com]

---

## **Licença**
Este projeto é licenciado sob a [MIT License](https://opensource.org/licenses/MIT).

