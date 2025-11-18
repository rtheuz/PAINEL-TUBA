# 🎯 PAINEL TUBA

Sistema completo de gestão empresarial desenvolvido com Google Apps Script e HTML, focado no gerenciamento de oficinas automotivas, manutenções e controle de materiais.

## 📋 Sobre o Projeto

O **Painel TUBA** é uma aplicação web robusta que oferece uma solução integrada para gestão de oficinas e serviços automotivos. O sistema utiliza Google Sheets como banco de dados, proporcionando uma solução acessível, escalável e fácil de manter.

## ✨ Funcionalidades Principais

### 🔐 Autenticação e Segurança
- Sistema de login seguro com controle de acesso
- Páginas protegidas com verificação de autenticação
- Gerenciamento de permissões de usuários

### 📊 Dashboard
- Painel de controle com métricas e indicadores
- Visualização consolidada de dados
- Acesso rápido às principais funcionalidades

### 🚗 Gestão de Veículos
- Cadastro completo de veículos
- Listagem e busca de veículos
- Histórico de manutenções por veículo
- Controle de informações técnicas

### 🔧 Manutenção
- Registro de ordens de serviço
- Acompanhamento de status de manutenções
- Histórico completo de serviços realizados
- Logs detalhados de manutenção

### 📋 Kanban
- Sistema visual de gerenciamento de tarefas
- Organização por status (A fazer, Em andamento, Concluído)
- Drag and drop para movimentação de cards
- Acompanhamento de progresso

### 💰 Gestão Financeira
- Criação e gerenciamento de orçamentos
- Controle de pedidos
- Geração de relatórios financeiros

### 📦 Controle de Materiais
- Cadastro de materiais e produtos
- Controle de estoque
- Gestão de fornecedores

### 🏷️ Sistema de Etiquetas
- Gerador de etiquetas personalizadas
- Impressão de etiquetas para identificação
- Templates customizáveis

### ⭐ Avaliações
- Sistema de avaliação de serviços
- Registro de feedback de clientes
- Análise de satisfação

### 📝 Formulários e Registros
- Formulários dinâmicos para entrada de dados
- Validação de informações
- Integração automática com banco de dados

### 📊 Logs e Auditoria
- Registro de todas as ações do sistema
- Rastreabilidade de operações
- Histórico de alterações

## 🛠️ Tecnologias Utilizadas

- **Frontend:**
  - HTML5
  - CSS3
  - JavaScript (Vanilla)
  - Bootstrap/Material Design (inferido pela estrutura)

- **Backend:**
  - Google Apps Script
  - Google Sheets (Banco de Dados)

- **Serviços Google:**
  - Google Drive
  - Google Sheets API
  - Google Apps Script Web App

## 📁 Estrutura do Projeto

```
PAINEL-TUBA/
├── Code.js                    # Backend principal (Google Apps Script)
├── login.html                 # Página de autenticação
├── dashboard.html             # Painel principal
├── paginasprotegidas.html     # Controle de acesso
│
├── 🚗 Veículos
│   ├── veiculos.html          # Cadastro de veículos
│   └── veiculos_list.html     # Listagem de veículos
│
├── 🔧 Manutenção
│   ├── manutencao.html        # Gestão de manutenções
│   └── manutencaologs.html    # Histórico de manutenções
│
├── 📋 Gestão
│   ├── kanban.html            # Board Kanban
│   ├── pedidos.html           # Controle de pedidos
│   └── orcamentos.html        # Gestão de orçamentos
│
├── 📦 Materiais
│   ├── materiais.html         # Controle de materiais
│   └── produtos.html          # Cadastro de produtos
│
├── 🏷️ Etiquetas
│   ├── geradoretiquetas.html  # Gerador de etiquetas
│   └── etiqueta.html          # Template de etiqueta
│
├── ⭐ Avaliações
│   ├── avaliacoes.html        # Sistema de avaliações
│   └── avaliacoespage.html    # Visualização de avaliações
│
├── 📝 Formulários e Logs
│   ├── formulario.html        # Formulários dinâmicos
│   └── logs.html              # Sistema de logs
│
└── README.md                  # Documentação
```

## 🚀 Como Usar

### Pré-requisitos

- Conta Google
- Acesso ao Google Drive
- Permissões para Google Apps Script

### Instalação

1. **Clone ou faça download do repositório:**
   ```bash
   git clone https://github.com/rtheuz/PAINEL-TUBA.git
   ```

2. **Configure o Google Apps Script:**
   - Acesse [script.google.com](https://script.google.com)
   - Crie um novo projeto
   - Cole o conteúdo de `Code.js` no editor
   - Configure as planilhas necessárias no Google Sheets

3. **Configure as planilhas:**
   - Crie uma planilha no Google Sheets
   - Configure as abas conforme necessário:
     - Usuários
     - Veículos
     - Manutenções
     - Materiais
     - Pedidos
     - Orçamentos
     - Avaliações
     - Logs

4. **Implante como Web App:**
   - No Apps Script, vá em: Implantar > Nova implantação
   - Selecione "Aplicativo da Web"
   - Configure "Executar como" para sua conta
   - Configure "Quem tem acesso" conforme necessário
   - Copie a URL da implantação

5. **Configure os arquivos HTML:**
   - Atualize as URLs de API nos arquivos HTML com a URL da sua implantação

## 🔧 Configuração

### Variáveis de Ambiente

No arquivo `Code.js`, configure:

```javascript
// ID da planilha principal
const SPREADSHEET_ID = 'seu-id-de-planilha-aqui';

// Abas da planilha
const SHEETS = {
  USUARIOS: 'Usuários',
  VEICULOS: 'Veículos',
  MANUTENCOES: 'Manutenções',
  MATERIAIS: 'Materiais',
  // ... outras abas
};
```

## 📖 Documentação das Páginas

### Login (`login.html`)
Página de autenticação com validação de credenciais.

### Dashboard (`dashboard.html`)
Painel central com métricas e acesso rápido às funcionalidades.

### Kanban (`kanban.html`)
Gestão visual de tarefas com funcionalidade drag-and-drop.

### Veículos (`veiculos.html`, `veiculos_list.html`)
Cadastro completo e listagem de veículos com filtros.

### Manutenção (`manutencao.html`, `manutencaologs.html`)
Controle de ordens de serviço e histórico de manutenções.

### Orçamentos (`orcamentos.html`)
Criação e gestão de orçamentos com cálculos automáticos.

## 🔒 Segurança

- Autenticação obrigatória para acesso
- Controle de sessão
- Validação de dados no backend
- Logs de auditoria para rastreabilidade
- Proteção contra acesso não autorizado

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:

1. Fork o projeto
2. Criar uma branch para sua feature (`git checkout -b feature/NovaFuncionalidade`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/NovaFuncionalidade`)
5. Abra um Pull Request

## 📝 Licença

Este projeto está sob licença livre para uso e modificação.

## 👤 Autor

**rtheuz**
- GitHub: [@rtheuz](https://github.com/rtheuz)

## 📞 Suporte

Para questões e suporte, abra uma [issue](https://github.com/rtheuz/PAINEL-TUBA/issues) no GitHub.

---

⭐ Se este projeto foi útil para você, considere dar uma estrela no repositório!

**Desenvolvido com ❤️ para gestão eficiente de oficinas automotivas**
