# Vieira Móveis Sob Medida — Sistema de Gestão

Sistema completo de gestão para marcenaria especializada em móveis sob medida, desenvolvido com **Google Apps Script** e **Google Sheets**.

## Sobre o Projeto

Sistema integrado para gerenciar todas as operações de uma marcenaria:

- **Gestão de Clientes** — Cadastro e histórico de clientes
- **Gestão de Pedidos** — Controle completo do ciclo de vida dos pedidos
- **Controle de Estoque** — Monitoramento de materiais e quantidades
- **Movimentações** — Rastreamento de entradas e saídas
- **Dashboard** — Visualização em tempo real de métricas
- **Alertas Automáticos** — Notificações de prazos e estoque mínimo

## Funcionalidades Principais

### Gestão de Clientes

- Cadastro com informações de contato (telefone, e-mail, endereço)
- Histórico de pedidos vinculado
- Data de cadastro e observações

### Gestão de Pedidos

- Criação de novos pedidos vinculados a clientes
- Rastreamento de: descrição, prazos, montagem e instalação
- Controle financeiro (valor total, recebido e saldo)
- Status do pedido e responsável
- Integração com Google Calendar para eventos automáticos

### Controle de Estoque

- Cadastro de materiais com código e descrição
- Rastreamento de quantidade (atual e mínima)
- Histórico de fornecedores e preços
- Alertas quando atingir quantidade mínima

### Dashboard

- Atualização automática ao editar pedidos ou estoque
- Visualização rápida de informações críticas

### Sistema de Alertas

- Alerta de prazos de entrega (configurável em dias)
- Verificação automática ao editar planilha
- Notificações via e-mail ao proprietário

## Estrutura do Projeto

```
Projeto Integrador Extensor/
├── README.md                          # Este arquivo
├── appScript/
│   ├── Código.gs                      # Código principal (Google Apps Script)
│   ├── FormPedido.html                # Formulário de novo pedido
│   ├── FormCliente.html               # Formulário de novo cliente
│   ├── FormEstoque.html               # Formulário de entrada/saída de material
│   └── BuscaCliente.html              # Interface de busca de clientes
└── images/                            # Imagens do projeto
```

## Instalação e Configuração

### 1 Pré-requisitos

- Conta Google (Gmail)
- Acesso ao Google Sheets e Google Apps Script
- Navegador atualizado

### 2 Criação da Planilha

1. Acesse [Google Sheets](https://sheets.google.com)
2. Crie uma nova planilha
3. Nomeie como "Vieira Móveis Sob Medida"

### 3 Instalação do Script

1. Na planilha, clique em **Extensões → Apps Script**
2. Delete o código padrão em `Code.gs`
3. Copie o conteúdo do arquivo `appScript/Código.gs` para `Code.gs`
4. Copie os arquivos HTML para arquivos separados em Apps Script:
   - **FormPedido.html**
   - **FormCliente.html**
   - **FormEstoque.html**
   - **BuscaCliente.html**

### 4 Configuração Inicial

1. **IMPORTANTE:** Abra o arquivo `Código.gs` e procure pela variável `CONFIG`:

   ```javascript
   var CONFIG = {
     EMAIL_PROPRIETARIO: "SEU_EMAIL@gmail.com", // ← ALTERE AQUI
     DIAS_ALERTA_PRAZO: 3,
     NOME_EMPRESA: "Vieira Móveis Sob Medida",
   };
   ```

   Substitua `SEU_EMAIL@gmail.com` pelo seu e-mail real.

2. **Salve** o script (Ctrl+S)
3. **Volte à planilha** e recarregue (F5)
4. No menu **Vieira Móveis** → **Configuração Inicial**
5. Clique em **Sim** para criar todas as abas automaticamente

## Como Usar

### Novo Pedido

1. Menu **Vieira Móveis** → **Novo Pedido**
2. Selecione ou crie um cliente
3. Preencha descrição, prazos e valores
4. Clique em **Salvar**

### Novo Cliente

1. Menu **Vieira Móveis** → **Novo Cliente**
2. Preencha nome, telefone, e-mail e endereço
3. Clique em **Salvar**

### Movimentação de Estoque

1. Menu **Vieira Móveis** → **Entrada de Material** ou **Saída de Material**
2. Selecione o material e quantidade
3. Clique em **Salvar**

### Buscar Cliente

1. Menu **Vieira Móveis** → **Buscar Cliente**
2. Digite nome ou código
3. Visualize informações e histórico

### Atualizar Dashboard

- Clique em **Vieira Móveis** → **Atualizar Dashboard**
- Atualiza automaticamente ao editar pedidos ou estoque

### Verificar Alertas

- Clique em **Vieira Móveis** → **Verificar Alertas Agora**
- Envia notificações para prazos próximos

## Estrutura das Abas do Google Sheets

| Aba             | Descrição                                           |
| --------------- | --------------------------------------------------- |
| **Pedidos**     | Nº, Cliente, Descrição, Prazos, Valores, Status     |
| **Clientes**    | Código, Nome, Contato, Endereço, Histórico          |
| **Estoque**     | Código, Descrição, Quantidades, Fornecedor, Preço   |
| **Mov_Estoque** | Movimentações (entradas/saídas com rastreabilidade) |
| **Dashboard**   | KPIs e visualizações em tempo real                  |

## Variáveis de Configuração

Edite o objeto `CONFIG` em `Código.gs` para customizar:

```javascript
var CONFIG = {
  EMAIL_PROPRIETARIO: "seu.email@gmail.com", // E-mail para alertas
  DIAS_ALERTA_PRAZO: 3, // Dias antes do prazo para alertar
  NOME_EMPRESA: "Vieira Móveis Sob Medida", // Nome da empresa
};
```

## Permissões Necessárias

O script solicita acesso a:

- Google Sheets (ler/escrever dados)
- Gmail (envio de alertas)
- Google Calendar (criar eventos de entrega)

Aceite todos os acessos solicitados na primeira execução.

## Troubleshooting

### Menu não aparece

- Recarregue a planilha (F5)
- Verifique se o script foi salvo corretamente
- Verifique as permissões do Apps Script

### E-mails de alerta não chegam

- Verifique se o `EMAIL_PROPRIETARIO` está correto em `CONFIG`
- Verifique a pasta de spam/lixo
- Execute manualmente: **Vieira Móveis** → **Verificar Alertas Agora**

### Não consegue criar abas

- Verifique se a planilha já não possui as abas necessárias
- Tente criar manualmente as abas com os nomes exatos

## Notas Importantes

- Todos os dados são salvos no Google Sheets
- Backups automáticos (verificar versões do Sheets)
- Sincronização em tempo real
- Certifique-se de usar em ambiente de produção com cautela
- Faça testes antes de usar com dados reais

## Contribuição

1. Giovanni B. Giovanelli
2. Lucas Gabriel Genovezi
3. Mateus Henrique Oliveira
4. Thiago Henrique Bonierski

## Licença

Desenvolvido para Vieira Móveis Sob Medida.

