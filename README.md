# CRUD-SERVICES
# 🐍 Sistema de Cadastro de Serviços

## 📋 Descrição do Projeto

O **Sistema de Cadastro de Serviços - VAC_OS** (Ordem de Serviço de Vacol) é uma aplicação de desktop desenvolvida em Python com a biblioteca **Tkinter** e um banco de dados **SQLite** local. Projetada para ser uma solução leve, autônoma e eficiente.

A aplicação oferece uma interface completa de Cadastro, Leitura, Atualização e Exclusão (CRUD) para gerenciar ordens de serviço. Vai além do básico com recursos essenciais para a operação, como **validação de CPF**, **verificação de duplicidade de endereço**, **paginação de dados** para alta performance e funcionalidades de exportação para **planilhas Excel** (via Pandas) e **documentos PDF** formatados (via ReportLab).

O objetivo é fornecer uma ferramenta leve, autônoma e confiável, valorizando a estabilidade e a clareza na arquitetura modular do código.

## ✨ Funcionalidades Principais

* **CRUD Completo:** Cadastro, Visualização, Edição e Exclusão de Ordens de Serviço.
* **Gerenciamento de Dados:** Utilização de SQLite local (`servicos.db`) para persistência de dados.
* **Interface Amigável:** Desenvolvido com Tkinter e estilizado com `ttk.Style`.
* **Filtros Dinâmicos:** Pesquisa por Nome, CPF, Status, Bairro e Rua na lista de serviços.
* **Paginação:** Controle de exibição de registros (10, 20, 50, 100 por página) para desempenho em grandes volumes.
* **Validação de Dados:** Utiliza a biblioteca `validate-docbr` para validação e formatação de CPF.
* **Prevenção de Erros:** Alerta automático sobre endereço duplicado antes de salvar o serviço.
* **Exportação de Relatórios:**
    * **Excel:** Exportação de todos os registros (ou filtrados) para arquivo `.xlsx` (via Pandas).
    * **PDF:** Geração e visualização de um PDF formatado da Ordem de Serviço individual (via ReportLab).

## 🚀 Tecnologias Envolvidas

* **Python 3.x**
* **Tkinter** (Interface Gráfica)
* **SQLite** (Banco de Dados Local)
* **Pandas** (Exportação para Excel)
* **ReportLab** (Geração de PDF)
* **validate-docbr** (Validação de CPF)

## ⚙️ Instalação e Execução

### Pré-requisitos

Certifique-se de ter o Python 3.x instalado em seu sistema operacional.

### 1. Clonar o Repositório

```bash
git clone <URL_DO_SEU_REPOSITORIO>
cd vac_os
