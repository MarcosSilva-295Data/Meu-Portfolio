# Recommendation Monitoring

Automação para consolidação, tratamento e monitoramento de recomendações utilizando Python, Pandas e Openpyxl.

O projeto realiza:
- tratamento de bases Excel
- comparação entre fechamentos
- cálculo de métricas
- geração automática de dashboards
- criação de relatórios formatados em Excel
- interface gráfica para seleção dos arquivos

---

# Tecnologias utilizadas

- Python
- Pandas
- Openpyxl
- Tkinter

---

# Funcionalidades

## 1. Consolidação de bases

O sistema realiza:
- leitura de múltiplas planilhas Excel
- cruzamento entre bases
- enriquecimento das informações
- cálculo de status e métricas

---

## 2. Controle de vencimentos

As recomendações são classificadas em:
- Overdue
- Due in 30 days
- Due in 60 days
- Due in 90 days
- Future Due Dates

---

## 3. Comparação entre fechamentos

O projeto identifica:
- novas recomendações
- recomendações concluídas
- cancelamentos
- reescalonamentos
- evidências recebidas
- evidências recusadas

---

## 4. Overview automatizado

O sistema gera automaticamente:
- tabelas resumo
- métricas consolidadas
- visão por área responsável
- overview de vencimentos
- gráficos no Excel

---

## 5. Interface gráfica

A aplicação possui interface em Tkinter para:
- selecionar os arquivos
- escolher a pasta de saída
- executar a geração do report

---

# Como executar

## Instalar dependências

```bash
pip install -r requirements.txt
```

---

## Executar aplicação

```bash
python interface_recos.py
```

---

# Arquivos necessários

O sistema utiliza:

- Base atual
- Base comparativa
- Recommendation file
- Inventory file
- Template Excel

---

# Exemplo de saída

O sistema gera automaticamente:

- base consolidada
- overview executivo
- métricas
- dashboards formatados em Excel

---

# Objetivo do projeto

O projeto foi desenvolvido com o objetivo de automatizar o processo de consolidação, análise e acompanhamento de recomendações, reduzindo atividades manuais e padronizando a geração de relatórios e métricas.

---

# Melhorias futuras

- Exportação para Power BI
- Dashboard web
- Banco de dados
- Atualização automática
- Agendamento de execução
- API de integração

