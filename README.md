# 📊 Sistema de Análise Psicossocial — NR-01

> Ferramenta desenvolvida em Microsoft Excel para coleta, processamento e análise estatística de pesquisas psicossociais conforme exigência da **NR-01 (Nova Redação)**, gerando médias segmentadas por setor e por pergunta para apoiar o diagnóstico organizacional de riscos psicossociais.

---

## 🎯 Objetivo

Automatizar a análise de respostas de questionários psicossociais obrigatórios pela NR-01, eliminando a necessidade de cálculos manuais e reduzindo erros de interpretação. O sistema processa até **300 respondentes** e **41 perguntas**, gerando indicadores por setor de forma dinâmica.

---

## 🗂️ Estrutura da Planilha

A planilha é composta por **3 abas** com funções distintas e integradas:

### 1. `RESPOSTAS` — Entrada de Dados Brutos
- Recebe as respostas do questionário psicossocial com **41 perguntas**
- Cada linha representa um respondente contendo:
  - Carimbo de data/hora
  - Data de resposta
  - Setor do respondente
  - Respostas em escala Likert de 1 a 5 (ex: *"3- Às vezes"*)
- Suporta até **300 respondentes**

### 2. `AUXILIAR` — Tratamento e Extração Numérica
- Camada intermediária de processamento entre os dados brutos e os cálculos finais
- Aplica filtro de extração numérica em cada resposta, convertendo o texto da escala Likert em número puro (1 a 5):
```excel
=VALOR(ESQUERDA(RESPOSTAS!D6;1))
```
- Elimina textos descritivos, mantendo apenas o valor numérico para cálculo
- Cobre toda a matriz de respondentes x perguntas (até 300 linhas)

### 3. `BASECÁLCULO` — Processamento e Resultados
- Aba principal de visualização dos indicadores
- Gera **4 níveis de análise** com tratamento automático de erros e células vazias

---

## 📐 Níveis de Análise

| Nível | Descrição | Fórmula Base |
|---|---|---|
| **Média Geral Total** | Média de todas as perguntas e todos os setores | `=ARRED(MÉDIA(C1:AQ1);2)` |
| **Média Geral por Setor** | Média de todas as perguntas filtrada por setor | `MÉDIA` com `SEERRO` |
| **Média por Pergunta (geral)** | Média individual de cada pergunta sem filtro de setor | `SOMARPRODUTO` / `CONT.NÚM` |
| **Média por Pergunta por Setor** | Média individual de cada pergunta filtrada por setor | `SOMASE` / `CONT.SE` com `SEERRO` |

---

## 🔧 Fórmulas Principais

### Média Geral Total
```excel
=ARRED(MÉDIA(C1:AQ1);2)
```

### Média Geral por Setor
```excel
=SEERRO(SE(ARRED(MÉDIA(C3:AQ3);2)=0;"";ARRED(MÉDIA(C3:AQ3);2));"")
```

### Média por Pergunta sem Filtro de Setor
```excel
=ARRED(SOMARPRODUTO(SEERRO(Auxiliar!D$1:D$300;0))/CONT.NÚM(Auxiliar!D$1:D$300);2)
```

### Média por Pergunta com Filtro de Setor
```excel
=SEERRO(SE(ARRED(SOMASE(RESPOSTAS!$C$2:$C$300;$B$3;Auxiliar!D$1:D$300)/CONT.SE(RESPOSTAS!$C$2:$C$300;$B$3);2)=0;"";ARRED(SOMASE(RESPOSTAS!$C$2:$C$300;$B$3;Auxiliar!D$1:D$300)/CONT.SE(RESPOSTAS!$C$2:$C$300;$B$3);2));"")
```

---

## 🛡️ Tratamentos de Qualidade

O sistema implementa os seguintes tratamentos automáticos para garantir a integridade dos dados:

- **Células sem valor** → ficam invisíveis (retornam `""` em vez de zero ou erro)
- **Erros de divisão** (`#DIV/0!`) → suprimidos via `SEERRO`
- **Erros de formatação** (`###`) → tratados com `SEERRO` e `VALOR()`
- **Zeros** → não são exibidos, evitando distorção visual
- **Formatação condicional** → células são coloridas apenas quando contêm valor válido
- **Contagem correta** → uso de `CONT.NÚM` para contar apenas valores numéricos, ignorando erros e células vazias

---

## 🔄 Fluxo de Dados

```
RESPOSTAS (dados brutos)
        ↓
   Escala Likert em texto
   "1- Nunca/quase nunca"
   "3- Às vezes"
   "5- Sempre"
        ↓
AUXILIAR (extração numérica)
   =VALOR(ESQUERDA(...;1))
   Resultado: 1, 2, 3, 4 ou 5
        ↓
BASECÁLCULO (resultados)
   ├── Média Geral Total
   ├── Média Geral por Setor
   ├── Média por Pergunta (geral)
   └── Média por Pergunta por Setor
```

---

## ⚖️ Contexto Legal

Este sistema foi desenvolvido para atender às exigências da **NR-01 — Disposições Gerais e Gerenciamento de Riscos Ocupacionais** do Ministério do Trabalho e Emprego do Brasil, especificamente no que tange à identificação e avaliação de **riscos psicossociais** no ambiente de trabalho.

A NR-01 (atualizada) passou a exigir que as empresas incluam os riscos psicossociais no Gerenciamento de Riscos Ocupacionais (GRO), tornando obrigatória a aplicação de questionários e análise de dados como os processados por esta ferramenta.

---

## 👤 Autor

Lucas Favaro
Desenvolvido para uso profissional em gestão de saúde ocupacional e compliance com a legislação trabalhista brasileira.

---

## 📄 Licença

Este projeto está protegido. Todos os direitos reservados ao autor.
