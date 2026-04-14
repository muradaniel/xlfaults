# ⚡ XLFaults – Simulador de Curto-Circuito no Excel

O **XLFaults** é uma ferramenta para análise de curto-circuito em sistemas elétricos de potência, desenvolvida por **Daniel Moraes de Freitas**, estudante de Engenharia Elétrica pela UERJ.

A proposta do projeto é tornar a modelagem e análise de sistemas elétricos mais acessível, utilizando o Microsoft Excel como interface principal, eliminando a necessidade de arquivos de entrada complexos.

---

## 🚀 Principais Diferenciais

- 📊 Interface baseada em **tabelas no Excel**
- 🔌 Integração com **Python** para cálculos elétricos
- ⚡ Simulação de diferentes tipos de curto-circuito
- 📈 Geração automática de **diagramas unifilares**
- 📑 Relatórios completos e organizados
- 🔍 Validação com software profissional como o Anafas

---

## 🛠️ Instalação

1. Baixe o projeto no GitHub  
2. Abra o Excel e ative a guia **Desenvolvedor**
3. Vá em:
4. Adicione o arquivo `.xlam` do projeto
5. Ative o suplemento

---

## 📦 Instalação de Dependências

O XLFaults utiliza bibliotecas Python para processamento.

➡️ Dentro do Excel, clique no botão:
**"Instalar Dependências"**

Isso instalará automaticamente os pacotes listados no arquivo `requirements.txt`.

---

## ⚙️ Configuração Inicial

Antes de iniciar:

- Gere o **template padrão**
- Defina:
- Potência base (ex: 100 MVA)
- Nome do caso
- Unidade dos resultados:
 - PU (por unidade)
 - Valores reais

---

## 🧩 Modelagem do Sistema

O sistema elétrico é montado através de tabelas no Excel:

### 🔹 Barramentos
- Níveis de tensão
- Identificação do sistema

### 🔹 Linhas de Transmissão
- Impedâncias (Z1, Z2, Z0)

### 🔹 Transformadores
- Relações de transformação
- Impedâncias

### 🔹 Máquinas
- Geradores ou motores
- Dados de sequência

✔️ Suporte completo para:
- Sequência positiva
- Sequência negativa
- Sequência zero

---

## ⚡ Tipos de Curto-Circuito

O simulador permite configurar:

- Curto trifásico
- Curto monofásico
- Curto bifásico
- Curto bifásico-terra

Com opção de:
- Impedância de falta
- Seleção da barra de aplicação

---

## 📊 Resultados

Após a simulação, o XLFaults gera:

### 📌 Diagrama Unifilar
- Colorido por nível de tensão
- Gerado automaticamente

### 📌 Matrizes do Sistema
- Matriz de Admitância (Ybarra)
- Matriz de Impedância (Zbarra)

### 📌 Relatório Completo
- Correntes de falta
- Tensões nos barramentos
- Contribuição de cada elemento

---

## ✅ Validação

Os resultados do XLFaults foram validados com:

- Software profissional (Anafas)
- Exemplos acadêmicos clássicos

✔️ Alta precisão nos cálculos

---

## 🔮 Próximas Funcionalidades

- 🔁 Transformadores de **três enrolamentos**
- 🎚️ Ajuste de **tap**
- 📊 Módulo de **fluxo de potência**

---

## 🎥 Demonstração

Confira o vídeo completo do projeto:

👉 https://youtu.be/DFxH0akx1NI

---

## 👨‍💻 Autor

**Daniel Moraes de Freitas**  
Estudante de Engenharia Elétrica – UERJ  
Foco em Sistemas Elétricos de Potência  

---

## 📌 Objetivo do Projeto

O XLFaults foi desenvolvido com o objetivo de:

- Facilitar a análise de curto-circuito
- Melhorar a organização de dados elétricos
- Utilizar ferramentas acessíveis como o Excel
- Integrar engenharia elétrica com programação
