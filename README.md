# 🛡️ TrataDoc Studio v8.2

### *Sanitização, Auditoria e Processamento de Documentos*

O **TrataDoc Studio** é uma ferramenta de automação voltada para a proteção de dados sensíveis e conformidade com a LGPD no âmbito do Ministério do Planejamento e Orçamento (MPO). O sistema utiliza Inteligência Artificial e Processamento de Linguagem Natural (NLP) para identificar, mapear e ocultar informações sensíveis em documentos públicos e processos administrativos da Corregedoria.

---

## 🚀 Principais Funcionalidades

* **⬛ Ocultação de Dados (Sanitização):** Mapeamento automatizado de CPFs, CNPJs, endereços, dados bancários, placas de veículos e nomes próprios. Permite a revisão manual flexível antes da aplicação definitiva das tarjas.
* **📄 OCR Pesquisável:** Motor de reconhecimento óptico de caracteres que transforma imagens e PDFs digitalizados em documentos pesquisáveis por texto, mantendo a integridade visual original.
* **🔗 Unificação de Arquivos:** Ferramenta para merge (junção) de múltiplos documentos PDF em um único volume, respeitando a ordem cronológica ou de importância definida pelo usuário.
* **👁️ Visualizador Integrado:** Interface para revisão de documentos em tempo real, com controle de zoom e ferramentas de tarjamento manual interativo.

---

## 🧠 Arquitetura Técnica e Inovação

Este projeto foi desenvolvido com uma arquitetura de **Motor Híbrido** para garantir máxima performance de acordo com o hardware do usuário:

1. **Motor Híbrido de OCR:**
   * **Ambiente Windows:** Utiliza o motor *Tesseract OCR* para garantir estabilidade e compatibilidade corporativa.
   * **Ambiente macOS:** Utiliza o *Apple Vision Framework* nativo. O sistema detecta a plataforma automaticamente e usa a aceleração de redes neurais dos chips Apple (M1/M2/M3) para uma leitura mais rápida e precisa, sem necessidade de dependências externas.
2. **Inteligência Artificial (NLP):** Implementado com a biblioteca *SpaCy* e o modelo `pt_core_news_md`, permitindo o reconhecimento avançado de entidades nomeadas (NER) para proteger nomes de servidores e interessados.
3. **Interface Fluida:** Desenvolvido em *CustomTkinter* com lógica de radar global para menus contextuais dinâmicos.

---

## 🛠️ Como Executar

### Versão Windows (.exe)
1. Baixe o instalador oficial (`Instalador_TrataDoc_v8_2.exe`) disponibilizado na aba de Releases.
2. Siga o assistente de instalação.
3. O software já configura automaticamente o motor OCR local e os modelos de Inteligência Artificial.

### Versão macOS (.app)
1. **Não requer instalação de motores externos.**
2. O sistema utiliza automaticamente o *Apple Vision Framework* nativo para processamento de OCR.
3. Basta baixar o aplicativo gerado, arrastar para a pasta de Aplicativos e executar.

---

## 👨‍💻 Desenvolvedor
**Anderson Guedes Francisco**
*Analista Técnico Administrativo - Corregedoria MPO*