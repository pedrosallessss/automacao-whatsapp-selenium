# 🤖 Bot de Monitoria - Automação de WhatsApp (RPA)

> **Status:** ✅ Concluído

Este projeto é uma ferramenta de **RPA (Robotic Process Automation)** desenvolvida para otimizar o processo de comunicação com alunos. O sistema lê uma base de dados em Excel, filtra os agendamentos do dia e envia lembretes personalizados via WhatsApp Web automaticamente, incluindo imagens e texto.

---

## 🎯 O Problema
O processo anterior envolvia verificar planilhas manualmente e enviar mensagens "uma a uma" para dezenas de alunos. Isso gerava:
- Alto consumo de tempo da equipe.
- Risco de erros humanos (esquecer alunos ou trocar horários).
- Falta de padronização na comunicação.

## 💡 A Solução
Desenvolvi um script em **Python** que atua como um assistente virtual. Ele automatiza 100% do fluxo de envio, mantendo um "humano no comando" apenas para a conferência final antes do disparo.

### Principais Funcionalidades:
- **📊 Leitura de Dados:** Integração com Excel (`Pandas`) para ler nomes, telefones e horários.
- **🔍 Filtro Inteligente:** Identifica automaticamente o dia da semana atual e filtra os alunos do próximo horário.
- **🖼️ Envio de Mídia via Clipboard:** Utiliza a API do Windows (`PyWin32`) para copiar imagens para a área de transferência e colálas no WhatsApp (Ctrl+V), garantindo compatibilidade e contornando falhas de botões de upload tradicionais.
- **💾 Sessão Persistente:** Salva o perfil do navegador (`Selenium Profiles`) para que o login no WhatsApp seja feito apenas uma vez, sem necessidade de escanear QR Code a cada execução.
- **⌨️ Simulação Humana:** Utiliza `ActionChains` para digitar mensagens e interagir com a interface de forma natural.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **Selenium WebDriver:** Orquestração e controle do navegador Chrome.
- **Pandas:** Manipulação e tratamento de dados (ETL).
- **PyWin32 (win32clipboard):** Manipulação da área de transferência do Windows.
- **Pillow (PIL):** Processamento de imagens.
- **Webdriver Manager:** Gerenciamento automático dos drivers do navegador.

---

## 📂 Estrutura do Projeto

```bash
📂 automacao-whatsapp
├── 📁 Lembretes           # Pasta contendo as imagens (.jpg/.png) para envio
├── 📁 Perfil_Zap          # Pasta criada automaticamente para salvar o login (não subir no git)
├── 📄 robo_final.py       # Código fonte principal
├── 📄 requirements.txt    # Lista de dependências
├── 📄 dados_monitoria.xlsx # (Arquivo real - Não versionado por segurança)
└── 📄 exemplo_base_dados.xlsx # Arquivo de exemplo para teste
🚀 Como Executar o Projeto
Pré-requisitos
Python instalado e configurado no PATH.

Google Chrome instalado.

Passo a Passo
Clone o repositório:

Bash

git clone [https://github.com/SEU-USUARIO/NOME-DO-REPO.git](https://github.com/SEU-USUARIO/NOME-DO-REPO.git)
Instale as bibliotecas necessárias:

Bash

pip install -r requirements.txt
Configure a Planilha:

Utilize o arquivo exemplo_base_dados.xlsx como modelo.

Crie um arquivo dados_monitoria.xlsx com os dados reais dos alunos.

Adicione as Imagens:

Coloque as imagens que deseja enviar na pasta Lembretes. O robô escolherá uma aleatoriamente a cada envio (ou a mesma, se só houver uma).

Execute o Script:

Bash

python robo_final.py
⚠️ Aviso Legal e Ética
Este software foi desenvolvido para fins de produtividade interna e aprendizado.

Não faça SPAM: O uso de automações para envio em massa não solicitado viola os termos de serviço do WhatsApp.

LGPD: Certifique-se de que os dados utilizados na planilha estão em conformidade com as leis de proteção de dados.

👤 Autor
Desenvolvido por Pedro Henrique Salles
