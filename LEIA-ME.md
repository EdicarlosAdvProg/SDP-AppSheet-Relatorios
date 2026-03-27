# Manual de Manutenção e Configuração: Sistema SDP OAB-GO

Este documento contém as instruções críticas para a operação, segurança e integração com inteligência artificial (Gemini) do **Sistema de Defesa das Prerrogativas (SDP)**.

## 1. Configuração da Inteligência Artificial (Gemini API)

O sistema utiliza o motor **Gemini 3 Flash (v1beta)** para análise de processos. Por segurança e conformidade com as normas do Google Cloud (2026), a integração exige um vínculo triplo: **Chave**, **Projeto Cloud** e **Permissão de Usuário**.

### Passo a Passo de Configuração (Obrigatório em Migrações):

1.  **Gerar a Chave (O Cérebro):**
    * Acesse o [Google AI Studio](https://aistudio.google.com/) com a conta da OAB-GO.
    * Clique em **"Get API key"** e **"Create API key in new project"**. Copie a chave (`AIza...`).

2.  **Vincular ao Google Cloud (A Identidade):**
    * No [Google Cloud Console](https://console.cloud.google.com/), selecione o projeto oficial (**SDP relatorios**).
    * Copie o **Número do Projeto** (Project Number) no Dashboard.
    * No Apps Script, vá em **Configurações do Projeto** ⚙️ > **Projeto do Google Cloud Platform (GCP)** > **Alterar projeto** e cole o número.

3.  **Ativar a API (O Disjuntor):**
    * No Cloud Console, pesquise por **"Gemini API"** no Marketplace e clique em **ATIVAR**. Sem isso, o sistema retornará *Erro 404*.

4.  **Autorizar Usuários (O Público-Alvo):**
    * No Cloud Console, vá em **APIs e Serviços** > **Tela de consentimento OAuth**.
    * Em **Usuários de teste**, adicione o e-mail de quem executará o script (ex: `secretaria@oabgo.org.br`). Sem isso, o sistema retornará *Erro 403 (Access Denied)*.

5.  **Gravar a Chave no Script:**
    * No Apps Script, em **Configurações do Projeto** ⚙️ > **Propriedades do script**, adicione:
        * **Propriedade:** `GEMINI_KEY`
        * **Valor:** (A chave copiada no passo 1).

## 2. Arquitetura de Ficheiros
O projeto segue um padrão rigoroso de organização para facilitar a manutenção:
* `*-macro.gs`: Lógica de servidor e comunicação com a API Gemini (Backend).
* `*-layout.html`: Estrutura visual baseada em Materialize CSS.
* `*-style.html`: Estilização e Identidade Visual institucional OAB-GO.
* `*-js.html`: Comportamento da interface e chamadas assíncronas (Frontend).
* `Startup.gs`: Definição de IDs de planilhas, variáveis globais e inicialização do menu.

## 3. Base de Dados
O sistema utiliza o Google Sheets como banco de dados relacional através das abas:
* `tabProcessos`: Registro de violações e dados brutos.
* `tabMembros`: Cadastro de advogados, cargos e comissões.
* `tabVotos`: Histórico de deliberações, pareceres e acórdãos.

## 4. Transferência de Propriedade (Migration)
Ao transferir este projeto para a OAB-GO ou nova conta:
1.  **Propriedade:** Transfira a pasta do Google Drive contendo o Doc, a Planilha e o Script.
2.  **Sincronia:** Verifique se o ID da planilha em `Startup.gs` (`PLANILHA_DADOS_ID`) aponta para o arquivo correto.
3.  **Reset de Chaves:** As "Propriedades do Script" **não são copiadas**. Você deve refazer o **Item 1** deste manual integralmente na nova conta.

## 5. Suporte Técnico e Cotas
* **Erro 429 (Quota Exceeded):** No nível gratuito, o Google limita o número de requisições por minuto. Caso ocorra, aguarde 30-60 segundos.
* **Modelo Homologado:** O sistema está configurado para usar `models/gemini-flash-latest`. Se este modelo for descontinuado, verifique a lista de modelos autorizados no log de diagnóstico do script.

---
**Desenvolvido para:** Sistema de Defesa das Prerrogativas - OAB-GO  
**Padrão Técnico:** Google Apps Script / Google Cloud Platform / Gemini AI 3.0  
**Data da última homologação:** 26 de Março de 2026