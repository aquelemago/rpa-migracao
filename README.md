# Bot de Automacao: Extracao de Certificados LMS Konviva (RPA High-Resilience)

Este robô de RPA (Robotic Process Automation) foi projetado para extração massiva, auditada e altamente resiliente de certificados do portal LMS Konviva. O foco principal deste projeto é a integridade absoluta dos dados e a capacidade de operação contínua sem intervenção humana, lidando com instabilidades do portal de forma autônoma.

---

## 🛠️ Especificações Técnicas e Workflow

O robô segue um fluxo de execução rigoroso para garantir que nenhum registro seja perdido:

1.  **Login e Autenticação**: Acessa a URL de login, insere credenciais e aguarda a transição para o portal.
2.  **Navegação e Preparação**:
    - Navega para a listagem de certificados.
    - Abre a busca avançada e expande recursivamente a árvore organizacional.
    - Identifica todas as unidades descendentes da unidade raiz configurada.
3.  **Processamento de Unidade (Loop Infinito de Sucesso)**:
    - **Seleção**: Seleciona a unidade específica na árvore e dispara a busca.
    - **Configuração**: Define a exibição para 50 registros por página para otimizar o download.
    - **Workflow de Página**:
        - **Selecionar Tudo**: Utiliza o checkbox de cabeçalho (`selectAll`) para marcar todos os registros visíveis.
        - **Validação de Seleção**: Compara o número de registros selecionados com a meta indicada pela paginação do sistema. Se houver divergência, realiza refresh e reinicia a página.
        - **Download e Processamento**: Aciona o modal de download, valida o total processado pelo servidor e aguarda a geração do arquivo ZIP.
4.  **Auditoria Física (Crítico)**:
    - Após o download, o bot abre o arquivo ZIP em memória.
    - Conta o número de arquivos PDF contidos no ZIP.
    - **Critério de Aceite**: O número de PDFs deve ser EXATAMENTE igual ao número de registros selecionados na página.
    - **Falha na Auditoria**: Se o ZIP vier incompleto, o bot remove o arquivo defeituoso, atualiza a página e tenta novamente. **Não há limite de tentativas; o bot só avança quando a auditoria física for 100% bem-sucedida.**
5.  **Checkpoint e Finalização**:
    - Salva o progresso em `checkpoint.json` após cada página confirmada.
    - Ao final, consolida todos os ZIPs organizados por pastas em um arquivo único `Resultado_Final_Bot.zip`.

---

## ⚙️ Configuração (`configs.json`)

```json
{
    "drive_letter": "D",  // Letra do drive para armazenamento temporário e final
    "ambientes": {
        "ava": {
            "ativo": true,
            "url_login": "https://...",
            "url_listagem": "https://...",
            "credenciais": {
                "login": "usuario",
                "senha": "senha"
            },
            "unidades": ["566"] // ID da unidade pai para iniciar a busca
        }
    }
}
```

---

## 📋 Regras de Ouro para Manutenção (Instruções para IA)

Para evitar alucinações ou erros de lógica em manutenções futuras, siga estas diretrizes:

- **Sem Limite de Retentativas**: Nunca implemente `MAX_TENTATIVAS`. O bot deve usar loops `while True` para garantir finalização apenas por sucesso.
- **Auditoria é Obrigatória**: Qualquer alteração no fluxo de download deve passar pela função `contar_pdfs_nos_zips`. O sucesso lógico do Selenium (clique no botão) não garante o sucesso físico (arquivo no disco).
- **Logs Detalhados**: Mantenha o padrão de logs com prefixos `[LOG]`, `[STATUS]`, `[AVISO]` e `[ERRO]`. Não utilize emojis no terminal para garantir compatibilidade com diferentes encodings.
- **Workflow de Seleção**: O bot deve sempre tentar o seletor `selectAll` no cabeçalho antes de tentar seletores individuais. A performance e estabilidade do portal são melhores com a seleção em massa nativa.
- **Resiliência do Driver**: Sempre envolva `driver.quit()` e `driver.refresh()` em blocos `try-except` para evitar quebras por perda de conexão com o socket do Chrome.

---

## 📦 Dependências

- Python 3.8+
- Selenium
- Webdriver-Manager (Automated Chrome Driver)

Instalação:
```bash
pip install selenium webdriver-manager
```

---
**Desenvolvido para garantir 100% de migração de dados em ambientes instáveis.**
