# Bot de Automação: Extração de Certificados LMS Konviva

Este robô de RPA (Robotic Process Automation) foi desenvolvido para realizar o download massivo, auditado e organizado de certificados do portal LMS Konviva. Ele garante 100% de integridade dos dados através de um sistema de auditoria física e lógica.

---

## 🚀 Funcionalidades Principais

- **Auditoria Física Real**: O robô valida o sucesso do download contando os PDFs dentro dos arquivos ZIP antes de avançar para a próxima página. Se houver divergência, a operação é repetida automaticamente.
- **Compatibilidade Chrome 117+**: Utiliza comandos **CDP (Chrome DevTools Protocol)** para forçar o download direto em modo incôgnito, ignorando restrições recentes do navegador.
- **Organização Hierárquica**: Reconstrói automaticamente a estrutura de pastas (Pai > Filho > Neto) conforme a árvore organizacional do sistema.
- **Sistema de Checkpoint**: Salva o progresso em `checkpoint.json`, permitindo retomar extrações longas exatamente de onde pararam em caso de interrupção ou erro.
- **Resiliência Avançada**: Implementa lógica de `retry` para lidar com quedas de sessão, timeouts e carregamentos lentos da interface.
- **Configuração Multi-Ambiente**: Permite alternar facilmente entre diferentes URLs (ex: EAD, AVA) via arquivo JSON.

---

## 🛠️ Pré-requisitos

1.  **Python 3.8+** instalado.
2.  **Google Chrome** atualizado.
3.  **Dependências Python**:
    ```bash
    pip install selenium webdriver-manager requests
    ```

---

## ⚙️ Configuração (`configs.json`)

O bot é controlado pelo arquivo `configs.json`. Certifique-se de preenchê-lo corretamente:

```json
{
    "drive_letter": "D",  // Letra do drive onde os arquivos serão salvos
    "ambientes": {
        "ava": {
            "ativo": true,
            "url_login": "https://...",
            "url_listagem": "https://...",
            "credenciais": {
                "login": "seu_email",
                "senha": "sua_senha"
            },
            "unidades": ["566"] // IDs das unidades raiz. Deixe vazio [] para pegar todas.
        }
    }
}
```

---

## 📂 Estrutura do Projeto

- `main.py`: Script principal com toda a lógica de automação e auditoria.
- `configs.json`: Arquivo de configuração de credenciais e ambientes.
- `checkpoint.json`: Arquivo gerado automaticamente para controle de progresso.
- `README.md`: Documentação técnica e guia de uso.

---

## 📋 Fluxo de Operação

1.  **Login e Navegação**: O bot acessa o portal, realiza o login e navega até a página de listagem de certificados.
2.  **Expansão de Árvore**: Expande recursivamente toda a estrutura de unidades para identificar os IDs alvos.
3.  **Processamento por Página**:
    - Configura a visualização para 50 registros.
    - Seleciona os itens e inicia o processamento do ZIP.
    - **Ação Crítica**: Monitora o sistema de arquivos até que o ZIP seja baixado.
    - **Auditoria**: Abre o ZIP em memória, conta os arquivos e compara com a meta da página.
4.  **Finalização**: Após processar todas as unidades, cria um arquivo consolidado `Resultado_Final_Bot.zip` na raiz do drive configurado.

---

## 🔧 Manutenção e Solução de Problemas

- **Divergência de Contagem**: Se o bot detectar que o ZIP contém menos arquivos que o esperado, ele fechará o modal e tentará o download novamente.
- **Logs**: O console exibe o status em tempo real: "Meta Final da Unidade", "Pagina X confirmada" ou "Tentativa X falhou".
- **Erros Fatais**: Em caso de falha crítica, o navegador permanecerá aberto por 60 segundos com o erro impresso no terminal para facilitar a depuração.

---
**Desenvolvido para garantir máxima confiabilidade na migração de dados.**
