# Documentação Técnica: Bot de Automação de Migração de Certificados

Este documento descreve o funcionamento, a arquitetura e as regras de negócio do robô de RPA (Robotic Process Automation) desenvolvido para a extração e organização de certificados do portal LMS Konviva.

---

## 1. Escopo e Objetivo
O robô foi projetado para realizar o download massivo de certificados de forma resiliente, garantindo que **100% dos registros** disponíveis no sistema sejam baixados e organizados localmente sem intervenção humana.

### Melhorias Implementadas (Contexto Atual):
- **Auditoria Física de Arquivos**: O robô não confia apenas no clique; ele conta fisicamente quantos PDFs existem na sua pasta antes e depois de cada download. O avanço de página só ocorre se o saldo de arquivos no HD bater com o número de registros selecionados.
- **Meta Dinâmica por Página**: O bot calcula automaticamente quantos registros existem em cada tela (ex: 100, 50 ou 28 na última página) através do texto informativo da tabela, garantindo precisão matemática em cada lote.
- **Hierarquia de Pastas (Árvore Organizacional)**: Organização automática de pastas seguindo a estrutura Pai > Filho > Neto do sistema original.
- **Extração e Limpeza Automática**: Baixa arquivos ZIP, extrai os PDFs soltos para a pasta da unidade e deleta os arquivos ZIP temporários imediatamente.
- **Supressão de Interrupções do Chrome**: Configuração de RPA para permitir múltiplos downloads automáticos sem exibir popups de confirmação do navegador.
- **Logs Limpos e Técnicos**: Mensagens de log em texto simples (sem emojis) para melhor compatibilidade com diferentes terminais e maior clareza visual.
- **Segurança contra Fechamento Súbito**: Em caso de erro fatal, o bot imprime o diagnóstico técnico completo e mantém a janela do navegador aberta por 60 segundos antes de encerrar, permitindo a inspeção da falha.

---

## 2. Fluxo de Execução Detalhado

1.  **Configuração Inicial**: Carrega credenciais e URLs do `configs.json`. Inicializa o Selenium com preferências de download automático.
2.  **Autenticação**: Realiza o login e trata possíveis redirecionamentos automáticos para o dashboard, garantindo a chegada na página de certificados.
3.  **Mapeamento de Unidades**: Expande recursivamente toda a árvore organizacional e identifica os IDs das unidades alvo (ou captura todas, se a lista estiver vazia).
4.  **Processamento de Unidade (Auditoria Tripla)**:
    - **Passo A**: Verifica se a unidade possui registros e reconstrói o caminho das pastas localmente.
    - **Passo B (Loop de Páginas)**:
        1. Identifica a meta da página atual (ex: registros 201 a 300 = meta 100).
        2. Seleciona os checkboxes e aguarda a renderização completa.
        3. Realiza o download do ZIP e extrai os PDFs.
        4. Compara o número de novos arquivos no HD com a meta calculada.
    - **Passo C**: Se a auditoria falhar, repete a mesma página. Se tiver sucesso, salva o checkpoint e avança.
5.  **Consolidação**: Ao final de todas as unidades, agrupa toda a estrutura de pastas em um arquivo único `Certificados_Final_Hierarquico.zip`.

---

## 3. Estrutura de Funções e Lógica

### 3.1 Inteligência de Lote e Auditoria
- `processar_paginas_da_unidade()`: Gerencia o contador físico de arquivos e a meta dinâmica. É o núcleo que garante o "zero perda" de dados.
- `contar_pdfs_na_pasta()`: Função auxiliar que interage com o Sistema Operacional para validar o sucesso real da operação de download.
- `obter_info_paginacao()`: Usa Regex para interpretar dinamicamente a contagem de registros exibida no elemento `DataTables_Table_0_info`.

### 3.2 Resiliência e Recuperação
- `executar_com_retry()`: Em caso de falha de carregamento ou rede, o bot tenta reiniciar o estado da busca até 5 vezes antes de reportar erro crítico.
- `ARQUIVO_CHECKPOINT`: Mantém o registro da última página e unidade concluída com sucesso, permitindo retomar o trabalho pesado sem duplicidade.

### 3.3 Navegação RPA
- `avancar_pagina()`: Utiliza múltiplos seletores para encontrar o botão "Próximo" e valida se a página mudou comparando o texto informativo anterior e atual.
- `expandir_arvore()`: Algoritmo recursivo que garante que todos os níveis da estrutura organizacional fiquem visíveis para captura de IDs.

---

## 4. Configurações e Variáveis Globais

| Variável | Valor Padrão | Descrição |
| :--- | :--- | :--- |
| `TIMEOUT_PADRAO` | 30s | Tempo de espera para elementos comuns. |
| `TIMEOUT_DOWNLOAD` | 300s | Tempo máximo para geração/baixa de ZIPs grandes. |
| `PAUSA_ENTRE_ACOES` | 2s | Delay de segurança para estabilidade da UI. |
| `Lote de Registros` | 50 | Quantidade fixa por página para máxima compatibilidade. |

---

## 5. Manutenção e Erros Comuns

- **O Navegador fecha rápido demais?** O bot agora espera 60 segundos após erros fatais. Verifique o log no terminal para ver o erro técnico.
- **Meta não bate?** O sistema de auditoria física reiniciará a página automaticamente se o download vier corrompido ou incompleto.
- **IDs não encontrados?** Certifique-se de que a unidade no `configs.json` está expandida na árvore organizacional.

---
**Documentação consolidada para garantir a rastreabilidade e integridade do processo de migração.**
