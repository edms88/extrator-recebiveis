# Extrator de Recebíveis — Streamlit + Gemini

Aplicação web em Python para transformar as tabelas visíveis de relatórios de
recebíveis em um arquivo Excel organizado, com uma aba para cada tabela.

O Excel mantém o padrão das entregas anteriores:

- título azul Fundacred `#00308F`;
- fonte e páginas de origem;
- cabeçalho cinza e linhas alternadas;
- filtros e cabeçalho congelado;
- moedas, percentuais, números e datas como valores tipados;
- subtotais e totais destacados;
- limitações visuais registradas na própria aba.

## Arquivos principais

| Arquivo | Finalidade |
|---|---|
| `app.py` | Interface web executada pelo Streamlit |
| `extrair_recebiveis_gemini.py` | Extração, validação e geração do Excel |
| `requirements.txt` | Dependências do projeto |
| `.streamlit/config.toml` | Tema visual e limite de upload |
| `exemplo_extracao.json` | Dados de exemplo para teste sem API |
| `tests/test_core.py` | Testes de conversão e geração do workbook |

## Executar localmente

Use Python 3.12 (versão usada na validação).

### Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:GEMINI_API_KEY="COLE_A_CHAVE_APENAS_NESTA_SESSAO"
streamlit run app.py
```

### Linux ou macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
export GEMINI_API_KEY="COLE_A_CHAVE_APENAS_NESTA_SESSAO"
streamlit run app.py
```

Não grave a chave no código, no GitHub, no AppSheet ou no Google Sheets.

## Publicar no Streamlit Community Cloud

1. Crie um repositório privado no GitHub.
2. Envie o conteúdo desta pasta para a raiz do repositório.
3. Acesse `https://share.streamlit.io/` e clique em **Create app**.
4. Selecione o repositório, a branch e o arquivo principal `app.py`.
5. Em **Advanced settings > Secrets**, cadastre:

```toml
GEMINI_API_KEY = "SUA_CHAVE_DO_GEMINI"
```

6. Publique o aplicativo.
7. Nas configurações de compartilhamento, mantenha o aplicativo privado e convide
   somente os e-mails autorizados.

O arquivo `.streamlit/secrets.toml` está bloqueado pelo `.gitignore`. Não remova
essa proteção e não coloque a chave diretamente no repositório.

## Colocar no Google Sites

Para aplicativo privado, prefira um botão no Google Sites que abra o endereço
`https://seu-app.streamlit.app` em uma nova aba. Incorporar o aplicativo em um
iframe pode causar falhas no login por bloqueio de cookies de terceiros.

O Google Sites deve ser restrito aos mesmos usuários autorizados no Streamlit.
Restringir somente o Site não protege a URL direta da aplicação.

## Testes

Os testes não chamam a Gemini API:

```bash
python -m unittest discover -s tests -v
```

Também é possível gerar um Excel sem consumir a API:

```bash
python extrair_recebiveis_gemini.py \
  --json exemplo_extracao.json \
  --output exemplo_saida.xlsx
```

## Limites e governança

**Status desta entrega: aprovado para teste anonimizado; uso de dados reais
bloqueado por governança até confirmação da modalidade e aprovação corporativa.**
Não ative billing ou assinatura para remover esse bloqueio sem autorização expressa.

O modelo é configurável por `GEMINI_MODEL` nos Secrets do Streamlit ou no ambiente.
O padrão é `gemini-3.8-flash`, documentado atualmente para PDF e JSON estruturado.
Se a conta não tiver acesso, o administrador pode trocar essa configuração sem
editar o código. A disponibilidade na conta ainda precisa de teste real.
As chamadas usam `store=False`, sem histórico de interações; isso não substitui
a aprovação dos termos de tratamento de dados da Gemini API.

Referências oficiais consultadas em 05/09/2026:
- https://ai.google.dev/gemini-api/docs/document-processing
- https://ai.google.dev/gemini-api/docs/structured-output
- https://ai.google.dev/gemini-api/docs/interactions-overview
- https://ai.google.dev/gemini-api/terms

Cadastre a chave pessoalmente no ambiente seguro; nunca a envie na conversa.
Na publicação, configure Secrets antes de concluir o deploy. Confirme o repositório
e o aplicativo como privados e não convide pessoas sem a lista autorizada.
O portal Google Sites permanece pendente da URL publicada e da indicação explícita
do site que poderá ser alterado. O botão deve se chamar “Extrair relatório de recebíveis”.

O piloto limita cada PDF a 10 páginas e 10 MB. Arquivos temporários são apagados
ao final de cada processamento e não há cache compartilhado dos documentos.

O schema valida a estrutura do resultado, mas não garante que todos os valores
extraídos estejam semanticamente corretos. Compare amostras do Excel com o PDF
antes de usar o resultado como dado oficial.

Antes de processar relatórios reais, confirme se o projeto associado à chave está
no nível pago/corporativo e se o tratamento foi aprovado pela Fundacred. A licença
corporativa do chat Gemini não garante, por si só, que a Gemini API esteja em
modalidade paga.

## Execução por linha de comando

A interface web é opcional. O backend continua aceitando execução direta:

```bash
python extrair_recebiveis_gemini.py "Relatorio Recebiveis.pdf" \
  --output "Relatorio Recebiveis_tabelas.xlsx" \
  --save-json "Relatorio Recebiveis_auditoria.json"
```
