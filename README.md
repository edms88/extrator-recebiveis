# Extrator de Recebíveis — sem IA ou com GPT

Aplicação Streamlit privada, com uma aba Excel por tabela, valores tipados,
filtros, congelamento de cabeçalho e proteção contra fórmulas em textos.

## Modo padrão: sem IA

Selecione **Sem IA — modelo Recebíveis**. Não requer chave e não chama nenhuma API.
O PDF é enviado ao servidor Streamlit e processado por pdfplumber. Os arquivos
temporários são apagados ao final; Excel e JSON ficam na sessão para download.

O perfil é específico do relatório detalhado Power BI fornecido. Requer texto
selecionável, marcadores reconhecíveis e a seção detalhada Agregado Mensal.
No modelo fictício de cinco páginas, produz 13 abas. A página de resumo não é
extraída pelo perfil; o aviso solicita sua revisão. Contratos e Parcelas possuem
colunas cortadas no PDF: somente as colunas completas são extraídas. Linhas fora
da rolagem não são recuperadas e totais oficiais não são recalculados.

Não é um extrator universal: outros layouts, PDFs escaneados ou alterações nos
cabeçalhos podem exigir ajuste das regras ou uso explícito do modo GPT. O modo
local nunca faz fallback automático para IA. Valide amostras antes de uso operacional.

## Modo opcional: GPT — OpenAI

A antiga chamada Gemini foi removida. Usa o SDK oficial OpenAI, Responses API,
PDF inline, JSON estruturado validado por Pydantic e store=False. Não utiliza
esta conversa do ChatGPT. Nenhuma assinatura ou billing é ativado pela aplicação.

Cadastre pessoalmente em App settings → Secrets:

```toml
OPENAI_API_KEY = "CHAVE_INSERIDA_DIRETAMENTE_PELO_USUARIO"
OPENAI_MODEL = "MODELO_GPT_HABILITADO_NA_SUA_CONTA"
```

Escolha um modelo com suporte a PDF e structured outputs. O nome é configurável;
não há modelo imposto para evitar alterar a escolha de qualidade/custo da conta.
Uma chave colocada em GEMINI_API_KEY não é reaproveitada. Não envie chaves na conversa.
Não é necessário cadastrar Secrets para o modo sem IA.

O PDF no modo GPT é enviado à OpenAI. Antes de usar dados reais, confirme a aprovação
corporativa do provedor e do tratamento de dados. store=False não substitui governança.

## Executar

Python 3.12:
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

CLI sem IA:
```powershell
python extrair_recebiveis_gemini.py relatorio.pdf --engine local -o saida.xlsx --save-json auditoria.json
```
O nome histórico do módulo foi mantido para compatibilidade com imports e comandos.
Ele contém validações e formatação do Excel; não contém mais chamadas Gemini.
Para GPT use --engine gpt e defina OPENAI_API_KEY e OPENAI_MODEL no ambiente.

## Verificações

```powershell
python -m py_compile app.py extrair_recebiveis_gemini.py extracao_local.py extracao_openai.py
python -m unittest discover -s tests -v
python extrair_recebiveis_gemini.py --json exemplo_extracao.json -o validacao_saida.xlsx
```

14 testes automatizados, incluindo SDK OpenAI real com transporte HTTP simulado,
erros sanitizados, ausência de fallback para chave Gemini, rejeição de layout
desconhecido e conversões do Excel. O PDF fictício foi validado separadamente sem
rede: 13 abas, totais, datas, identificadores e posições de percentuais vazios.
Integração real GPT permanece pendente dos Secrets corretos e teste na conta.

## Publicação e segurança

Repositório privado: https://github.com/edms88/extrator-recebiveis
Aplicação privada: https://extrator-recebiveis-bvyduh7minrzyzwtrgu9eo.streamlit.app/
Branch main, arquivo app.py, Python 3.12. Preservar ambos privados.
Não convidar colaboradores sem lista autorizada. Google Sites somente como portal,
após indicação explícita do site: botão “Extrair relatório de recebíveis” em nova aba.

Limites: 10 MB e 10 páginas. PDF deve ter assinatura %PDF-. Sem cache compartilhado
de documentos. .streamlit/secrets.toml, PDFs e arquivos temporários excluídos do Git.

## Referências oficiais

- https://developers.openai.com/api/docs/guides/file-inputs
- https://developers.openai.com/api/docs/guides/structured-outputs
