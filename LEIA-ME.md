# Painel de execução orçamentária – CBMSC (FUMCBM)

App Streamlit que lê os arquivos da pasta `dados/` e mostra o painel.

## Arquivos
- `app.py` – programa do painel (não precisa mexer).
- `painel_template.html` – visual do painel (não precisa mexer).
- `requirements.txt` – bibliotecas que o Streamlit instala sozinho.
- `dados/DD-MM-AAAA - analisedespesa.xlsx` – análise de despesa do SIGEF. A data no nome vira a "posição" do painel.
- `dados/preempenhos.xlsx` – pré-empenhos em aberto. Colunas: documento, subacao, fonte, natureza, a_empenhar, objeto.

## Para atualizar
1. No GitHub, abra a pasta `dados`.
2. Apague o xlsx antigo do SIGEF (ou deixe: o painel usa o de data mais recente no nome).
3. Clique em Add file > Upload files e envie o novo, com a data no nome.
4. Atualize `preempenhos.xlsx` da mesma forma.
5. Clique em Commit changes. Em cerca de 1 minuto o painel mostra os dados novos.
