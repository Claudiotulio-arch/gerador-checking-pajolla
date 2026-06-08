# Gerador de Checking Pajolla

Versão 1.1 corrigida: ajuste para gerar vários slides sem erro no PowerPoint.

Aplicativo web gratuito em Streamlit para gerar PPTX de checking a partir do modelo padrão da Pajolla.

## Como rodar no computador

1. Instale o Python 3.10 ou superior.
2. Abra o terminal dentro desta pasta.
3. Execute:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como publicar grátis no Streamlit Community Cloud

1. Crie uma conta no GitHub.
2. Crie um repositório, por exemplo `gerador-checking-pajolla`.
3. Envie estes arquivos para o repositório:
   - `app.py`
   - `requirements.txt`
   - `template_checking_pajolla.pptx`
4. Acesse o Streamlit Community Cloud.
5. Clique em **New app**.
6. Selecione o repositório e o arquivo `app.py`.
7. Clique em **Deploy**.

Depois disso a equipe acessa por um link, preenche os campos, sobe as fotos e baixa o PPTX pronto.

## Observações

- O app mantém a capa do modelo.
- Usa 1 foto por slide.
- As fotos são inseridas sem deformar.
- O layout base fica no arquivo `template_checking_pajolla.pptx`.
