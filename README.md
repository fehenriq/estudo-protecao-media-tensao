# SOFTWARE ESTUDO DE PROTEÇÃO DE MT

![Badge em Desenvolvimento](http://img.shields.io/static/v1?label=STATUS&message=EM%20DESENVOLVIMENTO&color=GREEN&style=for-the-badge)

## 🔨 Funcionalidades do projeto
- Ler a [Sheets do estudo](https://docs.google.com/spreadsheets/d/1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198/edit#gid=1455473048)
- Pegar os valores da aba Grafico e Curvas_fusíveis
- Criar um gráfico de escala logarítmica a partir desses valores
- Upar a imagem com o gráfico no Drive

## ✔️ Técnicas e tecnologias utilizadas
- `Python`
- `Matplotlib`
- `API do Google Sheets`
- `API do Google Drive`


## 🛠️ Para abrir e rodar o projeto

### Lembre-se de ter seu arquivo de credenciais do Google na raiz do projeto

```bash
git clone https://github.com/fehenriq/coordenograma-estudo-protecao-mt.git
```

```bash
cd coordenograma-estudo-protecao-mt
```

```bash
python3 -m venv .venv
```

```bash
source .venv/bin/activate
```

```bash
pip install -r requirements.txt
```

## 📝 Melhorias
- Acionador para ocultar e mostrar páginas ✔️
- Gráfico logarítmico ✔️
- Exportação e importação do projeto ✔️
- Bloquear acesso ao Google Apps Script ❌
