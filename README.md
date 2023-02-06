# SOFTWARE ESTUDO DE PROTEÇÃO DE MT

![Badge em Desenvolvimento](http://img.shields.io/static/v1?label=STATUS&message=EM%20DESENVOLVIMENTO&color=GREEN&style=for-the-badge)

## 🔨 Funcionalidades do projeto
- Autenticar usuários
- Ler a Planilha do Google do usuário logado
- Pegar os valores da aba Grafico e Curvas_fusíveis
- Criar um gráfico de escala logarítmica a partir desses valores
- Upar a imagem com o gráfico no Drive

## ✔️ Técnicas e tecnologias utilizadas
- `Python`
- `Matplotlib`
- `PySimpleGUI`
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
code .
```

### Agora no terminal integrado, execute os comandos:

```bash
python3 -m venv .venv #Windows: python -m venv .venv
```

```bash
source .venv/bin/activate #Windows: .venv\Scripts\activate
```

```bash
pip install -r requirements.txt
```