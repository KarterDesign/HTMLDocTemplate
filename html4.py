import pandas as pd
import os

# Função para criar o HTML de uma linha
def create_html(row, index, html_template):
    return html_template.format(
        title=f"Linha {index + 1}",
        nome_sku=row['_NomeSKU'],
        ean_sku=row['_EANSKU'],
        nome_produto=row['_NomeProduto'],
        cod_ref=row['_CodigoReferenciaProduto'],
        linktxt=row['_LinkTexto'],
        id_dep=row['_IDDepartamento'],
        nome_dep=row['_NomeDepartamento'],
        id_cat=row['_IDCategoria'],
        nome_cat=row['_NomeCategoria']
    )

# Carregar dados do Excel
file_path = r'C:\Users\Karter\Downloads\SKU Generator\Book4.xls'  # Substitua com o caminho do seu arquivo Excel
excel_data = pd.read_excel(file_path)

# Colunas a serem incluídas
columns_to_include = ['_NomeSKU', '_EANSKU', '_NomeProduto', '_CodigoReferenciaProduto', '_LinkTexto', '_IDDepartamento', '_NomeDepartamento', '_IDCategoria', '_NomeCategoria']
subset_data = excel_data[columns_to_include]

# Template HTML
html_template = """
        <div class="card">
            <h1>{nome_sku}</h1>
            <p><strong>EAN SKU:</strong> {ean_sku}</p>
            <p><strong>COD REF:</strong> {cod_ref}</p>
            <p><strong>LINK:</strong> {linktxt}</p>
            <p><strong>ID DEP:</strong> {id_dep}</p>
            <p><strong>NOME DEP:</strong> {nome_dep}</p>
            <p><strong>ID CAT:</strong> {id_cat}</p>
            <p><strong>NOME CAT:</strong> {nome_cat}</p>
        </div>
    """

html_content = """
<!DOCTYPE html>
<html>
<head>
    <title>Relatório Completo</title>
    <style>
        body {
            font-family: Barlow, sans-serif;
            margin: 0;
            padding: 0;
        }
        h1 {
        font-size: 23px;
        }
        .container {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            padding: 10px;
        }
        .card {
            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
            width: calc(30% - 100px); /* Ajuste a largura para 3 colunas */
            margin: 20px;
            border-radius: 5px;
            background-color: #ffffff;
            padding: 20px;
        }
        /* outros estilos */
    </style>
</head>
<body>
    <div class="container">
"""

# Acumular o conteúdo HTML de todas as linhas
for index, row in subset_data.iterrows():
    html_content += create_html(row, index, html_template)

# Fechar as tags HTML
html_content += """
</div>
</body>
</html>
"""

# Criar e salvar um único arquivo HTML
output_file = 'relatorio_completo.html'
with open(output_file, 'w', encoding='utf-8') as file:
    file.write(html_content)

print(f"Arquivo HTML criado com sucesso: {output_file}")