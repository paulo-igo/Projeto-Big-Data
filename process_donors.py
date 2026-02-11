import pandas as pd
import re

# Carregar a planilha original
file_path = '/home/ubuntu/upload/doadores23-25.xlsx'
df = pd.read_excel(file_path)

# Lista de bairros de Goiânia (baseada na pesquisa e conhecimento geral)
# Vou incluir os mais comuns e os que aparecem na planilha que são claramente de Goiânia.
# A lista da Wikipedia é extensa, vou usar uma abordagem de limpeza e verificação.

bairros_goiania = [
    "Adriana Park", "Aeroporto", "Aeroporto Sul", "Agua Branca", "Agua Santa", "Alcira De Resende", 
    "Alcira De Rezende", "Aldeia Do Vale", "Alice Barbosa", "Almerinda Rezende", "Alphaville", 
    "Alto Alegre", "Alto Boa Vista 2", "Alto Da Boa Vista 2", "Alto Da Gloria", "Alto Do Bom Fim", 
    "Alvina Paniago", "Alvorada", "American Park", "Amin Camargo", "Ana Rosa", "Angelina Peixoto", 
    "Anhanguera", "Antonio Rabelo", "Anuar Auad", "Aracy Amaral", "Araguaia Park", "Arco Iris", 
    "Assentamento Canudos", "Aurora Das Mansoes", "Bairro Da Vitoria", "Bairro Feliz", "Bairro Goiá",
    "Bairro Rodoviário", "Bela Vista", "Bueno", "Campinas", "Castelo Branco", "Centro", "Cidade Jardim",
    "Coimbra", "Criméia Leste", "Criméia Oeste", "Faiçalville", "Finsocial", "Gentil Meireles",
    "Goiânia 2", "Goiânia Viva", "Grajaú", "Guanabara", "Itatiaia", "Jaó", "Jardim América",
    "Jardim Balneário Meia Ponte", "Jardim Curitiba", "Jardim Europa", "Jardim Goiás", "Jardim Guanabara",
    "Jardim Novo Mundo", "Jardim Oliveira", "Marista", "Negrão de Lima", "Nova Suíça", "Oeste",
    "Parque Amazônia", "Parque Atheneu", "Parque das Laranjeiras", "Pedro Ludovico", "Setor Sul",
    "Setor Universitário", "Sudoeste", "Urias Magalhães", "Vila Nova", "Vila Rosa"
]

# Note: "Anápolis City" não é em Goiânia (é em Anápolis).
# Vou filtrar removendo os que contêm nomes de outras cidades conhecidas de GO.
outras_cidades = ["Anapolis", "Anápolis", "Aparecida", "Trindade", "Senador Canedo", "Guapó", "Goianira"]

def is_goiania(bairro):
    if pd.isna(bairro): return False
    bairro_str = str(bairro)
    for cidade in outras_cidades:
        if cidade.lower() in bairro_str.lower():
            return False
    # Se estiver na lista conhecida ou não for de outra cidade óbvia, mantemos por enquanto
    # (A lista completa de bairros de Goiânia tem mais de 600 nomes)
    return True

# Aplicar filtro
df_filtered = df[df['BAIRRO'].apply(is_goiania)].copy()

# Contagem de doadores por bairro (Top 10)
bairros_count = df_filtered['BAIRRO'].value_counts().head(15).reset_index()
bairros_count.columns = ['Bairro', 'Doadores']

# Contagem por tipo sanguíneo
blood_count = df_filtered['TIPO SANGUINEO'].value_counts().reset_index()
blood_count.columns = ['Tipo Sanguíneo', 'Doadores']

# Salvar dados processados para uso no Excel
with pd.ExcelWriter('/home/ubuntu/processed_data.xlsx') as writer:
    df_filtered.to_excel(writer, sheet_name='Dados Filtrados', index=False)
    bairros_count.to_excel(writer, sheet_name='Resumo Bairros', index=False)
    blood_count.to_excel(writer, sheet_name='Resumo Tipos Sanguineos', index=False)

print("Dados processados e salvos em /home/ubuntu/processed_data.xlsx")
