import matplotlib.pyplot as plt
from datetime import datetime
from lerplanilha import ler_planilha, pd

ARQUIVO_ENTRADA = "vendas.csv"
ARQUIVO_SAIDA = f"relatorio_vendas_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"


print("📊 Carregando planilha...")


df = ler_planilha(ARQUIVO_ENTRADA)

# Garante que a coluna de data esteja no formato correto

df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y')

# Calcula valor total da venda

df['Total_Venda'] = df['Quantidade'] * df['Preco_Unitario']
print(f"✅ {len(df)} vendas carregadas!")
print("🔢 Gerando análises...")

# Totais Gerais

total_geral = df['Total_Venda'].sum()
quantidade_total = df['Quantidade'].sum()
ticket_medio = total_geral / len(df)

# Totais por produto

vendas_produto = df.groupby('Produto').agg({
    'Total_Venda': 'sum',
    'Quantidade': 'sum'
}).sort_values('Total_Venda', ascending=False).reset_index()

# 3. Por vendedor

vendas_vendedor = df.groupby('Vendedor')['Total_Venda'].sum().sort_values(ascending=False).reset_index()

# 4. Por dia

vendas_dia = df.groupby(df['Data'].dt.date)['Total_Venda'].sum().reset_index()
vendas_dia.columns = ['Data', 'Total_Venda']

# ===================== GERA GRÁFICOS =====================

print("📈 Criando gráficos...")

plt.figure(figsize=(10, 6))
plt.bar(vendas_produto['Produto'][:5], vendas_produto['Total_Venda'][:5], color='#4CAF50')
plt.title('Top 5 Produtos - Vendas (R$)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('grafico_top_produtos.png')

plt.figure(figsize=(8, 6))
plt.pie(vendas_vendedor['Total_Venda'][:5], labels=vendas_vendedor['Vendedor'][:5], autopct='%1.1f%%')
plt.title('Participação de Vendas por Vendedor')
plt.savefig('grafico_vendedores.png')

plt.figure(figsize=(10, 5))
plt.plot(vendas_dia['Data'], vendas_dia['Total_Venda'], marker='o', color='#2196F3')
plt.title('Evolução de Vendas por Dia')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('grafico_evolucao.png')

# ===================== SALVA RELATÓRIO EXCEL =====================
print("💾 Salvando relatório Excel...")

with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:

    # Aba 1 - Resumo

    resumo = pd.DataFrame({
        'Métrica': ['Total de Vendas (R$)', 'Quantidade de Itens', 'Ticket Médio (R$)', 'Número de Vendas'],
        'Valor': [f"R$ {total_geral:,.2f}", quantidade_total, f"R$ {ticket_medio:,.2f}", len(df)]
    })
    resumo.to_excel(writer, sheet_name='Resumo', index=False)

    # Aba 2 - Vendas por Produto

    vendas_produto.to_excel(writer, sheet_name='Por Produto', index=False)

    # Aba 3 - Detalhes completos

    df.to_excel(writer, sheet_name='Detalhes', index=False)

print(f"🎉 Relatório pronto: {ARQUIVO_SAIDA}")
print("   Gráficos salvos: grafico_top_produtos.png, grafico_vendedores.png, grafico_evolucao.png")

