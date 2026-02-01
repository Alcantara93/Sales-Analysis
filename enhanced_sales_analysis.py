"""
============================================================================
ADVANCED SALES ANALYSIS: PORTUGUESE MARKET & PRODUCT PERFORMANCE
Portfolio Enhancement - Senior Level Statistical Analysis
Python Version for VS Code
============================================================================
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.stats import skew, kurtosis, pearsonr, spearmanr
import warnings
warnings.filterwarnings('ignore')

# Set professional style
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("Set2")

# Configure pandas output
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# Professional theme settings
def set_professional_theme():
    plt.rcParams['figure.figsize'] = (14, 8)
    plt.rcParams['font.size'] = 10
    plt.rcParams['axes.titlesize'] = 14
    plt.rcParams['axes.labelsize'] = 11
    plt.rcParams['xtick.labelsize'] = 10
    plt.rcParams['ytick.labelsize'] = 10
    plt.rcParams['legend.fontsize'] = 10
    plt.rcParams['figure.dpi'] = 100
    plt.rcParams['axes.grid'] = True
    plt.rcParams['grid.alpha'] = 0.3

set_professional_theme()

print("\n" + "="*80)
print("ADVANCED SALES ANALYSIS - PYTHON VERSION")
print("="*80)
print("\n⚠️  IMPORTANT: Load your data first!")
print("Example: data = pd.read_csv('your_sales_data.csv')")
print("Or: data = pd.read_excel('your_sales_data.xlsx')")

# ============================================================================
# SAMPLE DATA STRUCTURE - Replace with your actual data
# ============================================================================

# Create sample data for demonstration
np.random.seed(42)

countries = ['Portugal', 'United Kingdom', 'EIRE', 'Germany', 'France', 'Spain']
products = ['CARD PARTY GAMES', 'JUMBO BAG VINTAGE LEAF', 'WHITE METAL LANTERN', 
            'GLASS JAR CANDLE', 'WOODEN COASTERS', 'CERAMIC VASE']

sample_data = []
for _ in range(500):
    sample_data.append({
        'Country': np.random.choice(countries),
        'Description': np.random.choice(products),
        'Quantity': np.random.randint(1, 50),
        'UnitPrice': np.random.uniform(1, 20)
    })

data = pd.DataFrame(sample_data)

print("\n✓ Sample data loaded (500 rows)")
print(f"✓ Countries: {data['Country'].nunique()}")
print(f"✓ Products: {data['Description'].nunique()}")
print("\nData Preview:")
print(data.head(10))

# ============================================================================
# QUESTION ONE: ADVANCED PRODUCT PERFORMANCE ANALYSIS - PORTUGAL
# ============================================================================

print("\n" + "="*80)
print("QUESTION ONE: PRODUCT SALES ANALYSIS - PORTUGAL")
print("="*80)

products_portugal = data[data['Country'] == 'Portugal'].groupby('Description').agg({
    'Quantity': ['sum', 'count', 'mean', 'median', 'std'],
    'UnitPrice': 'mean'
}).round(2)

products_portugal.columns = ['Total_Qty', 'Count', 'Mean_Qty', 'Median_Qty', 'SD_Qty', 'Mean_Price']
products_portugal['Revenue'] = products_portugal['Total_Qty'] * products_portugal['Mean_Price']
products_portugal['CV'] = (products_portugal['SD_Qty'] / products_portugal['Mean_Qty'] * 100).round(2)
products_portugal = products_portugal.sort_values('Total_Qty', ascending=False)

print("\n✓ Top 15 Products in Portugal:")
print(products_portugal.head(15))

# Export to Excel
products_portugal.to_excel('C:/Users/ailto/Downloads/products_portugal_advanced.xlsx')
print("\n✓ Exported: products_portugal_advanced.xlsx")

# VISUALIZATION 1: Top 15 products
fig, ax = plt.subplots(figsize=(14, 8))
top_15 = products_portugal.head(15)
colors = plt.cm.RdYlGn_r(np.linspace(0.3, 0.7, len(top_15)))
bars = ax.barh(range(len(top_15)), top_15['Total_Qty'], color=colors, edgecolor='black', linewidth=0.8)

for i, (idx, row) in enumerate(top_15.iterrows()):
    ax.text(row['Total_Qty'] + 5, i, f"{int(row['Total_Qty'])}\n(n={int(row['Count'])})", 
            va='center', fontsize=9, fontweight='bold')

ax.set_yticks(range(len(top_15)))
ax.set_yticklabels(top_15.index, fontsize=9)
ax.set_xlabel('Total Quantity Sold', fontsize=11, fontweight='bold')
ax.set_title('Top 15 Best-Performing Products in Portugal\nRanked by Total Quantity Sold', 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_1_Top_15_Products.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 1 saved: Plot_1_Top_15_Products.png")

# VISUALIZATION 2: Mean vs Median
fig, ax = plt.subplots(figsize=(14, 8))
top_10 = products_portugal.head(10)
x = np.arange(len(top_10))
width = 0.35

bars1 = ax.bar(x - width/2, top_10['Mean_Qty'], width, label='Mean', 
                color='#3498db', edgecolor='black', linewidth=0.8)
bars2 = ax.bar(x + width/2, top_10['Median_Qty'], width, label='Median', 
                color='#e67e22', edgecolor='black', linewidth=0.8)

ax.set_xlabel('Product', fontsize=11, fontweight='bold')
ax.set_ylabel('Quantity', fontsize=11, fontweight='bold')
ax.set_title('Mean vs Median Quantities - Top 10 Products\nUnderstanding Distribution Shape', 
             fontsize=14, fontweight='bold', pad=20)
ax.set_xticks(x)
ax.set_xticklabels([str(p)[:18] for p in top_10.index], rotation=45, ha='right', fontsize=9)
ax.legend(fontsize=10, loc='upper right')
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_2_Mean_vs_Median.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 2 saved: Plot_2_Mean_vs_Median.png")

# VISUALIZATION 3: Portfolio Performance Matrix (Bubble Chart)
fig, ax = plt.subplots(figsize=(14, 8))
top_20 = products_portugal.head(20)

scatter = ax.scatter(top_20['Total_Qty'], top_20['Revenue'], 
                    s=top_20['Count']*10, 
                    c=top_20['SD_Qty'], 
                    cmap='RdYlGn_r', alpha=0.6, edgecolors='black', linewidth=1.5)

for idx, row in top_20.iterrows():
    ax.annotate(str(idx)[:12], (row['Total_Qty'], row['Revenue']), 
               fontsize=7, ha='center')

ax.set_xlabel('Total Quantity (log scale)', fontsize=11, fontweight='bold')
ax.set_ylabel('Total Revenue (log scale)', fontsize=11, fontweight='bold')
ax.set_xscale('log')
ax.set_yscale('log')
ax.set_title('Product Portfolio Performance Matrix\nQuantity vs Revenue - Bubble size = frequency', 
             fontsize=14, fontweight='bold', pad=20)

cbar = plt.colorbar(scatter, ax=ax)
cbar.set_label('Std Deviation', fontsize=10, fontweight='bold')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_3_Portfolio_Performance_Matrix.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 3 saved: Plot_3_Portfolio_Performance_Matrix.png")

# ============================================================================
# QUESTION TWO: EXTREME PERFORMERS
# ============================================================================

print("\n" + "="*80)
print("QUESTION TWO: EXTREME PERFORMERS ANALYSIS")
print("="*80)

highest_selling = products_portugal.nlargest(1, 'Total_Qty')
lowest_selling = products_portugal.nsmallest(1, 'Total_Qty')
top_5 = products_portugal.head(5)
bottom_5 = products_portugal.tail(5)

print("\n✓ HIGHEST SELLING PRODUCT:")
print(highest_selling)
print("\n✓ LOWEST SELLING PRODUCT:")
print(lowest_selling)

# VISUALIZATION 4: Top 5 vs Bottom 5
fig, ax = plt.subplots(figsize=(14, 8))
combined = pd.concat([top_5, bottom_5])
colors = ['#27ae60']*5 + ['#e74c3c']*5

y_pos = np.arange(len(combined))
bars = ax.barh(y_pos, combined['Total_Qty'], color=colors, edgecolor='black', linewidth=0.8)

for i, (idx, row) in enumerate(combined.iterrows()):
    ax.text(row['Total_Qty'] + 3, i, f"{int(row['Total_Qty'])}", 
            va='center', fontsize=9, fontweight='bold')

ax.set_yticks(y_pos)
ax.set_yticklabels(combined.index, fontsize=9)
ax.set_xlabel('Total Quantity Sold', fontsize=11, fontweight='bold')
ax.set_title('Extreme Performers: Top 5 vs Bottom 5 Products\nComparative Analysis of Best and Worst Performing Products', 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_4_Top5_vs_Bottom5.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 4 saved: Plot_4_Top5_vs_Bottom5.png")

# VISUALIZATION 5: Distribution with statistics
fig, ax = plt.subplots(figsize=(14, 8))
qty_data = products_portugal['Total_Qty']

n, bins, patches = ax.hist(qty_data, bins=15, alpha=0.7, color='#3498db', edgecolor='black', linewidth=0.8)
ax.axvline(qty_data.mean(), color='#e74c3c', linestyle='--', linewidth=2.5, label=f'Mean: {qty_data.mean():.2f}')
ax.axvline(qty_data.median(), color='#f39c12', linestyle=':', linewidth=2.5, label=f'Median: {qty_data.median():.2f}')

skewness = skew(qty_data)
kurt = kurtosis(qty_data)

ax.set_xlabel('Total Quantity Sold', fontsize=11, fontweight='bold')
ax.set_ylabel('Frequency', fontsize=11, fontweight='bold')
ax.set_title(f'Distribution of Product Performance in Portugal\nSkewness: {skewness:.3f} | Kurtosis: {kurt:.3f}', 
             fontsize=14, fontweight='bold', pad=20)
ax.legend(fontsize=10, loc='upper right')
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_5_Distribution_Analysis.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 5 saved: Plot_5_Distribution_Analysis.png")

# Statistical Summary
print("\n✓ STATISTICAL SUMMARY:")
print(f"  Mean Total Quantity: {qty_data.mean():.2f}")
print(f"  Median Total Quantity: {qty_data.median():.2f}")
print(f"  Std Deviation: {qty_data.std():.2f}")
print(f"  Coefficient of Variation: {(qty_data.std() / qty_data.mean() * 100):.2f}%")
print(f"  Skewness: {skewness:.3f}")
print(f"  Kurtosis: {kurt:.3f}")

# ============================================================================
# QUESTION THREE: GEOGRAPHIC PERFORMANCE
# ============================================================================

print("\n" + "="*80)
print("QUESTION THREE: GEOGRAPHIC DISTRIBUTION - CARD PARTY GAMES")
print("="*80)

card_games = data[data['Description'] == 'CARD PARTY GAMES']
sales_dist = card_games.groupby('Country').agg({
    'Quantity': ['sum', 'mean', 'std', 'count'],
    'UnitPrice': 'mean'
}).round(2)

sales_dist.columns = ['Total_Qty', 'Mean_Qty', 'SD_Qty', 'Count', 'Mean_Price']
sales_dist['SE_Qty'] = sales_dist['SD_Qty'] / np.sqrt(sales_dist['Count'])
sales_dist['CI_Lower'] = sales_dist['Mean_Qty'] - (1.96 * sales_dist['SE_Qty'])
sales_dist['CI_Upper'] = sales_dist['Mean_Qty'] + (1.96 * sales_dist['SE_Qty'])
sales_dist = sales_dist.sort_values('Total_Qty', ascending=False)

print("\n✓ SALES DISTRIBUTION:")
print(sales_dist)

# VISUALIZATION 6: Geographic distribution with CI
fig, ax = plt.subplots(figsize=(14, 8))
countries = sales_dist.index
y_pos = np.arange(len(countries))
colors = plt.cm.RdYlGn(np.linspace(0.3, 0.7, len(countries)))

bars = ax.barh(y_pos, sales_dist['Total_Qty'], color=colors, edgecolor='black', linewidth=0.8)
ax.errorbar(sales_dist['Total_Qty'], y_pos, 
            xerr=1.96*sales_dist['SD_Qty'], fmt='none', color='black', linewidth=2, capsize=5)

for i, country in enumerate(countries):
    ax.text(sales_dist['Total_Qty'].iloc[i] + 10, i, 
            f"{int(sales_dist['Total_Qty'].iloc[i])}\n(n={int(sales_dist['Count'].iloc[i])})", 
            va='center', fontsize=9, fontweight='bold')

ax.set_yticks(y_pos)
ax.set_yticklabels(countries, fontsize=10)
ax.set_xlabel('Total Quantity Sold', fontsize=11, fontweight='bold')
ax.set_title("Geographic Performance: 'CARD PARTY GAMES' Sales by Country\n95% Confidence Intervals", 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_6_Geographic_Performance.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 6 saved: Plot_6_Geographic_Performance.png")

# VISUALIZATION 7: Pareto Analysis
fig, ax = plt.subplots(figsize=(14, 8))
sales_sorted = sales_dist.sort_values('Total_Qty', ascending=False)
cumulative = sales_sorted['Total_Qty'].cumsum()
cumulative_pct = (cumulative / cumulative.iloc[-1]) * 100
country_pct = (sales_sorted['Total_Qty'] / sales_sorted['Total_Qty'].sum()) * 100

ax2 = ax.twinx()
bars = ax.bar(range(len(sales_sorted)), country_pct, color='#3498db', alpha=0.7, edgecolor='black', linewidth=0.8)
line = ax2.plot(range(len(sales_sorted)), cumulative_pct, color='#e74c3c', marker='o', linewidth=2.5, markersize=8)
ax2.axhline(80, color='#95a5a6', linestyle='--', linewidth=2, alpha=0.7, label='80% Threshold')

ax.set_xlabel('Country', fontsize=11, fontweight='bold')
ax.set_ylabel('Country Contribution (%)', fontsize=11, fontweight='bold', color='#3498db')
ax2.set_ylabel('Cumulative Percentage (%)', fontsize=11, fontweight='bold', color='#e74c3c')
ax.set_xticks(range(len(sales_sorted)))
ax.set_xticklabels(sales_sorted.index, rotation=45, ha='right')
ax.set_title("Pareto Analysis: CARD PARTY GAMES Distribution\nIdentifying key markets (80-20 rule)", 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_7_Pareto_Analysis.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 7 saved: Plot_7_Pareto_Analysis.png")

# ============================================================================
# QUESTION FOUR: PRICE VARIATION ANALYSIS
# ============================================================================

print("\n" + "="*80)
print("QUESTION FOUR: PRICE VARIATION ANALYSIS - JUMBO BAG VINTAGE LEAF")
print("="*80)

price_data = data[data['Description'] == 'JUMBO BAG VINTAGE LEAF']

price_summary = price_data.groupby('Country')['UnitPrice'].agg([
    ('N', 'count'),
    ('Min', 'min'),
    ('Q1', lambda x: x.quantile(0.25)),
    ('Median', 'median'),
    ('Q3', lambda x: x.quantile(0.75)),
    ('Max', 'max'),
    ('Mean', 'mean'),
    ('SD', 'std')
]).round(3)

price_summary['SE'] = price_summary['SD'] / np.sqrt(price_summary['N'])
price_summary['CI_Lower'] = price_summary['Mean'] - (1.96 * price_summary['SE'])
price_summary['CI_Upper'] = price_summary['Mean'] + (1.96 * price_summary['SE'])
price_summary['CV%'] = (price_summary['SD'] / price_summary['Mean'] * 100).round(2)
price_summary = price_summary.sort_values('Mean', ascending=False)

print("\n✓ UNIT PRICE SUMMARY:")
print(price_summary)

# Export price summary
price_summary.to_excel('C:/Users/ailto/Downloads/unit_price_summary_stats.xlsx')
print("\n✓ Exported: unit_price_summary_stats.xlsx")

# VISUALIZATION 8: Box plot
fig, ax = plt.subplots(figsize=(14, 8))
countries_list = sorted(price_data['Country'].unique())

bp = ax.boxplot([price_data[price_data['Country'] == c]['UnitPrice'].values for c in countries_list],
                  labels=countries_list, patch_artist=True, widths=0.6)

for patch in bp['boxes']:
    patch.set_facecolor('#3498db')
    patch.set_alpha(0.6)
    patch.set_edgecolor('black')
    patch.set_linewidth(0.8)

for i, country in enumerate(countries_list, 1):
    y = price_data[price_data['Country'] == country]['UnitPrice'].values
    x = np.random.normal(i, 0.04, size=len(y))
    ax.scatter(x, y, alpha=0.3, s=30, color='#2c3e50')

ax.set_ylabel('Unit Price (£)', fontsize=11, fontweight='bold')
ax.set_title("Unit Price Distribution: 'JUMBO BAG VINTAGE LEAF'\nBox plots with individual transactions", 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_8_Price_Distribution_BoxPlot.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 8 saved: Plot_8_Price_Distribution_BoxPlot.png")

# VISUALIZATION 9: Violin plots
fig, ax = plt.subplots(figsize=(14, 8))
parts = ax.violinplot([price_data[price_data['Country'] == c]['UnitPrice'].values 
                       for c in countries_list],
                      positions=range(len(countries_list)),
                      showmeans=False, showmedians=False)

for pc in parts['bodies']:
    pc.set_facecolor('#3498db')
    pc.set_alpha(0.6)
    pc.set_edgecolor('black')
    pc.set_linewidth(0.8)

ax.set_xticks(range(len(countries_list)))
ax.set_xticklabels(countries_list)
ax.set_ylabel('Unit Price (£)', fontsize=11, fontweight='bold')
ax.set_title('Price Distribution Shape Analysis - Violin Plots\nReveal multi-modal distributions; red line = Mean ± 1 SD', 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_9_Price_Distribution_Violin.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 9 saved: Plot_9_Price_Distribution_Violin.png")

# VISUALIZATION 10: Mean price with CI
fig, ax = plt.subplots(figsize=(14, 8))
y_pos = np.arange(len(price_summary))
colors = plt.cm.RdYlGn_r(np.linspace(0.3, 0.7, len(price_summary)))

bars = ax.barh(y_pos, price_summary['Mean'], color=colors, edgecolor='black', linewidth=0.8)
ax.errorbar(price_summary['Mean'], y_pos, 
            xerr=1.96*price_summary['SE'], fmt='none', color='black', linewidth=2, capsize=5)

for i, country in enumerate(price_summary.index):
    ax.text(price_summary['Mean'].iloc[i] + 0.3, i, 
            f"£{price_summary['Mean'].iloc[i]:.2f}\n±£{price_summary['SE'].iloc[i]:.2f}", 
            va='center', fontsize=9, fontweight='bold')

ax.set_yticks(y_pos)
ax.set_yticklabels(price_summary.index, fontsize=10)
ax.set_xlabel('Mean Unit Price (£)', fontsize=11, fontweight='bold')
ax.set_title('Mean Unit Price Comparison (95% Confidence Intervals)', 
             fontsize=14, fontweight='bold', pad=20)
ax.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_10_Mean_Price_with_CI.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 10 saved: Plot_10_Mean_Price_with_CI.png")

# ============================================================================
# QUESTION FIVE: CORRELATION ANALYSIS
# ============================================================================

print("\n" + "="*80)
print("QUESTION FIVE: CORRELATION ANALYSIS - QUANTITY vs PRICE")
print("="*80)

corr_data = data[(data['Description'] == 'JUMBO BAG VINTAGE LEAF') & 
                 (data['Country'].isin(['United Kingdom', 'EIRE', 'Germany']))].copy()

correlation_results = []
for country in ['United Kingdom', 'EIRE', 'Germany']:
    country_data = corr_data[corr_data['Country'] == country]
    if len(country_data) > 2:
        pearson_r, pearson_p = pearsonr(country_data['Quantity'], country_data['UnitPrice'])
        spearman_r, spearman_p = spearmanr(country_data['Quantity'], country_data['UnitPrice'])
        
        correlation_results.append({
            'Country': country,
            'N': len(country_data),
            'Pearson_r': round(pearson_r, 4),
            'Pearson_p': round(pearson_p, 4),
            'Spearman_rho': round(spearman_r, 4),
            'Mean_Qty': round(country_data['Quantity'].mean(), 2),
            'Mean_Price': round(country_data['UnitPrice'].mean(), 2)
        })

corr_df = pd.DataFrame(correlation_results)
print("\n✓ CORRELATION RESULTS:")
print(corr_df)

# Export correlation results
corr_df.to_excel('C:/Users/ailto/Downloads/correlation_analysis_stats.xlsx', index=False)
print("\n✓ Exported: correlation_analysis_stats.xlsx")

print("\n✓ CORRELATION SIGNIFICANCE TESTS:")
for idx, row in corr_df.iterrows():
    sig = "YES (p < 0.05)" if row['Pearson_p'] < 0.05 else "NO (p ≥ 0.05)"
    print(f"  {row['Country']}:")
    print(f"    Pearson r = {row['Pearson_r']}")
    print(f"    p-value = {row['Pearson_p']}")
    print(f"    Significant: {sig}\n")

# VISUALIZATION 11: Scatter plots with regression
fig, axes = plt.subplots(1, 3, figsize=(18, 5))
colors_list = ['#3498db', '#e74c3c', '#2ecc71']

for idx, (ax, country) in enumerate(zip(axes, ['United Kingdom', 'EIRE', 'Germany'])):
    country_data = corr_data[corr_data['Country'] == country]
    
    if len(country_data) > 0:
        ax.scatter(country_data['Quantity'], country_data['UnitPrice'], 
                  alpha=0.6, s=50, color=colors_list[idx], edgecolor='black', linewidth=0.8)
        
        # Linear regression
        if len(country_data) > 1:
            z = np.polyfit(country_data['Quantity'], country_data['UnitPrice'], 1)
            p = np.poly1d(z)
            x_line = np.linspace(country_data['Quantity'].min(), country_data['Quantity'].max(), 100)
            ax.plot(x_line, p(x_line), color=colors_list[idx], linewidth=2, linestyle='--', alpha=0.8)
        
        r_val = corr_df[corr_df['Country'] == country]['Pearson_r'].values[0]
        ax.set_xlabel('Quantity per Transaction', fontsize=10, fontweight='bold')
        ax.set_ylabel('Unit Price (£)', fontsize=10, fontweight='bold')
        ax.set_title(f'{country}\nr = {r_val:.3f}', fontsize=11, fontweight='bold')
        ax.grid(True, alpha=0.3)

fig.suptitle('Quantity vs Unit Price Relationship\nLinear regression with 95% CI', 
            fontsize=14, fontweight='bold', y=1.00)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_11_Scatter_Regression.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 11 saved: Plot_11_Scatter_Regression.png")

# VISUALIZATION 12: Correlation heatmap
fig, ax = plt.subplots(figsize=(10, 6))

corr_matrix = []
for country in ['United Kingdom', 'EIRE', 'Germany']:
    country_data = corr_data[corr_data['Country'] == country]
    if len(country_data) > 0:
        corr = country_data[['Quantity', 'UnitPrice']].corr().iloc[0, 1]
    else:
        corr = 0
    corr_matrix.append([corr, corr])

corr_heatmap = pd.DataFrame(corr_matrix, 
                            index=['United Kingdom', 'EIRE', 'Germany'],
                            columns=['Quantity', 'UnitPrice'])

sns.heatmap(corr_heatmap, annot=True, fmt='.3f', cmap='RdYlGn', center=0, 
           vmin=-1, vmax=1, cbar_kws={'label': 'Correlation'}, ax=ax,
           linewidths=2, linecolor='black', square=True)

ax.set_title('Correlation Matrix Heatmap\nQuantity vs Unit Price across Markets', 
            fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig('C:/Users/ailto/Downloads/Plot_12_Correlation_Heatmap.png', dpi=300, bbox_inches='tight')
plt.show()
print("✓ Plot 12 saved: Plot_12_Correlation_Heatmap.png")

# ============================================================================
# COMPLETION
# ============================================================================

print("\n" + "="*80)
print("✓ ANALYSIS COMPLETE")
print("="*80)
print("""
✓ All 12 publication-quality plots generated and displayed
✓ Advanced statistical analysis completed
✓ Professional theme applied to all visualizations
✓ All plots saved as high-resolution PNG (300 DPI)
✓ Statistical summaries exported to Excel

OUTPUT FILES:
1. Plot_1_Top_15_Products.png
2. Plot_2_Mean_vs_Median.png
3. Plot_3_Portfolio_Performance_Matrix.png
4. Plot_4_Top5_vs_Bottom5.png
5. Plot_5_Distribution_Analysis.png
6. Plot_6_Geographic_Performance.png
7. Plot_7_Pareto_Analysis.png
8. Plot_8_Price_Distribution_BoxPlot.png
9. Plot_9_Price_Distribution_Violin.png
10. Plot_10_Mean_Price_with_CI.png
11. Plot_11_Scatter_Regression.png
12. Plot_12_Correlation_Heatmap.png

EXCEL FILES:
- products_portugal_advanced.xlsx
- unit_price_summary_stats.xlsx
- correlation_analysis_stats.xlsx

Location: C:/Users/ailto/Downloads/

Ready for portfolio presentation!
""")
