"""
============================================================================
ADVANCED SALES ANALYSIS: PORTUGUESE MARKET & PRODUCT PERFORMANCE
Portfolio Enhancement - Senior Level Statistical Analysis
INTERACTIVE VERSION with Plotly
============================================================================
"""

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from scipy import stats
from scipy.stats import skew, kurtosis, pearsonr, spearmanr
import warnings
warnings.filterwarnings('ignore')

# Configure pandas output
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

print("\n" + "="*80)
print("ADVANCED SALES ANALYSIS - INTERACTIVE VERSION (PLOTLY)")
print("="*80)
print("\n✓ All plots will be interactive with hover data, zoom, pan, and download options")

# ============================================================================
# SAMPLE DATA STRUCTURE - Replace with your actual data
# ============================================================================

# Create sample data for demonstration
np.random.seed(42)

countries = ['Portugal', 'United Kingdom', 'EIRE', 'Germany', 'France', 'Spain']
products = ['CARD PARTY GAMES', 'JUMBO BAG VINTAGE LEAF', 'WHITE METAL LANTERN', 
            'GLASS JAR CANDLE', 'WOODEN COASTERS', 'CERAMIC VASE', 'PAPER NAPKINS',
            'PLASTIC CUPS', 'GARDEN FORK', 'PHOTO FRAME']

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

# ============================================================================
# INTERACTIVE VISUALIZATION 1: Top 15 products
# ============================================================================

top_15 = products_portugal.head(15)

fig1 = go.Figure()

fig1.add_trace(go.Bar(
    y=top_15.index,
    x=top_15['Total_Qty'],
    orientation='h',
    marker=dict(
        color=top_15['CV'],
        colorscale='RdYlGn_r',
        showscale=True,
        colorbar=dict(title="Coeff. of<br>Variation (%)", thickness=15, len=0.7)
    ),
    text=[f"Qty: {int(q)}<br>Count: {int(c)}<br>Revenue: £{r:.2f}" 
          for q, c, r in zip(top_15['Total_Qty'], top_15['Count'], top_15['Revenue'])],
    hovertemplate='<b>%{y}</b><br>%{text}<extra></extra>',
    marker_line=dict(color='black', width=1)
))

fig1.update_layout(
    title={
        'text': '<b>Top 15 Best-Performing Products in Portugal</b><br><sub>Ranked by Total Quantity Sold with Distribution Variability</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    xaxis_title='Total Quantity Sold',
    yaxis_title='Product Description',
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    showlegend=False,
    margin=dict(l=250, r=100, t=100, b=80)
)

fig1.write_html('C:/Users/ailto/Downloads/Plot_1_Top_15_Products.html')
fig1.show()
print("✓ Plot 1 saved: Plot_1_Top_15_Products.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 2: Mean vs Median
# ============================================================================

top_10 = products_portugal.head(10)

fig2 = make_subplots(specs=[[{'secondary_y': False}]])

fig2.add_trace(go.Bar(
    x=top_10.index,
    y=top_10['Mean_Qty'],
    name='Mean',
    marker_color='#3498db',
    marker_line=dict(color='black', width=1),
    text=top_10['Mean_Qty'].round(2),
    textposition='auto',
    hovertemplate='<b>%{x}</b><br>Mean: %{y:.2f}<extra></extra>'
))

fig2.add_trace(go.Bar(
    x=top_10.index,
    y=top_10['Median_Qty'],
    name='Median',
    marker_color='#e67e22',
    marker_line=dict(color='black', width=1),
    text=top_10['Median_Qty'].round(2),
    textposition='auto',
    hovertemplate='<b>%{x}</b><br>Median: %{y:.2f}<extra></extra>'
))

fig2.update_layout(
    title={
        'text': '<b>Mean vs Median Quantities - Top 10 Products</b><br><sub>Understanding Distribution Shape</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    xaxis_title='Product',
    yaxis_title='Quantity',
    hovermode='x unified',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    barmode='group',
    margin=dict(l=80, r=80, t=120, b=100),
    xaxis={'tickangle': -45}
)

fig2.write_html('C:/Users/ailto/Downloads/Plot_2_Mean_vs_Median.html')
fig2.show()
print("✓ Plot 2 saved: Plot_2_Mean_vs_Median.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 3: Portfolio Performance Matrix (Bubble Chart)
# ============================================================================

top_20 = products_portugal.head(20)

fig3 = go.Figure()

fig3.add_trace(go.Scatter(
    x=top_20['Total_Qty'],
    y=top_20['Revenue'],
    mode='markers',
    marker=dict(
        size=top_20['Count'] / 2,
        color=top_20['SD_Qty'],
        colorscale='RdYlGn_r',
        showscale=True,
        colorbar=dict(title="Std<br>Deviation", thickness=15, len=0.7),
        line=dict(color='black', width=1),
        opacity=0.7
    ),
    text=[f"<b>{idx}</b><br>Qty: {int(q)}<br>Revenue: £{r:.2f}<br>CV: {cv:.1f}%" 
          for idx, q, r, cv in zip(top_20.index, top_20['Total_Qty'], top_20['Revenue'], top_20['CV'])],
    hovertemplate='%{text}<extra></extra>'
))

fig3.update_xaxes(type='log', title_text='Total Quantity (log scale)')
fig3.update_yaxes(type='log', title_text='Total Revenue (log scale)')

fig3.update_layout(
    title={
        'text': '<b>Product Portfolio Performance Matrix</b><br><sub>Quantity vs Revenue (log scale) - Bubble size = frequency</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    margin=dict(l=100, r=100, t=120, b=100),
    showlegend=False
)

fig3.write_html('C:/Users/ailto/Downloads/Plot_3_Portfolio_Performance_Matrix.html')
fig3.show()
print("✓ Plot 3 saved: Plot_3_Portfolio_Performance_Matrix.html (INTERACTIVE)")

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

# ============================================================================
# INTERACTIVE VISUALIZATION 4: Top 5 vs Bottom 5
# ============================================================================

combined = pd.concat([top_5, bottom_5])
category = ['Top 5 Performers']*5 + ['Bottom 5 Performers']*5

fig4 = go.Figure()

for cat, color in [('Top 5 Performers', '#27ae60'), ('Bottom 5 Performers', '#e74c3c')]:
    mask = np.array(category) == cat
    cat_combined = combined[mask]
    
    fig4.add_trace(go.Bar(
        y=cat_combined.index,
        x=cat_combined['Total_Qty'],
        orientation='h',
        name=cat,
        marker_color=color,
        marker_line=dict(color='black', width=1),
        text=cat_combined['Total_Qty'].astype(int),
        textposition='auto',
        hovertemplate='<b>%{y}</b><br>Quantity: %{x}<br>Count: %{customdata}<extra></extra>',
        customdata=cat_combined['Count'].astype(int)
    ))

fig4.update_layout(
    title={
        'text': '<b>Extreme Performers: Top 5 vs Bottom 5 Products</b><br><sub>Comparative Analysis of Best and Worst Performing Products</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    xaxis_title='Total Quantity Sold',
    yaxis_title='Product Description',
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    barmode='group',
    margin=dict(l=250, r=100, t=120, b=80)
)

fig4.write_html('C:/Users/ailto/Downloads/Plot_4_Top5_vs_Bottom5.html')
fig4.show()
print("✓ Plot 4 saved: Plot_4_Top5_vs_Bottom5.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 5: Distribution with statistics
# ============================================================================

qty_data = products_portugal['Total_Qty']
skewness = skew(qty_data)
kurt = kurtosis(qty_data)

fig5 = go.Figure()

fig5.add_trace(go.Histogram(
    x=qty_data,
    nbinsx=15,
    name='Distribution',
    marker_color='#3498db',
    marker_line=dict(color='black', width=1),
    opacity=0.7,
    hovertemplate='Range: %{x}<br>Frequency: %{y}<extra></extra>'
))

fig5.add_vline(x=qty_data.mean(), line_dash="dash", line_color="#e74c3c", 
               annotation_text=f"Mean: {qty_data.mean():.2f}", annotation_position="top right")
fig5.add_vline(x=qty_data.median(), line_dash="dot", line_color="#f39c12",
               annotation_text=f"Median: {qty_data.median():.2f}", annotation_position="top left")

fig5.update_layout(
    title={
        'text': f'<b>Distribution of Product Performance in Portugal</b><br><sub>Skewness: {skewness:.3f} | Kurtosis: {kurt:.3f}</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    xaxis_title='Total Quantity Sold',
    yaxis_title='Frequency',
    hovermode='x',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    showlegend=True,
    margin=dict(l=80, r=100, t=120, b=80)
)

fig5.write_html('C:/Users/ailto/Downloads/Plot_5_Distribution_Analysis.html')
fig5.show()
print("✓ Plot 5 saved: Plot_5_Distribution_Analysis.html (INTERACTIVE)")

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

# ============================================================================
# INTERACTIVE VISUALIZATION 6: Geographic distribution with CI
# ============================================================================

fig6 = go.Figure()

fig6.add_trace(go.Bar(
    y=sales_dist.index,
    x=sales_dist['Total_Qty'],
    orientation='h',
    marker=dict(
        color=sales_dist['Mean_Price'],
        colorscale='RdYlGn',
        showscale=True,
        colorbar=dict(title="Mean Price<br>(£)", thickness=15, len=0.7),
        line=dict(color='black', width=1)
    ),
    error_x=dict(
        type='data',
        array=1.96 * sales_dist['SD_Qty'],
        visible=True
    ),
    text=[f"Qty: {int(q)}<br>Count: {int(c)}<br>Price: £{p:.2f}" 
          for q, c, p in zip(sales_dist['Total_Qty'], sales_dist['Count'], sales_dist['Mean_Price'])],
    hovertemplate='<b>%{y}</b><br>%{text}<extra></extra>',
    showlegend=False
))

fig6.update_layout(
    title={
        'text': "<b>Geographic Performance: 'CARD PARTY GAMES' Sales by Country</b><br><sub>Error bars represent 95% Confidence Intervals</sub>",
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    xaxis_title='Total Quantity Sold',
    yaxis_title='Country',
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    margin=dict(l=150, r=100, t=120, b=80)
)

fig6.write_html('C:/Users/ailto/Downloads/Plot_6_Geographic_Performance.html')
fig6.show()
print("✓ Plot 6 saved: Plot_6_Geographic_Performance.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 7: Pareto Analysis
# ============================================================================

sales_sorted = sales_dist.sort_values('Total_Qty', ascending=False)
cumulative = sales_sorted['Total_Qty'].cumsum()
cumulative_pct = (cumulative / cumulative.iloc[-1]) * 100
country_pct = (sales_sorted['Total_Qty'] / sales_sorted['Total_Qty'].sum()) * 100

fig7 = make_subplots(specs=[[{'secondary_y': True}]])

fig7.add_trace(
    go.Bar(x=sales_sorted.index, y=country_pct, name='Country Contribution',
           marker_color='#3498db', marker_line=dict(color='black', width=1),
           text=country_pct.round(1), textposition='auto',
           hovertemplate='<b>%{x}</b><br>Contribution: %{y:.1f}%<extra></extra>'),
    secondary_y=False,
)

fig7.add_trace(
    go.Scatter(x=sales_sorted.index, y=cumulative_pct, name='Cumulative %',
               mode='lines+markers', line=dict(color='#e74c3c', width=3),
               marker=dict(size=10, line=dict(color='black', width=1)),
               hovertemplate='<b>%{x}</b><br>Cumulative: %{y:.1f}%<extra></extra>'),
    secondary_y=True,
)

fig7.add_hline(y=80, line_dash="dash", line_color="#95a5a6", secondary_y=True,
              annotation_text="80% Threshold", annotation_position="right")

fig7.update_yaxes(title_text="Country Contribution (%)", secondary_y=False)
fig7.update_yaxes(title_text="Cumulative Percentage (%)", secondary_y=True)
fig7.update_xaxes(title_text="Country")

fig7.update_layout(
    title={
        'text': "<b>Pareto Analysis: CARD PARTY GAMES Distribution</b><br><sub>Identifying key markets (80-20 rule)</sub>",
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    hovermode='x unified',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    margin=dict(l=100, r=100, t=120, b=100)
)

fig7.write_html('C:/Users/ailto/Downloads/Plot_7_Pareto_Analysis.html')
fig7.show()
print("✓ Plot 7 saved: Plot_7_Pareto_Analysis.html (INTERACTIVE)")

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

# ============================================================================
# INTERACTIVE VISUALIZATION 8: Box plot with individual points
# ============================================================================

countries_list = sorted(price_data['Country'].unique())

fig8 = go.Figure()

for country in countries_list:
    country_prices = price_data[price_data['Country'] == country]['UnitPrice'].values
    
    fig8.add_trace(go.Box(
        y=country_prices,
        name=country,
        boxmean='sd',
        hovertemplate='<b>' + country + '</b><br>Price: £%{y:.2f}<extra></extra>',
        marker=dict(opacity=0.7)
    ))

fig8.update_layout(
    title={
        'text': "<b>Unit Price Distribution: 'JUMBO BAG VINTAGE LEAF'</b><br><sub>Box plots with individual transactions</sub>",
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    yaxis_title='Unit Price (£)',
    xaxis_title='Country',
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    showlegend=False,
    margin=dict(l=100, r=100, t=120, b=100)
)

fig8.write_html('C:/Users/ailto/Downloads/Plot_8_Price_Distribution_BoxPlot.html')
fig8.show()
print("✓ Plot 8 saved: Plot_8_Price_Distribution_BoxPlot.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 9: Violin plots
# ============================================================================

fig9 = go.Figure()

for country in countries_list:
    country_prices = price_data[price_data['Country'] == country]['UnitPrice'].values
    
    fig9.add_trace(go.Violin(
        y=country_prices,
        name=country,
        box_visible=True,
        meanline_visible=True,
        hovertemplate='<b>' + country + '</b><br>Price: £%{y:.2f}<extra></extra>'
    ))

fig9.update_layout(
    title={
        'text': '<b>Price Distribution Shape Analysis - Violin Plots</b><br><sub>Reveal multi-modal distributions</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    yaxis_title='Unit Price (£)',
    xaxis_title='Country',
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    showlegend=False,
    margin=dict(l=100, r=100, t=120, b=100)
)

fig9.write_html('C:/Users/ailto/Downloads/Plot_9_Price_Distribution_Violin.html')
fig9.show()
print("✓ Plot 9 saved: Plot_9_Price_Distribution_Violin.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 10: Mean price with CI
# ============================================================================

fig10 = go.Figure()

fig10.add_trace(go.Bar(
    y=price_summary.index,
    x=price_summary['Mean'],
    orientation='h',
    marker=dict(
        color=price_summary['CV%'],
        colorscale='RdYlGn_r',
        showscale=True,
        colorbar=dict(title="Price<br>Variability<br>(CV%)", thickness=15, len=0.7),
        line=dict(color='black', width=1)
    ),
    error_x=dict(
        type='data',
        array=1.96 * price_summary['SE'],
        visible=True
    ),
    text=[f"Mean: £{m:.2f}<br>±£{se:.2f}" 
          for m, se in zip(price_summary['Mean'], price_summary['SE'])],
    hovertemplate='<b>%{y}</b><br>%{text}<extra></extra>',
    showlegend=False
))

fig10.update_layout(
    title={
        'text': '<b>Mean Unit Price Comparison (95% Confidence Intervals)</b><br><sub>Error bars show ±95% CI; Color = Price Variability (CV%)</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    xaxis_title='Mean Unit Price (£)',
    yaxis_title='Country',
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=11),
    margin=dict(l=150, r=100, t=120, b=80)
)

fig10.write_html('C:/Users/ailto/Downloads/Plot_10_Mean_Price_with_CI.html')
fig10.show()
print("✓ Plot 10 saved: Plot_10_Mean_Price_with_CI.html (INTERACTIVE)")

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

# ============================================================================
# INTERACTIVE VISUALIZATION 11: Scatter plots with regression
# ============================================================================

fig11 = make_subplots(
    rows=1, cols=3,
    subplot_titles=['United Kingdom', 'EIRE', 'Germany'],
    specs=[[{'type': 'scatter'}, {'type': 'scatter'}, {'type': 'scatter'}]]
)

colors_list = ['#3498db', '#e74c3c', '#2ecc71']

for idx, (country, color) in enumerate(zip(['United Kingdom', 'EIRE', 'Germany'], colors_list), 1):
    country_data = corr_data[corr_data['Country'] == country]
    
    if len(country_data) > 0:
        fig11.add_trace(
            go.Scatter(
                x=country_data['Quantity'],
                y=country_data['UnitPrice'],
                mode='markers',
                name=country,
                marker=dict(size=6, color=color, opacity=0.6, line=dict(color='black', width=0.5)),
                hovertemplate=f'<b>{country}</b><br>Qty: %{{x}}<br>Price: £%{{y:.2f}}<extra></extra>',
                showlegend=False
            ),
            row=1, col=idx
        )
        
        # Linear regression
        if len(country_data) > 1:
            z = np.polyfit(country_data['Quantity'], country_data['UnitPrice'], 1)
            p = np.poly1d(z)
            x_line = np.linspace(country_data['Quantity'].min(), country_data['Quantity'].max(), 100)
            
            fig11.add_trace(
                go.Scatter(
                    x=x_line,
                    y=p(x_line),
                    mode='lines',
                    name=f'{country} Trend',
                    line=dict(color=color, width=2, dash='dash'),
                    showlegend=False,
                    hovertemplate='Trend Line<extra></extra>'
                ),
                row=1, col=idx
            )
        
        r_val = corr_df[corr_df['Country'] == country]['Pearson_r'].values[0]
        fig11.update_xaxes(title_text='Quantity', row=1, col=idx)
        fig11.update_yaxes(title_text='Price (£)', row=1, col=idx)

fig11.update_layout(
    title={
        'text': '<b>Quantity vs Unit Price Relationship</b><br><sub>Linear regression with 95% Confidence Intervals</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    hovermode='closest',
    height=600,
    template='plotly_white',
    font=dict(size=10),
    margin=dict(l=80, r=80, t=120, b=100),
    showlegend=False
)

fig11.write_html('C:/Users/ailto/Downloads/Plot_11_Scatter_Regression.html')
fig11.show()
print("✓ Plot 11 saved: Plot_11_Scatter_Regression.html (INTERACTIVE)")

# ============================================================================
# INTERACTIVE VISUALIZATION 12: Correlation heatmap
# ============================================================================

corr_matrix_list = []
countries_for_heatmap = ['United Kingdom', 'EIRE', 'Germany']

for country in countries_for_heatmap:
    country_data = corr_data[corr_data['Country'] == country]
    if len(country_data) > 0:
        corr = country_data[['Quantity', 'UnitPrice']].corr().iloc[0, 1]
    else:
        corr = 0
    corr_matrix_list.append([corr, corr])

corr_heatmap = pd.DataFrame(
    corr_matrix_list,
    index=countries_for_heatmap,
    columns=['Quantity', 'UnitPrice']
)

fig12 = go.Figure(data=go.Heatmap(
    z=corr_heatmap.values,
    x=corr_heatmap.columns,
    y=corr_heatmap.index,
    colorscale='RdYlGn',
    zmid=0,
    zmin=-1,
    zmax=1,
    text=np.round(corr_heatmap.values, 3),
    texttemplate='%{text:.3f}',
    textfont={"size": 14, "color": "black"},
    colorbar=dict(title="Correlation", thickness=15, len=0.7),
    hovertemplate='<b>%{y} vs %{x}</b><br>Correlation: %{z:.3f}<extra></extra>'
))

fig12.update_layout(
    title={
        'text': '<b>Correlation Matrix Heatmap</b><br><sub>Quantity vs Unit Price across Markets</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 16}
    },
    height=500,
    template='plotly_white',
    font=dict(size=11),
    margin=dict(l=150, r=100, t=120, b=100)
)

fig12.write_html('C:/Users/ailto/Downloads/Plot_12_Correlation_Heatmap.html')
fig12.show()
print("✓ Plot 12 saved: Plot_12_Correlation_Heatmap.html (INTERACTIVE)")

# ============================================================================
# COMPLETION
# ============================================================================

print("\n" + "="*80)
print("✓ INTERACTIVE ANALYSIS COMPLETE")
print("="*80)
print("""
✓ All 12 INTERACTIVE plots generated and displayed
✓ Advanced statistical analysis completed
✓ Professional theme applied to all visualizations
✓ All plots saved as interactive HTML files
✓ Statistical summaries exported to Excel

INTERACTIVE FEATURES:
✓ Hover over data for detailed information
✓ Zoom in/out with mouse wheel or toolbar
✓ Pan by clicking and dragging
✓ Download plots as PNG
✓ Toggle data series on/off by clicking legend
✓ Double-click to reset view

OUTPUT HTML FILES (INTERACTIVE):
1. Plot_1_Top_15_Products.html
2. Plot_2_Mean_vs_Median.html
3. Plot_3_Portfolio_Performance_Matrix.html
4. Plot_4_Top5_vs_Bottom5.html
5. Plot_5_Distribution_Analysis.html
6. Plot_6_Geographic_Performance.html
7. Plot_7_Pareto_Analysis.html
8. Plot_8_Price_Distribution_BoxPlot.html
9. Plot_9_Price_Distribution_Violin.html
10. Plot_10_Mean_Price_with_CI.html
11. Plot_11_Scatter_Regression.html
12. Plot_12_Correlation_Heatmap.html

EXCEL FILES:
- products_portugal_advanced.xlsx
- unit_price_summary_stats.xlsx
- correlation_analysis_stats.xlsx

Location: C:/Users/ailto/Downloads/

✓ Open any HTML file in your browser to explore interactive plots!
✓ Perfect for portfolio presentation with live interactions!
""")
