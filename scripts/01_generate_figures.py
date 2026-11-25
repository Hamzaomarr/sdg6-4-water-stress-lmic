import os
import numpy as np
import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.colors as mcolors
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize, LinearSegmentedColormap
from matplotlib.lines import Line2D
from shapely.ops import unary_union

# Setup directories
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
DATA_PATH = os.path.join(PROJECT_DIR, 'data', 'AQUASTAT_Project1_Data.xlsx')
GEO_PATH = os.path.join(PROJECT_DIR, 'data', 'geo', 'ne_50m_admin_0_countries.geojson')
PDF_DIR = os.path.join(PROJECT_DIR, 'figures', 'pdf')
PNG_DIR = os.path.join(PROJECT_DIR, 'figures', 'png')
FONT_PATH = os.path.join(PROJECT_DIR, 'fonts', 'Inter-Variable.ttf')

os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(PNG_DIR, exist_ok=True)

if os.path.exists(FONT_PATH):
    fm.fontManager.addfont(FONT_PATH)
    plt.rcParams['font.family'] = 'Inter'

# Strict Publication Typography & Styling (Inter Regular Only)
plt.rcParams['font.weight'] = 'normal'
plt.rcParams['axes.labelweight'] = 'normal'
plt.rcParams['axes.titleweight'] = 'normal'
plt.rcParams['figure.titleweight'] = 'normal'
plt.rcParams['font.size'] = 8.5
plt.rcParams['axes.labelsize'] = 8.5
plt.rcParams['xtick.labelsize'] = 8.0
plt.rcParams['ytick.labelsize'] = 8.0
plt.rcParams['axes.linewidth'] = 0.5
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = 'white'

region_map = {
    'Afghanistan': 'Central & Southern Asia', 'Angola': 'Sub-Saharan Africa', 'Bangladesh': 'Central & Southern Asia',
    'Benin': 'Sub-Saharan Africa', 'Bhutan': 'Central & Southern Asia', 'Burkina Faso': 'Sub-Saharan Africa',
    'Burundi': 'Sub-Saharan Africa', 'Cambodia': 'Eastern & South-Eastern Asia', 'Cameroon': 'Sub-Saharan Africa',
    'Central African Republic': 'Sub-Saharan Africa', 'Chad': 'Sub-Saharan Africa', 'Comoros': 'Sub-Saharan Africa',
    'Congo': 'Sub-Saharan Africa', "Côte d'Ivoire": 'Sub-Saharan Africa',
    "Democratic People's Republic of Korea": 'Eastern & South-Eastern Asia',
    'Democratic Republic of the Congo': 'Sub-Saharan Africa', 'Djibouti': 'Sub-Saharan Africa',
    'Eritrea': 'Sub-Saharan Africa', 'Ethiopia': 'Sub-Saharan Africa', 'Gambia': 'Sub-Saharan Africa',
    'Ghana': 'Sub-Saharan Africa', 'Guinea': 'Sub-Saharan Africa', 'Guinea-Bissau': 'Sub-Saharan Africa',
    'Haiti': 'Latin America & Caribbean', 'India': 'Central & Southern Asia', 'Kenya': 'Sub-Saharan Africa',
    'Kyrgyzstan': 'Central & Southern Asia', "Lao People's Democratic Republic": 'Eastern & South-Eastern Asia',
    'Lesotho': 'Sub-Saharan Africa', 'Liberia': 'Sub-Saharan Africa', 'Madagascar': 'Sub-Saharan Africa',
    'Malawi': 'Sub-Saharan Africa', 'Mali': 'Sub-Saharan Africa', 'Mauritania': 'Sub-Saharan Africa',
    'Mozambique': 'Sub-Saharan Africa', 'Myanmar': 'Eastern & South-Eastern Asia', 'Nepal': 'Central & Southern Asia',
    'Nicaragua': 'Latin America & Caribbean', 'Niger': 'Sub-Saharan Africa', 'Rwanda': 'Sub-Saharan Africa',
    'Sao Tome and Principe': 'Sub-Saharan Africa', 'Senegal': 'Sub-Saharan Africa', 'Sierra Leone': 'Sub-Saharan Africa',
    'Somalia': 'Sub-Saharan Africa', 'South Sudan': 'Sub-Saharan Africa', 'Sudan': 'Sub-Saharan Africa',
    'Syrian Arab Republic': 'Northern Africa & Western Asia', 'Tajikistan': 'Central & Southern Asia',
    'Timor-Leste': 'Eastern & South-Eastern Asia', 'Togo': 'Sub-Saharan Africa', 'Uganda': 'Sub-Saharan Africa',
    'United Republic of Tanzania': 'Sub-Saharan Africa', 'Uzbekistan': 'Central & Southern Asia',
    'Viet Nam': 'Eastern & South-Eastern Asia', 'Yemen': 'Northern Africa & Western Asia',
    'Zambia': 'Sub-Saharan Africa', 'Zimbabwe': 'Sub-Saharan Africa'
}

def load_data():
    df = pd.read_excel(DATA_PATH, sheet_name='Data')
    df_countries = df[df['IsAggregate'] == False].copy()
    df_countries['Region'] = df_countries['Area'].map(region_map)
    
    df_2023 = df_countries[df_countries['Year'] == 2023].copy()
    piv_2023 = df_2023.pivot(index=['Area', 'Region', 'm49'], columns='Variable', values='Value').reset_index()
    
    gdf_world = gpd.read_file(GEO_PATH)
    gdf_world_no_ant = gdf_world[gdf_world['ADMIN'] != 'Antarctica'].copy()
    
    som_match = gdf_world_no_ant[gdf_world_no_ant['NAME'] == 'Somalia']
    sml_match = gdf_world_no_ant[gdf_world_no_ant['NAME'] == 'Somaliland']
    if len(som_match) > 0 and len(sml_match) > 0:
        som_idx = som_match.index[0]
        sml_idx = sml_match.index[0]
        geom_som = gdf_world_no_ant.loc[som_idx, 'geometry']
        geom_sml = gdf_world_no_ant.loc[sml_idx, 'geometry']
        gdf_world_no_ant.loc[som_idx, 'geometry'] = unary_union([geom_som, geom_sml])
        gdf_world_no_ant = gdf_world_no_ant.drop(index=sml_idx).reset_index(drop=True)
        
    name_map_geo = {
        'Democratic Republic of the Congo': 'Dem. Rep. Congo',
        "Democratic People's Republic of Korea": 'North Korea',
        "Côte d'Ivoire": "Côte d'Ivoire",
        'Central African Republic': 'Central African Rep.',
        "Lao People's Democratic Republic": 'Laos',
        'United Republic of Tanzania': 'Tanzania',
        'Viet Nam': 'Vietnam',
        'South Sudan': 'S. Sudan',
        'Timor-Leste': 'Timor-Leste',
        'Syrian Arab Republic': 'Syria',
        'Gambia': 'Gambia',
        'Sao Tome and Principe': 'São Tomé and Principe',
        'Somalia': 'Somalia'
    }
    piv_2023['GeoName'] = piv_2023['Area'].map(lambda x: name_map_geo.get(x, x))
    
    # Project to Robinson for seamless, un-clipped global mapping
    gdf_world_rob = gdf_world_no_ant.to_crs('+proj=robin')
    gdf_merged_rob = gdf_world_rob.merge(piv_2023, left_on='NAME', right_on='GeoName', how='left')
    
    return df, df_countries, piv_2023, gdf_world_rob, gdf_merged_rob

# ==============================================================================
# FIGURE 1: Global Water Stress (Modern Robinson Map) & Sectoral Breakdown
# ==============================================================================
def generate_figure1(piv_2023, gdf_world_rob, gdf_merged_rob):
    fig = plt.figure(figsize=(11.5, 5.0), dpi=300)
    gs1 = fig.add_gridspec(1, 2, width_ratios=[1.22, 1.0], wspace=0.32, top=0.86, bottom=0.12, left=0.03, right=0.97)
    
    # Panel A: Clean Continuous Robinson Map (zero sharp rectangular cuts)
    ax1a = fig.add_subplot(gs1[0])
    ax1a.set_axis_off()
    
    gdf_world_rob.plot(ax=ax1a, facecolor='#F8FAFC', edgecolor='#E2E8F0', linewidth=0.35)
    
    CMAP_STRESS = LinearSegmentedColormap.from_list(
        "stress_cont", ["#E0F2FE", "#38BDF8", "#0284C7", "#D97706", "#DC2626", "#7F1D1D"]
    )
    norm_stress = Normalize(vmin=0, vmax=300)
    
    stress_col = 'SDG 6.4.2. Water Stress'
    study_gdf = gdf_merged_rob.dropna(subset=[stress_col])
    study_gdf.plot(column=stress_col, ax=ax1a, cmap=CMAP_STRESS, norm=norm_stress,
                   edgecolor='#1E293B', linewidth=0.45)
    
    # Sleek Horizontal Flat Colorbar anchored directly below map
    fig.canvas.draw()
    bounds = gdf_world_rob.total_bounds
    p_bl = ax1a.transData.transform((bounds[0], bounds[1]))
    p_tr = ax1a.transData.transform((bounds[2], bounds[3]))
    inv = fig.transFigure.inverted()
    fig_bl = inv.transform(p_bl)
    fig_tr = inv.transform(p_tr)
    
    map_w = fig_tr[0] - fig_bl[0]
    map_x0 = fig_bl[0]
    map_y0 = fig_bl[1]
    
    cax = fig.add_axes([map_x0 + map_w * 0.12, map_y0 - 0.050, map_w * 0.76, 0.016])
    cb = fig.colorbar(ScalarMappable(cmap=CMAP_STRESS, norm=norm_stress), cax=cax, orientation='horizontal', extend='neither')
    cb.set_label('SDG 6.4.2 Water stress (%)', fontsize=8.0, color='#334155', labelpad=4)
    cb.outline.set_linewidth(0.5)
    cb.ax.tick_params(labelsize=7.5, length=2, width=0.5, color='#334155')
    
    # Panel B: Sectoral Breakdown Bar Chart
    ax1b = fig.add_subplot(gs1[1])
    top12 = piv_2023.dropna(subset=[stress_col]).sort_values(by=stress_col, ascending=True).tail(12)
    y_pos = np.arange(len(top12))
    h = 0.58
    
    c_agri = '#2B5C8F'
    c_serv = '#1B9E77'
    c_ind = '#D95F02'
    
    v_agri = top12['SDG 6.4.2. Agricultural Sector Contribution to Water Stress'].values
    v_serv = top12['SDG 6.4.2. Municipal Sector Contribution to Water Stress'].values
    v_ind = top12['SDG 6.4.2. Industrial Sector Contribution to Water Stress'].values
    
    ax1b.barh(y_pos, v_agri, color=c_agri, height=h, label='Agriculture')
    ax1b.barh(y_pos, v_serv, left=v_agri, color=c_serv, hatch='///', height=h, label='Municipal / Services')
    ax1b.barh(y_pos, v_ind, left=v_agri + v_serv, color=c_ind, hatch='xxx', height=h, label='Industry')
    
    ax1b.axvline(100, color='#64748B', linestyle='--', linewidth=0.7)
    
    display_names = {
        'Democratic Republic of the Congo': 'Dem. Rep. Congo',
        "Lao People's Democratic Republic": 'Lao PDR',
        'Syrian Arab Republic': 'Syria',
        'United Republic of Tanzania': 'Tanzania'
    }
    y_labels = [display_names.get(name, name) for name in top12['Area']]
    
    ax1b.set_yticks(y_pos)
    ax1b.set_yticklabels(y_labels)
    ax1b.set_xlabel('SDG 6.4.2 Water stress (%)')
    ax1b.set_xlim([0, 300])
    ax1b.text(0.98, 0.04, f'n = {piv_2023[stress_col].notna().sum()} countries mapped; 12 highest shown', transform=ax1b.transAxes, ha='right', fontsize=6.8, color='#475569')
    ax1b.spines['top'].set_visible(False)
    ax1b.spines['right'].set_visible(False)
    ax1b.xaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax1b.yaxis.grid(False)
    ax1b.set_axisbelow(True)
    
    ax1b.legend(loc='lower left', bbox_to_anchor=(-0.02, 1.02), ncol=3, frameon=False,
                fontsize=7.2, columnspacing=1.8, handletextpad=0.3)
    
    fig.text(0.02, 0.92, 'A', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    fig.text(0.53, 0.92, 'B', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    
    fig.savefig(os.path.join(PDF_DIR, 'fig1_water_stress_global_and_sectoral.pdf'), bbox_inches='tight')
    fig.savefig(os.path.join(PNG_DIR, 'fig1_water_stress_global_and_sectoral.png'), dpi=350, bbox_inches='tight')
    plt.close(fig)
    print("✓ Project 1 Figure 1 exported.")

# ==============================================================================
# FIGURE 2: Sectoral Efficiency Disparities & Stress Decoupling
# ==============================================================================
def generate_figure2(piv_2023):
    fig = plt.figure(figsize=(11.5, 5.0), dpi=300)
    gs2 = fig.add_gridspec(1, 2, width_ratios=[1.0, 1.25], wspace=0.34, top=0.84, bottom=0.14, left=0.06, right=0.97)
    ax2a = fig.add_subplot(gs2[0])
    ax2b = fig.add_subplot(gs2[1])
    
    # Panel A: Sectoral Efficiency Boxplots
    data_agri = piv_2023['SDG 6.4.1. Irrigated Agriculture Water Use Efficiency'].dropna()
    data_ind = piv_2023['SDG 6.4.1. Industrial Water Use Efficiency'].dropna()
    data_serv = piv_2023['SDG 6.4.1. Services Water Use Efficiency'].dropna()
    
    box_colors = ['#2B5C8F', '#D95F02', '#1B9E77']
    bp = ax2a.boxplot([data_agri, data_ind, data_serv], patch_artist=True, widths=0.48,
                      showfliers=False,
                      medianprops=dict(color='black', linewidth=1.0),
                      whiskerprops=dict(color='#64748B', linewidth=0.7),
                      capprops=dict(color='#64748B', linewidth=0.7))
    
    for patch, color in zip(bp['boxes'], box_colors):
        patch.set_facecolor(mcolors.to_rgba(color, alpha=0.35))
        patch.set_edgecolor(color)
        patch.set_linewidth(1.0)
        
    np.random.seed(42)
    for i, (d, col) in enumerate(zip([data_agri, data_ind, data_serv], box_colors), start=1):
        jitter = np.random.normal(0, 0.04, size=len(d))
        ax2a.scatter(i + jitter, d, color=col, alpha=0.6, s=14, edgecolors='none', zorder=3)
        
    ax2a.set_yscale('log')
    ax2a.set_xticks([1, 2, 3])
    ax2a.set_xticklabels(['Agriculture', 'Industry', 'Services'])
    ax2a.set_ylabel('Water use efficiency (USD / m³, log scale)')
    ax2a.set_ylim([0.01, 10000])
    ax2a.text(0.98, 0.03, f'n = {len(data_agri)} / {len(data_ind)} / {len(data_serv)}', transform=ax2a.transAxes, ha='right', fontsize=6.8, color='#475569')
    ax2a.spines['top'].set_visible(False)
    ax2a.spines['right'].set_visible(False)
    ax2a.yaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax2a.xaxis.grid(False)
    ax2a.set_axisbelow(True)
    
    # Panel B: Stress vs Overall Efficiency Scatter
    reg_colors = {
        'Sub-Saharan Africa': '#2B5C8F',
        'Central & Southern Asia': '#D95F02',
        'Eastern & South-Eastern Asia': '#1B9E77',
        'Northern Africa & Western Asia': '#7570B3',
        'Latin America & Caribbean': '#E7298A'
    }
    reg_markers = {
        'Sub-Saharan Africa': 'o',
        'Central & Southern Asia': 's',
        'Eastern & South-Eastern Asia': '^',
        'Northern Africa & Western Asia': 'D',
        'Latin America & Caribbean': 'v'
    }
    
    for reg, grp in piv_2023.groupby('Region'):
        sub = grp.dropna(subset=['SDG 6.4.2. Water Stress', 'SDG 6.4.1. Water Use Efficiency'])
        ax2b.scatter(sub['SDG 6.4.2. Water Stress'], sub['SDG 6.4.1. Water Use Efficiency'],
                     color=reg_colors.get(reg, '#64748B'),
                     marker=reg_markers.get(reg, 'o'),
                     s=32, alpha=0.85, edgecolors='white', linewidth=0.5,
                     label=reg)
        
    ax2b.set_xscale('log')
    ax2b.set_yscale('log')
    ax2b.set_xlabel('SDG 6.4.2 Water stress (%, log scale)')
    ax2b.set_ylabel('SDG 6.4.1 Water use efficiency (USD / m³, log scale)')
    ax2b.set_xlim([0.01, 400])
    ax2b.text(0.98, 0.03, 'n = %d countries' % piv_2023.dropna(subset=['SDG 6.4.2. Water Stress', 'SDG 6.4.1. Water Use Efficiency']).shape[0], transform=ax2b.transAxes, ha='right', fontsize=6.8, color='#475569')
    ax2b.set_ylim([0.1, 200])
    ax2b.spines['top'].set_visible(False)
    ax2b.spines['right'].set_visible(False)
    ax2b.xaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax2b.yaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax2b.set_axisbelow(True)
    
    raw_handles = [
        Line2D([0], [0], marker=reg_markers[r], color='w', markerfacecolor=reg_colors[r], markersize=5.0, label=r)
        for r in ['Sub-Saharan Africa', 'Central & Southern Asia', 'Eastern & South-Eastern Asia', 'Northern Africa & Western Asia', 'Latin America & Caribbean']
    ]
    legend_handles = [raw_handles[0], raw_handles[3], raw_handles[1], raw_handles[4], raw_handles[2]]
    ax2b.legend(handles=legend_handles, loc='lower left', bbox_to_anchor=(-0.03, 1.02), ncol=3, frameon=False, fontsize=7.0, columnspacing=0.5, handletextpad=0.2)
    
    fig.text(0.06, 0.92, 'A', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    fig.text(0.48, 0.92, 'B', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    
    fig.savefig(os.path.join(PDF_DIR, 'fig2_water_efficiency_sectoral_disparity.pdf'), bbox_inches='tight')
    fig.savefig(os.path.join(PNG_DIR, 'fig2_water_efficiency_sectoral_disparity.png'), dpi=350, bbox_inches='tight')
    plt.close(fig)
    print("✓ Project 1 Figure 2 exported.")

# ==============================================================================
# FIGURE 3: Agricultural Withdrawal Intensity (Modern Robinson Map) & Sustainability
# ==============================================================================
def generate_figure3(piv_2023, gdf_world_rob, gdf_merged_rob):
    fig = plt.figure(figsize=(11.5, 5.0), dpi=300)
    gs3 = fig.add_gridspec(1, 2, width_ratios=[1.22, 1.0], wspace=0.32, top=0.86, bottom=0.12, left=0.03, right=0.97)
    
    # Panel A: Agricultural Withdrawal Share Map (zero sharp rectangular cuts)
    ax3a = fig.add_subplot(gs3[0])
    ax3a.set_axis_off()
    
    gdf_world_rob.plot(ax=ax3a, facecolor='#F8FAFC', edgecolor='#E2E8F0', linewidth=0.35)
    
    CMAP_AGRI = LinearSegmentedColormap.from_list(
        "agri_cont", ["#ECFDF5", "#A7F3D0", "#34D399", "#059669", "#047857", "#064E3B"]
    )
    norm_agri = Normalize(vmin=0, vmax=160)
    
    var_agri_pct = 'Agricultural water withdrawal as % of total renewable water resources'
    study_agri = gdf_merged_rob.dropna(subset=[var_agri_pct])
    study_agri.plot(column=var_agri_pct,
                    ax=ax3a, cmap=CMAP_AGRI, norm=norm_agri, edgecolor='#1E293B', linewidth=0.45)
    
    fig.canvas.draw()
    bounds = gdf_world_rob.total_bounds
    p_bl = ax3a.transData.transform((bounds[0], bounds[1]))
    p_tr = ax3a.transData.transform((bounds[2], bounds[3]))
    inv = fig.transFigure.inverted()
    fig_bl = inv.transform(p_bl)
    fig_tr = inv.transform(p_tr)
    
    map_w = fig_tr[0] - fig_bl[0]
    map_x0 = fig_bl[0]
    map_y0 = fig_bl[1]
    
    cax3 = fig.add_axes([map_x0 + map_w * 0.12, map_y0 - 0.050, map_w * 0.76, 0.016])
    cb3 = fig.colorbar(ScalarMappable(cmap=CMAP_AGRI, norm=norm_agri), cax=cax3, orientation='horizontal', extend='neither')
    cb3.set_label('Agricultural withdrawal share (%)', fontsize=8.0, color='#334155', labelpad=4)
    cb3.outline.set_linewidth(0.5)
    cb3.ax.tick_params(labelsize=7.5, length=2, width=0.5, color='#334155')
    
    # Panel B: Agricultural Withdrawal vs Water Stress
    ax3b = fig.add_subplot(gs3[1])
    reg_colors = {
        'Sub-Saharan Africa': '#2B5C8F',
        'Central & Southern Asia': '#D95F02',
        'Eastern & South-Eastern Asia': '#1B9E77',
        'Northern Africa & Western Asia': '#7570B3',
        'Latin America & Caribbean': '#E7298A'
    }
    reg_markers = {
        'Sub-Saharan Africa': 'o',
        'Central & Southern Asia': 's',
        'Eastern & South-Eastern Asia': '^',
        'Northern Africa & Western Asia': 'D',
        'Latin America & Caribbean': 'v'
    }
    
    for reg, grp in piv_2023.groupby('Region'):
        sub = grp.dropna(subset=[var_agri_pct, 'SDG 6.4.2. Water Stress'])
        ax3b.scatter(sub[var_agri_pct], sub['SDG 6.4.2. Water Stress'],
                     color=reg_colors.get(reg, '#64748B'),
                     marker=reg_markers.get(reg, 'o'),
                     s=34, alpha=0.85, edgecolors='white', linewidth=0.5,
                     label=reg)
        
    ax3b.set_yscale('log')
    ax3b.set_xlabel('Agricultural withdrawal (% of total renewable water)')
    ax3b.set_ylabel('SDG 6.4.2 Water stress (%, log scale)')
    ax3b.set_xlim([-5, 175])
    ax3b.text(0.98, 0.03, 'n = %d countries' % piv_2023.dropna(subset=[var_agri_pct, 'SDG 6.4.2. Water Stress']).shape[0], transform=ax3b.transAxes, ha='right', fontsize=6.8, color='#475569')
    ax3b.set_ylim([0.1, 400])
    ax3b.axvline(25, color='#64748B', linestyle=':', linewidth=0.7)
    ax3b.axhline(100, color='#64748B', linestyle='--', linewidth=0.7)
    
    ax3b.spines['top'].set_visible(False)
    ax3b.spines['right'].set_visible(False)
    ax3b.xaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax3b.yaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax3b.set_axisbelow(True)
    
    raw_handles_f3 = [
        Line2D([0], [0], marker=reg_markers[r], color='w', markerfacecolor=reg_colors[r], markersize=5.0, label=r)
        for r in ['Sub-Saharan Africa', 'Central & Southern Asia', 'Eastern & South-Eastern Asia', 'Northern Africa & Western Asia', 'Latin America & Caribbean']
    ]
    legend_handles_f3 = [raw_handles_f3[0], raw_handles_f3[3], raw_handles_f3[1], raw_handles_f3[4], raw_handles_f3[2]]
    ax3b.legend(handles=legend_handles_f3, loc='lower left', bbox_to_anchor=(-0.03, 1.02), ncol=3, frameon=False, fontsize=7.0, columnspacing=0.5, handletextpad=0.2)
    
    fig.text(0.02, 0.92, 'A', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    fig.text(0.53, 0.92, 'B', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    
    fig.savefig(os.path.join(PDF_DIR, 'fig3_agricultural_pressure_sustainability.pdf'), bbox_inches='tight')
    fig.savefig(os.path.join(PNG_DIR, 'fig3_agricultural_pressure_sustainability.png'), dpi=350, bbox_inches='tight')
    plt.close(fig)
    print("✓ Project 1 Figure 3 exported.")

# ==============================================================================
# FIGURE 4: Longitudinal Trajectories & Decoupling Rates
# ==============================================================================
def generate_figure4(df_countries):
    fig = plt.figure(figsize=(11.5, 5.0), dpi=300)
    gs4 = fig.add_gridspec(1, 2, width_ratios=[1.15, 1.0], wspace=0.32, top=0.84, bottom=0.14, left=0.06, right=0.97)
    ax4a = fig.add_subplot(gs4[0])
    ax4b = fig.add_subplot(gs4[1])
    
    # Panel A: Trajectory Indices (2021=100)
    df_ts = df_countries.dropna(subset=['Year', 'Variable', 'Value']).copy()
    piv_ts = df_ts.pivot_table(index=['Year'], columns='Variable', values='Value', aggfunc='median').reset_index()
    
    eff_2021 = piv_ts[piv_ts['Year'] == 2021]['SDG 6.4.1. Water Use Efficiency'].values[0]
    str_2021 = piv_ts[piv_ts['Year'] == 2021]['SDG 6.4.2. Water Stress'].values[0]
    
    idx_eff = (piv_ts['SDG 6.4.1. Water Use Efficiency'] / eff_2021) * 100
    idx_str = (piv_ts['SDG 6.4.2. Water Stress'] / str_2021) * 100
    
    ax4a.plot(piv_ts['Year'], idx_eff, color='#1B9E77', marker='o', markersize=4.5, linewidth=1.5, label='Water use efficiency (SDG 6.4.1)')
    ax4a.plot(piv_ts['Year'], idx_str, color='#D95F02', marker='s', markersize=4.5, linewidth=1.5, linestyle='--', label='Water stress (SDG 6.4.2)')
    
    ax4a.axhline(100, color='#64748B', linestyle=':', linewidth=0.7)
    ax4a.set_xticks(piv_ts['Year'])
    ax4a.set_xlabel('Year')
    ax4a.set_ylabel('Trajectory index (2021 = 100)')
    ax4a.set_ylim([85, 125])
    ax4a.spines['top'].set_visible(False)
    ax4a.spines['right'].set_visible(False)
    ax4a.xaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax4a.yaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax4a.set_axisbelow(True)
    
    ax4a.legend(loc='lower left', bbox_to_anchor=(-0.02, 1.02), ncol=2, frameon=False,
                fontsize=7.2, columnspacing=1.8, handletextpad=0.3)
    
    # Panel B: Regional Growth Rates
    df_growth = df_countries[df_countries['Year'].isin([2021, 2023])].copy()
    piv_g = df_growth.pivot_table(index=['Region', 'Year'], columns='Variable', values='Value', aggfunc='median').reset_index()
    
    regions = ['Sub-Saharan Africa', 'Central & Southern Asia', 'Eastern & South-Eastern Asia']
    eff_growth = []
    str_growth = []
    
    for r in regions:
        sub_r = piv_g[piv_g['Region'] == r]
        v_eff_21 = sub_r[sub_r['Year'] == 2021]['SDG 6.4.1. Water Use Efficiency'].values[0]
        v_eff_23 = sub_r[sub_r['Year'] == 2023]['SDG 6.4.1. Water Use Efficiency'].values[0]
        v_str_21 = sub_r[sub_r['Year'] == 2021]['SDG 6.4.2. Water Stress'].values[0]
        v_str_23 = sub_r[sub_r['Year'] == 2023]['SDG 6.4.2. Water Stress'].values[0]
        
        eff_growth.append(((v_eff_23 - v_eff_21) / v_eff_21) * 100 / 2)
        str_growth.append(((v_str_23 - v_str_21) / v_str_21) * 100 / 2)
        
    y_pos = np.arange(len(regions))
    h = 0.32
    ax4b.barh(y_pos + h/2, eff_growth, height=h, color='#1B9E77', label='Efficiency annualized growth (%)')
    ax4b.barh(y_pos - h/2, str_growth, height=h, color='#D95F02', hatch='//', label='Water stress annualized growth (%)')
    
    ax4b.axvline(0, color='#64748B', linewidth=0.7)
    for yy, sg in zip(y_pos - h/2, str_growth):
        ax4b.text(0.15, yy, f'{sg:.1f}%', va='center', fontsize=6.8, color='#9A3412')
    ax4b.text(0.98, 0.03, 'n = %d countries with stress data' % df_growth[df_growth['Variable'] == 'SDG 6.4.2. Water Stress']['Area'].nunique(), transform=ax4b.transAxes, ha='right', fontsize=6.8, color='#475569')
    ax4b.set_yticks(y_pos)
    ax4b.set_yticklabels(['Sub-Saharan\nAfrica', 'Central &\nSouthern Asia', 'Eastern &\nSE Asia'])
    ax4b.set_xlabel('Annualized growth rate (% / year, 2021–2023)')
    ax4b.spines['top'].set_visible(False)
    ax4b.spines['right'].set_visible(False)
    ax4b.xaxis.grid(True, linestyle='-', linewidth=0.4, color='#E2E8F0', alpha=0.9)
    ax4b.yaxis.grid(False)
    ax4b.set_axisbelow(True)
    
    ax4b.legend(loc='lower left', bbox_to_anchor=(-0.02, 1.02), ncol=2, frameon=False,
                fontsize=7.2, columnspacing=1.8, handletextpad=0.3)
    
    fig.text(0.04, 0.92, 'A', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    fig.text(0.53, 0.92, 'B', fontsize=10.5, fontweight='normal', verticalalignment='bottom')
    
    fig.savefig(os.path.join(PDF_DIR, 'fig4_temporal_dynamics_ldc_trajectories.pdf'), bbox_inches='tight')
    fig.savefig(os.path.join(PNG_DIR, 'fig4_temporal_dynamics_ldc_trajectories.png'), dpi=350, bbox_inches='tight')
    plt.close(fig)
    print("✓ Project 1 Figure 1, 2, 3, 4 exported.")

def main():
    print("=== Generating Elevated Project 1 Figures ===")
    df, df_countries, piv_2023, gdf_world_rob, gdf_merged_rob = load_data()
    generate_figure1(piv_2023, gdf_world_rob, gdf_merged_rob)
    generate_figure2(piv_2023)
    generate_figure3(piv_2023, gdf_world_rob, gdf_merged_rob)
    generate_figure4(df_countries)
    print("=== Elevated Project 1 figures complete! ===")

if __name__ == '__main__':
    main()
