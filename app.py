"""
KONE EQUIPMENT MAINTENANCE DASHBOARD - NEW PROJECT
Version: 1.0 - Multi-Excel Integration with Cascading Filters

KEY FEATURES:
✅ Merge two Excel files (Equipment Data + Maintenance Data)
✅ Cascading filters: Equipment Code → Type → Module → Components
✅ Clean number formatting (43397068 not 43,397,068)
✅ Same beautiful design as v1.3
✅ Production ready for demo

FILTER FLOW:
  Excel 1 (Equipment Data):
    ├─ Equipment Code (43397068)
    ├─ Type (KONE KCE)
    └─ (Link key: Equipment Code)

  Excel 2 (Maintenance Data):
    ├─ Module
    ├─ Components
    ├─ Preparation/Finalization time
    ├─ Activity time
    ├─ Total time
    └─ No of man power

  User Selection:
    1. Select Equipment Code
    2. Type auto-populated (from Equipment Data)
    3. Select Module
    4. Select Components
    5. View data & export
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import re

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="Equipment Maintenance Dashboard",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================================
# CUSTOM CSS - SAME BEAUTIFUL DESIGN AS v1.3
# ============================================================================
st.markdown("""
    <style>
        /* Main category card styling */
        .equipment-card {
            position: relative;
            width: 100%;
            aspect-ratio: 1 / 1;
            border-radius: 8px;
            overflow: hidden;
            cursor: pointer;
            transition: all 0.3s ease;
            border: 3px solid transparent;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 150px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        
        .equipment-card:hover {
            transform: scale(1.05);
            border-color: #2E7D9E;
            box-shadow: 0 4px 12px rgba(46, 125, 158, 0.3);
        }
        
        .equipment-card.selected {
            border-color: #28a745;
            box-shadow: 0 4px 12px rgba(40, 167, 69, 0.4);
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }
        
        .equipment-card-content {
            position: relative;
            z-index: 2;
            text-align: center;
            background: rgba(255, 255, 255, 0.95);
            padding: 1rem;
            border-radius: 6px;
            backdrop-filter: blur(10px);
        }
        
        .equipment-code {
            font-size: 20px;
            font-weight: bold;
            color: #2E7D9E;
            margin: 0.5rem 0;
        }
        
        .equipment-type {
            font-size: 12px;
            color: #666;
            margin: 0;
        }
        
        /* Stats metrics */
        .stat-card {
            flex: 1;
            min-width: 200px;
            background: white;
            border-radius: 8px;
            padding: 1.5rem;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-top: 4px solid #2E7D9E;
        }
        
        .stat-card-value {
            font-size: 28px;
            font-weight: bold;
            color: #2E7D9E;
            margin: 0.5rem 0;
        }
        
        .stat-card-label {
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .filter-section {
            background: white;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            margin: 1rem 0;
        }
    </style>
""", unsafe_allow_html=True)

# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================
"""
SESSION STATE for Cascading Filters:
  - equipment_data: Merged data from two Excel files
  - equipment_codes: List of unique equipment codes
  - selected_equipment: Currently selected equipment code
  - selected_type: Type (auto-populated from equipment data)
  - selected_module: Currently selected module
  - selected_components: List of selected components
"""

if 'equipment_data' not in st.session_state:
    st.session_state.equipment_data = None
if 'selected_equipment' not in st.session_state:
    st.session_state.selected_equipment = None
if 'selected_type' not in st.session_state:
    st.session_state.selected_type = None
if 'selected_module' not in st.session_state:
    st.session_state.selected_module = None
if 'selected_components' not in st.session_state:
    st.session_state.selected_components = []

# ============================================================================
# UTILITY FUNCTIONS - TIME CONVERSION
# ============================================================================

def time_str_to_seconds(time_str):
    """Convert HH:MM:SS to seconds"""
    if pd.isna(time_str):
        return 0
    try:
        parts = str(time_str).split(':')
        if len(parts) == 3:
            h, m, s = map(int, parts)
            return h * 3600 + m * 60 + s
    except:
        pass
    return 0

def seconds_to_time_str(seconds):
    """Convert seconds to HH:MM:SS"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

def time_str_to_hours(time_str):
    """Convert HH:MM:SS to decimal hours"""
    return time_str_to_seconds(time_str) / 3600

def clean_equipment_code(code):
    """
    Clean equipment code: Remove commas and extra formatting
    Example: 43,397,068 → 43397068
    """
    if pd.isna(code):
        return None
    try:
        # Remove commas
        cleaned = str(code).replace(',', '')
        # Convert to integer and back to string (ensures it's a proper number)
        return str(int(cleaned))
    except:
        return str(code)

# ============================================================================
# DATA LOADING & MERGING
# ============================================================================

def merge_excel_files(equipment_file, maintenance_file):
    """
    Merge two Excel files:
    
    Equipment File should have:
      - Equipment Code
      - Type
    
    Maintenance File should have:
      - Equipment Code (link key)
      - Module
      - Components
      - Preparation/Finalization time
      - Activity time
      - Total time
      - No of man power
    """
    try:
        # Read both files
        equipment_df = pd.read_excel(equipment_file)
        maintenance_df = pd.read_excel(maintenance_file)
        
        # Clean equipment codes in both files
        equipment_df['Equipment Code'] = equipment_df['Equipment Code'].apply(clean_equipment_code)
        maintenance_df['Equipment Code'] = maintenance_df['Equipment Code'].apply(clean_equipment_code)
        
        # Merge on Equipment Code
        merged_df = maintenance_df.merge(
            equipment_df[['Equipment Code', 'Type']],
            on='Equipment Code',
            how='left'
        )
        
        # Forward fill Type and Module (handle merged cells)
        merged_df['Type'] = merged_df['Type'].fillna(method='ffill')
        if 'Module' in merged_df.columns:
            merged_df['Module'] = merged_df['Module'].fillna(method='ffill')
        
        st.session_state.equipment_data = merged_df
        return merged_df, None
    
    except Exception as e:
        return None, f"❌ Error merging files: {str(e)}"

# ============================================================================
# FILTER FUNCTIONS - CASCADING LOGIC
# ============================================================================

def get_equipment_codes(df):
    """Get unique equipment codes"""
    if df is None:
        return []
    try:
        codes = sorted(df['Equipment Code'].dropna().unique().tolist())
        return [str(code) for code in codes]
    except:
        return []

def get_type_for_equipment(df, equipment_code):
    """Get type for selected equipment code"""
    if df is None or equipment_code is None:
        return None
    try:
        filtered = df[df['Equipment Code'] == equipment_code]
        types = filtered['Type'].dropna().unique()
        return str(types[0]) if len(types) > 0 else None
    except:
        return None

def get_modules(df, equipment_code):
    """Get unique modules for selected equipment code"""
    if df is None or equipment_code is None:
        return []
    try:
        filtered = df[df['Equipment Code'] == equipment_code]
        modules = sorted(filtered['Module'].dropna().unique().tolist())
        return [str(m) for m in modules]
    except:
        return []

def get_components(df, equipment_code, module):
    """Get components for selected equipment and module"""
    if df is None or equipment_code is None or module is None:
        return []
    try:
        filtered = df[
            (df['Equipment Code'] == equipment_code) &
            (df['Module'] == module)
        ]
        components = sorted(filtered['Components'].dropna().unique().tolist())
        return [str(c) for c in components]
    except:
        return []

def filter_data(df, equipment_code, module, components):
    """Final data filter"""
    if df is None:
        return None
    try:
        filtered = df[df['Equipment Code'] == equipment_code].copy()
        filtered = filtered[filtered['Module'] == module]
        
        if components:
            filtered = filtered[filtered['Components'].isin(components)]
        
        return filtered
    except:
        return None

def calculate_stats(df):
    """Calculate statistics"""
    if df is None or df.empty:
        return {'records': 0, 'total_time': 0, 'avg_manpower': 0}
    
    stats = {
        'records': len(df),
        'total_time': 0,
        'avg_manpower': 0,
    }
    
    try:
        if 'Total time' in df.columns:
            total_secs = sum(time_str_to_seconds(t) for t in df['Total time'])
            stats['total_time'] = total_secs
    except:
        pass
    
    try:
        if 'No of man power' in df.columns:
            manpower = pd.to_numeric(df['No of man power'], errors='coerce').dropna()
            if len(manpower) > 0:
                stats['avg_manpower'] = float(manpower.mean())
    except:
        pass
    
    return stats

# ============================================================================
# SIDEBAR - FILE UPLOAD
# ============================================================================

with st.sidebar:
    st.title("🔧 Equipment Dashboard Control")
    st.markdown("---")
    
    st.subheader("📁 Upload Excel Files")
    st.info("You need 2 Excel files to get started")
    
    # Equipment file upload
    equipment_file = st.file_uploader(
        "Equipment File (Equipment Code + Type)",
        type=['xlsx', 'xls'],
        key='equipment_file'
    )
    
    # Maintenance file upload
    maintenance_file = st.file_uploader(
        "Maintenance File (Components + Times + Manpower)",
        type=['xlsx', 'xls'],
        key='maintenance_file'
    )
    
    # Merge files
    if equipment_file is not None and maintenance_file is not None:
        with st.spinner("📥 Merging Excel files..."):
            merged_data, error = merge_excel_files(equipment_file, maintenance_file)
            
            if error:
                st.error(error)
            else:
                st.success(f"✅ Merged {len(merged_data)} records!")
                with st.expander("📋 Data Preview"):
                    st.dataframe(merged_data.head(), use_container_width=True)
    
    st.markdown("---")
    
    # View selection
    st.subheader("👁️ View Mode")
    view_mode = st.radio(
        "Select View",
        ["🔍 Equipment Selection", "📊 Analytics", "📈 Summary"],
    )
    
    st.markdown("---")
    
    # Export
    st.subheader("💾 Export Data")
    if (st.session_state.equipment_data is not None and 
        st.session_state.selected_components):
        try:
            filtered = filter_data(
                st.session_state.equipment_data,
                st.session_state.selected_equipment,
                st.session_state.selected_module,
                st.session_state.selected_components
            )
            
            if filtered is not None and len(filtered) > 0:
                csv = filtered.to_csv(index=False)
                st.download_button(
                    label="📥 Download CSV",
                    data=csv,
                    file_name=f"equipment_{st.session_state.selected_equipment}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
        except:
            pass

# ============================================================================
# MAIN CONTENT
# ============================================================================

st.title("🔧 Equipment Maintenance Dashboard")
st.markdown("Upload two Excel files → Select Equipment → Filter Data → Analyze → Export")

if st.session_state.equipment_data is None:
    st.info("👈 Upload both Excel files in the sidebar to get started")
    st.stop()

# ============================================================================
# VIEW 1: EQUIPMENT SELECTION
# ============================================================================

if view_mode == "🔍 Equipment Selection":
    
    st.header("Equipment Selection & Data Preview")
    
    # FILTER 1: Equipment Code Selection (Square Cards)
    equipment_codes = get_equipment_codes(st.session_state.equipment_data)
    
    if equipment_codes:
        st.subheader("📦 Step 1: Select Equipment Code")
        
        cols = st.columns(3)
        for idx, code in enumerate(equipment_codes):
            with cols[idx % 3]:
                # Create beautiful equipment card
                is_selected = st.session_state.selected_equipment == code
                
                # Get type for this equipment
                equipment_type = get_type_for_equipment(
                    st.session_state.equipment_data,
                    code
                )
                
                card_html = f"""
                <div class="equipment-card {'selected' if is_selected else ''}">
                    <div class="equipment-card-content">
                        <p class="equipment-code">{code}</p>
                        <p class="equipment-type">{equipment_type if equipment_type else 'N/A'}</p>
                    </div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)
                
                if st.button(f"Select {code}", key=f"equip_{code}", use_container_width=True):
                    st.session_state.selected_equipment = code
                    st.session_state.selected_type = equipment_type
                    st.session_state.selected_module = None
                    st.session_state.selected_components = []
                    st.rerun()
        
        st.markdown("---")
        
        # FILTER 2: Type Display (Auto-populated)
        if st.session_state.selected_equipment:
            st.subheader("🏷️ Step 2: Type (Auto-populated)")
            col1, col2 = st.columns([2, 1])
            with col1:
                st.write(f"**Type:** {st.session_state.selected_type}")
            with col2:
                if st.button("🔄 Change Equipment"):
                    st.session_state.selected_equipment = None
                    st.session_state.selected_type = None
                    st.rerun()
            
            st.markdown("---")
            
            # FILTER 3: Module Selection
            modules = get_modules(st.session_state.equipment_data, st.session_state.selected_equipment)
            
            if modules:
                st.subheader("📊 Step 3: Select Module")
                
                # Create columns for module selection
                module_cols = st.columns(min(3, len(modules)))
                
                for idx, module in enumerate(modules):
                    with module_cols[idx % len(module_cols)]:
                        if st.button(
                            f"📊 {module}",
                            key=f"module_{module}",
                            use_container_width=True,
                            help=f"Click to select {module}"
                        ):
                            st.session_state.selected_module = module
                            st.session_state.selected_components = []
                            st.rerun()
                
                st.markdown("---")
                
                # FILTER 4: Components Selection
                if st.session_state.selected_module:
                    components = get_components(
                        st.session_state.equipment_data,
                        st.session_state.selected_equipment,
                        st.session_state.selected_module
                    )
                    
                    if components:
                        st.subheader("🔹 Step 4: Select Components")
                        
                        selected_components = st.multiselect(
                            "Choose components (Selections persist when switching views)",
                            options=components,
                            default=st.session_state.selected_components if st.session_state.selected_components else components[0:1],
                            key="component_select"
                        )
                        st.session_state.selected_components = selected_components
                        
                        st.markdown("---")
                        
                        # DATA DISPLAY & STATISTICS
                        if st.session_state.selected_components:
                            filtered_df = filter_data(
                                st.session_state.equipment_data,
                                st.session_state.selected_equipment,
                                st.session_state.selected_module,
                                st.session_state.selected_components
                            )
                            
                            if filtered_df is not None and not filtered_df.empty:
                                st.subheader(f"📊 Data Preview ({len(filtered_df)} records)")
                                st.dataframe(filtered_df, use_container_width=True, height=400)
                                
                                # Statistics
                                stats = calculate_stats(filtered_df)
                                
                                col1, col2, col3 = st.columns(3)
                                
                                with col1:
                                    st.markdown(f"""
                                    <div class="stat-card">
                                        <div class="stat-card-label">Records</div>
                                        <div class="stat-card-value">{stats['records']}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                
                                with col2:
                                    time_str = seconds_to_time_str(int(stats['total_time'])) if stats['total_time'] > 0 else "N/A"
                                    st.markdown(f"""
                                    <div class="stat-card">
                                        <div class="stat-card-label">Total Time</div>
                                        <div class="stat-card-value">{time_str}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                
                                with col3:
                                    st.markdown(f"""
                                    <div class="stat-card">
                                        <div class="stat-card-label">Avg Manpower</div>
                                        <div class="stat-card-value">{stats['avg_manpower']:.1f}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.warning("No data found for selected filters")

# ============================================================================
# VIEW 2: ANALYTICS
# ============================================================================

elif view_mode == "📊 Analytics":
    
    if not st.session_state.selected_equipment or not st.session_state.selected_components:
        st.warning("Please select equipment and components first")
        st.stop()
    
    filtered_df = filter_data(
        st.session_state.equipment_data,
        st.session_state.selected_equipment,
        st.session_state.selected_module,
        st.session_state.selected_components
    )
    
    if filtered_df is None or filtered_df.empty:
        st.error("No data found")
        st.stop()
    
    tab1, tab2, tab3 = st.tabs(["📋 Table", "⏱️ Time", "👥 Manpower"])
    
    with tab1:
        st.subheader("Complete Data")
        st.dataframe(filtered_df, use_container_width=True, height=500)
    
    with tab2:
        st.subheader("Time Analysis")
        if 'Preparation/Finalization (h:mm:ss)' in filtered_df.columns and 'Activity (h:mm:ss)' in filtered_df.columns:
            try:
                prep = filtered_df['Preparation/Finalization (h:mm:ss)'].apply(time_str_to_hours)
                activity = filtered_df['Activity (h:mm:ss)'].apply(time_str_to_hours)
                
                fig = go.Figure(data=[
                    go.Bar(name='Preparation', x=filtered_df['Components'], y=prep),
                    go.Bar(name='Activity', x=filtered_df['Components'], y=activity)
                ])
                
                fig.update_layout(barmode='stack', height=500, hovermode='x unified')
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Could not generate chart: {str(e)}")
    
    with tab3:
        st.subheader("Manpower Requirements")
        if 'No of man power' in filtered_df.columns:
            try:
                fig = px.bar(
                    filtered_df,
                    x='Components',
                    y='No of man power',
                    title="Manpower Needed",
                    color='No of man power',
                    color_continuous_scale='Viridis'
                )
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Could not generate chart: {str(e)}")

# ============================================================================
# VIEW 3: SUMMARY
# ============================================================================

elif view_mode == "📈 Summary":
    
    st.header("Summary Report")
    
    equipment_codes = get_equipment_codes(st.session_state.equipment_data)
    stats = calculate_stats(st.session_state.equipment_data)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-card-label">Equipment Codes</div>
            <div class="stat-card-value">{len(equipment_codes)}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-card-label">Total Records</div>
            <div class="stat-card-value">{len(st.session_state.equipment_data)}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-card-label">Average Manpower</div>
            <div class="stat-card-value">{stats['avg_manpower']:.1f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Equipment Summary")
    
    summary_data = []
    for equip_code in equipment_codes:
        equip_df = st.session_state.equipment_data[
            st.session_state.equipment_data['Equipment Code'] == equip_code
        ]
        equip_type = get_type_for_equipment(st.session_state.equipment_data, equip_code)
        
        summary_data.append({
            'Equipment Code': equip_code,
            'Type': equip_type,
            'Records': len(equip_df),
        })
    
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True)

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    st.caption("🔧 Equipment Maintenance Dashboard v1.0")
with col2:
    if st.session_state.equipment_data is not None:
        st.caption(f"Records: {len(st.session_state.equipment_data)}")
with col3:
    st.caption("✅ Multi-Excel Integration | 📊 Cascading Filters")
