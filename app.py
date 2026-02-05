"""
KONE EQUIPMENT MAINTENANCE DASHBOARD - UPDATED v1.1
Version: 1.1 - Single Excel File with Multiple Sheets

KEY CHANGE:
✅ Now uses ONE Excel file instead of two
✅ Reads from Sheet 2 (user selectable)
✅ All data in one sheet
✅ Future-ready: Can scale to two separate sheets later

CURRENT FLOW:
  Excel File (Multiple Sheets):
    └─ Sheet 2 (or any selected sheet)
       ├─ Equipment Code
       ├─ Type
       ├─ Module
       ├─ Components
       ├─ Preparation/Finalization time
       ├─ Activity time
       ├─ Total time
       └─ No of man power

USER SELECTION:
  1. Upload single Excel file
  2. Select which sheet to use (default: Sheet 2)
  3. Select Equipment Code
  4. Type auto-populated
  5. Select Module
  6. Select Components
  7. View data & export
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

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

if 'equipment_data' not in st.session_state:
    st.session_state.equipment_data = None
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None
if 'selected_equipment' not in st.session_state:
    st.session_state.selected_equipment = None
if 'selected_type' not in st.session_state:
    st.session_state.selected_type = None
if 'selected_module' not in st.session_state:
    st.session_state.selected_module = None
if 'selected_components' not in st.session_state:
    st.session_state.selected_components = []
if 'available_sheets' not in st.session_state:
    st.session_state.available_sheets = []

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
        cleaned = str(code).replace(',', '').strip()
        # Convert to integer and back to string (ensures it's a proper number)
        return str(int(cleaned))
    except:
        return str(code).strip()

# ============================================================================
# DATA LOADING - SINGLE EXCEL FILE WITH SHEET SELECTION
# ============================================================================

def get_excel_sheets(excel_file):
    """Get list of sheet names from Excel file"""
    try:
        xls = pd.ExcelFile(excel_file)
        return xls.sheet_names
    except Exception as e:
        return None, f"❌ Error reading Excel: {str(e)}"

def load_sheet_data(excel_file, sheet_name):
    """
    Load data from selected sheet
    
    Expected columns:
      - Equipment Code
      - Type
      - Module
      - Components
      - Preparation/Finalization time (or similar)
      - Activity time (or similar)
      - Total time (or similar)
      - No of man power
    """
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        if df.empty:
            return None, f"❌ Sheet '{sheet_name}' is empty!"
        
        # Clean equipment codes
        if 'Equipment Code' in df.columns:
            df['Equipment Code'] = df['Equipment Code'].apply(clean_equipment_code)
        
        # Forward fill Type, Module (handle merged cells)
        if 'Equipment Code' in df.columns:
            df['Equipment Code'] = df['Equipment Code'].fillna(method='ffill')
        if 'Type' in df.columns:
            df['Type'] = df['Type'].fillna(method='ffill')
        if 'Module' in df.columns:
            df['Module'] = df['Module'].fillna(method='ffill')
        
        st.session_state.equipment_data = df
        st.session_state.selected_sheet = sheet_name
        
        return df, None
    
    except Exception as e:
        return None, f"❌ Error loading sheet: {str(e)}"

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
    
    # Find the total time column (flexible naming)
    total_time_col = None
    for col in df.columns:
        if 'total' in col.lower() and 'time' in col.lower():
            total_time_col = col
            break
    
    try:
        if total_time_col and total_time_col in df.columns:
            total_secs = sum(time_str_to_seconds(t) for t in df[total_time_col])
            stats['total_time'] = total_secs
    except:
        pass
    
    # Find the manpower column (flexible naming)
    manpower_col = None
    for col in df.columns:
        if 'man' in col.lower() and 'power' in col.lower():
            manpower_col = col
            break
    
    try:
        if manpower_col and manpower_col in df.columns:
            manpower = pd.to_numeric(df[manpower_col], errors='coerce').dropna()
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
    
    st.subheader("📁 Upload Excel File")
    st.info("✅ Upload ONE Excel file with multiple sheets")
    
    # File upload
    excel_file = st.file_uploader(
        "Choose your Excel file",
        type=['xlsx', 'xls'],
        key='excel_file'
    )
    
    # Get sheets and let user select
    if excel_file is not None:
        sheets = get_excel_sheets(excel_file)
        
        if sheets:
            st.subheader("📄 Select Sheet")
            
            # Default to Sheet 2 if it exists, otherwise first sheet
            default_sheet = sheets[1] if len(sheets) > 1 else sheets[0]
            default_index = sheets.index(default_sheet) if default_sheet in sheets else 0
            
            selected_sheet = st.selectbox(
                "Which sheet contains the data?",
                options=sheets,
                index=default_index,
                help=f"Default: {default_sheet}"
            )
            
            # Load the selected sheet
            with st.spinner(f"📥 Loading '{selected_sheet}'..."):
                df, error = load_sheet_data(excel_file, selected_sheet)
                
                if error:
                    st.error(error)
                else:
                    st.success(f"✅ Loaded {len(df)} records from '{selected_sheet}'!")
                    
                    with st.expander("📋 Data Preview"):
                        st.dataframe(df.head(), use_container_width=True)
                        st.write(f"**Columns found**: {', '.join(df.columns.tolist())}")
    
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
st.markdown("Upload Excel → Select Sheet → Select Equipment → Filter Data → Analyze → Export")

if st.session_state.equipment_data is None:
    st.info("👈 Upload your Excel file in the sidebar to get started")
    st.stop()

df = st.session_state.equipment_data

# Verify required columns exist
required_cols = ['Equipment Code', 'Type', 'Module', 'Components']
missing_cols = [col for col in required_cols if col not in df.columns]

if missing_cols:
    st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
    st.info("**Expected columns:**")
    st.write("- Equipment Code")
    st.write("- Type")
    st.write("- Module")
    st.write("- Components")
    st.write("- Preparation/Finalization (h:mm:ss)")
    st.write("- Activity (h:mm:ss)")
    st.write("- Total time (h:mm:ss)")
    st.write("- No of man power")
    st.stop()

# ============================================================================
# VIEW 1: EQUIPMENT SELECTION
# ============================================================================

if view_mode == "🔍 Equipment Selection":
    
    st.header("Equipment Selection & Data Preview")
    st.caption(f"Sheet: **{st.session_state.selected_sheet}**")
    
    # FILTER 1: Equipment Code Selection (Square Cards)
    equipment_codes = get_equipment_codes(df)
    
    if equipment_codes:
        st.subheader("📦 Step 1: Select Equipment Code")
        
        cols = st.columns(3)
        for idx, code in enumerate(equipment_codes):
            with cols[idx % 3]:
                # Get type for this equipment
                equipment_type = get_type_for_equipment(df, code)
                
                is_selected = st.session_state.selected_equipment == code
                
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
            modules = get_modules(df, st.session_state.selected_equipment)
            
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
                        df,
                        st.session_state.selected_equipment,
                        st.session_state.selected_module
                    )
                    
                    if components:
                        st.subheader("🔹 Step 4: Select Components")
                        
                        selected_components = st.multiselect(
                            "Choose components",
                            options=components,
                            default=st.session_state.selected_components if st.session_state.selected_components else components[0:1],
                            key="component_select"
                        )
                        st.session_state.selected_components = selected_components
                        
                        st.markdown("---")
                        
                        # DATA DISPLAY & STATISTICS
                        if st.session_state.selected_components:
                            filtered_df = filter_data(
                                df,
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
        st.warning("Please select equipment and components first from Equipment Selection view")
        st.stop()
    
    filtered_df = filter_data(
        df,
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
        # Find time columns
        prep_col = None
        activity_col = None
        
        for col in df.columns:
            if 'preparation' in col.lower() or 'finalization' in col.lower():
                prep_col = col
            if 'activity' in col.lower():
                activity_col = col
        
        if prep_col and activity_col:
            try:
                prep = filtered_df[prep_col].apply(time_str_to_hours)
                activity = filtered_df[activity_col].apply(time_str_to_hours)
                
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
        manpower_col = None
        for col in df.columns:
            if 'man' in col.lower() and 'power' in col.lower():
                manpower_col = col
                break
        
        if manpower_col:
            try:
                fig = px.bar(
                    filtered_df,
                    x='Components',
                    y=manpower_col,
                    title="Manpower Needed",
                    color=manpower_col,
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
    st.caption(f"Sheet: **{st.session_state.selected_sheet}**")
    
    equipment_codes = get_equipment_codes(df)
    stats = calculate_stats(df)
    
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
            <div class="stat-card-value">{len(df)}</div>
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
        equip_df = df[df['Equipment Code'] == equip_code]
        equip_type = get_type_for_equipment(df, equip_code)
        
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
    st.caption("🔧 Equipment Maintenance Dashboard v1.1")
with col2:
    if st.session_state.equipment_data is not None:
        st.caption(f"Records: {len(st.session_state.equipment_data)} | Sheet: {st.session_state.selected_sheet}")
with col3:
    st.caption("✅ Single Excel with Sheets | 📊 Cascading Filters")
