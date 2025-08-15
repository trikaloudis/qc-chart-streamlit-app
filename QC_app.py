# QC_app.py
# To run this app, save the code as 'QC_app.py' and run the command: streamlit run QC_app.py
#
# Required libraries:
# streamlit
# pandas
# plotly
# openpyxl (for reading Excel files)

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
from collections import defaultdict

# --- Page Configuration ---
st.set_page_config(
    page_title="Advanced Quality Control Chart Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Helper Functions ---

@st.cache_data(ttl=600) # Cache data for 10 minutes
def load_data_from_excel(uploaded_file):
    """
    Loads data from an uploaded Excel file.
    Expects three sheets: 'QC data', 'Historical limits', 'Specification limits'.
    """
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            required_sheets = ['QC data', 'Historical limits', 'Specification limits']
            if not all(sheet in xls.sheet_names for sheet in required_sheets):
                st.error("Error: The Excel file must contain the sheets: 'QC data', 'Historical limits', and 'Specification limits'.")
                return None, None, None

            qc_data = pd.read_excel(xls, sheet_name='QC data')
            historical_limits = pd.read_excel(xls, sheet_name='Historical limits')
            spec_limits = pd.read_excel(xls, sheet_name='Specification limits')
            
            return qc_data, historical_limits, spec_limits
    except Exception as e:
        st.error(f"An error occurred while reading the Excel file: {e}")
        return None, None, None

@st.cache_data(ttl=60) # Cache data for 1 minute for refresh capability
def load_data_from_gsheet(url):
    """
    Loads data from a public Google Sheet URL.
    Constructs CSV export links for the required sheets.
    """
    try:
        match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
        if not match:
            st.error("Invalid Google Sheets URL. Please provide a valid URL.")
            return None, None, None
        sheet_id = match.group(1)

        base_csv_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet='

        qc_data = pd.read_csv(base_csv_url + 'QC%20data')
        historical_limits = pd.read_csv(base_csv_url + 'Historical%20limits')
        spec_limits = pd.read_csv(base_csv_url + 'Specification%20limits')

        return qc_data, historical_limits, spec_limits
    except Exception as e:
        st.error(f"Failed to load data from Google Sheets. Ensure the link is correct and the sheet is public. Error: {e}")
        return None, None, None

def process_limits_df(df):
    """
    Processes the limits dataframes to have a consistent structure.
    """
    df_processed = df.T
    df_processed.columns = ['mean', 'std']
    df_processed.index.name = 'parameter'
    return df_processed.astype(float)

def apply_westgard_rules(series, mean, std_dev):
    """
    Applies Westgard rules to a data series and returns points that violate them.
    """
    violations = defaultdict(list)
    s_p1 = mean + std_dev
    s_p2 = mean + 2 * std_dev
    s_p3 = mean + 3 * std_dev
    s_m1 = mean - std_dev
    s_m2 = mean - 2 * std_dev
    s_m3 = mean - 3 * std_dev

    # Rule 1-3s: One point outside ±3s
    violations['1-3s'] = series[(series > s_p3) | (series < s_m3)].index.tolist()

    # Rule 2-2s: Two consecutive points on same side, outside ±2s
    for i in range(1, len(series)):
        if (series.iloc[i] > s_p2 and series.iloc[i-1] > s_p2) or \
           (series.iloc[i] < s_m2 and series.iloc[i-1] < s_m2):
            violations['2-2s'].extend([series.index[i-1], series.index[i]])

    # Rule R-4s: Range between two consecutive points > 4s
    for i in range(1, len(series)):
        if abs(series.iloc[i] - series.iloc[i-1]) > 4 * std_dev:
            violations['R-4s'].extend([series.index[i-1], series.index[i]])

    # Rule 4-1s: Four consecutive points on same side, outside ±1s
    for i in range(3, len(series)):
        last_four = series.iloc[i-3:i+1]
        if all(last_four > s_p1) or all(last_four < s_m1):
            violations['4-1s'].extend(last_four.index)

    # Rule 10-x: Ten consecutive points on same side of the mean
    for i in range(9, len(series)):
        last_ten = series.iloc[i-9:i+1]
        if all(last_ten > mean) or all(last_ten < mean):
            violations['10-x'].extend(last_ten.index)

    # NEW Rule 7-T: Seven consecutive points trending in one direction
    for i in range(6, len(series)):
        last_seven = series.iloc[i-6:i+1]
        # Check for increasing trend
        if all(last_seven.iloc[j] > last_seven.iloc[j-1] for j in range(1, len(last_seven))):
            violations['7-T'].extend(last_seven.index)
        # Check for decreasing trend
        if all(last_seven.iloc[j] < last_seven.iloc[j-1] for j in range(1, len(last_seven))):
            violations['7-T'].extend(last_seven.index)
            
    # Remove duplicates
    for rule in violations:
        violations[rule] = sorted(list(set(violations[rule])))
        
    return violations

def create_qc_chart(data, date_col, param_col, cl, ucl, lcl, std_dev, violations, applied_rules):
    """
    Generates an interactive I-Chart using Plotly, highlighting rule violations.
    """
    fig = go.Figure()
    
    # Define zones for plotting (+-2s)
    zones = {
        '+2s': cl + 2 * std_dev,
        '-2s': cl - 2 * std_dev
    }

    # Add zone lines (+-2s dashed)
    for zone_name, zone_val in zones.items():
        fig.add_hline(y=zone_val, line_dash="dash", line_color="orange", opacity=0.7,
                      annotation_text=zone_name, annotation_position="bottom right")

    # Add control limit lines (+-3s solid)
    fig.add_hline(y=ucl, line_dash="solid", line_color="red", annotation_text=f"UCL (+3s): {ucl:.2f}")
    fig.add_hline(y=cl, line_dash="solid", line_color="green", annotation_text=f"Center: {cl:.2f}")
    fig.add_hline(y=lcl, line_dash="solid", line_color="red", annotation_text=f"LCL (-3s): {lcl:.2f}")

    # Add data points trace
    fig.add_trace(go.Scatter(
        x=data[date_col], y=data[param_col], mode='lines+markers', name=param_col,
        marker=dict(color='#1f77b4'), line=dict(color='#1f77b4')
    ))

    # Highlight violations
    rule_colors = {'1-3s': 'red', '2-2s': 'orange', 'R-4s': 'purple', '4-1s': 'brown', '10-x': 'pink', '7-T': 'cyan'}
    rule_symbols = {'1-3s': 'x', '2-2s': 'diamond', 'R-4s': 'star', '4-1s': 'square', '10-x': 'triangle-up', '7-T': 'hourglass'}

    for rule in applied_rules:
        points = violations.get(rule, [])
        if points:
            violation_data = data.loc[points]
            fig.add_trace(go.Scatter(
                x=violation_data[date_col], y=violation_data[param_col],
                mode='markers', name=f'Violation: {rule}',
                marker=dict(color=rule_colors[rule], size=12, symbol=rule_symbols[rule], line=dict(width=2, color='DarkSlateGrey'))
            ))

    # Update layout with larger fonts
    fig.update_layout(
        title=dict(text=f'Individual Control Chart (I-Chart) for {param_col}', font=dict(size=22)),
        xaxis_title=dict(text='Date', font=dict(size=18)),
        yaxis_title=dict(text='Measurement Value', font=dict(size=18)),
        legend=dict(font=dict(size=14), orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    # Update X-axis to show all dates vertically
    fig.update_xaxes(
        tickfont=dict(size=12),
        tickangle=-90,
        tickmode='array',
        tickvals=data[date_col],
        tickformat='%Y-%m-%d'
    )
    fig.update_yaxes(tickfont=dict(size=14))
    
    return fig

# --- Main Application ---
def main():
    st.title("📊 Advanced Quality Control Chart Generator")
    st.markdown("Generate **Individual Control Charts (I-Charts)** with interactive **Westgard Rules** analysis.")

    with st.sidebar:
        st.header("1. Data Source")
        input_method = st.radio("Choose input method:", ("Upload Excel File", "Google Sheets URL"))

        qc_data, historical_limits, spec_limits = None, None, None

        if input_method == "Upload Excel File":
            uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], help="Must contain 'QC data', 'Historical limits', and 'Specification limits' sheets.")
            if uploaded_file:
                qc_data, historical_limits, spec_limits = load_data_from_excel(uploaded_file)

        elif input_method == "Google Sheets URL":
            gsheet_url = st.text_input("Enter your public Google Sheets URL", st.session_state.get("gsheet_url", ""))
            if gsheet_url:
                st.session_state["gsheet_url"] = gsheet_url
                if st.button("Refresh Data"):
                    st.cache_data.clear()
                qc_data, historical_limits, spec_limits = load_data_from_gsheet(gsheet_url)
        
        # --- Add logo and hyperlink at the bottom of the sidebar ---
        st.markdown("---")
        # Replace with the raw URL of your logo from your GitHub repository
        logo_url = "Aquomixlab Logo v2 white font.jpg"
        st.image(logo_url, use_container_width=True)
        st.markdown(
            "<div style='text-align: center;'><a href='https://www.aquomixlab.com/'>https://www.aquomixlab.com/</a></div>",
            unsafe_allow_html=True
        )


    if qc_data is not None and historical_limits is not None and spec_limits is not None:
        try:
            date_column = qc_data.columns[0]
            qc_data[date_column] = pd.to_datetime(qc_data[date_column])
            parameters = qc_data.columns[1:].tolist()
            historical_limits_processed = process_limits_df(historical_limits)
            spec_limits_processed = process_limits_df(spec_limits)

            st.header("2. Select Parameters for Charting")
            col1, col2 = st.columns([1, 2])
            with col1:
                selected_parameter = st.selectbox("Select a parameter:", options=parameters)
                limit_method = st.radio("Calculate control limits using:", ("Calculate from QC data", "Use Historical limits", "Use Specification limits"))
            
            with col2:
                westgard_rules_options = ['1-3s', '2-2s', 'R-4s', '4-1s', '10-x', '7-T']
                applied_rules = st.multiselect("Apply Westgard Rules:", options=westgard_rules_options, default=westgard_rules_options)

            if selected_parameter:
                st.header(f"3. Control Chart for {selected_parameter}")
                
                mean, std_dev = 0, 0
                
                if limit_method == "Calculate from QC data":
                    mean = qc_data[selected_parameter].mean()
                    std_dev = qc_data[selected_parameter].std()
                elif limit_method == "Use Historical limits":
                    mean = historical_limits_processed.loc[selected_parameter, 'mean']
                    std_dev = historical_limits_processed.loc[selected_parameter, 'std']
                elif limit_method == "Use Specification limits":
                    mean = spec_limits_processed.loc[selected_parameter, 'mean']
                    std_dev = spec_limits_processed.loc[selected_parameter, 'std']

                st.info(f"Control limits for **{selected_parameter}** based on **{limit_method}**: Mean = {mean:.3f}, Std Dev = {std_dev:.3f}")

                if pd.isna(std_dev) or std_dev == 0:
                    st.warning("Standard deviation is zero or missing. Cannot calculate control limits or apply Westgard rules.")
                else:
                    center_line = mean
                    upper_control_limit = mean + 3 * std_dev
                    lower_control_limit = mean - 3 * std_dev

                    violations = apply_westgard_rules(qc_data[selected_parameter], mean, std_dev)
                    
                    fig = create_qc_chart(
                        qc_data, date_column, selected_parameter,
                        center_line, upper_control_limit, lower_control_limit,
                        std_dev, violations, applied_rules
                    )
                    st.plotly_chart(fig, use_container_width=True)

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
            st.warning("Please ensure your data is formatted correctly.")
    else:
        st.info("Awaiting data... Please upload a file or provide a Google Sheets URL in the sidebar.")
        # ... (Instructions remain the same) ...

if __name__ == "__main__":
    main()

