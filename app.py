import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import io
import base64
from datetime import datetime

# Set page configuration
st.set_page_config(
    page_title="LTR Data Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constants
DATA_PATH = "processed_data"
EPICS_FILE = "processed_epics.csv"
MAINTENANCE_FILE = "processed_maintenance.csv"
UTILIZATION_FILE = "processed_utilization.csv"
CORRELATION_FILE = "correlation_data.csv"

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 600;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.8rem;
        font-weight: 500;
        color: #0D47A1;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #f5f5f5;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        text-align: center;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 600;
        color: #1E88E5;
    }
    .metric-label {
        font-size: 1rem;
        color: #616161;
    }
    .insight-box {
        background-color: #e3f2fd;
        border-left: 5px solid #1E88E5;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 0 5px 5px 0;
    }
    .recommendation-box {
        background-color: #e8f5e9;
        border-left: 5px solid #43a047;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 0 5px 5px 0;
    }
    .warning-box {
        background-color: #fff3e0;
        border-left: 5px solid #ff9800;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 0 5px 5px 0;
    }
    .divider {
        margin: 2rem 0;
        border-top: 1px solid #e0e0e0;
    }
    /* Custom navigation styles are now handled directly in the sidebar_nav function */
    .report-type-button {
        display: inline-block;
        padding: 1rem 2rem;
        margin: 0.5rem;
        border-radius: 10px;
        background-color: #f0f2f6;
        border: 2px solid #e6e9ef;
        transition: all 0.3s ease;
        cursor: pointer;
        text-align: center;
        width: 100%;
    }
    .report-type-button:hover {
        background-color: #e6e9ef;
        border-color: #d0d4db;
        transform: translateY(-2px);
    }
    .report-type-button.selected {
        background-color: #1E88E5;
        color: white;
        border-color: #1E88E5;
    }
    .report-type-icon {
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }
    .report-type-title {
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    .report-type-desc {
        font-size: 0.9rem;
        color: #666;
    }
    .selected .report-type-desc {
        color: #e6e9ef;
    }
    .report-type-container {
        display: flex;
        gap: 1rem;
        margin: 1rem 0;
    }
    .report-type-card {
        flex: 1;
        padding: 1.5rem;
        border-radius: 10px;
        background-color: #f0f2f6;
        border: 2px solid #e6e9ef;
        transition: all 0.3s ease;
        cursor: pointer;
        text-align: center;
    }
    .report-type-card:hover {
        background-color: #e6e9ef;
        border-color: #d0d4db;
        transform: translateY(-2px);
    }
    .report-type-card.selected {
        background-color: #1E88E5;
        color: white;
        border-color: #1E88E5;
    }
    .report-type-icon {
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }
    .report-type-title {
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    .report-type-desc {
        font-size: 0.9rem;
        color: #666;
    }
    .selected .report-type-desc {
        color: #e6e9ef;
    }
</style>
""", unsafe_allow_html=True)
# Helper functions
@st.cache_data
def load_data():
    """Load processed data files"""
    try:
        epics_df = pd.read_csv(os.path.join(DATA_PATH, EPICS_FILE))
        maintenance_df = pd.read_csv(os.path.join(DATA_PATH, MAINTENANCE_FILE))
        utilization_df = pd.read_csv(os.path.join(DATA_PATH, UTILIZATION_FILE))
        correlation_df = pd.read_csv(os.path.join(DATA_PATH, CORRELATION_FILE))
        
        # Convert date columns
        date_cols = {
            'epics': ['created', 'updated', 'duedate', 'Completed Date', 'Start Date'],
            'maintenance': ['created', 'updated', 'duedate', 'Updated Completed Date', 'Start Date', 'Completed Date'],
            'utilization': ['Created On', 'Start', 'End', 'Date', 'Start UTC', 'End UTC']
        }
        
        for col in date_cols['epics']:
            if col in epics_df.columns:
                epics_df[col] = pd.to_datetime(epics_df[col], errors='coerce')
                
        for col in date_cols['maintenance']:
            if col in maintenance_df.columns:
                maintenance_df[col] = pd.to_datetime(maintenance_df[col], errors='coerce')
                
        for col in date_cols['utilization']:
            if col in utilization_df.columns:
                utilization_df[col] = pd.to_datetime(utilization_df[col], errors='coerce')
                
        return {
            'epics': epics_df,
            'maintenance': maintenance_df,
            'utilization': utilization_df,
            'correlation': correlation_df
        }
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

def create_metric_card(label, value, prefix="", suffix="", color="#1E88E5", help_text=None):
    """Create a styled metric card"""
    # Ensure label is never empty for accessibility
    if not label or label.strip() == "":
        label = "Metric"
        
    if help_text:
        label = f"{label} ℹ️"
        
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: {color};">{prefix}{value}{suffix}</div>
        <div class="metric-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)
    
    if help_text and st.checkbox(f"More info on {label}", False, key=f"help_{label}_{value}".replace(" ", "_")):
        st.info(help_text)

def create_downloadable_excel(dataframes, filename="dashboard_data.xlsx"):
    """Generate a downloadable Excel file with multiple sheets"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    buffer.seek(0)
    b64 = base64.b64encode(buffer.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel File</a>'
    return href

def generate_email_template(kpis, insights, recommendations):
    """Generate an email template for management reporting"""
    current_date = datetime.now().strftime("%B %d, %Y")
    
    email_template = f"""
    Subject: LTR Weekly Performance Report - Week 16, 2025
    
    Dear Management Team,
    
    I hope this email finds you well. Please find below the key performance indicators and insights from our automation and project management systems for the week ending April 21, 2025.
    
    **KEY PERFORMANCE INDICATORS:**
    
    """
    
    # Add KPIs
    for kpi, value in kpis.items():
        email_template += f"- {kpi}: {value}\n"
    
    email_template += """
    
    **KEY INSIGHTS:**
    
    """
    
    # Add insights
    for insight in insights:
        email_template += f"- {insight}\n"
    
    email_template += """
    
    **RECOMMENDATIONS:**
    
    """
    
    # Add recommendations
    for recommendation in recommendations:
        email_template += f"- {recommendation}\n"
    
    email_template += """
    
    The complete dashboard with detailed analytics is available for your review. Please let me know if you need any clarification or have questions about specific metrics.
    
    Best regards,
    
    [Your Name]
    Automation Team Lead
    """
    
    return email_template

# Sidebar navigation
def sidebar_nav():
    st.sidebar.title("Navigation")
    
    # Initialize session state for selected page if not exists
    if 'selected_page' not in st.session_state:
        st.session_state.selected_page = "Summary Overview"
    
    # Navigation options with icons
    nav_options = {
        "Summary Overview": "📊",
        "JIRA Epics Analysis": "📈",
        "Maintenance Analysis": "🔧",
        "Machine Utilization": "⚙️",
        "Report Generator": "📝"
    }
    
    # Create navigation tiles using Streamlit columns
    for page, icon in nav_options.items():
        tile_container = st.sidebar.container()
        is_active = st.session_state.selected_page == page
        
        # Style the container based on active state
        tile_style = """
            <style>
                div[data-testid="stVerticalBlock"]:has(button#nav-{}) {{
                    background-color: {};
                    border-radius: 0.5rem;
                    margin: 0.25rem 0;
                    transition: all 0.3s ease;
                }}
                div[data-testid="stVerticalBlock"]:has(button#nav-{}):hover {{
                    transform: translateX(5px);
                }}
            </style>
        """.format(
            page.replace(" ", "-"),
            "#e3f2fd" if is_active else "#f8f9fa",
            page.replace(" ", "-")
        )
        st.markdown(tile_style, unsafe_allow_html=True)
        
        # Create the button with icon and text
        if tile_container.button(
            f"{icon} {page}",
            key=f"nav-{page.replace(' ', '-')}",
            use_container_width=True
        ):
            st.session_state.selected_page = page
            st.rerun()
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filters")
    
    # Global date filter if we have time series data
    data = load_data()
    if data and 'epics' in data:
        # Get min and max dates from epics data
        min_date = data['epics']['created'].min().date()
        max_date = data['epics']['created'].max().date()
        date_range = st.date_input(
            "Filter by Date Range",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date,
            help="Select a date range to filter the dashboard data"
        )
    
    st.sidebar.markdown("---")
    st.sidebar.info("This dashboard provides insights into automation project status, bot performance, and maintenance activities.")
    
    return st.session_state.selected_page

# Page functions
def summary_overview_page():
    st.markdown('<div class="main-header">LTR Data Dashboard - Summary Overview</div>', unsafe_allow_html=True)
    
    data = load_data()
    if not data:
        st.error("Failed to load data. Please check the data files.")
        return
    
    # Extract key metrics
    epics_df = data['epics']
    maintenance_df = data['maintenance']
    utilization_df = data['utilization']
    correlation_df = data['correlation']
    
    # KPIs Row
    st.markdown('<div class="sub-header">Key Performance Indicators</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_epics = len(epics_df)
        completed_epics = len(epics_df[epics_df['Status'] == 'Done'])
        completion_rate = round((completed_epics / total_epics) * 100, 1) if total_epics > 0 else 0
        create_metric_card("Completion Rate", f"{completion_rate}%", color="#43a047", 
                          help_text="Percentage of completed epics relative to total epics")
    
    with col2:
        # Machine utilization
        avg_utilization = utilization_df['Machine_Utilization__'].mean() * 100
        create_metric_card("Avg Machine Utilization", f"{avg_utilization:.1f}%", color="#1e88e5",
                          help_text="Average percentage of time machines are actively running tasks")
    
    with col3:
        # Maintenance allocation
        if 'Maintenance_Time_Allocation_Percentage' in correlation_df.columns:
            maint_allocation = correlation_df['Maintenance_Time_Allocation_Percentage'].mean() * 100
            create_metric_card("Maintenance Allocation", f"{maint_allocation:.1f}%", color="#ff9800",
                              help_text="Percentage of time allocated to maintenance activities")
        else:
            create_metric_card("Maintenance Allocation", "N/A", color="#ff9800")
    
    with col4:
        # Success rate
        if 'Desktop_Run_Percent_Success' in correlation_df.columns:
            success_rate = correlation_df['Desktop_Run_Percent_Success'].mean() * 100
            create_metric_card("Bot Success Rate", f"{success_rate:.1f}%", color="#43a047",
                              help_text="Percentage of bot runs that completed successfully")
        else:
            success_rate = utilization_df[utilization_df['desktop_taskstatus'] == 'Succeeded']['desktop_taskstatus'].count() / len(utilization_df) * 100 if len(utilization_df) > 0 else 0
            create_metric_card("Bot Success Rate", f"{success_rate:.1f}%", color="#43a047")
    
    # Project Status Chart
    st.markdown('<div class="sub-header">Project Status Overview</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Status distribution
        status_counts = epics_df['Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Count']
        
        fig = px.bar(status_counts, x='Status', y='Count', 
                    title='Epic Status Distribution',
                    color='Status',
                    color_discrete_map={
                        'Done': '#4caf50',
                        'In Progress': '#2196f3',
                        'Development': '#2196f3',
                        'Backlog': '#9e9e9e',
                        'To Do': '#9e9e9e'
                    })
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Priority distribution
        priority_counts = epics_df['priority'].value_counts().reset_index()
        priority_counts.columns = ['Priority', 'Count']
        
        fig = px.pie(priority_counts, names='Priority', values='Count',
                    title='Epic Priority Distribution',
                    color='Priority',
                    color_discrete_map={
                        'Highest (P1)': '#f44336',
                        'High (P2)': '#ff9800',
                        'Medium (P3)': '#ffeb3b',
                        'Low (P4)': '#4caf50',
                        'Lowest (P5)': '#2196f3'
                    })
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # Machine Utilization Over Time
    st.markdown('<div class="sub-header">Machine Utilization Overview</div>', unsafe_allow_html=True)
    
    # Group by date and calculate average utilization
    if 'Created On' in utilization_df.columns:
        utilization_by_date = utilization_df.groupby(utilization_df['Created On'].dt.date).agg({
            'Machine_Utilization__': 'mean',
            'SumRuntime_duration__mins_': 'sum',
            'desktop_taskstatus': 'count'
        }).reset_index()
        
        utilization_by_date['Machine_Utilization__'] = utilization_by_date['Machine_Utilization__'] * 100
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig.add_trace(
            go.Scatter(
                x=utilization_by_date['Created On'],
                y=utilization_by_date['Machine_Utilization__'],
                name="Machine Utilization (%)",
                line=dict(color="#1e88e5", width=3)
            ),
            secondary_y=False
        )
        
        fig.add_trace(
            go.Bar(
                x=utilization_by_date['Created On'],
                y=utilization_by_date['desktop_taskstatus'],
                name="Number of Runs",
                marker_color="#90caf9"
            ),
            secondary_y=True
        )
        
        fig.update_layout(
            title="Machine Utilization and Run Count Over Time",
            xaxis_title="Date",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400
        )
        
        fig.update_yaxes(title_text="Utilization (%)", secondary_y=False)
        fig.update_yaxes(title_text="Number of Runs", secondary_y=True)
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Key Insights and Recommendations
    st.markdown('<div class="sub-header">Key Insights & Recommendations</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
        st.markdown("<b>Key Insights:</b>", unsafe_allow_html=True)
        st.markdown("""
        - Streamlit Bot Monitoring Dashboard is in active development
        - SAFE-T CLAIMS automation was completed in Week 16
        - Machine utilization is currently at 2.03%, indicating capacity for additional workflows
        - Maintenance activities consume approximately 3% of total time allocation
        """)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="recommendation-box">', unsafe_allow_html=True)
        st.markdown("<b>Recommendations:</b>", unsafe_allow_html=True)
        st.markdown("""
        - Increase machine utilization by expanding automation workflows
        - Complete the Streamlit Bot Monitoring Dashboard for improved visibility
        - Evaluate potential ROI of the new test project (RPA-1179)
        - Continue integration of Streamlit with Snowflake for enhanced reporting
        """)
        st.markdown('</div>', unsafe_allow_html=True)

def jira_epics_analysis_page():
    st.markdown('<div class="main-header">LTR Data Dashboard - JIRA Epics Analysis</div>', unsafe_allow_html=True)
    
    data = load_data()
    if not data:
        st.error("Failed to load data. Please check the data files.")
        return
    
    epics_df = data['epics']

    # Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        status_filter = st.multiselect(
            "Filter by Status",
            options=sorted(epics_df['Status'].unique()),
            default=sorted(epics_df['Status'].unique()),
            key="status_filter"
        )
    with col2:
        priority_filter = st.multiselect(
            "Filter by Priority",
            options=sorted(epics_df['priority'].unique()),
            default=sorted(epics_df['priority'].unique()),
            key="priority_filter"
        )
    with col3:
        assignee_filter = st.multiselect(
            "Filter by Assignee",
            options=sorted(epics_df['Assignee'].unique()),
            default=sorted(epics_df['Assignee'].unique()),
            key="assignee_filter"
        )
    
    # Apply filters
    filtered_df = epics_df[
        epics_df['Status'].isin(status_filter) &
        epics_df['priority'].isin(priority_filter) &
        epics_df['Assignee'].isin(assignee_filter)
    ]
    
    # JIRA Epics Timeline
    st.markdown('<div class="sub-header">JIRA Epics Timeline</div>', unsafe_allow_html=True)
    
    if not filtered_df.empty:
        # Create timeline chart
        filtered_df = filtered_df.sort_values('created')
        
        # Determine task completion
        filtered_df['completed'] = filtered_df['Status'] == 'Done'
        
        # Create color mapping for status
        color_map = {
            'Done': '#4caf50',
            'In Progress': '#2196f3',
            'Development': '#2196f3',
            'Backlog': '#9e9e9e',
            'To Do': '#9e9e9e'
        }
        
        fig = px.timeline(
            filtered_df,
            x_start='Start Date',
            x_end='Completed Date',
            y='Key',
            color='Status',
            hover_name='summary',
            hover_data=['priority', 'Assignee'],
            color_discrete_map=color_map,
            title='JIRA Epics Timeline'
        )
        
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No epics match the selected filters.")
    
    # Epics Details Table
    st.markdown('<div class="sub-header">JIRA Epics Details</div>', unsafe_allow_html=True)
    
    # Select columns for display
    display_cols = ['Key', 'summary', 'Status', 'priority', 'Assignee', 'created', 'Completed Date', 'Estimated Financial Impact']
    display_df = filtered_df[display_cols].sort_values('created', ascending=False)
    
    # Add styling to the table
    st.dataframe(
        display_df.style.apply(
            lambda row: [
                'background-color: #e8f5e9' if row['Status'] == 'Done' else 
                'background-color: #e3f2fd' if row['Status'] in ['In Progress', 'Development'] else 
                'background-color: #fff3e0' if row['Status'] in ['Backlog', 'To Do'] else ''
                for _ in display_cols
            ],
            axis=1
        ),
        height=400,
        use_container_width=True
    )
    
    # Financial Impact Analysis
    st.markdown('<div class="sub-header">Financial Impact Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Financial Impact by Status
        fin_impact_by_status = filtered_df.groupby('Status')['Estimated Financial Impact'].sum().reset_index()
        fin_impact_by_status = fin_impact_by_status.sort_values('Estimated Financial Impact', ascending=False)
        
        fig = px.bar(
            fin_impact_by_status,
            x='Status',
            y='Estimated Financial Impact',
            color='Status',
            color_discrete_map=color_map,
            title='Estimated Financial Impact by Status'
        )
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Financial Impact by Assignee
        fin_impact_by_assignee = filtered_df.groupby('Assignee')['Estimated Financial Impact'].sum().reset_index()
        fin_impact_by_assignee = fin_impact_by_assignee.sort_values('Estimated Financial Impact', ascending=False)
        
        fig = px.pie(
            fin_impact_by_assignee,
            names='Assignee',
            values='Estimated Financial Impact',
            title='Estimated Financial Impact by Assignee'
        )
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)

def maintenance_analysis_page():
    st.markdown('<div class="main-header">LTR Data Dashboard - Maintenance Analysis</div>', unsafe_allow_html=True)
    
    data = load_data()
    if not data:
        st.error("Failed to load data. Please check the data files.")
        return
    
    maintenance_df = data['maintenance']
    
    # Maintenance KPIs
    st.markdown('<div class="sub-header">Maintenance Key Metrics</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_maintenance_hours = maintenance_df['SumMaintenance_Hours'].sum()
        create_metric_card("Total Maintenance Hours", f"{total_maintenance_hours:.1f}", color="#ff9800")
    
    with col2:
        avg_maintenance_allocation = maintenance_df['Maintenance_Time_Allocation_Percentage'].mean() * 100
        create_metric_card("Avg Time Allocation", f"{avg_maintenance_allocation:.1f}%", color="#ff9800")
    
    with col3:
        total_maintenance_tickets = maintenance_df['Total_Maintenance_Tickets_By_Week'].sum()
        create_metric_card("Total Tickets", f"{int(total_maintenance_tickets)}", color="#ff9800")
    
    with col4:
        if 'Bug_Completed_Count_Last_Week' in maintenance_df.columns:
            bug_completion = maintenance_df['Bug_Completed_Count_Last_Week'].sum()
            create_metric_card("Bugs Fixed Last Week", f"{int(bug_completion)}", color="#ff9800")
        else:
            create_metric_card("Bugs Fixed Last Week", "N/A", color="#ff9800")
    
    # Maintenance Activity Analysis
    st.markdown('<div class="sub-header">Maintenance Activity Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Maintenance by Issue Type
        if 'Issue Type' in maintenance_df.columns:
            issue_type_counts = maintenance_df['Issue Type'].value_counts().reset_index()
            issue_type_counts.columns = ['Issue Type', 'Count']
            
            fig = px.pie(
                issue_type_counts,
                names='Issue Type',
                values='Count',
                title='Maintenance by Issue Type'
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Issue Type data not available")
    
    with col2:
        # Maintenance Hours by Priority
        if 'priority' in maintenance_df.columns:
            maintenance_by_priority = maintenance_df.groupby('priority')['SumMaintenance_Hours'].sum().reset_index()
            maintenance_by_priority = maintenance_by_priority.sort_values('SumMaintenance_Hours', ascending=False)
            
            fig = px.bar(
                maintenance_by_priority,
                x='priority',
                y='SumMaintenance_Hours',
                color='priority',
                title='Maintenance Hours by Priority'
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Priority data not available")
    
    # Maintenance Tickets Details
    st.markdown('<div class="sub-header">Maintenance Tickets Details</div>', unsafe_allow_html=True)
    
    # Select relevant columns
    display_cols = [col for col in ['Key', 'summary', 'Issue Type', 'Status', 'priority', 
                                  'created', 'Completed Date', 'SumMaintenance_Hours', 'Reason for Failure'] 
                   if col in maintenance_df.columns]
    
    if display_cols:
        display_df = maintenance_df[display_cols].sort_values('created', ascending=False)
        st.dataframe(display_df, height=400, use_container_width=True)
    else:
        st.warning("No suitable data for display")

    # Maintenance Recommendations
    st.markdown('<div class="sub-header">Maintenance Insights & Recommendations</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="recommendation-box">', unsafe_allow_html=True)
    st.markdown("<b>Maintenance Recommendations:</b>", unsafe_allow_html=True)
    st.markdown("""
    - Focus on resolving the highest priority maintenance items first
    - Schedule routine maintenance to reduce unplanned downtime
    - Document common failure patterns to prevent recurring issues
    - Implement automated monitoring for early issue detection
    """)
    st.markdown('</div>', unsafe_allow_html=True)

def machine_utilization_page():
    st.markdown('<div class="main-header">LTR Data Dashboard - Machine Utilization</div>', unsafe_allow_html=True)
    
    data = load_data()
    if not data:
        st.error("Failed to load data. Please check the data files.")
        return
    
    utilization_df = data['utilization']
    
    # Utilization KPIs
    st.markdown('<div class="sub-header">Machine Utilization Key Metrics</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        avg_utilization = utilization_df['Machine_Utilization__'].mean() * 100
        create_metric_card("Avg Machine Utilization", f"{avg_utilization:.1f}%", color="#1e88e5")
    
    with col2:
        avg_idle = utilization_df['Idle_Percentage__'].mean() * 100
        create_metric_card("Avg Idle Time", f"{avg_idle:.1f}%", color="#ff9800")
    
    with col3:
        total_runtime = utilization_df['SumRuntime_duration__mins_'].sum()
        runtime_hours = total_runtime / 60
        create_metric_card("Total Runtime", f"{runtime_hours:.1f}", suffix=" hours", color="#1e88e5")
    
    with col4:
        if 'desktop_taskstatus' in utilization_df.columns:
            success_count = utilization_df[utilization_df['desktop_taskstatus'] == 'Succeeded']['desktop_taskstatus'].count()
            total_count = len(utilization_df)
            success_rate = (success_count / total_count) * 100 if total_count > 0 else 0
            create_metric_card("Success Rate", f"{success_rate:.1f}%", color="#43a047")
        else:
            success_rate = 0
            create_metric_card("Success Rate", "N/A", color="#43a047")
    
    # Success/Failure Analysis
    st.markdown('<div class="sub-header">Success/Failure Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if 'desktop_taskstatus' in utilization_df.columns:
            # Success/Failure distribution
            status_counts = utilization_df['desktop_taskstatus'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            
            fig = px.pie(
                status_counts,
                names='Status',
                values='Count',
                title='Bot Run Status Distribution',
                color='Status',
                color_discrete_map={
                    'Succeeded': '#4caf50',
                    'Failed': '#f44336',
                    'Cancelled': '#ff9800',
                    'Terminated': '#9e9e9e'
                }
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Task status data not available")
    
    with col2:
        if 'ErrorCode' in utilization_df.columns:
            # Error distribution (exclude empty error codes)
            error_df = utilization_df[utilization_df['ErrorCode'].notna() & (utilization_df['ErrorCode'] != 'Unknown')]
            if not error_df.empty:
                error_counts = error_df['ErrorCode'].value_counts().reset_index()
                error_counts.columns = ['Error Code', 'Count']
                
                fig = px.bar(
                    error_counts.head(10),  # Top 10 errors
                    x='Error Code',
                    y='Count',
                    title='Top Error Codes',
                    color='Count',
                    color_continuous_scale='Reds'
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No error codes recorded in the data")
        else:
            st.warning("Error code data not available")
    
    # Machine-specific metrics
    st.markdown('<div class="sub-header">Machine-Specific Metrics</div>', unsafe_allow_html=True)
    
    if 'Flow Machine Group' in utilization_df.columns:
        # Group by machine
        machine_metrics = utilization_df.groupby('Flow Machine Group').agg({
            'Machine_Utilization__': 'mean',
            'SumRuntime_duration__mins_': 'sum',
            'desktop_taskstatus': 'count'
        }).reset_index()
        
        # Calculate success rate per machine
        machine_success = utilization_df[utilization_df['desktop_taskstatus'] == 'Succeeded'].groupby('Flow Machine Group').size().reset_index()
        machine_success.columns = ['Flow Machine Group', 'SuccessCount']
        
        machine_metrics = pd.merge(machine_metrics, machine_success, on='Flow Machine Group', how='left')
        machine_metrics['SuccessCount'] = machine_metrics['SuccessCount'].fillna(0)
        machine_metrics['SuccessRate'] = (machine_metrics['SuccessCount'] / machine_metrics['desktop_taskstatus']) * 100
        
        # Multiply utilization to get percentage
        machine_metrics['Machine_Utilization__'] = machine_metrics['Machine_Utilization__'] * 100
        
        # Convert runtime to hours
        machine_metrics['Runtime_Hours'] = machine_metrics['SumRuntime_duration__mins_'] / 60
        
        # Sort by utilization
        machine_metrics = machine_metrics.sort_values('Machine_Utilization__', ascending=False)
        
        # Display as a bar chart
        fig = px.bar(
            machine_metrics,
            x='Flow Machine Group',
            y='Machine_Utilization__',
            color='SuccessRate',
            color_continuous_scale='RdYlGn',
            title='Machine Utilization by Group',
            hover_data=['Runtime_Hours', 'desktop_taskstatus']
        )
        
        fig.update_layout(
            yaxis_title="Utilization (%)",
            xaxis_title="Machine Group",
            height=500
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Show detailed metrics table
        st.markdown('<div class="sub-header">Machine Details</div>', unsafe_allow_html=True)
        
        # Rename columns for display
        display_cols = {
            'Flow Machine Group': 'Machine Group',
            'Machine_Utilization__': 'Utilization (%)',
            'Runtime_Hours': 'Runtime (Hours)',
            'desktop_taskstatus': 'Total Runs',
            'SuccessCount': 'Successful Runs',
            'SuccessRate': 'Success Rate (%)'
        }
        
        display_df = machine_metrics.rename(columns=display_cols)
        display_df = display_df[list(display_cols.values())]
        
        # Format numeric columns
        display_df['Utilization (%)'] = display_df['Utilization (%)'].round(2)
        display_df['Runtime (Hours)'] = display_df['Runtime (Hours)'].round(2)
        display_df['Success Rate (%)'] = display_df['Success Rate (%)'].round(2)
        
        st.dataframe(display_df, use_container_width=True)
    else:
        st.warning("Machine group data not available")
    
    # Utilization Trends
    st.markdown('<div class="sub-header">Utilization Trends</div>', unsafe_allow_html=True)
    
    if 'Created On' in utilization_df.columns and 'Hour_of_Day' in utilization_df.columns:
        # Analyze utilization by hour of day
        hourly_utilization = utilization_df.groupby('Hour_of_Day').agg({
            'Machine_Utilization__': 'mean',
            'desktop_taskstatus': 'count'
        }).reset_index()
        
        hourly_utilization['Machine_Utilization__'] = hourly_utilization['Machine_Utilization__'] * 100
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig.add_trace(
            go.Scatter(
                x=hourly_utilization['Hour_of_Day'],
                y=hourly_utilization['Machine_Utilization__'],
                name="Utilization (%)",
                line=dict(color="#1e88e5", width=3)
            ),
            secondary_y=False
        )
        
        fig.add_trace(
            go.Bar(
                x=hourly_utilization['Hour_of_Day'],
                y=hourly_utilization['desktop_taskstatus'],
                name="Run Count",
                marker_color="#90caf9"
            ),
            secondary_y=True
        )
        
        fig.update_layout(
            title="Utilization and Run Count by Hour of Day",
            xaxis_title="Hour of Day",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400
        )
        
        fig.update_yaxes(title_text="Utilization (%)", secondary_y=False)
        fig.update_yaxes(title_text="Number of Runs", secondary_y=True)
        
        st.plotly_chart(fig, use_container_width=True)
    elif 'Created On' in utilization_df.columns:
        # Extract hour from Created On
        utilization_df['Hour'] = utilization_df['Created On'].dt.hour
        
        # Analyze utilization by hour of day
        hourly_utilization = utilization_df.groupby('Hour').agg({
            'Machine_Utilization__': 'mean',
            'desktop_taskstatus': 'count'
        }).reset_index()
        
        hourly_utilization['Machine_Utilization__'] = hourly_utilization['Machine_Utilization__'] * 100
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig.add_trace(
            go.Scatter(
                x=hourly_utilization['Hour'],
                y=hourly_utilization['Machine_Utilization__'],
                name="Utilization (%)",
                line=dict(color="#1e88e5", width=3)
            ),
            secondary_y=False
        )
        
        fig.add_trace(
            go.Bar(
                x=hourly_utilization['Hour'],
                y=hourly_utilization['desktop_taskstatus'],
                name="Run Count",
                marker_color="#90caf9"
            ),
            secondary_y=True
        )
        
        fig.update_layout(
            title="Utilization and Run Count by Hour of Day",
            xaxis_title="Hour of Day",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400
        )
        
        fig.update_yaxes(title_text="Utilization (%)", secondary_y=False)
        fig.update_yaxes(title_text="Number of Runs", secondary_y=True)
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Time-based analysis data not available")
    
    # Utilization Insights
    st.markdown('<div class="sub-header">Utilization Insights & Recommendations</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
    st.markdown("<b>Utilization Insights:</b>", unsafe_allow_html=True)
    st.markdown("""
    - Current machine utilization is at {:.1f}%, indicating significant available capacity
    - Success rate is currently at {:.1f}%, exceeding the 85% goal
    - Runtime patterns suggest optimal processing during off-hours
    - Error patterns suggest focus areas for reliability improvements
    """.format(avg_utilization, success_rate))
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="recommendation-box">', unsafe_allow_html=True)
    st.markdown("<b>Utilization Recommendations:</b>", unsafe_allow_html=True)
    st.markdown("""
    - Increase workflow volume to optimize machine utilization
    - Implement load balancing across machine groups
    - Schedule resource-intensive jobs during low-utilization periods
    - Address recurring errors to improve overall success rate
    """)
    st.markdown('</div>', unsafe_allow_html=True)

def save_report_to_file(report, output_dir="reports"):
    """
    Save the report to both HTML and Excel formats in the specified directory
    """
    try:
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save HTML report
        html_report = generate_report_html(report)
        html_filename = f"ltr_report_{timestamp}.html"
        with open(os.path.join(output_dir, html_filename), 'w') as f:
            f.write(html_report)
        
        # Save Excel report
        excel_data = {
            'Data_Sources': pd.DataFrame(report['data_sources']),
            'Metrics': pd.DataFrame([report['metrics']]),
            'Insights': pd.DataFrame({'Insights': report['insights']}),
            'Recommendations': pd.DataFrame({'Recommendations': report['recommendations']})
        }
        excel_filename = f"ltr_report_{timestamp}.xlsx"
        with pd.ExcelWriter(os.path.join(output_dir, excel_filename), engine='openpyxl') as writer:
            for sheet_name, df in excel_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return html_filename, excel_filename
    except Exception as e:
        st.error(f"Error saving report: {e}")
        return None, None

def generate_automated_report():
    """
    Generate and save a report automatically
    """
    try:
        # Load data
        data = load_data()
        if data is None:
            st.error("Failed to load data. Please check the data files.")
            return
        
        # Generate report
        report = generate_comprehensive_report(data)
        
        # Save report
        html_file, excel_file = save_report_to_file(report)
        
        if html_file and excel_file:
            st.success(f"Report generated successfully!")
            st.info(f"HTML report saved as: {html_file}")
            st.info(f"Excel report saved as: {excel_file}")
        else:
            st.error("Failed to save report files.")
    except Exception as e:
        st.error(f"Error generating automated report: {e}")

def generate_comprehensive_report(data):
    """
    Generate a comprehensive report of all data sources and their analysis
    """
    report = {
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'data_sources': {},
        'metrics': {},
        'insights': [],
        'recommendations': []
    }
    
    # Analyze JIRA Epics Data
    epics_df = data['epics']
    report['data_sources']['jira_epics'] = {
        'total_records': len(epics_df),
        'date_range': {
            'start': epics_df['created'].min().strftime("%Y-%m-%d"),
            'end': epics_df['created'].max().strftime("%Y-%m-%d")
        },
        'completion_rate': (epics_df['Status'] == 'Done').mean() * 100,
        'avg_cycle_time': (pd.to_datetime(epics_df['Completed Date']) - pd.to_datetime(epics_df['Start Date'])).mean().days
    }
    
    # Analyze Maintenance Data
    maintenance_df = data['maintenance']
    report['data_sources']['maintenance'] = {
        'total_records': len(maintenance_df),
        'total_hours': maintenance_df['SumMaintenance_Hours'].sum(),
        'avg_tickets_per_week': maintenance_df['Total_Maintenance_Tickets_By_Week'].mean(),
        'maintenance_allocation': maintenance_df['Maintenance_Time_Allocation_Percentage'].mean()
    }
    
    # Analyze Machine Utilization Data
    utilization_df = data['utilization']
    report['data_sources']['utilization'] = {
        'total_records': len(utilization_df),
        'avg_utilization_rate': utilization_df['Machine_Utilization__'].mean() * 100,
        'total_available_hours': utilization_df['SumRuntime_duration__mins_'].sum() / 60,
        'total_utilized_hours': utilization_df['SumRuntime_duration__mins_'].sum() / 60
    }
    
    # Calculate Key Metrics
    report['metrics'] = {
        'overall_efficiency': report['data_sources']['utilization']['avg_utilization_rate'],
        'maintenance_impact': (report['data_sources']['maintenance']['total_hours'] / report['data_sources']['utilization']['total_available_hours']) * 100,
        'epic_completion_rate': report['data_sources']['jira_epics']['completion_rate']
    }
    
    # Generate Insights
    if report['metrics']['overall_efficiency'] < 80:
        report['insights'].append("Machine utilization efficiency is below target (80%). Consider optimizing scheduling.")
    
    if report['metrics']['maintenance_impact'] > 20:
        report['insights'].append("Maintenance activities are consuming more than 20% of available time. Review maintenance schedules.")
    
    if report['data_sources']['jira_epics']['avg_cycle_time'] > 14:
        report['insights'].append("Average epic cycle time exceeds two weeks. Review workflow bottlenecks.")
    
    # Generate Recommendations
    report['recommendations'] = [
        "Implement predictive maintenance scheduling to reduce unplanned downtime",
        "Review and optimize epic workflow processes to reduce cycle time",
        "Consider capacity planning based on current utilization patterns"
    ]
    
    return report

def generate_report_html(report):
    """
    Convert the report dictionary to HTML format
    """
    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            .section {{ margin-bottom: 30px; }}
            .metric {{ margin: 10px 0; padding: 10px; background-color: #f5f5f5; }}
            .insight {{ margin: 10px 0; padding: 10px; background-color: #e3f2fd; }}
            .recommendation {{ margin: 10px 0; padding: 10px; background-color: #e8f5e9; }}
            h1, h2, h3 {{ color: #1E88E5; }}
        </style>
    </head>
    <body>
        <h1>LTR Dashboard Comprehensive Report</h1>
        <p>Generated on: {report['timestamp']}</p>
        
        <div class="section">
            <h2>Data Sources Analysis</h2>
            <h3>JIRA Epics</h3>
            <div class="metric">
                <p>Total Records: {report['data_sources']['jira_epics']['total_records']}</p>
                <p>Date Range: {report['data_sources']['jira_epics']['date_range']['start']} to {report['data_sources']['jira_epics']['date_range']['end']}</p>
                <p>Completion Rate: {report['data_sources']['jira_epics']['completion_rate']:.2f}%</p>
                <p>Average Cycle Time: {report['data_sources']['jira_epics']['avg_cycle_time']:.1f} days</p>
            </div>
            
            <h3>Maintenance Data</h3>
            <div class="metric">
                <p>Total Records: {report['data_sources']['maintenance']['total_records']}</p>
                <p>Total Maintenance Hours: {report['data_sources']['maintenance']['total_hours']:.1f}</p>
                <p>Average Tickets per Week: {report['data_sources']['maintenance']['avg_tickets_per_week']:.1f}</p>
                <p>Maintenance Allocation: {report['data_sources']['maintenance']['maintenance_allocation']:.1f}%</p>
            </div>
            
            <h3>Machine Utilization</h3>
            <div class="metric">
                <p>Total Records: {report['data_sources']['utilization']['total_records']}</p>
                <p>Average Utilization Rate: {report['data_sources']['utilization']['avg_utilization_rate']:.1f}%</p>
                <p>Total Available Hours: {report['data_sources']['utilization']['total_available_hours']:.1f}</p>
                <p>Total Utilized Hours: {report['data_sources']['utilization']['total_utilized_hours']:.1f}</p>
            </div>
        </div>
        
        <div class="section">
            <h2>Key Metrics</h2>
            <div class="metric">
                <p>Overall Efficiency: {report['metrics']['overall_efficiency']:.1f}%</p>
                <p>Maintenance Impact: {report['metrics']['maintenance_impact']:.1f}%</p>
                <p>Epic Completion Rate: {report['metrics']['epic_completion_rate']:.1f}%</p>
            </div>
        </div>
        
        <div class="section">
            <h2>Insights</h2>
            {''.join([f'<div class="insight"><p>{insight}</p></div>' for insight in report['insights']])}
        </div>
        
        <div class="section">
            <h2>Recommendations</h2>
            {''.join([f'<div class="recommendation"><p>{recommendation}</p></div>' for recommendation in report['recommendations']])}
        </div>
    </body>
    </html>
    """
    return html

def report_generator_page():
    """Enhanced report generator page with comprehensive analysis"""
    st.title("📝 Report Generator")
    
    # Add custom CSS for the report type buttons
    st.markdown("""
    <style>
    .report-type-container {
        display: flex;
        gap: 1.5rem;
        margin: 1.5rem 0;
        padding: 1rem;
    }
    .report-type-card {
        flex: 1;
        padding: 2rem;
        border-radius: 15px;
        background-color: #f0f2f6;
        border: 2px solid #e6e9ef;
        transition: all 0.3s ease;
        cursor: pointer;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .report-type-card:hover {
        background-color: #e6e9ef;
        border-color: #d0d4db;
        transform: translateY(-3px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .report-type-card.selected {
        background-color: #1E88E5;
        color: white;
        border-color: #1E88E5;
        box-shadow: 0 6px 8px rgba(30, 136, 229, 0.3);
    }
    .report-type-icon {
        font-size: 2.5rem;
        margin-bottom: 1rem;
    }
    .report-type-title {
        font-size: 1.4rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
    }
    .report-type-desc {
        font-size: 1rem;
        color: #666;
        margin-bottom: 1.2rem;
        line-height: 1.4;
    }
    .selected .report-type-desc {
        color: #e6e9ef;
    }
    .report-type-features {
        text-align: left;
        margin-top: 1rem;
        padding: 0.8rem;
        background-color: rgba(255, 255, 255, 0.1);
        border-radius: 8px;
    }
    .report-type-features ul {
        margin: 0;
        padding-left: 1rem;
    }
    .report-type-features li {
        margin-bottom: 0.5rem;
        font-size: 0.9rem;
    }
    .report-type-button {
        margin-top: 1rem;
        padding: 0.8rem 1.5rem;
        border-radius: 8px;
        background-color: #1E88E5;
        color: white;
        border: none;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .report-type-button:hover {
        background-color: #1976D2;
        transform: translateY(-2px);
    }
    .selected .report-type-button {
        background-color: white;
        color: #1E88E5;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Initialize session state for selected report type if not exists
    if 'selected_report_type' not in st.session_state:
        st.session_state.selected_report_type = "Standard Report"
    
    # Report type selection
    st.markdown("### Select Report Type")
    
    # Create a container for the cards
    st.markdown('<div class="report-type-container">', unsafe_allow_html=True)
    
    # Standard Report Card
    col1, col2, col3 = st.columns(3)
    with col1:
        is_standard = st.session_state.selected_report_type == "Standard Report"
        st.markdown(f"""
        <div class="report-type-card {'selected' if is_standard else ''}">
            <div class="report-type-icon">📊</div>
            <div class="report-type-title">Standard Report</div>
            <div class="report-type-desc">Comprehensive analysis with detailed metrics and insights</div>
            <div class="report-type-features">
                <ul>
                    <li>📈 Performance metrics and KPIs</li>
                    <li>📊 Data source analysis</li>
                    <li>💡 Key insights and trends</li>
                    <li>🎯 Actionable recommendations</li>
                    <li>📋 Detailed breakdowns</li>
                </ul>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Select Standard Report", key="btn_standard", use_container_width=True):
            st.session_state.selected_report_type = "Standard Report"
            st.rerun()
    
    # RPA LTR Report Card
    with col2:
        is_rpa = st.session_state.selected_report_type == "RPA LTR Report"
        st.markdown(f"""
        <div class="report-type-card {'selected' if is_rpa else ''}">
            <div class="report-type-icon">🤖</div>
            <div class="report-type-title">RPA LTR Report</div>
            <div class="report-type-desc">RPA-focused updates with detailed roadmap and metrics</div>
            <div class="report-type-features">
                <ul>
                    <li>🤖 Bot performance metrics</li>
                    <li>📅 Project timeline updates</li>
                    <li>📈 ROI and efficiency analysis</li>
                    <li>🔧 Maintenance insights</li>
                    <li>🎯 Future roadmap items</li>
                </ul>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Select RPA Report", key="btn_rpa", use_container_width=True):
            st.session_state.selected_report_type = "RPA LTR Report"
            st.rerun()
    
    # Kaluza LTR Report Card
    with col3:
        is_kaluza = st.session_state.selected_report_type == "Kaluza LTR Report"
        st.markdown(f"""
        <div class="report-type-card {'selected' if is_kaluza else ''}">
            <div class="report-type-icon">📈</div>
            <div class="report-type-title">Kaluza LTR Report</div>
            <div class="report-type-desc">Kaluza-specific metrics with detailed analysis</div>
            <div class="report-type-features">
                <ul>
                    <li>📊 Kaluza-specific KPIs</li>
                    <li>📈 Performance trends</li>
                    <li>💡 Strategic insights</li>
                    <li>🎯 Implementation updates</li>
                    <li>📋 Detailed metrics</li>
                </ul>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Select Kaluza Report", key="btn_kaluza", use_container_width=True):
            st.session_state.selected_report_type = "Kaluza LTR Report"
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    data = load_data()
    if data is None:
        st.error("Failed to load data. Please check the data files.")
        return
    
    # Generate report based on selected type
    if st.session_state.selected_report_type == "Standard Report":
        report = generate_comprehensive_report(data)
        display_standard_report(report)
        download_standard_report(report)
    elif st.session_state.selected_report_type == "RPA LTR Report":
        report = generate_rpa_ltr_report(data)
        display_rpa_ltr_report(report)
        download_rpa_ltr_report(report)
    else:  # Kaluza LTR Report
        report = generate_kaluza_ltr_report(data)
        display_kaluza_ltr_report(report)
        download_kaluza_ltr_report(report)

def download_standard_report(report):
    """Download standard report in HTML and Excel formats"""
    st.markdown("### Download Report")
    col1, col2 = st.columns(2)
    
    with col1:
        # Download as HTML
        html_report = generate_report_html(report)
        st.download_button(
            label="Download HTML Report",
            data=html_report,
            file_name="ltr_standard_report.html",
            mime="text/html"
        )
    
    with col2:
        # Download as Excel
        excel_data = {
            'Data_Sources': pd.DataFrame(report['data_sources']),
            'Metrics': pd.DataFrame([report['metrics']]),
            'Insights': pd.DataFrame({'Insights': report['insights']}),
            'Recommendations': pd.DataFrame({'Recommendations': report['recommendations']})
        }
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for sheet_name, df in excel_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        excel_buffer.seek(0)
        st.download_button(
            label="Download Excel Report",
            data=excel_buffer,
            file_name="ltr_standard_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def download_rpa_ltr_report(report):
    """Download RPA LTR report in HTML and Excel formats"""
    st.markdown("### Download Report")
    col1, col2 = st.columns(2)
    
    with col1:
        # Download as HTML
        html_report = generate_rpa_ltr_html(report)
        st.download_button(
            label="Download HTML Report",
            data=html_report,
            file_name="ltr_rpa_report.html",
            mime="text/html"
        )
    
    with col2:
        # Download as Excel
        excel_data = {
            'Key_Updates': pd.DataFrame(report['key_updates']),
            'Exec_LTR_View': pd.DataFrame([report['exec_ltr_view']]),
            'Roadmap_Updates': pd.DataFrame(report['roadmap_updates']),
            'Detail_Notes': pd.DataFrame({'Notes': report['detail_notes']}),
            'Deep_Dives': pd.DataFrame(report['deep_dives']),
            'Files': pd.DataFrame({'Files': report['files']})
        }
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for sheet_name, df in excel_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        excel_buffer.seek(0)
        st.download_button(
            label="Download Excel Report",
            data=excel_buffer,
            file_name="ltr_rpa_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def download_kaluza_ltr_report(report):
    """Download Kaluza LTR report in HTML and Excel formats"""
    st.markdown("### Download Report")
    col1, col2 = st.columns(2)
    
    with col1:
        # Download as HTML
        html_report = generate_kaluza_ltr_html(report)
        st.download_button(
            label="Download HTML Report",
            data=html_report,
            file_name="ltr_kaluza_report.html",
            mime="text/html"
        )
    
    with col2:
        # Download as Excel
        excel_data = {
            'Line_Items': pd.DataFrame(report['line_items']),
            'Key_Updates': pd.DataFrame(report['key_updates']),
            'Exec_LTR_View': pd.DataFrame([report['exec_ltr_view']]),
            'Roadmap_Updates': pd.DataFrame(report['roadmap_updates']),
            'Detail_Notes': pd.DataFrame({'Notes': report['detail_notes']}),
            'Deep_Dives': pd.DataFrame(report['deep_dives']),
            'Files': pd.DataFrame({'Files': report['files']})
        }
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for sheet_name, df in excel_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        excel_buffer.seek(0)
        st.download_button(
            label="Download Excel Report",
            data=excel_buffer,
            file_name="ltr_kaluza_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def generate_rpa_ltr_html(report):
    """Generate HTML for RPA LTR report"""
    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            .section {{ margin-bottom: 30px; }}
            .metric {{ margin: 10px 0; padding: 10px; background-color: #f5f5f5; }}
            .update {{ margin: 10px 0; padding: 10px; background-color: #e3f2fd; }}
            .roadmap {{ margin: 10px 0; padding: 10px; background-color: #e8f5e9; }}
            h1, h2, h3 {{ color: #1E88E5; }}
        </style>
    </head>
    <body>
        <h1>RPA LTR Report</h1>
        <p>Generated on: {report['timestamp']}</p>
        
        <div class="section">
            <h2>Key Updates</h2>
            {''.join([f'''
            <div class="update">
                <h3>{update['title']}</h3>
                <p>Impact: {update['impact']}</p>
                <p>Date: {update['date']}</p>
            </div>
            ''' for update in report['key_updates']])}
        </div>
        
        <div class="section">
            <h2>Exec LTR View</h2>
            <div class="metric">
                <p>Completion Rate: {report['exec_ltr_view']['completion_rate']:.1f}%</p>
                <p>Average Utilization: {report['exec_ltr_view']['avg_utilization']:.1f}%</p>
                <p>Maintenance Impact: {report['exec_ltr_view']['maintenance_impact']:.1f} hours</p>
                <p>Success Rate: {report['exec_ltr_view']['success_rate']:.1f}%</p>
            </div>
        </div>
        
        <div class="section">
            <h2>Roadmap & Next Steps</h2>
            {''.join([f'''
            <div class="roadmap">
                <h3>{update['title']}</h3>
                <p>Due Date: {update['due_date']}</p>
                <p>Owner: {update['owner']}</p>
                <p>Impact: {update['impact']}</p>
            </div>
            ''' for update in report['roadmap_updates']])}
        </div>
        
        <div class="section">
            <h2>Detail Notes</h2>
            {''.join([f'<p>{note}</p>' for note in report['detail_notes']])}
        </div>
        
        <div class="section">
            <h2>Deep Dives</h2>
            {''.join([f'''
            <div class="update">
                <h3>{dive['title']}</h3>
                <p>Status: {dive['status']}</p>
                <p>Priority: {dive['priority']}</p>
            </div>
            ''' for dive in report['deep_dives']])}
        </div>
        
        <div class="section">
            <h2>Files</h2>
            {''.join([f'<p>{file}</p>' for file in report['files']])}
        </div>
    </body>
    </html>
    """
    return html

def generate_kaluza_ltr_html(report):
    """Generate HTML for Kaluza LTR report"""
    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            .section {{ margin-bottom: 30px; }}
            .metric {{ margin: 10px 0; padding: 10px; background-color: #f5f5f5; }}
            .line-item {{ margin: 10px 0; padding: 10px; background-color: #e3f2fd; }}
            .update {{ margin: 10px 0; padding: 10px; background-color: #e8f5e9; }}
            h1, h2, h3 {{ color: #1E88E5; }}
        </style>
    </head>
    <body>
        <h1>Kaluza LTR Report</h1>
        <p>Generated on: {report['timestamp']}</p>
        
        <div class="section">
            <h2>Line Items</h2>
            {''.join([f'''
            <div class="line-item">
                <h3>{item['description']}</h3>
                <p>Definition: {item['definition']}</p>
                <p>Owner: {item['owner']}</p>
                <p>Goal: {item['goal']}%</p>
                <p>Current Value: {item['current_value']:.1f}%</p>
            </div>
            ''' for item in report['line_items']])}
        </div>
        
        <div class="section">
            <h2>Key Updates</h2>
            {''.join([f'''
            <div class="update">
                <h3>{update['title']}</h3>
                <p>Impact: {update['impact']}</p>
                <p>Date: {update['date']}</p>
            </div>
            ''' for update in report['key_updates']])}
        </div>
        
        <div class="section">
            <h2>Exec LTR View</h2>
            <div class="metric">
                <p>Completion Rate: {report['exec_ltr_view']['completion_rate']:.1f}%</p>
                <p>Average Utilization: {report['exec_ltr_view']['avg_utilization']:.1f}%</p>
                <p>Maintenance Impact: {report['exec_ltr_view']['maintenance_impact']:.1f} hours</p>
                <p>Success Rate: {report['exec_ltr_view']['success_rate']:.1f}%</p>
            </div>
        </div>
        
        <div class="section">
            <h2>Roadmap & Next Steps</h2>
            {''.join([f'''
            <div class="update">
                <h3>{update['title']}</h3>
                <p>Due Date: {update['due_date']}</p>
                <p>Owner: {update['owner']}</p>
                <p>Impact: {update['impact']}</p>
            </div>
            ''' for update in report['roadmap_updates']])}
        </div>
        
        <div class="section">
            <h2>Detail Notes</h2>
            {''.join([f'<p>{note}</p>' for note in report['detail_notes']])}
        </div>
        
        <div class="section">
            <h2>Deep Dives</h2>
            {''.join([f'''
            <div class="update">
                <h3>{dive['title']}</h3>
                <p>Status: {dive['status']}</p>
                <p>Priority: {dive['priority']}</p>
            </div>
            ''' for dive in report['deep_dives']])}
        </div>
        
        <div class="section">
            <h2>Files</h2>
            {''.join([f'<p>{file}</p>' for file in report['files']])}
        </div>
    </body>
    </html>
    """
    return html

def display_standard_report(report):
    """Display standard report in Streamlit"""
    st.markdown("## 📊 Comprehensive Data Analysis Report")
    st.markdown(f"*Generated on: {report['timestamp']}*")
    
    # Data Sources Analysis
    with st.expander("Data Sources Analysis", expanded=True):
        st.markdown("### JIRA Epics")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Records", report['data_sources']['jira_epics']['total_records'])
            st.metric("Completion Rate", f"{report['data_sources']['jira_epics']['completion_rate']:.1f}%")
        with col2:
            st.metric("Date Range", f"{report['data_sources']['jira_epics']['date_range']['start']} to {report['data_sources']['jira_epics']['date_range']['end']}")
            st.metric("Average Cycle Time", f"{report['data_sources']['jira_epics']['avg_cycle_time']:.1f} days")
        
        st.markdown("### Maintenance Data")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Records", report['data_sources']['maintenance']['total_records'])
            st.metric("Total Maintenance Hours", f"{report['data_sources']['maintenance']['total_hours']:.1f}")
        with col2:
            st.metric("Average Tickets/Week", f"{report['data_sources']['maintenance']['avg_tickets_per_week']:.1f}")
            st.metric("Maintenance Allocation", f"{report['data_sources']['maintenance']['maintenance_allocation']:.1f}%")
        
        st.markdown("### Machine Utilization")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Records", report['data_sources']['utilization']['total_records'])
            st.metric("Average Utilization", f"{report['data_sources']['utilization']['avg_utilization_rate']:.1f}%")
        with col2:
            st.metric("Total Available Hours", f"{report['data_sources']['utilization']['total_available_hours']:.1f}")
            st.metric("Total Utilized Hours", f"{report['data_sources']['utilization']['total_utilized_hours']:.1f}")
    
    # Key Metrics
    with st.expander("Key Metrics", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Overall Efficiency", f"{report['metrics']['overall_efficiency']:.1f}%")
        with col2:
            st.metric("Maintenance Impact", f"{report['metrics']['maintenance_impact']:.1f}%")
        with col3:
            st.metric("Epic Completion Rate", f"{report['metrics']['epic_completion_rate']:.1f}%")
    
    # Insights
    with st.expander("Insights", expanded=True):
        for insight in report['insights']:
            st.info(insight)
    
    # Recommendations
    with st.expander("Recommendations", expanded=True):
        for recommendation in report['recommendations']:
            st.success(recommendation)

def display_rpa_ltr_report(report):
    """Display RPA LTR report in Streamlit"""
    st.markdown("## 📊 RPA LTR Report")
    st.markdown(f"*Generated on: {report['timestamp']}*")
    
    # Key Updates
    with st.expander("Key Updates", expanded=True):
        for update in report['key_updates']:
            st.markdown(f"### {update['title']}")
            st.markdown(f"- Impact: {update['impact']}")
            st.markdown(f"- Date: {update['date']}")
    
    # Exec LTR View
    with st.expander("Exec LTR View", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Completion Rate", f"{report['exec_ltr_view']['completion_rate']:.1f}%")
        with col2:
            st.metric("Avg Utilization", f"{report['exec_ltr_view']['avg_utilization']:.1f}%")
        with col3:
            st.metric("Maintenance Impact", f"{report['exec_ltr_view']['maintenance_impact']:.1f} hours")
        with col4:
            st.metric("Success Rate", f"{report['exec_ltr_view']['success_rate']:.1f}%")
    
    # Roadmap Updates
    with st.expander("Roadmap & Next Steps", expanded=True):
        for update in report['roadmap_updates']:
            st.markdown(f"### {update['title']}")
            st.markdown(f"- Due Date: {update['due_date']}")
            st.markdown(f"- Owner: {update['owner']}")
            st.markdown(f"- Impact: {update['impact']}")
    
    # Detail Notes
    with st.expander("Detail Notes", expanded=True):
        for note in report['detail_notes']:
            st.markdown(f"- {note}")
    
    # Deep Dives
    with st.expander("Deep Dives", expanded=True):
        for dive in report['deep_dives']:
            st.markdown(f"### {dive['title']}")
            st.markdown(f"- Status: {dive['status']}")
            st.markdown(f"- Priority: {dive['priority']}")
    
    # Files
    with st.expander("Files", expanded=True):
        for file in report['files']:
            st.markdown(f"- {file}")

def display_kaluza_ltr_report(report):
    """Display Kaluza LTR report in Streamlit"""
    st.markdown("## 📊 Kaluza LTR Report")
    st.markdown(f"*Generated on: {report['timestamp']}*")
    
    # Line Items
    with st.expander("Line Items", expanded=True):
        for item in report['line_items']:
            st.markdown(f"### {item['description']}")
            st.markdown(f"- Definition: {item['definition']}")
            st.markdown(f"- Owner: {item['owner']}")
            st.markdown(f"- Goal: {item['goal']}%")
            st.markdown(f"- Current Value: {item['current_value']:.1f}%")
    
    # Key Updates
    with st.expander("Key Updates", expanded=True):
        for update in report['key_updates']:
            st.markdown(f"### {update['title']}")
            st.markdown(f"- Impact: {update['impact']}")
            st.markdown(f"- Date: {update['date']}")
    
    # Exec LTR View
    with st.expander("Exec LTR View", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Completion Rate", f"{report['exec_ltr_view']['completion_rate']:.1f}%")
        with col2:
            st.metric("Avg Utilization", f"{report['exec_ltr_view']['avg_utilization']:.1f}%")
        with col3:
            st.metric("Maintenance Impact", f"{report['exec_ltr_view']['maintenance_impact']:.1f} hours")
        with col4:
            st.metric("Success Rate", f"{report['exec_ltr_view']['success_rate']:.1f}%")
    
    # Roadmap Updates
    with st.expander("Roadmap & Next Steps", expanded=True):
        for update in report['roadmap_updates']:
            st.markdown(f"### {update['title']}")
            st.markdown(f"- Due Date: {update['due_date']}")
            st.markdown(f"- Owner: {update['owner']}")
            st.markdown(f"- Impact: {update['impact']}")
    
    # Detail Notes
    with st.expander("Detail Notes", expanded=True):
        for note in report['detail_notes']:
            st.markdown(f"- {note}")
    
    # Deep Dives
    with st.expander("Deep Dives", expanded=True):
        for dive in report['deep_dives']:
            st.markdown(f"### {dive['title']}")
            st.markdown(f"- Status: {dive['status']}")
            st.markdown(f"- Priority: {dive['priority']}")
    
    # Files
    with st.expander("Files", expanded=True):
        for file in report['files']:
            st.markdown(f"- {file}")

def generate_rpa_ltr_report(data):
    """Generate RPA LTR report based on the RPA_LTR_Template.txt format"""
    report = {
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'key_updates': [],
        'exec_ltr_view': {},
        'roadmap_updates': [],
        'detail_notes': [],
        'deep_dives': [],
        'files': []
    }
    
    # Extract data for key updates
    epics_df = data['epics']
    maintenance_df = data['maintenance']
    utilization_df = data['utilization']
    
    # Key Updates (up to 3 items)
    if len(epics_df) > 0:
        recent_completed = epics_df[epics_df['Status'] == 'Done'].sort_values('Completed Date', ascending=False).head(3)
        for _, epic in recent_completed.iterrows():
            report['key_updates'].append({
                'title': epic['summary'],
                'impact': epic.get('Estimated Financial Impact', 'N/A'),
                'date': epic['Completed Date']
            })
    
    # Exec LTR View
    report['exec_ltr_view'] = {
        'completion_rate': (epics_df['Status'] == 'Done').mean() * 100,
        'avg_utilization': utilization_df['Machine_Utilization__'].mean() * 100,
        'maintenance_impact': maintenance_df['SumMaintenance_Hours'].sum(),
        'success_rate': (utilization_df['desktop_taskstatus'] == 'Succeeded').mean() * 100
    }
    
    # Roadmap Updates
    upcoming_epics = epics_df[epics_df['Status'] != 'Done'].sort_values('duedate').head(4)
    for _, epic in upcoming_epics.iterrows():
        report['roadmap_updates'].append({
            'title': epic['summary'],
            'due_date': epic['duedate'],
            'owner': epic['Assignee'],
            'impact': epic.get('Estimated Financial Impact', 'N/A')
        })
    
    # Detail Notes
    report['detail_notes'].extend([
        f"Total Epics: {len(epics_df)}",
        f"Total Maintenance Hours: {maintenance_df['SumMaintenance_Hours'].sum():.1f}",
        f"Average Machine Utilization: {utilization_df['Machine_Utilization__'].mean() * 100:.1f}%"
    ])
    
    # Deep Dives
    if len(epics_df) > 0:
        recent_epics = epics_df.sort_values('created', ascending=False).head(3)
        for _, epic in recent_epics.iterrows():
            report['deep_dives'].append({
                'title': epic['summary'],
                'status': epic['Status'],
                'priority': epic['priority']
            })
    
    # Files
    report['files'] = [
        "Week16_MainKPI.png",
        "Week16_WBR_Notes.txt",
        "SQL_SP_UOW_ALL_2025-04-21 10_01_46 AM.csv"
    ]
    
    return report

def generate_kaluza_ltr_report(data):
    """Generate Kaluza LTR report based on the LTR_Template_Kaluza.txt format"""
    report = {
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'line_items': [],
        'key_updates': [],
        'exec_ltr_view': {},
        'roadmap_updates': [],
        'detail_notes': [],
        'deep_dives': [],
        'files': []
    }
    
    # Extract data
    epics_df = data['epics']
    maintenance_df = data['maintenance']
    utilization_df = data['utilization']
    
    # Line Items (UDEs)
    report['line_items'] = [
        {
            'description': 'Epic Completion Rate',
            'definition': 'Percentage of epics completed on time',
            'owner': 'Project Manager',
            'goal': 95,
            'current_value': (epics_df['Status'] == 'Done').mean() * 100
        },
        {
            'description': 'Machine Utilization',
            'definition': 'Percentage of time machines are actively running tasks',
            'owner': 'Operations Manager',
            'goal': 80,
            'current_value': utilization_df['Machine_Utilization__'].mean() * 100
        },
        {
            'description': 'Maintenance Impact',
            'definition': 'Total hours spent on maintenance activities',
            'owner': 'Maintenance Lead',
            'goal': 20,
            'current_value': maintenance_df['SumMaintenance_Hours'].sum()
        }
    ]
    
    # Key Updates (up to 3 items)
    if len(epics_df) > 0:
        recent_completed = epics_df[epics_df['Status'] == 'Done'].sort_values('Completed Date', ascending=False).head(3)
        for _, epic in recent_completed.iterrows():
            report['key_updates'].append({
                'title': epic['summary'],
                'impact': epic.get('Estimated Financial Impact', 'N/A'),
                'date': epic['Completed Date']
            })
    
    # Exec LTR View
    report['exec_ltr_view'] = {
        'completion_rate': (epics_df['Status'] == 'Done').mean() * 100,
        'avg_utilization': utilization_df['Machine_Utilization__'].mean() * 100,
        'maintenance_impact': maintenance_df['SumMaintenance_Hours'].sum(),
        'success_rate': (utilization_df['desktop_taskstatus'] == 'Succeeded').mean() * 100
    }
    
    # Roadmap Updates
    upcoming_epics = epics_df[epics_df['Status'] != 'Done'].sort_values('duedate').head(3)
    for _, epic in upcoming_epics.iterrows():
        report['roadmap_updates'].append({
            'title': epic['summary'],
            'due_date': epic['duedate'],
            'owner': epic['Assignee'],
            'impact': epic.get('Estimated Financial Impact', 'N/A')
        })
    
    # Detail Notes
    report['detail_notes'].extend([
        f"Total Epics: {len(epics_df)}",
        f"Total Maintenance Hours: {maintenance_df['SumMaintenance_Hours'].sum():.1f}",
        f"Average Machine Utilization: {utilization_df['Machine_Utilization__'].mean() * 100:.1f}%"
    ])
    
    # Deep Dives
    if len(epics_df) > 0:
        recent_epics = epics_df.sort_values('created', ascending=False).head(5)
        for _, epic in recent_epics.iterrows():
            report['deep_dives'].append({
                'title': epic['summary'],
                'status': epic['Status'],
                'priority': epic['priority']
            })
    
    # Files
    report['files'] = [
        "Week16_MainKPI.png",
        "Week16_WBR_Notes.txt",
        "SQL_SP_UOW_ALL_2025-04-21 10_01_46 AM.csv"
    ]
    
    return report

# Main execution
def main():
    # Display sidebar and get selected page
    selected_page = sidebar_nav()
    
    # Display the selected page
    try:
        if selected_page == "Summary Overview":
            summary_overview_page()
        elif selected_page == "JIRA Epics Analysis":
            jira_epics_analysis_page()
        elif selected_page == "Maintenance Analysis":
            maintenance_analysis_page()
        elif selected_page == "Machine Utilization":
            machine_utilization_page()
        elif selected_page == "Report Generator":
            report_generator_page()
        else:
            # Default to summary page
            summary_overview_page()
    except Exception as e:
        st.error(f"An error occurred while loading the {selected_page} page: {e}")
        st.info("Try refreshing the page or selecting a different dashboard view.")

# Run the app
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"Application error: {e}")
        st.info("Please check the data files and try again.")
