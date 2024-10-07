import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from google.colab import files
import io

# Custom color schemes for pie and bar charts
colors_pie = ['#636EFA', '#EF553B', '#00CC96', '#AB63FA', '#FFA15A']
colors_bar = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', 
              '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']

def load_data():
    """Load data from uploaded Excel file, skipping the first 6 rows."""
    print("Please upload your Excel file.")
    uploaded = files.upload()
    file_name = next(iter(uploaded))
    return pd.read_excel(io.BytesIO(uploaded[file_name]), skiprows=6)

def process_data(df):
    """Process the dataframe to extract required information."""
    print("Processing data...")
    # Rename columns for clarity
    df.columns = ['Security', 'Weight', 'Rating', 'Country', 'Issuer', 'SecurityType']
    
    # Filter for corporate bonds (now checking for "Corporate Bonds")
    df = df[df['SecurityType'] == 'Corporate Bonds']
    print(f"Number of corporate bonds: {len(df)}")
    
    # Ensure Weight column is numeric
    df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce')
    
    # Calculate top holdings
    top_holdings = df.nlargest(10, 'Weight')[['Security', 'Weight']]
    
    # Calculate country exposure
    country_exposure = df.groupby('Country')['Weight'].sum().sort_values(ascending=False)
    
    # Calculate rating exposure
    rating_exposure = df.groupby('Rating')['Weight'].sum().sort_values(ascending=False)
    
    # Calculate top issuers
    top_issuers = df.groupby('Issuer')['Weight'].sum().nlargest(10).sort_values(ascending=True)
    
    return top_holdings, country_exposure, rating_exposure, top_issuers

def create_pie_chart(data, title, colors):
    """Create a pie chart using plotly with improved labels and custom colors."""
    return go.Pie(
        labels=data.index,
        values=data.values,
        name=title,
        textposition='inside',
        textinfo='percent+label',
        insidetextorientation='radial',
        showlegend=False,
        pull=[0.1 if i < 3 else 0 for i in range(len(data))],  # Pull out the top 3 slices
        hoverinfo='label+percent+value',
        hovertemplate='%{label}: %{value:.2f}%<extra></extra>',
        marker=dict(colors=colors)  # Apply custom colors
    )

def create_bar_chart(data, title, is_horizontal=True, colors=colors_bar):
    """Create a bar chart using plotly with custom colors and hover information."""
    return go.Bar(
        y=data.index if is_horizontal else data.values,
        x=data.values if is_horizontal else data.index,
        orientation='h' if is_horizontal else 'v',
        name=title,
        text=data.values,
        texttemplate='%{text:.2f}%',
        textposition='outside',
        hovertemplate='%{y}: %{x:.2f}%<extra></extra>' if is_horizontal else '%{x}: %{y:.2f}%<extra></extra>',
        marker=dict(color=colors)  # Apply custom colors
    )

def create_dashboard():
    """Create the main dashboard with aesthetic improvements."""
    try:
        print("Creating dashboard...")
        df = load_data()
        top_holdings, country_exposure, rating_exposure, top_issuers = process_data(df)
        
        # Create subplots with more vertical space and adjusted column widths
        fig = make_subplots(
            rows=3, cols=2, 
            specs=[[{'type':'domain'}, {'type':'domain'}],
                   [{'type':'xy', 'colspan': 2}, None],
                   [{'type':'xy', 'colspan': 2}, None]],
            subplot_titles=['Country Exposure', 'Rating Exposure', 
                            'Top 10 Issuers', 'Top 10 Holdings by Weight'],
            vertical_spacing=0.12,  # Increase spacing between subplots
            row_heights=[0.4, 0.3, 0.3], 
            column_widths=[0.6, 0.4]
        )
        
        # Add charts with custom colors
        fig.add_trace(create_pie_chart(country_exposure, 'Country Exposure', colors_pie), 1, 1)
        fig.add_trace(create_pie_chart(rating_exposure, 'Rating Exposure', colors_pie), 1, 2)
        fig.add_trace(create_bar_chart(top_issuers, 'Top 10 Issuers'), 2, 1)
        fig.add_trace(create_bar_chart(top_holdings.set_index('Security')['Weight'], 'Top 10 Holdings', False), 3, 1)
        
        # Update layout with fonts, background color, and spacing
        fig.update_layout(
            title_text="Fixed Income Fund Analysis Dashboard",
            height=1600,  # Increase height for better spacing
            width=1200,
            showlegend=False,
            font=dict(family="Arial", size=14, color="black"),  # Customize font
            plot_bgcolor='#f2f2f2',  # Light background for charts
            paper_bgcolor='#ffffff',  # White background for dashboard
            margin=dict(l=60, r=60, t=160, b=80),  # Adjust margins for breathing space
            title_x=0.5,  # Center the main title
            title_font=dict(size=22, color='darkblue'),  # Larger title with color
            title_pad=dict(t=50, b=20),  # Add padding above and below dashboard title
        )
        
        # Update axes and titles
        fig.update_yaxes(title_text="Issuer", row=2, col=1, automargin=True)
        fig.update_xaxes(title_text="Weight (%)", row=2, col=1, automargin=True)
        fig.update_xaxes(title_text="Security", row=3, col=1, automargin=True)
        fig.update_yaxes(title_text="Weight (%)", row=3, col=1, automargin=True)
        
        # Remove custom annotations for chart titles (use subplot titles for correct spacing)
        
        print("Dashboard created. Displaying plot...")
        # Show plot
        fig.show()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        print("Please make sure you've uploaded the correct Excel file and that it has the expected structure.")

# Run the dashboard
create_dashboard()
