import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

def create_output_directory(output_dir="processed_data"):
    """
    Create an output directory if it doesn't exist
    
    Args:
        output_dir (str): Name of the output directory
    
    Returns:
        str: Path to the output directory
    """
    try:
        # Get current directory
        current_dir = os.getcwd()
        # Create path for output directory
        output_path = os.path.join(current_dir, output_dir)
        
        # Create directory if it doesn't exist
        if not os.path.exists(output_path):
            os.makedirs(output_path)
            print(f"Created output directory: {output_path}")
        else:
            print(f"Output directory already exists: {output_path}")
            
        return output_path
    except Exception as e:
        print(f"Error creating output directory: {e}")
        return None

def load_csv_files():
    """
    Load the three CSV files from the current directory
    
    Returns:
        tuple: Three pandas DataFrames containing the data from each CSV file
    """
    try:
        # Get current directory
        current_dir = os.getcwd()
        
        # Define file paths
        epics_file = None
        maintenance_file = None
        utilization_file = None
        
        # Find the files based on partial names
        for file in os.listdir(current_dir):
            if file.endswith(".csv"):
                if "API_JIRA_Data_Epics" in file:
                    epics_file = os.path.join(current_dir, file)
                elif "API_JIRA_Data_Maintenance_Query" in file:
                    maintenance_file = os.path.join(current_dir, file)
                elif "Dataverse Desktop Machine Utilizations" in file:
                    utilization_file = os.path.join(current_dir, file)
        
        # Check if files were found
        if not (epics_file and maintenance_file and utilization_file):
            missing_files = []
            if not epics_file:
                missing_files.append("API_JIRA_Data_Epics")
            if not maintenance_file:
                missing_files.append("API_JIRA_Data_Maintenance_Query")
            if not utilization_file:
                missing_files.append("Dataverse Desktop Machine Utilizations")
            
            raise FileNotFoundError(f"Could not find these files: {', '.join(missing_files)}")
        
        # Load CSV files
        print(f"Loading file: {os.path.basename(epics_file)}")
        epics_df = pd.read_csv(epics_file)
        
        print(f"Loading file: {os.path.basename(maintenance_file)}")
        maintenance_df = pd.read_csv(maintenance_file)
        
        print(f"Loading file: {os.path.basename(utilization_file)}")
        utilization_df = pd.read_csv(utilization_file)
        
        return epics_df, maintenance_df, utilization_df
    
    except FileNotFoundError as e:
        print(f"File not found: {e}")
        return None, None, None
    except Exception as e:
        print(f"Error loading CSV files: {e}")
        return None, None, None

def clean_column_names(df):
    """
    Clean column names by removing prefixes and brackets
    
    Args:
        df (pandas.DataFrame): DataFrame with columns to clean
    
    Returns:
        pandas.DataFrame: DataFrame with cleaned column names
    """
    # Create a copy of the DataFrame to avoid modifying the original
    cleaned_df = df.copy()
    
    # Clean column names
    new_columns = {}
    for col in cleaned_df.columns:
        # Remove prefixes like "API_JIRA_Data_Epics[" and "]"
        new_col = col
        if '[' in col and ']' in col:
            new_col = col.split('[')[-1].replace(']', '')
        
        # Remove "Dataverse_Desktop Machines Utilizations[" prefix
        if 'Dataverse_Desktop Machines Utilizations[' in col:
            new_col = col.replace('Dataverse_Desktop Machines Utilizations[', '').replace(']', '')
            
        # Remove "API_JIRA_Data_Maintenance[" prefix
        if 'API_JIRA_Data_Maintenance[' in col:
            new_col = col.replace('API_JIRA_Data_Maintenance[', '').replace(']', '')
        
        # Create mapping from old column name to new column name
        new_columns[col] = new_col
    
    # Rename columns
    cleaned_df = cleaned_df.rename(columns=new_columns)
    
    return cleaned_df

def convert_date_columns(df, date_columns):
    """
    Convert date columns to proper datetime format
    
    Args:
        df (pandas.DataFrame): DataFrame with date columns
        date_columns (list): List of column names containing dates
    
    Returns:
        pandas.DataFrame: DataFrame with converted date columns
    """
    df_copy = df.copy()
    
    for col in date_columns:
        if col in df_copy.columns:
            try:
                # Convert column to datetime, handling errors
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce')
                print(f"Converted column '{col}' to datetime")
            except Exception as e:
                print(f"Error converting column '{col}' to datetime: {e}")
    
    return df_copy

def handle_missing_values(df, strategy='median'):
    """
    Handle missing values in the DataFrame
    
    Args:
        df (pandas.DataFrame): DataFrame with missing values
        strategy (str): Strategy for handling missing values ('median', 'mean', or 'drop')
    
    Returns:
        pandas.DataFrame: DataFrame with handled missing values
    """
    df_copy = df.copy()
    
    # Count missing values per column
    missing_count = df_copy.isnull().sum()
    print("\nMissing values per column:")
    for col in missing_count[missing_count > 0].index:
        print(f"  {col}: {missing_count[col]}")
    
    # Handle missing values based on strategy
    if strategy == 'drop':
        # Drop rows with missing values
        df_copy = df_copy.dropna()
        print(f"Dropped rows with missing values. Remaining rows: {len(df_copy)}")
    else:
        # Fill missing numeric values with median or mean
        numeric_cols = df_copy.select_dtypes(include=np.number).columns
        for col in numeric_cols:
            if df_copy[col].isnull().sum() > 0:
                if strategy == 'median':
                    df_copy[col] = df_copy[col].fillna(df_copy[col].median())
                    print(f"  Filled '{col}' missing values with median")
                elif strategy == 'mean':
                    df_copy[col] = df_copy[col].fillna(df_copy[col].mean())
                    print(f"  Filled '{col}' missing values with mean")
        
        # Fill missing categorical values with most frequent value or 'Unknown'
        categorical_cols = df_copy.select_dtypes(include=['object']).columns
        for col in categorical_cols:
            if df_copy[col].isnull().sum() > 0:
                df_copy[col] = df_copy[col].fillna('Unknown')
                print(f"  Filled '{col}' missing values with 'Unknown'")
    
    return df_copy

def create_week_column(df, date_column):
    """
    Create a week column based on a date column
    
    Args:
        df (pandas.DataFrame): DataFrame with a date column
        date_column (str): Name of the date column
    
    Returns:
        pandas.DataFrame: DataFrame with an added week column
    """
    df_copy = df.copy()
    
    if date_column in df_copy.columns:
        if pd.api.types.is_datetime64_dtype(df_copy[date_column]):
            # Extract year and week from the date column
            df_copy['Year'] = df_copy[date_column].dt.isocalendar().year
            df_copy['Week'] = df_copy[date_column].dt.isocalendar().week
            df_copy['YearWeek'] = df_copy['Year'].astype(str) + df_copy['Week'].astype(str).str.zfill(2)
            print(f"Created week columns based on '{date_column}'")
        else:
            print(f"Column '{date_column}' is not in datetime format")
    else:
        print(f"Column '{date_column}' not found in DataFrame")
    
    return df_copy

def calculate_correlation_metrics(maintenance_df, utilization_df):
    """
    Calculate correlation metrics between maintenance activities and bot utilization
    
    Args:
        maintenance_df (pandas.DataFrame): DataFrame with maintenance data
        utilization_df (pandas.DataFrame): DataFrame with utilization data
    
    Returns:
        pandas.DataFrame: DataFrame with correlation metrics
    """
    try:
        # Ensure we have YearWeek columns in both DataFrames
        if 'YearWeek' not in maintenance_df.columns or 'YearWeek' not in utilization_df.columns:
            print("YearWeek column missing in one or both DataFrames")
            return None
        
        # Group by YearWeek and calculate metrics
        maintenance_weekly = maintenance_df.groupby('YearWeek').agg({
            'SumMaintenance_Hours': 'sum',
            'Total_Maintenance_Tickets_By_Week': 'mean',
            'Maintenance_Time_Allocation_Percentage': 'mean'
        }).reset_index()
        
        utilization_weekly = utilization_df.groupby('YearWeek').agg({
            'SumRuntime_duration__mins_': 'sum',
            'Machine_Utilization__': 'mean',
            'Idle_Percentage__': 'mean',
            'Desktop_Flow_Success_Rate_Goal': 'mean',
            'Desktop_Run_Percent_Success': 'mean'
        }).reset_index()
        
        # Merge DataFrames on YearWeek
        correlation_df = pd.merge(maintenance_weekly, utilization_weekly, on='YearWeek', how='inner')
        
        # Calculate correlation matrix
        numeric_cols = correlation_df.select_dtypes(include=np.number).columns
        correlation_matrix = correlation_df[numeric_cols].corr()
        
        print("\nCorrelation Analysis Complete")
        return correlation_df, correlation_matrix
    
    except Exception as e:
        print(f"Error calculating correlation metrics: {e}")
        return None, None

def plot_correlation_matrix(correlation_matrix, output_path):
    """
    Plot the correlation matrix as a heatmap
    
    Args:
        correlation_matrix (pandas.DataFrame): Correlation matrix to plot
        output_path (str): Path to save the plot
    """
    try:
        plt.figure(figsize=(12, 10))
        sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', vmin=-1, vmax=1, fmt=".2f")
        plt.title('Correlation Matrix: Maintenance vs. Utilization Metrics')
        plt.tight_layout()
        
        # Save plot
        plot_path = os.path.join(output_path, 'correlation_matrix.png')
        plt.savefig(plot_path)
        print(f"Saved correlation matrix plot to: {plot_path}")
        
        # Close the plot to free memory
        plt.close()
    
    except Exception as e:
        print(f"Error plotting correlation matrix: {e}")

def main():
    """
    Main function to process the data
    """
    print("Starting data processing...")
    
    # Create output directory
    output_path = create_output_directory("processed_data")
    if not output_path:
        return
    
    # Load CSV files
    epics_df, maintenance_df, utilization_df = load_csv_files()
    if epics_df is None or maintenance_df is None or utilization_df is None:
        return
    
    print("\nProcessing JIRA Epics data...")
    # Clean JIRA Epics data
    epics_df = clean_column_names(epics_df)
    epics_date_columns = ['created', 'updated', 'duedate', 'Completed Date', 'Start Date']
    epics_df = convert_date_columns(epics_df, epics_date_columns)
    epics_df = handle_missing_values(epics_df)
    epics_df = create_week_column(epics_df, 'created')
    
    print("\nProcessing JIRA Maintenance data...")
    # Clean JIRA Maintenance data
    maintenance_df = clean_column_names(maintenance_df)
    maintenance_date_columns = ['created', 'updated', 'duedate', 'Updated Completed Date', 'Start Date', 'Completed Date']
    maintenance_df = convert_date_columns(maintenance_df, maintenance_date_columns)
    maintenance_df = handle_missing_values(maintenance_df)
    maintenance_df = create_week_column(maintenance_df, 'created')
    
    print("\nProcessing Machine Utilization data...")
    # Clean Machine Utilization data
    utilization_df = clean_column_names(utilization_df)
    utilization_date_columns = ['Created On', 'Start', 'End', 'Date', 'Start UTC', 'End UTC']
    utilization_df = convert_date_columns(utilization_df, utilization_date_columns)
    utilization_df = handle_missing_values(utilization_df)
    utilization_df = create_week_column(utilization_df, 'Created On')
    
    print("\nCalculating correlation metrics...")
    # Calculate correlation metrics
    correlation_df, correlation_matrix = calculate_correlation_metrics(maintenance_df, utilization_df)
    
    if correlation_df is not None and correlation_matrix is not None:
        # Plot correlation matrix
        plot_correlation_matrix(correlation_matrix, output_path)
        
        # Export processed data
        try:
            epics_df.to_csv(os.path.join(output_path, 'processed_epics.csv'), index=False)
            maintenance_df.to_csv(os.path.join(output_path, 'processed_maintenance.csv'), index=False)
            utilization_df.to_csv(os.path.join(output_path, 'processed_utilization.csv'), index=False)
            correlation_df.to_csv(os.path.join(output_path, 'correlation_data.csv'), index=False)
            
            print("\nExported processed data to:")
            print(f"  {os.path.join(output_path, 'processed_epics.csv')}")
            print(f"  {os.path.join(output_path, 'processed_maintenance.csv')}")
            print(f"  {os.path.join(output_path, 'processed_utilization.csv')}")
            print(f"  {os.path.join(output_path, 'correlation_data.csv')}")
        
        except Exception as e:
            print(f"Error exporting processed data: {e}")

if __name__ == "__main__":
    main()
