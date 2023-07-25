import pandas as pd
from datetime import datetime


def get_current_phase(df_phases, current_date):
    """Determine the current phase based on the current date using the phases DataFrame."""
    current_phase_row = df_phases[df_phases['Date'] == current_date]
    current_phase = current_phase_row['Phase'].values[0]
    return current_phase


def load_and_prune_data(filenames, df_phases):
    """Load data from CSV files and prune based on the current phase."""
    dataframes = []  # List to hold dataframes

    # Get the current date
    current_date = datetime.today().date()

    # Calculate the current phase
    current_phase = get_current_phase(df_phases, current_date)

    for filename in filenames:
        # Load the data from the CSV file
        df = pd.read_csv(filename)

        # Prune the DataFrame by retaining only the rows where the 'Phase' matches the current phase
        df_pruned = df[df['Phase'] == current_phase]

        # Append the pruned dataframe to the list
        dataframes.append(df_pruned)

    return dataframes


# Filenames of the CSV files
filenames = ['filename1.csv', 'filename2.csv', 'filename3.csv', 'filename4.csv']

# Load the phases data from the Excel file
df_phases = pd.read_excel('/mnt/data/phases.xlsx')

# Convert the 'Date' column from string to datetime
df_phases['Date'] = pd.to_datetime(df_phases['Date'])

# Load and prune the data
dataframes = load_and_prune_data(filenames, df_phases)


