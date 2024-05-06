import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import os
import base64

def plot_altimetry(data):
    # Convert distance from meters to kilometers and round to 2 decimal places
    data['Distance_km'] = (data['Distance'] / 1000).round(2)

    # Create a new column for interval in kilometers
    data['Interval'] = (data['Distance_km'] // 1).astype(int) * 1

    # Group data by kilometer intervals and calculate average altitudes
    data_grouped = data.groupby('Interval').mean()
    data_grouped = data_grouped[['Altitude', 'Distance_km']]
    data_grouped.rename(columns={'Altitude': 'Altitude_m'}, inplace=True)
    data_grouped = data_grouped[['Distance_km', 'Altitude_m']].round(2)

    return data_grouped

def create_excel_file(data, filename):
    """
    Function to create an Excel file with two sheets: 'Données' and 'Graphique'.
    'Graphique' sheet contains a 3D Area chart representing the data from the 'Données' sheet.
    """
    excel_output_file = f'static/{filename}.xlsx'
    writer = pd.ExcelWriter(excel_output_file, engine='xlsxwriter')

    # Write 'Données' sheet
    data.to_excel(writer, sheet_name='Données', index=False)

    # Create 'Graphique' sheet
    workbook = writer.book
    worksheet = workbook.add_worksheet('Graphique')

    # Add chart to 'Graphique' sheet
    chart = workbook.add_chart({'type': 'area'})
    chart.add_series({
        'categories': ['Données', 1, 0, len(data), 0],
        'values': ['Données', 1, 1, len(data), 1],
        'fill': {'color': '#8B4513'},  # Brown color for the series
    })
    chart.set_size({'width': 720, 'height': 576})
    chart.set_title({'name': 'Profil altimétrique'})
    worksheet.insert_chart('A1', chart)

    writer.save()  # Save the ExcelWriter object
    writer.close()  # Close the ExcelWriter object

    return excel_output_file




def get_excel_download_link(file_path, filename):
    """
    Function to get the download link for the Excel file.
    """
    with open(file_path, 'rb') as f:
        file_content = f.read()
    base64_encoded = base64.b64encode(file_content).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{base64_encoded}" download="{filename}.xlsx">Télécharger le fichier Excel</a>'
    return href

# Create static folder if it doesn't exist
if not os.path.exists('static'):
    os.makedirs('static')

st.title('Analyse d\'altimétrie')

# Section to upload CSV file
st.header('Télécharger le fichier CSV')
uploaded_file = st.file_uploader("Télécharger un fichier CSV", type=['csv'])

if uploaded_file is not None:
    # Read CSV data
    data = pd.read_csv(uploaded_file)

    st.header('Données chargées:')
    st.write(data)

    # Display converted data
    st.header('Données converties:')
    converted_data = plot_altimetry(data)
    st.write(converted_data)

    # Plot altimetry profile
    st.header("Profil altimétrique")
    fig, ax = plt.subplots(figsize=(10, 6))

    # Plot the topographic profile
    ax.fill_between(converted_data['Distance_km'], converted_data['Altitude_m'], color='red', alpha=0.5)
    ax.plot(converted_data['Distance_km'], converted_data['Altitude_m'], color='black', label='Topography')

    # Customize the ticks on the y-axis to show altitude in meters
    ax.set_yticks(range(0, int(max(converted_data['Altitude_m'])) + 1, 100))

    # Label axes
    ax.set_xlabel('Distance (km)')
    ax.set_ylabel('Altitude (m)')
    ax.set_title('Profil altimétrique')
    ax.grid(True)
    ax.legend()

    # Display the plot
    st.pyplot(fig)

    # Create Excel file with two sheets: 'Données' and 'Graphique'
    excel_output_file = create_excel_file(converted_data, 'profil_altimetry')

    # Display the link to download the processed data as Excel
    st.markdown("### Télécharger les données converties:")
    st.markdown(get_excel_download_link(excel_output_file, 'profil_altimetry'), unsafe_allow_html=True)
