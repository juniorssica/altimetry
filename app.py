import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import os
import base64

def plot_altimetry(data):
    # Convert distance from meters to kilometers
    data['Distance_km'] = data['Distance'] / 1000  

    # Create a new column for interval in kilometers
    data['Interval'] = (data['Distance_km'] // 1).astype(int) * 1

    # Group data by kilometer intervals and calculate average altitudes
    data_grouped = data.groupby('Interval').mean()
    data_grouped = data_grouped[['Altitude', 'Distance_km']]
    data_grouped.rename(columns={'Altitude': 'Altitude_m'}, inplace=True)
    data_grouped = data_grouped[['Distance_km', 'Altitude_m']]

    return data_grouped

def get_excel_download_link(data, filename):
    """
    Function to get the download link for the Excel file.
    """
    excel_output_file = f'static/{filename}.xlsx'
    data.to_excel(excel_output_file, index=False)
    with open(excel_output_file, 'rb') as f:
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

    # Display the link to download the processed data as Excel
    st.markdown("### Télécharger les données converties:")
    st.markdown(get_excel_download_link(converted_data, 'profil_altimetry'), unsafe_allow_html=True)
