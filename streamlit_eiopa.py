import streamlit as st
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import calendar
import os
import zipfile
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image


from Jonathan_automatische_visualisatie import download_and_extract_excel, create_and_save_plot

def get_download_link_for_date(url, target_year, target_month):
    """Get the download link for specific month and year"""
    try:
        last_day = calendar.monthrange(target_year, target_month)[1]
        filename = f"EIOPA_RFR_{target_year}{target_month:02d}{last_day}"
        
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        for link in soup.find_all('a', href=True):
            href = link.get('href')
            if ('download' in href.lower() and 
                f'_en?filename={filename}.zip' in href):
                return f"https://www.eiopa.europa.eu{href}" if not href.startswith('http') else href
        
        return None
        
    except Exception as e:
        st.error(f"Error retrieving page: {e}")
        return None

def main():
    st.title("EIOPA Rates Extraction Tool")
    
    st.sidebar.header("Select Dates")
    
    # Create a container for year-month selections
    date_selections = []
    
    # Allow adding multiple year-month combinations
    num_selections = st.sidebar.number_input("Number of date selections", min_value=1, value=1)
    
    for i in range(num_selections):
        st.sidebar.markdown(f"**Selection {i+1}**")
        
        # Year selection for each combination
        year = st.sidebar.selectbox(
            f"Select Year #{i+1}",
            range(2020, datetime.now().year + 1),
            key=f"year_{i}"
        )
        
        # Month selection for each combination
        month = st.sidebar.selectbox(
            f"Select Month #{i+1}",
            range(1, 13),
            format_func=lambda x: calendar.month_name[x],
            key=f"month_{i}"
        )
        
        date_selections.append((year, month))
    
    # Rate type selection
    rate_types = st.sidebar.multiselect(
        "Select Rate Type(s)",
        options=[
            'Euro_Rate_No_VA',
            'Euro_Rate_With_VA',
            'Euro_Rate_No_VA_Up',
            'Euro_Rate_No_VA_Down',
            'Euro_Rate_With_VA_Up',
            'Euro_Rate_With_VA_Down'
        ],
        default=['Euro_Rate_No_VA', 'Euro_Rate_With_VA']
    )
    
    if st.sidebar.button("Extract Rates"):
        url = "https://www.eiopa.europa.eu/tools-and-data/risk-free-interest-rate-term-structures_en"
        
        all_rates = []  # Store rates for all selected dates
        
        for year, month in date_selections:
            with st.spinner(f"Fetching EIOPA rates for {calendar.month_name[month]} {year}..."):
                zip_file_url = get_download_link_for_date(url, year, month)
                
                if zip_file_url:
                    os.makedirs('downloads', exist_ok=True)
                    excel_file = download_and_extract_excel(zip_file_url, 'downloads')
                    
                    if excel_file:
                        # Read rates
                        df_no_va = pd.read_excel(excel_file, sheet_name='RFR_spot_no_VA')
                        df_with_va = pd.read_excel(excel_file, sheet_name='RFR_spot_with_VA')
                        
                        euro_rates_no_va = df_no_va.iloc[9:, 2].reset_index(drop=True)
                        euro_rates_with_va = df_with_va.iloc[9:, 2].reset_index(drop=True)
                        
                        # Create DataFrame with selected rate types
                        rates_dict = {
                            'Maturity': range(1, len(euro_rates_no_va) + 1)
                        }
                        
                        if 'Euro_Rate_No_VA' in rate_types:
                            rates_dict[f'No_VA_{year}_{month:02d}'] = euro_rates_no_va
                        if 'Euro_Rate_With_VA' in rate_types:
                            rates_dict[f'With_VA_{year}_{month:02d}'] = euro_rates_with_va
                        
                        month_df = pd.DataFrame(rates_dict)
                        all_rates.append(month_df)
                else:
                    st.error(f"No rates found for {calendar.month_name[month]} {year}")
        
        if all_rates:
            # Merge all rates
            final_df = all_rates[0]
            for df in all_rates[1:]:
                final_df = pd.merge(final_df, df, on='Maturity')
            
            st.subheader("Euro Rates")
            st.dataframe(final_df)
            
            # Create and display plot
            fig = plt.figure(figsize=(12, 6))
            for col in final_df.columns:
                if col != 'Maturity':
                    if col.startswith('No_VA_'):
                        year = col[-6:-2]
                        month = int(col[-2:])
                        scenario = 'Base' if 'shock' not in col else col.split('_')[3]
                        label = f"{calendar.month_name[month]} {year} Without VA {scenario}"
                    elif col.startswith('With_VA_'):
                        year = col[-6:-2]
                        month = int(col[-2:])
                        scenario = 'Base' if 'shock' not in col else col.split('_')[3]
                        label = f"{calendar.month_name[month]} {year} With VA {scenario}"
                    else:
                        label = col
                    
                    plt.plot(final_df['Maturity'], final_df[col], 
                            label=label, linewidth=1)
                            
            plt.title('EIOPA Euro Rates by Maturity')
            plt.xlabel('Maturity')
            plt.ylabel('Rate (%)')
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            st.pyplot(fig)
            
            # Save and provide download link
            filename = 'EIOPA_Rates_' + '_'.join(f"{y}{m:02d}" for y, m in date_selections)
            output_path = os.path.join('downloads', f'{filename}.xlsx')
            final_df.to_excel(output_path, index=False)
            
            with open(output_path, 'rb') as f:
                st.download_button(
                    label="Download Excel file",
                    data=f.read(),
                    file_name=f'{filename}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

if __name__ == "__main__":
    main()