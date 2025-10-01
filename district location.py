import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time

def get_district_from_coords(lat, lon):
    """
    Get district from latitude and longitude using Nominatim
    """
    geolocator = Nominatim(user_agent="district_finder")
    
    try:
        # Reverse geocode the coordinates
        location = geolocator.reverse((lat, lon), exactly_one=True)
        
        if location:
            address = location.raw.get('address', {})
            # Try different possible keys for district
            district = (address.get('district') or 
                       address.get('county') or 
                       address.get('city') or 
                       address.get('town') or 
                       address.get('village'))
            return district
        return "Not found"
    except GeocoderTimedOut:
        return "Timeout Error"
    except Exception as e:
        return f"Error: {str(e)}"

def process_excel_file(file_path, lat_col, lon_col, output_file):
    """
    Process Excel file and add district column
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        print(f"Columns in Excel file: {df.columns.tolist()}")
        print(f"First few rows:\n{df.head()}")
        
        districts = []
        
        for index, row in df.iterrows():
            lat = row[lat_col]
            lon = row[lon_col]
            
            print(f"Processing row {index+1}: Lat={lat}, Lon={lon}")
            
            # Get district
            district = get_district_from_coords(lat, lon)
            districts.append(district)
            
            # Add delay to respect API limits
            time.sleep(1)
            
            print(f"Processed {index+1}/{len(df)}: {district}")
        
        # Add districts to dataframe
        df['district'] = districts
        
        # Save to new Excel file
        df.to_excel(output_file, index=False)
        print(f"Results saved to {output_file}")
        
    except Exception as e:
        print(f"Error reading file: {e}")

# First, let's check what columns are in your Excel file
def check_columns(file_path):
    """
    Check the column names in the Excel file
    """
    try:
        df = pd.read_excel(file_path)
        print("Available columns in your Excel file:")
        for col in df.columns:
            print(f"  - {col}")
        print(f"\nFirst 3 rows:")
        print(df.head(3))
        return df.columns.tolist()
    except Exception as e:
        print(f"Error: {e}")
        return None

# Usage
if __name__ == "__main__":
    file_path = r""
    output_file = r""
    
    # First, check the columns in your Excel file
    print("Checking your Excel file structure...")
    columns = check_columns(file_path)
    
    if columns:
        # You need to replace "Latitude" and "Longitude" with your actual column names
        # Based on what you see in the output above
        lat_column = "Latitude"  # Change this to your actual latitude column name
        lon_column = "Longitude" # Change this to your actual longitude column name
        
        # Check if the columns exist
        if lat_column in columns and lon_column in columns:
            print(f"\nUsing columns: {lat_column} and {lon_column}")
            process_excel_file(file_path, lat_column, lon_column, output_file)
        else:
            print(f"\nError: Columns '{lat_column}' and/or '{lon_column}' not found in Excel file.")
            print("Please update the 'lat_column' and 'lon_column' variables with the correct column names from the list above.")