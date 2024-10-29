import pandas as pd
import simplekml

def read_excel_data(file_path):
    df = pd.read_excel(file_path)
    print(df.columns)
    df['Lat A'] = pd.to_numeric(df['Lat A'], errors='coerce')
    df['Long A'] = pd.to_numeric(df['Long A'], errors='coerce')
    df['Lat B'] = pd.to_numeric(df['Lat B'], errors='coerce')
    df['Long B'] = pd.to_numeric(df['Long B'], errors='coerce')
    
    return df

def create_kml(df, output_file):
    kml = simplekml.Kml()

    for index, row in df.iterrows():
        try:
            link_name = row['Link Name Label']
            lat_a = float(row['Lat A'])
            lon_a = float(row['Long A'])
            lat_b = float(row['Lat B'])
            lon_b = float(row['Long B'])

            if pd.isna(lat_a) or pd.isna(lon_a) or pd.isna(lat_b) or pd.isna(lon_b):
                print(f"Skipping row {index} due to missing coordinates.")
                continue

            print(f"Plotting line from A({lat_a}, {lon_a}) to B({lat_b}, {lon_b}) for Link: {link_name}")

            linestring = kml.newlinestring(name=link_name)
            linestring.coords = [(lon_a, lat_a), (lon_b, lat_b)]
            linestring.style.linestyle.width = 2
            linestring.style.linestyle.color = simplekml.Color.white
            
            linestring.extrude = 5
            linestring.altitudemode = simplekml.AltitudeMode.relativetoground

            description = (
                f"Link Name: {link_name}<br>"
                f"A End NSSID: {row['A End NSSID']}<br>"
                f"LatLong A: ({lat_a}, {lon_a})<br>"
                f"A End Generic Name: {row['Link Name Label']}<br>"
                #f"A NE: {row['A NE']}<br>"
                f"B End NSSID: {row['B End NSSID']}<br>"
                f"LatLong B: ({lat_b}, {lon_b})<br>"
                f"B End Generic Name: {row['B End Generic Name']}<br>"
                #f"B NE: {row['B NE']}<br>"
                #f"Layer Rate: {row['Layer Rate']}<br>"
                #f"NE/VNE: {row['NE/VNE']}<br>"  
            )
            linestring.description = description

        except ValueError as e:
            print(f"Skipping row {index} due to invalid data: {e}")
            continue

    kml.save(output_file)

def excel_to_kml(excel_file, kml_output):
    df = read_excel_data(excel_file)
    create_kml(df, kml_output)

excel_file = r"GUJ_SECTION _EXCEL_29_OCT.xlsx"  
kml_output = "GUJ_KML_Excel_29_OCT.kml"  
excel_to_kml(excel_file, kml_output)
 