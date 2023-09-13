import pandas as pd
from geopy.geocoders import Nominatim


geolocator = Nominatim(user_agent="city_locator")


def get_city_location(city_name):
    # Initialize the geocoder

    try:
        # Geocode the city
        location = geolocator.geocode(city_name)

        if location:
            return {
                "address": location.address,
                "latitude": location.latitude,
                "longitude": location.longitude
            }
        else:
            return None

    except Exception as e:
        return f"Error: {str(e)}"


def getExcel(path):
    input_excel_file = path
    print("🔥 Loading Excel file...")
    # Load the Excel file into a DataFrame
    return pd.read_excel(input_excel_file)


def saveExcel(path, df):
    output_excel_file = path
    # Save the modified DataFrame back to an Excel file
    df.to_excel(output_excel_file, index=False)
    print("🔥 Excel file saved successfully")


def addCoordinates(df):

    # Create empty lists to store latitude and longitude
    latitudes = []
    longitudes = []

    for index, row in df.iterrows():
        city_name = row["city"]
        location = get_city_location(city_name)
        if location == None:
            print(f"🍅 {index + 1} 🍅 {city_name} location: None")
            latitudes.append("None 🍅 XXX")
            longitudes.append("None 🍅 XXX")
        else:
            try:

                print(
                    f"🍅 {index + 1} 🍅 {city_name} location: {location} longitude: {location['longitude']} latitude: {location['latitude']}")

                latitudes.append(location["latitude"])
                longitudes.append(location["longitude"])

            except Exception as e:
                print(f"🍅 {index + 1} 🍅 {city_name} Error: {str(e)}")
                latitudes.append("None 🍅 XXX")
                longitudes.append("None 🍅 XXX")

    df["latitude"] = latitudes
    df["longitude"] = longitudes

    return df


def main():
    inFile = "in.xlsx"
    outFile = "out.xlsx"
    df = getExcel(inFile)
    df = addCoordinates(df)
    saveExcel(outFile, df)


main()
