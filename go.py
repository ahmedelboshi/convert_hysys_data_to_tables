import pandas as pd
import requests
from bs4 import BeautifulSoup

# Define the base URL for the website
base_url = "https://www.brnenskeovzdusi.cz/brno-"

# Define the different station names
stations = ["detska-nemocnice", "arboretum", "lany", "svatoplukova", "vystaviste", "masna", "lisen", "uvoz", "turany"]

for station in stations:
    # Create the full URL for the station
    url = base_url + station

    # Make a request to the website
    response = requests.get(url)

    # Parse the HTML content
    soup = BeautifulSoup(response.content, "html.parser")

    # Find the table containing the data
    table = soup.find("table", {"class": "detail"})

    # Extract the data from the table
    data = []
    for row in table.find_all("tr"):
        columns = row.find_all("td")
        data.append([column.text for column in columns])

    print(data)
    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=["Time", "PM10", "PM2.5"])

    # Add a column for the station name
    df["Station"] = station

    # Save the data to a CSV file
    df.to_csv(f"{station}.csv", index=False)
