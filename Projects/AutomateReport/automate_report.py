import requests

# Function to download a CSV file from a given URL and save it locally
def download_csv(url, local_filename):
    # Send a GET request to the specified URL
    response = requests.get(url)
    
    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Open the local file in write-binary mode and write the content to it
        with open(local_filename, 'wb') as file:
            file.write(response.content)
        print(f"CSV file downloaded and saved as {local_filename}")
    else:
        # Print an error message if the request was unsuccessful
        print(f"Failed to download file: {response.status_code}")

# Example usage to download the Titanic dataset
csv_url = 'https://raw.githubusercontent.com/datasciencedojo/datasets/master/titanic.csv'
local_csv_filename = 'titanic.csv'
download_csv(csv_url, local_csv_filename)


import pandas as pd

# Function to preprocess the Titanic dataset
def preprocess_data(df):
    # Drop rows where crucial information (Age, Fare, Embarked) is missing
    df = df.dropna(subset=['Age', 'Fare', 'Embarked'])
    
    # Fill missing values for 'Age' with the median age of the dataset
    df.loc[:, 'Age'] = df['Age'].fillna(df['Age'].median())
    
    # Fill missing values for 'Embarked' with the most common value (mode)
    df.loc[:, 'Embarked'] = df['Embarked'].fillna(df['Embarked'].mode()[0])
    
    # Filter out rows where 'Fare' is less than or equal to 0 or excessively high
    df = df[(df['Fare'] > 0) & (df['Fare'] < 500)]  # Adjust the range as needed

    return df

# Function to load and preprocess data, then perform some basic calculations
def extract_data(file_path):
    # Load the dataset from a CSV file
    df = pd.read_csv(file_path)
    
    # Preprocess the dataset (cleaning and filtering)
    df = preprocess_data(df)
    
    # Calculate various metrics and convert them to appropriate types
    total_passengers = int(len(df))  # Total number of passengers
    survived_passengers = int(df['Survived'].sum())  # Total number of survivors
    survival_rate = float(survived_passengers / total_passengers)  # Survival rate
    avg_age = float(df['Age'].mean())  # Average age of passengers
    max_fare = float(df['Fare'].max())  # Maximum fare paid
    avg_fare = float(df['Fare'].mean())  # Average fare paid
    male_passengers = int(len(df[df['Sex'] == 'male']))  # Number of male passengers
    female_passengers = int(len(df[df['Sex'] == 'female']))  # Number of female passengers
    avg_age_male = float(df[df['Sex'] == 'male']['Age'].mean())  # Average age of male passengers
    avg_age_female = float(df[df['Sex'] == 'female']['Age'].mean())  # Average age of female passengers
    
    # Return the calculated metrics as a dictionary
    return {
        'Total Passengers': total_passengers,
        'Survived Passengers': survived_passengers,
        'Survival Rate': survival_rate,
        'Average Age': avg_age,
        'Maximum Fare': max_fare,
        'Average Fare': avg_fare,
        'Male Passengers': male_passengers,
        'Female Passengers': female_passengers,
        'Average Age (Male)': avg_age_male,
        'Average Age (Female)': avg_age_female
    }

# Example usage to extract and calculate metrics from the Titanic dataset
local_csv_filename = 'titanic.csv'
data = extract_data(local_csv_filename)

# Print each metric on a new line
for key, value in data.items():
    print(f'{key}: {value}\n')


import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define the scope for Google Sheets and Google Drive API access
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load the credentials from the JSON key file and create a client for Google Sheets
creds = ServiceAccountCredentials.from_json_keyfile_name('automatereport-853e83e794f7.json', scope)
client = gspread.authorize(creds)

# Function to upload the extracted data to Google Sheets
def upload_to_google_sheets(data, sheet_name='Titanic Data Analysis'):
    # Try to open the Google Sheet by name; if it doesn't exist, create it
    try:
        sheet = client.open(sheet_name).sheet1
    except gspread.SpreadsheetNotFound:
        sheet = client.create(sheet_name).sheet1
    
    # Clear any existing content in the sheet
    sheet.clear()
    
    # Write headers for the data (Metric and Value)
    sheet.append_row(['Metric', 'Value'])
    
    # Write each metric and its value to the sheet
    for key, value in data.items():
        sheet.append_row([key, value])
    
    print("Data successfully uploaded to Google Sheets.")

# Example usage to upload the extracted metrics to Google Sheets
local_csv_filename = 'titanic.csv'
data = extract_data(local_csv_filename)
upload_to_google_sheets(data)
