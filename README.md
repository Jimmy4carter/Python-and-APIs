# Python-and-APIs
Various program to showcase simple API connections

# OPENWEATHERMAP API
This Weather Report Application fetches weather data for a specific city from the OpenWeatherMap API and displays the weather forecast in a tabular format within a graphical user interface (GUI). Additionally, it allows users to export the weather data to an Excel file for further analysis or storage.

Features
Input: Users can input the name of the city for which they want to retrieve weather data.
Display: The application displays the weather forecast for the provided city, including details like temperature, humidity, wind speed, and description.
Export: Users can export the displayed weather data to an Excel file for offline use or analysis.
Requirements
Python 3.x
requests library (to make HTTP requests)
pandas library (for data manipulation and exporting to Excel)
OpenWeatherMap API key (Register for a free API key here)
Installation
Clone or download the repository.

Install the required libraries using pip:

Copy code
pip install requests pandas
Replace API_KEY in the code with your OpenWeatherMap API key.

Usage
Run the weather_report.py file.
Enter the name of the city for which you want to fetch the weather data.
Click on the "Fetch Weather" button to retrieve the weather forecast.
The weather forecast will be displayed in the GUI in a tabular format.
Click on the "Export to Excel" button to export the displayed data to an Excel file.
A message box will appear confirming the successful export, showing the name of the exported file.
Notes
Ensure a stable internet connection for fetching weather data from the API.
The program may encounter errors if the API key is invalid or if the city name is misspelled or not found in the database.
