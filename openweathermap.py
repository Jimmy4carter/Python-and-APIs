import requests
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
import pandas as pd


API_KEY = "b8d470586f59072f5b7c5e482b609015"

def fetch_weather():
    city = entry_city.get()

    base_url = "http://api.openweathermap.org/data/2.5/forecast"
    params = {
        "q": city,
        "appid": API_KEY,
        "units": "metric",
    }

    try:
        response = requests.get(base_url, params=params)
        data = response.json()

        weather_data = []
        for forecast in data['list']:
            timestamp = datetime.fromtimestamp(forecast['dt'])
            weather_info = {
                'Date': timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                'Temperature (C)': forecast['main']['temp'],
                'Feels Like (C)': forecast['main']['feels_like'],
                'Min Temp (C)': forecast['main']['temp_min'],
                'Max Temp (C)': forecast['main']['temp_max'],
                'Pressure (hPa)': forecast['main']['pressure'],
                'Humidity (%)': forecast['main']['humidity'],
                'Description': forecast['weather'][0]['description'],
                'Wind Speed (m/s)': forecast['wind']['speed'],
            }
            weather_data.append(weather_info)

        df = pd.DataFrame(weather_data)
        display_text.delete(1.0, tk.END)
        display_text.insert(tk.END, df.to_string(index=False))

        file_name = f"{city}_weather.xlsx"
        df.to_excel(file_name, index=False)
        messagebox.showinfo("Export Successful", f"File exported successfully: {file_name}")

    except requests.RequestException as e:
        messagebox.showerror("Error", f"Failed to fetch data: {str(e)}")

# Create a modern styled theme
style = ttk.Style()
style.theme_use('clam')  # Choose a different theme if preferred

root = tk.Tk()
root.title("Weather Report")
root.geometry("500x400")

label = tk.Label(root, text="Enter City:", font=("Arial", 12))
label.pack()

entry_city = tk.Entry(root, font=("Arial", 12))
entry_city.pack()

fetch_button = tk.Button(root, text="Fetch Weather", command=fetch_weather, font=("Arial", 12), bg="#4CAF50", fg="white")
fetch_button.pack(pady=10)

display_frame = tk.Frame(root)
display_frame.pack()

display_text = tk.Text(display_frame, height=30, width=120, font=("Arial", 10))
display_text.pack()

export_button = tk.Button(root, text="Export to Excel", command=fetch_weather, font=("Arial", 12), bg="#2196F3", fg="white")
export_button.pack(pady=10)

root.mainloop()
