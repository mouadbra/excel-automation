import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    API_KEY = os.getenv('API_KEY', '')
    EXCEL_FILE = os.getenv('EXCEL_FILE', 'weather_data.xlsx')
    UPDATE_INTERVAL = int(os.getenv('UPDATE_INTERVAL', 3600))
    
    # OpenWeatherMap API (Free tier)
    WEATHER_API_URL = "https://api.openweathermap.org/data/2.5/weather"
    FORECAST_API_URL = "https://api.openweathermap.org/data/2.5/forecast"
    
    # Cities to monitor
    CITIES = [
        {"name": "Paris", "country": "FR"},
        {"name": "London", "country": "GB"},
        {"name": "New York", "country": "US"},
        {"name": "Tokyo", "country": "JP"},
        {"name": "Sydney", "country": "AU"}
    ]