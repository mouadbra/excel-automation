import requests
from config import Config

class WeatherAPIClient:
    def __init__(self, api_key):
        self.api_key = api_key
        self.base_url = Config.WEATHER_API_URL
    
    def get_weather_data(self, city, country):
        """Get current weather data for a city"""
        try:
            params = {
                'q': f"{city},{country}",
                'appid': self.api_key,
                'units': 'metric'
            }
            
            response = requests.get(self.base_url, params=params)
            response.raise_for_status()
            
            data = response.json()
            
            return {
                'city': data['name'],
                'country': data['sys']['country'],
                'temperature': data['main']['temp'],
                'feels_like': data['main']['feels_like'],
                'humidity': data['main']['humidity'],
                'pressure': data['main']['pressure'],
                'weather': data['weather'][0]['main'],
                'description': data['weather'][0]['description'],
                'wind_speed': data['wind']['speed']
            }
        
        except requests.RequestException as e:
            print(f"Error fetching weather data for {city}: {e}")
            return None
        except KeyError as e:
            print(f"Error parsing weather data for {city}: {e}")
            return None
    
    def get_multiple_cities_weather(self, cities):
        """Get weather data for multiple cities"""
        weather_data = []
        
        for city_info in cities:
            data = self.get_weather_data(city_info['name'], city_info['country'])
            if data:
                weather_data.append(data)
        
        return weather_data