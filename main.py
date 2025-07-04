from config import Config
from excel_manager import ExcelManager
from api_client import WeatherAPIClient
import schedule
import time
from datetime import datetime

def update_weather_data():
    """Main function to update weather data"""
    print(f"Starting weather data update at {datetime.now()}")
    
    # Initialize components
    api_client = WeatherAPIClient(Config.API_KEY)
    excel_manager = ExcelManager(Config.EXCEL_FILE)
    
    # Load or create workbook
    excel_manager.load_workbook()
    
    # Fetch weather data
    weather_data = api_client.get_multiple_cities_weather(Config.CITIES)
    
    if weather_data:
        # Add data to Excel
        excel_manager.add_weather_data(weather_data)
        
        # Clean old data (keep last 7 days)
        excel_manager.clear_old_data(keep_days=7)
        
        # Create/update summary sheet
        excel_manager.create_summary_sheet()
        
        # Save workbook
        excel_manager.save_workbook()
        
        print(f"Successfully updated weather data for {len(weather_data)} cities")
    else:
        print("No weather data retrieved")

def run_scheduler():
    """Run the scheduler"""
    # Schedule the job
    schedule.every(1).hours.do(update_weather_data)
    
    # Run once immediately
    update_weather_data()
    
    print("Weather data updater started. Press Ctrl+C to stop.")
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
    except KeyboardInterrupt:
        print("\nScheduler stopped.")

if __name__ == "__main__":
    run_scheduler()
