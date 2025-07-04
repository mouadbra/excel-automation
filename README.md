# Excel Automation with Python and OpenPyXL

This project demonstrates how to automate Excel file operations using Python and the OpenPyXL library, with real-time data from APIs.

## Features

- Create and modify Excel files programmatically
- Fetch data from APIs (OpenWeatherMap in this example)
- Automated data updates with scheduling
- Data cleaning and maintenance
- Summary sheet generation
- Cross-platform compatibility (Windows, Mac, Linux)

## Setup

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Get a free API key from [OpenWeatherMap](https://openweathermap.org/api)

4. Create a `.env` file with your configuration:
   ```
   API_KEY=your_openweathermap_api_key
   EXCEL_FILE=weather_data.xlsx
   UPDATE_INTERVAL=3600
   ```

## Usage

### Basic Excel Operations
```bash
python examples/basic_excel_operations.py
```

### Weather Data Automation
```bash
python main.py
```

### Manual Data Update
```python
from main import update_weather_data
update_weather_data()
```

## Project Structure

```
excel-automation/
├── main.py                    # Main application
├── config.py                  # Configuration management
├── excel_manager.py           # Excel operations
├── api_client.py              # API client
├── examples/
│   └── basic_excel_operations.py
├── requirements.txt
├── .env
└── README.md

