**Garden Helper Google Sheets Add-on Overview**
Garden Helper is a Google Apps Script add-on for Google Sheets that helps home gardeners track their plants and get actionable, weather-aware care recommendations. By integrating real-time weather data through weatherAPI.com API calls with detailed plant profiles, it generates a personalized garden dashboard inside your spreadsheet, making garden management easier and smarter.

**Key Features**
Live Weather Integration:
Fetches current conditions and a 7-day forecast for your location (using WeatherAPI.com).

**Plant Database:**
Contains care profiles for common garden plants (watering, sun, fertilizer, pH, pests, diseases, advanced tips, and planting date).

**Automated Recommendations:**
For each plant, calculates days since planting and generates:

**Current growth status**
- Today's care advice (based on weather)
- Weekly outlook
- Special notes (e.g., pest/disease risk, weather warnings)

**Spreadsheet Automation:**
- With one click, clears and refreshes two sheets:
- Weather Data: Current conditions and 7-day forecast

**Plant Recommendations: **
- Up-to-date care advice for all your plants


**How It Works**

**Menu Integration:**
- On opening your Google Sheet, a custom menu ("Garden Helper") appears.
- Select "Update Garden Information" to run the add-on.

**Weather Fetching:**
- The script queries the WeatherAPI for your ZIP code, retrieving current weather and a 7-day forecast.

**Sheet Management:**
- Deletes any old "Weather Data" or "Plant Recommendations" sheets.
- Creates fresh sheets and populates them with new data.

**Weather Sheet:**
- Displays location, current weather, and a color-coded 7-day forecast table.

**Plant Recommendations Sheet:**
- For each plant in your garden, calculates days since planting.
- Determines the plant's current growth phase.
- Generates daily care recommendations based on current and forecasted weather.
- Provides a weekly care outlook and special notes for potential issues.

**Formatting:**
- Auto-sizes columns and applies formatting for readability.
- Uses color highlights for headers and wraps text for long recommendations.
