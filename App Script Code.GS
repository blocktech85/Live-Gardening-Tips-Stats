// API key for WeatherAPI.com
const API_KEY = "af435ec887dc4b438e392552250305";
const ZIP_CODE = "28540";

// Plant information with planting date
const PLANTS = {
  "Strawberries": {
    waterNeeds: "1–1.5 inches/week; keep soil consistently moist, especially during fruiting.",
    sunNeeds: "Full sun (6–10+ hours/day).",
    fertilizer: "Balanced, low-nitrogen; side-dress with compost or 10-10-10 after planting and after harvest.",
    idealPH: "5.6–6.5",
    commonBugs: "Aphids, spider mites, slugs, root weevils",
    commonDiseases: "Red stele, root rot, gray mold, leaf spot",
    advancedCare: "Mulch with straw, rotate crops, avoid planting after nightshades, remove runners for bigger berries.",
    plantingDate: new Date("April 12, 2025")
  },
  "Fortex Beans": {
    waterNeeds: "1 inch/week; keep soil evenly moist, especially during flowering and pod set.",
    sunNeeds: "Full sun (at least 6–8 hours/day).",
    fertilizer: "Low nitrogen; use bean inoculant, side-dress with compost or 5-10-10 at planting.",
    idealPH: "5.5–6.8 (optimal 6.0–6.5)",
    commonBugs: "Aphids, bean beetles, cutworms",
    commonDiseases: "Rust, bacterial blight, mosaic virus",
    advancedCare: "Provide trellis, mulch for moisture, harvest frequently for continuous production.",
    plantingDate: new Date("April 12, 2025")
  },
  "Sweet Peppers": {
    waterNeeds: "1–2 inches/week; keep soil evenly moist, especially during fruiting.",
    sunNeeds: "Full sun (6–8+ hours/day).",
    fertilizer: "Balanced (5-10-10 or tomato fertilizer); side-dress with nitrogen at fruit set.",
    idealPH: "6.0–6.8",
    commonBugs: "Aphids, pepper maggots, flea beetles",
    commonDiseases: "Bacterial spot, anthracnose, mosaic virus",
    advancedCare: "Mulch, stake/cage plants, avoid overhead watering, prune lower leaves for airflow.",
    plantingDate: new Date("April 12, 2025")
  },
  "Banana Peppers": {
    waterNeeds: "1–2 inches/week; consistent moisture, avoid letting soil dry out.",
    sunNeeds: "Full sun (6–8+ hours/day).",
    fertilizer: "Balanced (5-10-10); side-dress with nitrogen at fruit set.",
    idealPH: "6.0–6.8",
    commonBugs: "Aphids, thrips, pepper maggots",
    commonDiseases: "Bacterial leaf spot, powdery mildew, mosaic virus",
    advancedCare: "Mulch, stake for heavy fruit, prune for airflow.",
    plantingDate: new Date("April 12, 2025")
  },
  "Marigolds": {
    waterNeeds: "Moderate; water when top inch of soil is dry. Avoid overwatering.",
    sunNeeds: "Full sun (6+ hours/day); tolerate partial shade.",
    fertilizer: "Light feeding with balanced fertilizer; avoid excess nitrogen.",
    idealPH: "6.0–7.0",
    commonBugs: "Aphids, spider mites, thrips",
    commonDiseases: "Powdery mildew, botrytis, damping-off",
    advancedCare: "Deadhead to prolong bloom, good for pest control in vegetable gardens.",
    plantingDate: new Date("April 12, 2025")
  },
  "Onions": {
    waterNeeds: "1 inch/week; keep soil moist, especially during bulb formation.",
    sunNeeds: "Full sun (6+ hours/day).",
    fertilizer: "High phosphorus and potassium; avoid excess nitrogen late in season.",
    idealPH: "6.0–7.0",
    commonBugs: "Onion maggot, thrips",
    commonDiseases: "Downy mildew, pink root, botrytis",
    advancedCare: "Mulch to retain moisture, rotate crops, avoid planting with beans or peas.",
    plantingDate: new Date("April 12, 2025")
  },
  "Morning Glory": {
    waterNeeds: "Low to moderate; drought-tolerant once established, water during dry spells.",
    sunNeeds: "Full sun (6+ hours/day).",
    fertilizer: "Minimal; too much fertilizer = more leaves, fewer flowers.",
    idealPH: "6.0–7.5",
    commonBugs: "Aphids, spider mites",
    commonDiseases: "Fungal leaf spot, rust",
    advancedCare: "Soak seeds overnight before planting, provide trellis, deadhead for more blooms.",
    plantingDate: new Date("April 12, 2025")
  },
  "Cantaloupe": {
    waterNeeds: "1–2 inches/week; keep soil moist until fruit sets, then reduce slightly for sweeter fruit.",
    sunNeeds: "Full sun (8+ hours/day).",
    fertilizer: "Balanced at planting; side-dress with nitrogen at vining.",
    idealPH: "6.0–6.5",
    commonBugs: "Cucumber beetles, aphids, squash bugs",
    commonDiseases: "Powdery mildew, fusarium wilt, anthracnose",
    advancedCare: "Mulch with black plastic for warmth, hand-pollinate if pollinators are scarce.",
    plantingDate: new Date("April 12, 2025")
  },
  "Cucumbers": {
    waterNeeds: "1–2 inches/week; keep soil evenly moist, especially during fruiting.",
    sunNeeds: "Full sun (6–8+ hours/day).",
    fertilizer: "Balanced at planting; side-dress with nitrogen at flowering.",
    idealPH: "6.0–6.8",
    commonBugs: "Cucumber beetles, aphids, spider mites",
    commonDiseases: "Downy mildew, powdery mildew, bacterial wilt",
    advancedCare: "Grow on trellis for airflow and cleaner fruit, mulch to retain moisture.",
    plantingDate: new Date("April 12, 2025")
  },
  "Tomatoes": {
    waterNeeds: "1–2 inches/week; deep, consistent watering, avoid wetting leaves.",
    sunNeeds: "Full sun (6–10+ hours/day).",
    fertilizer: "Balanced at planting; side-dress with calcium nitrate to prevent blossom end rot.",
    idealPH: "6.0–6.8",
    commonBugs: "Tomato hornworm, aphids, whiteflies",
    commonDiseases: "Early blight, late blight, fusarium wilt, blossom end rot",
    advancedCare: "Stake/cage for support, mulch, prune suckers for airflow, rotate crops.",
    plantingDate: new Date("April 12, 2025")
  }
};




/**
 * Main function to update garden information
 */
function updateGardenInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing sheets if they exist
  try {
    sheet.deleteSheet(sheet.getSheetByName("Weather Data"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  try {
    sheet.deleteSheet(sheet.getSheetByName("Plant Recommendations"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  // Create new sheets
  const weatherSheet = sheet.insertSheet("Weather Data");
  const recommendationsSheet = sheet.insertSheet("Plant Recommendations");
  
  // Get weather data
  const weatherData = fetchWeatherData();
  
  // Populate weather sheet
  populateWeatherSheet(weatherSheet, weatherData);
  
  // Generate and populate recommendations
  populateRecommendationsSheet(recommendationsSheet, weatherData);
  
  // Format sheets
  formatSheets(weatherSheet, recommendationsSheet);
  
  // Notify user
  SpreadsheetApp.getUi().alert('Garden information has been updated!');
}

/**
 * Fetches weather data from WeatherAPI.com
 */
function fetchWeatherData() {
  const url = `https://api.weatherapi.com/v1/forecast.json?key=${API_KEY}&q=${ZIP_CODE}&days=7&aqi=yes&alerts=yes`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log("Error fetching weather data: " + error);
    SpreadsheetApp.getUi().alert("Error fetching weather data: " + error);
    return null;
  }
}

/**
 * Populates the weather sheet with current and forecast data
 */
function populateWeatherSheet(sheet, data) {
  if (!data) return;
  
  // Current weather section
  sheet.getRange("A1:B1").merge().setValue("CURRENT WEATHER CONDITIONS FOR ZIP CODE " + ZIP_CODE);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  
  const current = data.current;
  const location = data.location;
  
  // Location info
  sheet.getRange("A2").setValue("Location:");
  sheet.getRange("B2").setValue(`${location.name}, ${location.region}, ${location.country}`);
  
  sheet.getRange("A3").setValue("Local Time:");
  sheet.getRange("B3").setValue(location.localtime);
  
  // Current weather details
  sheet.getRange("A4").setValue("Temperature:");
  sheet.getRange("B4").setValue(`${current.temp_f}°F (${current.temp_c}°C)`);
  
  sheet.getRange("A5").setValue("Feels Like:");
  sheet.getRange("B5").setValue(`${current.feelslike_f}°F (${current.feelslike_c}°C)`);
  
  sheet.getRange("A6").setValue("Condition:");
  sheet.getRange("B6").setValue(current.condition.text);
  
  sheet.getRange("A7").setValue("Wind:");
  sheet.getRange("B7").setValue(`${current.wind_mph} mph from ${current.wind_dir}`);
  
  sheet.getRange("A8").setValue("Humidity:");
  sheet.getRange("B8").setValue(`${current.humidity}%`);
  
  sheet.getRange("A9").setValue("Precipitation:");
  sheet.getRange("B9").setValue(`${current.precip_in} inches`);
  
  sheet.getRange("A10").setValue("UV Index:");
  sheet.getRange("B10").setValue(current.uv);
  
  // Forecast section
  sheet.getRange("A12:H12").merge().setValue("7-DAY WEATHER FORECAST");
  sheet.getRange("A12:H12").setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  
  // Headers for forecast table
  const forecastHeaders = ["Date", "Min Temp (°F)", "Max Temp (°F)", "Condition", "Precipitation", "Chance of Rain", "Humidity", "UV Index"];
  for (let i = 0; i < forecastHeaders.length; i++) {
    sheet.getRange(13, i + 1).setValue(forecastHeaders[i]);
  }
  
  // Populate forecast data
  const forecast = data.forecast.forecastday;
  for (let i = 0; i < forecast.length; i++) {
    const day = forecast[i].day;
    const date = forecast[i].date;
    
    sheet.getRange(14 + i, 1).setValue(date);
    sheet.getRange(14 + i, 2).setValue(day.mintemp_f);
    sheet.getRange(14 + i, 3).setValue(day.maxtemp_f);
    sheet.getRange(14 + i, 4).setValue(day.condition.text);
    sheet.getRange(14 + i, 5).setValue(`${day.totalprecip_in} inches`);
    sheet.getRange(14 + i, 6).setValue(`${day.daily_chance_of_rain}%`);
    sheet.getRange(14 + i, 7).setValue(`${day.avghumidity}%`);
    sheet.getRange(14 + i, 8).setValue(day.uv);
  }
  
  // Auto-size columns
  sheet.autoResizeColumns(1, 8);
}

/**
 * Generates plant-specific recommendations based on weather data
 */
function populateRecommendationsSheet(sheet, data) {
  if (!data) return;
  
  // Title
  sheet.getRange("A1:F1").merge().setValue("GARDEN RECOMMENDATIONS BASED ON WEATHER FORECAST");
  sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  
  // Headers
  sheet.getRange("A2").setValue("Plant");
  sheet.getRange("B2").setValue("Days Since Planting");
  sheet.getRange("C2").setValue("Current Status");
  sheet.getRange("D2").setValue("Today's Care");
  sheet.getRange("E2").setValue("Weekly Outlook");
  sheet.getRange("F2").setValue("Special Notes");
  
  // Style headers
  sheet.getRange("A2:F2").setFontWeight("bold").setBackground("#E8EAED");
  
  // Current weather and forecast data
  const current = data.current;
  const forecast = data.forecast.forecastday;
  
  // Calculate days since planting
  const today = new Date();
  
  // Row counter
  let row = 3;
  
  // Generate recommendations for each plant
  for (const [plantName, plantInfo] of Object.entries(PLANTS)) {
    const daysSincePlanting = Math.floor((today - plantInfo.plantingDate) / (1000 * 60 * 60 * 24));
    
    // Current status based on days since planting
    let status = getPlantStatus(plantName, daysSincePlanting);
    
    // Today's care recommendation based on current weather
    let todayCare = generateDailyCare(plantName, current, forecast[0].day);
    
    // Weekly outlook based on forecast
    let weeklyOutlook = generateWeeklyOutlook(plantName, forecast);
    
    // Special notes based on current conditions and forecast
    let specialNotes = generateSpecialNotes(plantName, current, forecast, daysSincePlanting);
    
    // Add to sheet
    sheet.getRange(row, 1).setValue(plantName);
    sheet.getRange(row, 2).setValue(daysSincePlanting);
    sheet.getRange(row, 3).setValue(status);
    sheet.getRange(row, 4).setValue(todayCare);
    sheet.getRange(row, 5).setValue(weeklyOutlook);
    sheet.getRange(row, 6).setValue(specialNotes);
    
    row++;
  }
  
  // Auto-size columns
  sheet.autoResizeColumns(1, 6);
}

/**
 * Formats the sheets for better readability
 */
function formatSheets(weatherSheet, recommendationsSheet) {
  // Format weather sheet
  weatherSheet.getRange("A1:H1").setFontSize(14);
  weatherSheet.getRange("A12:H12").setFontSize(14);
  weatherSheet.getRange("A13:H13").setFontWeight("bold").setBackground("#E8EAED");
  
  // Format recommendations sheet
  recommendationsSheet.getRange("A1:F1").setFontSize(14);
  
  // Set column widths for recommendations
  recommendationsSheet.setColumnWidth(1, 150); // Plant name
  recommendationsSheet.setColumnWidth(2, 100); // Days since planting
  recommendationsSheet.setColumnWidth(3, 200); // Current status
  recommendationsSheet.setColumnWidth(4, 300); // Today's care
  recommendationsSheet.setColumnWidth(5, 300); // Weekly outlook
  recommendationsSheet.setColumnWidth(6, 300); // Special notes
  
  // Wrap text in recommendation columns
  recommendationsSheet.getRange("C3:F100").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

/**
 * Determines the current status of a plant based on days since planting
 */
function getPlantStatus(plantName, daysSincePlanting) {
  // Early growth phase (0-14 days)
  if (daysSincePlanting < 14) {
    return "Early establishment phase. Focus on consistent moisture and protection from extreme conditions.";
  }
  
  // Mid growth phase (14-45 days)
  if (daysSincePlanting < 45) {
    switch(plantName) {
      case "Strawberries":
        return "Vegetative growth phase. Plants should be developing leaves and possibly early flowers.";
      case "Fortex Beans":
        return "Vegetative growth phase. Plants should be climbing and developing first flowers.";
      case "Sweet Peppers":
      case "Banana Peppers":
        return "Vegetative growth phase. Plants should be developing strong stems and leaves.";
      case "Marigolds":
        return "Early flowering phase. Plants should be developing buds and first flowers.";
      case "Onions":
        return "Leaf development phase. Focus on foliage growth before bulb formation.";
      case "Morning Glory":
        return "Vine development phase. Plants should be climbing and developing first buds.";
      case "Cantaloupe":
      case "Cucumbers":
        return "Vine development phase. Plants should be spreading and developing first flowers.";
      case "Tomatoes":
        return "Vegetative growth phase. Plants should be developing strong stems and leaves.";
      default:
        return "Mid-growth phase. Focus on supporting healthy development.";
    }
  }
  
  // Late growth/fruiting phase (45+ days)
  switch(plantName) {
    case "Strawberries":
      return "Fruiting phase. Plants should be producing berries.";
    case "Fortex Beans":
      return "Fruiting phase. Plants should be producing beans.";
    case "Sweet Peppers":
    case "Banana Peppers":
      return "Early fruiting phase. Plants should be developing first peppers.";
    case "Marigolds":
      return "Full bloom phase. Plants should be flowering abundantly.";
    case "Onions":
      return "Bulb formation phase. Focus on consistent moisture and reduced nitrogen.";
    case "Morning Glory":
      return "Flowering phase. Vines should be well-established with regular blooms.";
    case "Cantaloupe":
      return "Fruiting phase. Plants should be developing melons.";
    case "Cucumbers":
      return "Fruiting phase. Plants should be producing cucumbers.";
    case "Tomatoes":
      return "Fruiting phase. Plants should be developing and ripening tomatoes.";
    default:
      return "Mature phase. Focus on supporting production and preventing disease.";
  }
}

/**
 * Generates daily care recommendations based on current weather
 */
function generateDailyCare(plantName, current, forecast) {
  let care = [];
  
  // Temperature considerations
  if (current.temp_f > 90) {
    care.push("🌡️ High temperature alert! Provide extra water and consider shade cloth for protection.");
  } else if (current.temp_f < 50) {
    care.push("🌡️ Cool temperatures may slow growth. Monitor plants closely.");
  }
  
  // Precipitation considerations
  if (current.precip_in > 0.5) {
    care.push("💧 Significant rainfall today. Skip watering unless plants show signs of drought.");
  } else if (current.precip_in < 0.1 && forecast.daily_chance_of_rain < 30) {
    // Plant-specific watering needs
    switch(plantName) {
      case "Strawberries":
      case "Sweet Peppers":
      case "Banana Peppers":
      case "Tomatoes":
      case "Cucumbers":
      case "Cantaloupe":
        care.push("💧 Water deeply today as conditions are dry.");
        break;
      case "Fortex Beans":
      case "Onions":
        care.push("💧 Check soil moisture and water if the top inch is dry.");
        break;
      case "Marigolds":
      case "Morning Glory":
        care.push("💧 These plants are somewhat drought-tolerant, but check soil moisture if it's been several days without rain.");
        break;
    }
  }
  
  // Wind considerations
  if (current.wind_mph > 15) {
    switch(plantName) {
      case "Tomatoes":
      case "Sweet Peppers":
      case "Banana Peppers":
        care.push("💨 Windy conditions. Check stakes and supports.");
        break;
      case "Fortex Beans":
      case "Morning Glory":
      case "Cucumbers":
        care.push("💨 Windy conditions. Ensure trellises are secure.");
        break;
    }
  }
  
  // Humidity considerations
  if (current.humidity > 80) {
    switch(plantName) {
      case "Tomatoes":
      case "Sweet Peppers":
      case "Banana Peppers":
      case "Strawberries":
        care.push("💦 High humidity may increase disease risk. Ensure good air circulation and avoid wetting foliage.");
        break;
    }
  } else if (current.humidity < 40) {
    care.push("💦 Low humidity may increase water needs. Monitor soil moisture closely.");
  }
  
  // UV considerations
  if (current.uv > 8) {
    care.push("☀️ High UV levels. Newly transplanted plants may benefit from temporary shade.");
  }
  
  // If no specific recommendations, provide general care
  if (care.length === 0) {
    care.push("Conditions are favorable. Maintain regular care routine.");
  }
  
  return care.join("\n");
}

/**
 * Generates weekly outlook based on forecast
 */
function generateWeeklyOutlook(plantName, forecast) {
  // Calculate weekly metrics
  let totalRain = 0;
  let rainDays = 0;
  let highTemp = -Infinity;
  let lowTemp = Infinity;
  let avgHumidity = 0;
  
  for (let i = 0; i < forecast.length; i++) {
    const day = forecast[i].day;
    totalRain += day.totalprecip_in;
    if (day.daily_chance_of_rain > 50) rainDays++;
    highTemp = Math.max(highTemp, day.maxtemp_f);
    lowTemp = Math.min(lowTemp, day.mintemp_f);
    avgHumidity += day.avghumidity;
  }
  
  avgHumidity = avgHumidity / forecast.length;
  
  // Generate outlook
  let outlook = [];
  
  // Temperature outlook
  if (highTemp > 90) {
    outlook.push(`🌡️ Heat stress possible with temperatures reaching ${highTemp}°F this week.`);
    
    switch(plantName) {
      case "Lettuce":
      case "Spinach":
      case "Peas":
        outlook.push("Consider providing afternoon shade for heat-sensitive crops.");
        break;
      case "Tomatoes":
        outlook.push("Watch for blossom drop if temperatures stay above 90°F.");
        break;
    }
  }
  
  if (lowTemp < 45) {
    outlook.push(`🌡️ Cool nights with temperatures dropping to ${lowTemp}°F may slow growth.`);
    
    switch(plantName) {
      case "Tomatoes":
      case "Peppers":
      case "Cucumbers":
      case "Cantaloupe":
        outlook.push("Be prepared to protect warm-season crops if temperatures drop further.");
        break;
    }
  }
  
  // Precipitation outlook
  if (totalRain < 0.5) {
    outlook.push("📅 Dry week ahead. Plan for regular watering.");
    
    switch(plantName) {
      case "Strawberries":
        outlook.push("Maintain consistent soil moisture for berry development.");
        break;
      case "Tomatoes":
        outlook.push("Inconsistent watering can lead to blossom end rot and fruit splitting.");
        break;
      case "Cucumbers":
      case "Cantaloupe":
        outlook.push("Consistent moisture is critical during fruit development.");
        break;
    }
  } else if (totalRain > 2) {
    outlook.push(`📅 Wet week with ${totalRain.toFixed(1)} inches of rain expected. Monitor for disease.`);
    
    switch(plantName) {
      case "Tomatoes":
      case "Strawberries":
        outlook.push("Watch for signs of fungal diseases in wet conditions.");
        break;
      case "Onions":
        outlook.push("Ensure good drainage to prevent rot in wet conditions.");
        break;
    }
  } else {
    outlook.push(`📅 Moderate rainfall (${totalRain.toFixed(1)} inches) expected this week.`);
  }
  
  // Humidity outlook
  if (avgHumidity > 75) {
    outlook.push("💦 High humidity conditions may increase disease pressure.");
    
    switch(plantName) {
      case "Tomatoes":
      case "Cucumbers":
        outlook.push("Maintain good air circulation and consider preventative fungicide if disease-prone.");
        break;
    }
  }
  
  return outlook.join("\n");
}

/**
 * Generates special notes based on plant type, weather, and growth stage
 */
function generateSpecialNotes(plantName, current, forecast, daysSincePlanting) {
  let notes = [];
  
  // Growth stage specific notes
  if (daysSincePlanting < 14) {
    notes.push("🌱 Early establishment phase: Focus on consistent moisture and protection.");
  } else if (daysSincePlanting < 45) {
    notes.push("🌿 Growth phase: Support vegetative development with appropriate nutrients.");
  } else {
    notes.push("🌸 Production phase: Focus on supporting fruit/flower development.");
  }
  
  // Plant-specific notes
  switch(plantName) {
    case "Strawberries":
      if (daysSincePlanting > 45) {
        notes.push("🍓 Apply straw mulch to keep berries clean and prevent soil-borne diseases.");
        notes.push("🍓 Watch for birds and consider netting if they become a problem.");
      }
      break;
      
    case "Tomatoes":
      if (daysSincePlanting > 30) {
        notes.push("🍅 Prune suckers for indeterminate varieties to improve air circulation.");
        notes.push("🍅 Consider calcium supplementation if blossom end rot appears.");
      }
      break;
      
    case "Cucumbers":
    case "Cantaloupe":
      if (daysSincePlanting > 30) {
        notes.push("🥒 Train vines onto trellis to save space and improve fruit quality.");
        notes.push("🥒 Harvest cucumbers regularly to encourage continued production.");
      }
      break;
      
    case "Sweet Peppers":
    case "Banana Peppers":
      if (daysSincePlanting > 45) {
        notes.push("🌶️ Support plants with stakes or cages as fruit develops.");
      }
      break;
      
    case "Fortex Beans":
      if (daysSincePlanting > 40) {
        notes.push("🌱 Harvest beans regularly to encourage continued production.");
      }
      break;
      
    case "Onions":
      if (daysSincePlanting > 60) {
        notes.push("🧅 Reduce watering when tops begin to fall over to cure bulbs.");
      }
      break;
      
    case "Marigolds":
      notes.push("🌼 Deadhead spent blooms to encourage more flowers and prevent self-seeding.");
      notes.push("🌼 Marigolds help repel some garden pests - a natural companion plant.");
      break;
      
    case "Morning Glory":
      notes.push("🌸 Provide sturdy support for climbing. Can become aggressive, so monitor spread.");
      break;
  }
  
  // Weather pattern notes
  let rainInNext3Days = false;
  let highHeatInNext3Days = false;
  
  for (let i = 0; i < Math.min(3, forecast.length); i++) {
    if (forecast[i].day.daily_chance_of_rain > 60) rainInNext3Days = true;
    if (forecast[i].day.maxtemp_f > 90) highHeatInNext3Days = true;
  }
  
  if (rainInNext3Days) {
    notes.push("☔ Rain expected soon. Hold off on irrigation and consider delaying fertilizer application.");
  }
  
  if (highHeatInNext3Days) {
    notes.push("🔥 Heat stress possible. Ensure adequate water and consider temporary shade for sensitive plants.");
  }
  
  // Disease pressure notes based on conditions
  if (current.humidity > 75 && current.temp_f > 75) {
    notes.push("⚠️ Current conditions favor fungal diseases. Monitor plants closely and ensure good air circulation.");
  }
  
  return notes.join("\n");
}



/**
 * Calculates the water needs based on precipitation and temperature
 */
function calculateWaterNeeds(plantName, current, forecast) {
  // Base water needs in inches per week
  let baseWaterNeeds = 1;
  
  switch(plantName) {
    case "Strawberries":
    case "Fortex Beans":
    case "Onions":
      baseWaterNeeds = 1;
      break;
    case "Sweet Peppers":
    case "Banana Peppers":
    case "Tomatoes":
    case "Cucumbers":
    case "Cantaloupe":
      baseWaterNeeds = 1.5;
      break;
    case "Marigolds":
    case "Morning Glory":
      baseWaterNeeds = 0.75;
      break;
  }
  
  // Adjust for temperature
  if (current.temp_f > 85) {
    baseWaterNeeds *= 1.3; // Increase water needs by 30% in hot weather
  } else if (current.temp_f < 65) {
    baseWaterNeeds *= 0.8; // Decrease water needs by 20% in cool weather
  }
  
  // Adjust for humidity
  if (current.humidity < 40) {
    baseWaterNeeds *= 1.2; // Increase water needs by 20% in dry conditions
  } else if (current.humidity > 80) {
    baseWaterNeeds *= 0.9; // Decrease water needs by 10% in humid conditions
  }
  
  // Calculate expected rainfall for the next 7 days
  let expectedRainfall = 0;
  for (let i = 0; i < Math.min(7, forecast.length); i++) {
    expectedRainfall += forecast[i].day.totalprecip_in;
  }
  
  // Calculate supplemental water needed
  let supplementalWater = Math.max(0, baseWaterNeeds - expectedRainfall);
  
  return {
    baseNeeds: baseWaterNeeds.toFixed(2),
    expectedRain: expectedRainfall.toFixed(2),
    supplementalNeeds: supplementalWater.toFixed(2)
  };
}

/**
 * Creates a dashboard with summary information
 */
function createDashboard(weatherData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    sheet.deleteSheet(sheet.getSheetByName("Garden Dashboard"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  const dashboardSheet = sheet.insertSheet("Garden Dashboard");
  
  // Title
  dashboardSheet.getRange("A1:E1").merge().setValue("GARDEN DASHBOARD - ZIP CODE " + ZIP_CODE);
  dashboardSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#4285F4").setFontColor("white").setFontSize(16);
  
  // Current conditions summary
  dashboardSheet.getRange("A3:B3").merge().setValue("CURRENT CONDITIONS");
  dashboardSheet.getRange("A3:B3").setFontWeight("bold").setBackground("#E8EAED");
  
  const current = weatherData.current;
  
  dashboardSheet.getRange("A4").setValue("Temperature:");
  dashboardSheet.getRange("B4").setValue(`${current.temp_f}°F (${current.temp_c}°C)`);
  
  dashboardSheet.getRange("A5").setValue("Condition:");
  dashboardSheet.getRange("B5").setValue(current.condition.text);
  
  dashboardSheet.getRange("A6").setValue("Humidity:");
  dashboardSheet.getRange("B6").setValue(`${current.humidity}%`);
  
  dashboardSheet.getRange("A7").setValue("Wind:");
  dashboardSheet.getRange("B7").setValue(`${current.wind_mph} mph from ${current.wind_dir}`);
  
  dashboardSheet.getRange("A8").setValue("Precipitation:");
  dashboardSheet.getRange("B8").setValue(`${current.precip_in} inches`);
  
  // Weekly summary
  dashboardSheet.getRange("D3:E3").merge().setValue("WEEKLY OUTLOOK");
  dashboardSheet.getRange("D3:E3").setFontWeight("bold").setBackground("#E8EAED");
  
  const forecast = weatherData.forecast.forecastday;
  
  // Calculate weekly metrics
  let totalRain = 0;
  let rainDays = 0;
  let highTemp = -Infinity;
  let lowTemp = Infinity;
  
  for (let i = 0; i < forecast.length; i++) {
    const day = forecast[i].day;
    totalRain += day.totalprecip_in;
    if (day.daily_chance_of_rain > 50) rainDays++;
    highTemp = Math.max(highTemp, day.maxtemp_f);
    lowTemp = Math.min(lowTemp, day.mintemp_f);
  }
  
  dashboardSheet.getRange("D4").setValue("Total Expected Rain:");
  dashboardSheet.getRange("E4").setValue(`${totalRain.toFixed(2)} inches`);
  
  dashboardSheet.getRange("D5").setValue("Rain Days:");
  dashboardSheet.getRange("E5").setValue(rainDays);
  
  dashboardSheet.getRange("D6").setValue("High Temperature:");
  dashboardSheet.getRange("E6").setValue(`${highTemp}°F`);
  
  dashboardSheet.getRange("D7").setValue("Low Temperature:");
  dashboardSheet.getRange("E7").setValue(`${lowTemp}°F`);
  
  // Garden tasks section
  dashboardSheet.getRange("A10:E10").merge().setValue("PRIORITY GARDEN TASKS");
  dashboardSheet.getRange("A10:E10").setFontWeight("bold").setBackground("#E8EAED");
  
  let row = 11;
  
  // Generate priority tasks based on conditions
  let tasks = [];
  
  // Watering tasks
  if (totalRain < 0.5) {
    tasks.push("💧 Water deeply this week as conditions are dry.");
  } else if (totalRain > 2) {
    tasks.push("💧 Check for proper drainage in garden beds due to heavy rain forecast.");
  }
  
  // Temperature-related tasks
  if (highTemp > 90) {
    tasks.push("🌡️ Provide afternoon shade for heat-sensitive plants.");
    tasks.push("🌡️ Water in the early morning to help plants cope with heat.");
  }
  
  if (lowTemp < 45) {
    tasks.push("🌡️ Be prepared to protect tender plants from cool temperatures.");
  }
  
  // Disease prevention
  if (totalRain > 1.5 || current.humidity > 75) {
    tasks.push("🌿 Monitor for fungal diseases due to wet/humid conditions.");
    tasks.push("🌿 Ensure good air circulation around plants.");
  }
  
  // Seasonal tasks based on date
  const today = new Date();
  const month = today.getMonth(); // 0-11
  
  if (month >= 4 && month <= 8) { // May through September
    tasks.push("🌱 Monitor for insect pests as activity increases in warm weather.");
    tasks.push("🌱 Apply mulch to conserve soil moisture and suppress weeds.");
  }
  
  // Add tasks to dashboard
  for (let i = 0; i < tasks.length; i++) {
    dashboardSheet.getRange(row + i, 1, 1, 5).merge();
    dashboardSheet.getRange(row + i, 1).setValue(tasks[i]);
  }
  
  row += tasks.length + 2;
  
  // Plant-specific alerts
  dashboardSheet.getRange(row, 1, 1, 5).merge().setValue("PLANT-SPECIFIC ALERTS");
  dashboardSheet.getRange(row, 1, 1, 5).setFontWeight("bold").setBackground("#E8EAED");
  row++;
  
  let alerts = [];
  
  // Generate alerts based on plant needs and conditions
  if (current.temp_f > 85) {
    alerts.push("🍅 Tomatoes: Watch for blossom drop in high temperatures.");
    alerts.push("🥒 Cucumbers & Cantaloupe: Ensure consistent moisture in heat.");
  }
  
  if (current.humidity > 75) {
    alerts.push("🍓 Strawberries: Monitor for gray mold in humid conditions.");
    alerts.push("🍅 Tomatoes: Watch for early blight symptoms in humid weather.");
  }
  
  if (totalRain > 2) {
    alerts.push("🧅 Onions: Check for signs of rot with excessive moisture.");
  }
  
  // Add alerts to dashboard
  for (let i = 0; i < alerts.length; i++) {
    dashboardSheet.getRange(row + i, 1, 1, 5).merge();
    dashboardSheet.getRange(row + i, 1).setValue(alerts[i]);
  }
  
  // Format dashboard
  dashboardSheet.autoResizeColumns(1, 5);
}

/**
 * Main function to update all garden information
 */
function updateGardenInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing sheets if they exist
  try {
    sheet.deleteSheet(sheet.getSheetByName("Weather Data"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  try {
    sheet.deleteSheet(sheet.getSheetByName("Plant Recommendations"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  try {
    sheet.deleteSheet(sheet.getSheetByName("Garden Dashboard"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  // Create new sheets
  const weatherSheet = sheet.insertSheet("Weather Data");
  const recommendationsSheet = sheet.insertSheet("Plant Recommendations");
  
  // Get weather data
  const weatherData = fetchWeatherData();
  
  if (!weatherData) {
    SpreadsheetApp.getUi().alert("Failed to fetch weather data. Please try again later.");
    return;
  }
  
  // Populate weather sheet
  populateWeatherSheet(weatherSheet, weatherData);
  
  // Generate and populate recommendations
  populateRecommendationsSheet(recommendationsSheet, weatherData);
  
  // Create dashboard
  createDashboard(weatherData);
  
  // Format sheets
  formatSheets(weatherSheet, recommendationsSheet);
  
  // Set dashboard as active sheet
  sheet.setActiveSheet(sheet.getSheetByName("Garden Dashboard"));
  
  // Notify user
  SpreadsheetApp.getUi().alert('Garden information has been updated!');
}

/**
 * Detailed debugging function to help troubleshoot issues with the Garden Helper script
 * This function will test each component of the script and provide detailed logs
 */
function debugGardenHelper() {
  // Create a debug log sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    sheet.deleteSheet(sheet.getSheetByName("Debug Log"));
  } catch (e) {
    // Sheet doesn't exist, continue
  }
  
  const debugSheet = sheet.insertSheet("Debug Log");
  
  // Initialize log array and row counter
  let logs = [];
  let testsPassed = 0;
  let testsFailed = 0;
  
  // Helper function to add log entries
  function addLog(component, test, result, details) {
    logs.push({
      component: component,
      test: test,
      result: result,
      details: details,
      timestamp: new Date().toISOString()
    });
    
    if (result === "PASS") {
      testsPassed++;
    } else if (result === "FAIL") {
      testsFailed++;
    }
  }
  
  // 1. Test API Connection
  try {
    addLog("API Connection", "Testing connection to WeatherAPI.com", "RUNNING", "Attempting to fetch current weather data...");
    
    const testUrl = `https://api.weatherapi.com/v1/current.json?key=${API_KEY}&q=${ZIP_CODE}`;
    const response = UrlFetchApp.fetch(testUrl);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      addLog("API Connection", "Testing connection to WeatherAPI.com", "PASS", 
             `Successfully connected to API. Location: ${data.location.name}, ${data.location.region}. Current temp: ${data.current.temp_f}°F`);
    } else {
      addLog("API Connection", "Testing connection to WeatherAPI.com", "FAIL", 
             `API returned status code ${responseCode}`);
    }
  } catch (error) {
    addLog("API Connection", "Testing connection to WeatherAPI.com", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 2. Test Forecast API
  try {
    addLog("Forecast API", "Testing forecast data retrieval", "RUNNING", "Attempting to fetch 7-day forecast...");
    
    const testUrl = `https://api.weatherapi.com/v1/forecast.json?key=${API_KEY}&q=${ZIP_CODE}&days=7`;
    const response = UrlFetchApp.fetch(testUrl);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const forecastDays = data.forecast.forecastday.length;
      
      addLog("Forecast API", "Testing forecast data retrieval", "PASS", 
             `Successfully retrieved ${forecastDays} days of forecast data.`);
    } else {
      addLog("Forecast API", "Testing forecast data retrieval", "FAIL", 
             `API returned status code ${responseCode}`);
    }
  } catch (error) {
    addLog("Forecast API", "Testing forecast data retrieval", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 3. Test Plant Data Structure
  try {
    addLog("Plant Data", "Validating plant data structure", "RUNNING", "Checking plant data for completeness...");
    
    const missingFields = [];
    const requiredFields = ["waterNeeds", "sunNeeds", "fertilizer", "idealPH", "commonBugs", "commonDiseases", "advancedCare", "plantingDate"];
    
    for (const [plantName, plantInfo] of Object.entries(PLANTS)) {
      const plantMissingFields = [];
      
      for (const field of requiredFields) {
        if (!plantInfo.hasOwnProperty(field)) {
          plantMissingFields.push(field);
        }
      }
      
      if (plantMissingFields.length > 0) {
        missingFields.push(`${plantName} is missing: ${plantMissingFields.join(", ")}`);
      }
    }
    
    if (missingFields.length === 0) {
      addLog("Plant Data", "Validating plant data structure", "PASS", 
             `All ${Object.keys(PLANTS).length} plants have complete data.`);
    } else {
      addLog("Plant Data", "Validating plant data structure", "FAIL", 
             `Found missing fields: ${missingFields.join("; ")}`);
    }
  } catch (error) {
    addLog("Plant Data", "Validating plant data structure", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 4. Test Date Calculations
  try {
    addLog("Date Calculations", "Testing days since planting calculation", "RUNNING", "Calculating days since planting...");
    
    const today = new Date();
    const plantingDate = new Date("April 12, 2025");
    const daysSincePlanting = Math.floor((today - plantingDate) / (1000 * 60 * 60 * 24));
    
    if (isNaN(daysSincePlanting)) {
      addLog("Date Calculations", "Testing days since planting calculation", "FAIL", 
             "Invalid date calculation result");
    } else {
      addLog("Date Calculations", "Testing days since planting calculation", "PASS", 
             `Days since planting (April 12, 2025): ${daysSincePlanting}`);
    }
  } catch (error) {
    addLog("Date Calculations", "Testing days since planting calculation", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 5. Test Recommendation Functions
  try {
    addLog("Recommendation Functions", "Testing plant status function", "RUNNING", "Generating plant status...");
    
    // Mock data for testing
    const mockPlantName = "Tomatoes";
    const mockDaysSincePlanting = 30;
    
    const status = getPlantStatus(mockPlantName, mockDaysSincePlanting);
    
    if (status && status.length > 0) {
      addLog("Recommendation Functions", "Testing plant status function", "PASS", 
             `Successfully generated status: "${status}"`);
    } else {
      addLog("Recommendation Functions", "Testing plant status function", "FAIL", 
             "Failed to generate plant status");
    }
  } catch (error) {
    addLog("Recommendation Functions", "Testing plant status function", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 6. Test Weather Data Processing
  try {
    addLog("Weather Processing", "Testing weather data processing", "RUNNING", "Fetching and processing weather data...");
    
    const weatherData = fetchWeatherData();
    
    if (weatherData && weatherData.current && weatherData.forecast) {
      // Test water needs calculation
      const waterNeeds = calculateWaterNeeds("Tomatoes", weatherData.current, weatherData.forecast.forecastday);
      
      addLog("Weather Processing", "Testing weather data processing", "PASS", 
             `Successfully processed weather data. Water needs calculation: Base=${waterNeeds.baseNeeds}", Expected rain=${waterNeeds.expectedRain}", Supplemental=${waterNeeds.supplementalNeeds}"`);
    } else {
      addLog("Weather Processing", "Testing weather data processing", "FAIL", 
             "Failed to retrieve or process weather data");
    }
  } catch (error) {
    addLog("Weather Processing", "Testing weather data processing", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 7. Test Sheet Creation
  try {
    addLog("Sheet Creation", "Testing sheet creation capabilities", "RUNNING", "Creating test sheet...");
    
    const testSheetName = "Debug Test Sheet";
    
    try {
      sheet.deleteSheet(sheet.getSheetByName(testSheetName));
    } catch (e) {
      // Sheet doesn't exist, continue
    }
    
    const testSheet = sheet.insertSheet(testSheetName);
    testSheet.getRange("A1").setValue("Debug Test");
    
    addLog("Sheet Creation", "Testing sheet creation capabilities", "PASS", 
           "Successfully created and modified test sheet");
    
    // Clean up test sheet
    sheet.deleteSheet(testSheet);
  } catch (error) {
    addLog("Sheet Creation", "Testing sheet creation capabilities", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 8. Test Script Permissions
  try {
    addLog("Script Permissions", "Checking script permissions", "RUNNING", "Verifying script has necessary permissions...");
    
    const permissionsNeeded = [];
    
    // Test URL Fetch permission
    try {
      UrlFetchApp.fetch("https://www.google.com");
    } catch (e) {
      if (e.toString().includes("Authorization")) {
        permissionsNeeded.push("UrlFetchApp");
      }
    }
    
    // Test Spreadsheet permission
    try {
      SpreadsheetApp.getActiveSpreadsheet().getName();
    } catch (e) {
      if (e.toString().includes("Authorization")) {
        permissionsNeeded.push("SpreadsheetApp");
      }
    }
    
    if (permissionsNeeded.length === 0) {
      addLog("Script Permissions", "Checking script permissions", "PASS", 
             "Script has all necessary permissions");
    } else {
      addLog("Script Permissions", "Checking script permissions", "FAIL", 
             `Missing permissions: ${permissionsNeeded.join(", ")}`);
    }
  } catch (error) {
    addLog("Script Permissions", "Checking script permissions", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 9. Test API Key Validity
  try {
    addLog("API Key", "Validating API key", "RUNNING", "Checking if API key is valid...");
    
    const testUrl = `https://api.weatherapi.com/v1/current.json?key=${API_KEY}&q=London`;
    const response = UrlFetchApp.fetch(testUrl);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      addLog("API Key", "Validating API key", "PASS", 
             "API key is valid");
    } else if (responseText.includes("Invalid API key")) {
      addLog("API Key", "Validating API key", "FAIL", 
             "Invalid API key");
    } else {
      addLog("API Key", "Validating API key", "WARNING", 
             `Unexpected response: ${responseText}`);
    }
  } catch (error) {
    addLog("API Key", "Validating API key", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // 10. Test Memory Usage
  try {
    addLog("Performance", "Checking memory usage", "RUNNING", "Estimating memory usage...");
    
    // Create a large object to test memory limits
    const memoryTest = [];
    let memoryExceeded = false;
    
    try {
      for (let i = 0; i < 1000; i++) {
        memoryTest.push(new Array(1000).fill("test"));
      }
    } catch (e) {
      memoryExceeded = true;
      addLog("Performance", "Checking memory usage", "WARNING", 
             "Memory limits may be reached with large datasets");
    }
    
    if (!memoryExceeded) {
      addLog("Performance", "Checking memory usage", "PASS", 
             "Memory usage appears to be within acceptable limits");
    }
  } catch (error) {
    addLog("Performance", "Checking memory usage", "FAIL", 
           `Error: ${error.toString()}`);
  }
  
  // Output logs to debug sheet
  debugSheet.getRange("A1:D1").setValues([["Component", "Test", "Result", "Details"]]);
  debugSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#E8EAED");
  
  const logValues = logs.map(log => [log.component, log.test, log.result, log.details]);
  if (logValues.length > 0) {
    debugSheet.getRange(2, 1, logValues.length, 4).setValues(logValues);
  }
  
  // Color-code results
  for (let i = 0; i < logs.length; i++) {
    const resultCell = debugSheet.getRange(i + 2, 3);
    if (logs[i].result === "PASS") {
      resultCell.setBackground("#CEEAD6"); // Light green
    } else if (logs[i].result === "FAIL") {
      resultCell.setBackground("#F8CCC8"); // Light red
    } else if (logs[i].result === "WARNING") {
      resultCell.setBackground("#FEF7E0"); // Light yellow
    } else {
      resultCell.setBackground("#E8EAED"); // Light gray
    }
  }
  
  // Add summary
  const summaryRow = logs.length + 3;
  debugSheet.getRange(summaryRow, 1, 1, 4).merge();
  debugSheet.getRange(summaryRow, 1).setValue(`DEBUG SUMMARY: ${testsPassed} tests passed, ${testsFailed} tests failed, ${logs.length - testsPassed - testsFailed} informational`);
  debugSheet.getRange(summaryRow, 1).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  
  // Format debug sheet
  debugSheet.autoResizeColumns(1, 4);
  
  // Set debug sheet as active
  sheet.setActiveSheet(debugSheet);
  
  return {
    passed: testsPassed,
    failed: testsFailed,
    total: logs.length
  };
}

/**
 * Add debug option to menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Garden Helper')
      .addItem('Update Garden Information', 'updateGardenInfo')
      .addSeparator()
      .addItem('Run Diagnostics', 'debugGardenHelper')
      .addToUi();
}


