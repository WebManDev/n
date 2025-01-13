const fs = require('fs');
const XLSX = require('xlsx');
const axios = require('axios');

// File paths
const excelFilePath = './data.xlsx'; // Path to the Excel file
const geojsonFilePath = './border.geojson'; // Path to the GeoJSON file
const outputHtmlFile = './index.html'; // Output HTML file

// Google Maps API key
const apiKey = 'AIzaSyBwKPP8_WTiIoRTfWmjmRKiLYlcXnxEo_E'; // Replace with your API key

// Load data from the Excel file
const workbook = XLSX.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Function to geocode an address
async function geocodeAddress(address) {
    try {
        const response = await axios.get('https://maps.googleapis.com/maps/api/geocode/json', {
            params: {
                address: address,
                key: apiKey,
            },
        });
        if (response.data.results.length > 0) {
            const location = response.data.results[0].geometry.location;
            return { lat: location.lat, lng: location.lng };
        }
        return null;
    } catch (error) {
        console.error(`Error geocoding ${address}:`, error.message);
        return null;
    }
}

// Main function to process data and generate map
async function processAndGenerateMap() {
    let geocodedData = [];
    console.log('Geocoding addresses...');

    // Geocode all addresses
    for (const row of rawData) {
        const address = `${row.Address}, ${row.City}, ${row.State} ${row.Zip}`;
        const location = await geocodeAddress(address);

        if (location) {
            geocodedData.push({
                ...row,
                Latitude: location.lat,
                Longitude: location.lng,
            });
            console.log(`Geocoded: ${row.Address}`);
        }
    }

    console.log(`Successfully geocoded ${geocodedData.length} addresses.`);

    // Load GeoJSON for border
    const geojson = JSON.parse(fs.readFileSync(geojsonFilePath, 'utf-8'));

    // Generate Leaflet map HTML
    generateMap(geocodedData, geojson);
}

// Generate the Leaflet map
function generateMap(data, geojson) {
    const markers = data
        .map((row) => {
            const iconSize = Math.min(Math.max(row['Site EUI (kBtu/sq ft)'] / 5, 20), 80); // Scale icon size
            return `
                L.marker([${row.Latitude}, ${row.Longitude}], {
                    icon: L.icon({
                        iconUrl: 'smoke.png',
                        iconSize: [${iconSize}, ${iconSize}],
                    }),
                }).bindPopup(
                    '<b>Building Name:</b> ${row['Building Name']}<br>' +
                    '<b>Address:</b> ${row.Address}<br>' +
                    '<b>Site EUI:</b> ${row['Site EUI (kBtu/sq ft)'].toFixed(2)} kBtu/sq ft'
                ).addTo(map);
            `;
        })
        .join('\n');

    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Geocoded Map</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
</head>
<body>
    <div id="map" style="width: 100%; height: 100vh;"></div>
    <script>
        // Initialize map
        const map = L.map('map').setView([39.1, -77.2], 10);

        // Add base layer
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
        }).addTo(map);

        // Add GeoJSON overlay
        const geojson = ${JSON.stringify(geojson)};
        L.geoJSON(geojson, {
            style: {
                color: 'black',
                weight: 2,
                fillColor: 'blue',
                fillOpacity: 0.3,
            },
        }).addTo(map);

        // Add markers
        ${markers}
    </script>
</body>
</html>
    `;

    // Save HTML file
    fs.writeFileSync(outputHtmlFile, htmlContent);
    console.log(`Map generated and saved to ${outputHtmlFile}`);
}

// Run the process
processAndGenerateMap();
