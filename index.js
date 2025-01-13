const fs = require('fs');
const XLSX = require('xlsx');
const axios = require('axios');

// File paths
const excelFilePath = '2023 Disclosed Benchmarking Data for All Covered Buildings.xlsx'; // Path to the Excel file
const geojsonFilePath = './h.geojson'; // Path to the GeoJSON file
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
        console.warn(`No results found for address: ${address}`);
        return null;
    } catch (error) {
        console.error(
            `Error geocoding ${address}:`,
            error.response ? error.response.data : error.message
        );
        return null;
    }
}

// Retry mechanism for geocoding
async function geocodeAddressWithRetry(address, retries = 3) {
    for (let i = 0; i < retries; i++) {
        const result = await geocodeAddress(address);
        if (result) return result;
        console.warn(`Retrying (${i + 1}/${retries}) for address: ${address}`);
    }
    console.error(`Failed to geocode address after ${retries} attempts: ${address}`);
    return null;
}

// Main function to process data and generate map
async function processAndGenerateMap() {
    let geocodedData = [];
    console.log('Geocoding all addresses...');

    for (const row of rawData) {
        if (row.Address && row.City && row.State) {
            const address = `${row.Address}, ${row.City}, ${row.State} ${row.Zip || ''}`;
            console.log(`Geocoding address: ${address}`);
            const location = await geocodeAddressWithRetry(address);

            if (location) {
                geocodedData.push({
                    ...row,
                    Latitude: location.lat,
                    Longitude: location.lng,
                });
                console.log(`Geocoded: ${row.Address}`);
            } else {
                console.error(`Failed to geocode: ${row.Address}`);
            }
        } else {
            console.warn(`Skipping row due to missing fields: ${JSON.stringify(row)}`);
        }
    }

    console.log(`Successfully geocoded ${geocodedData.length} addresses.`);

    // Load GeoJSON for border
    const geojson = JSON.parse(fs.readFileSync(geojsonFilePath, 'utf-8'));

    // Generate Leaflet map HTML with marker clustering
    generateMap(geocodedData, geojson);
}

// Generate the Leaflet map with marker clustering
function generateMap(data, geojson) {
    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Geocoded Map with Smoke Icon (Lazy Loading)</title>
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

        // Custom smoke icon
        const smokeIcon = L.icon({
            iconUrl: 'smoke.png',
            iconSize: [30, 30], // Adjust size as needed
        });

        // Lazy load markers based on visible area
        const allData = ${JSON.stringify(
            data.map((row) => ({
                lat: row.Latitude,
                lng: row.Longitude,
                popup: `<b>Building Name:</b> ${row['Building Name'] || 'N/A'}<br>` +
                       `<b>Address:</b> ${row.Address}<br>` +
                       `<b>Site EUI (kBtu/sq ft):</b> ${(row['Site EUI (kBtu/sq ft)'] || 0).toFixed(2)} kBtu/sq ft`,
            }))
        )};

        const markerLayer = L.featureGroup().addTo(map);

        function loadVisibleMarkers() {
            const bounds = map.getBounds();
            markerLayer.clearLayers();
            allData.forEach((row) => {
                if (bounds.contains([row.lat, row.lng])) {
                    L.marker([row.lat, row.lng], { icon: smokeIcon }).bindPopup(row.popup).addTo(markerLayer);
                }
            });
        }

        map.on('moveend', loadVisibleMarkers);
        loadVisibleMarkers(); // Load markers initially
    </script>
</body>
</html>
    `;

    // Save HTML file
    fs.writeFileSync(outputHtmlFile, htmlContent);
    console.log(`Map with smoke icon and lazy loading generated and saved to ${outputHtmlFile}`);
}


// Run the process
processAndGenerateMap();
