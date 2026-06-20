
const map = L.map('map').setView([20, 0], 2);
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; OpenStreetMap contributors'
}).addTo(map);
const markerGroup = L.layerGroup().addTo(map);
const whiteIcon = new L.Icon({
    iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-grey.png',
    shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-shadow.png',
    iconSize: [25, 41],
    iconAnchor: [12, 41],
    popupAnchor: [1, -34],
    shadowSize: [41, 41]
});
async function loadData() {
    // fetch the local Excel file

    const res = await fetch('data.xlsx');
    const arrayBuffer = await res.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetSPBE = workbook.Sheets['SPBE'];
    if (!sheetSPBE) {
        console.error('Sheet "SPBE" not found. Available:', workbook.SheetNames);
    }
    const rowsSPBE = XLSX.utils.sheet_to_json(sheetSPBE);
    console.log(rowsSPBE); // check your column names here

    const bounds = [];
    rowsSPBE.forEach(row => {
        const lat = parseFloat(row.lat ?? row.Lat ?? row.latitude ?? row.Latitude);
        const lng = parseFloat(row.lng ?? row.Lng ?? row.lon ?? row.Longitude ?? row.longitude);
        const label = row.SPBE ?? row.SPBE ?? '';

        if (!isNaN(lat) && !isNaN(lng)) {

            L.circleMarker([lat, lng], {
                radius: 7,
                color: '#1565c0',        // border
                fillColor: '#42a5f5',    // fill
                fillOpacity: 0.85,
                weight: 2
            }).addTo(markerGroup).bindPopup(String(label));
            bounds.push([lat, lng]);

        }
    });

    const sheetTLPG = workbook.Sheets['TLPG'];
    if (!sheetTLPG) {
        console.error('Sheet "TLPG" not found. Available:', workbook.SheetNames);
    }
    const rowsTLPG = XLSX.utils.sheet_to_json(sheetTLPG);
    console.log(rowsTLPG); // check your column names here

    rowsTLPG.forEach(row => {
        const lat = parseFloat(row.lat ?? row.Lat ?? row.latitude ?? row.Latitude);
        const lng = parseFloat(row.lng ?? row.Lng ?? row.lon ?? row.Longitude ?? row.longitude);
        const label = row.TLPG ?? row.TLPG ?? '';

        if (!isNaN(lat) && !isNaN(lng)) {


            L.marker([lat, lng], { icon: whiteIcon }).addTo(markerGroup).bindPopup(String(label));
            bounds.push([lat, lng]);


        }
    });

    if (bounds.length) map.fitBounds(bounds);
}

loadData().catch(err => console.error('Failed to load data:', err));