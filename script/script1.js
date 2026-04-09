// ========== KONFIGURASI LANGSUNG ==========
    // Link spreadsheet yang diberikan (format CSV/Excel)
    const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1SVG5tqs6TUktTN_7wUXDpUYGr0eXMPo63eyktQA2eWg/edit?usp=sharing";
    
    // Untuk multi-sheet, karena link di atas adalah CSV export dari sheet pertama saja,
    // kita perlu mengakses file asli XLSX. Google Spreadsheet dengan ID publik bisa diunduh sebagai XLSX.
    // Kita extract ID dari link publik yang diberikan.
    // Format link: /e/2PACX-1vSsT7WVKJTlaKHzmaBxDOCmTJmK-5WUuLIIiSmhwKcgtMNKy9BRuLdxc9ZAPXbUB7VQAxZTI81lFYuG/
    function getXlsxUrlFromPublicCsv(csvUrl) {
        // Extract ID dari pola /e/.../
        const match = csvUrl.match(/\/e\/([a-zA-Z0-9_-]+)/);
        if (match) {
            const docId = match[1];
            // Gunakan endpoint export XLSX untuk mendapatkan semua sheet
            return `https://docs.google.com/spreadsheets/d/e/${docId}/export?format=xlsx`;
        }
        // Fallback: coba dari pattern lain
        return csvUrl.replace('pub?output=csv', 'export?format=xlsx');
    }
    
    // Global variables
    let currentChart = null;
    let allSheetsData = {};      // { sheetName: { years, emas, perak, perunggu } }
    let sheetNames = [];
    
    const loadingDiv = document.getElementById('loadingIndicator');
    const errorDiv = document.getElementById('errorMsg');
    const sheetPanel = document.getElementById('sheetPanel');
    const sheetButtonsDiv = document.getElementById('sheetButtons');
    const canvas = document.getElementById('medalChart');
    
    function showError(msg) {
        errorDiv.style.display = 'block';
        errorDiv.innerHTML = `<strong>❌ Error:</strong> ${msg}`;
        loadingDiv.style.display = 'none';
        sheetPanel.style.display = 'none';
    }
    
    function hideError() {
        errorDiv.style.display = 'none';
    }
    
    // Parsing data dari worksheet (array of array)
    function parseSheetData(worksheetRows, sheetName) {
        if (!worksheetRows || worksheetRows.length < 2) return null;
        
        // Cari baris header - cari baris yang mengandung kata tahun/year dan emas/gold
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(worksheetRows.length, 15); i++) {
            const row = worksheetRows[i];
            if (!row || row.length === 0) continue;
            const rowStr = row.join(' ').toLowerCase();
            if ((rowStr.includes('tahun') || rowStr.includes('year')) && 
                (rowStr.includes('emas') || rowStr.includes('gold') || rowStr.includes('perak'))) {
                headerRowIndex = i;
                break;
            }
        }
        
        if (headerRowIndex === -1) {
            // Asumsikan baris pertama sebagai header
            headerRowIndex = 0;
        }
        
        const headers = worksheetRows[headerRowIndex].map(cell => (cell || '').toString().toLowerCase().trim());
        
        // Cari indeks kolom
        let colNama = headers.findIndex(h => h.includes('nama') || h === 'name' || h.includes('negara') || h.includes('kontingen'));
        let colTahun = headers.findIndex(h => h.includes('tahun') || h === 'year');
        let colEmas = headers.findIndex(h => h.includes('emas') || h === 'gold');
        let colPerak = headers.findIndex(h => h.includes('perak') || h === 'silver');
        let colPerunggu = headers.findIndex(h => h.includes('perunggu') || h === 'bronze');
        
        // Fallback berdasarkan posisi umum
        if (colTahun === -1 && headers.length > 1) colTahun = 1;
        if (colEmas === -1 && headers.length > 2) colEmas = 2;
        if (colPerak === -1 && headers.length > 3) colPerak = 3;
        if (colPerunggu === -1 && headers.length > 4) colPerunggu = 4;
        if (colNama === -1) colNama = 0;
        
        const dataRows = worksheetRows.slice(headerRowIndex + 1);
        const tahunMap = new Map(); // tahun -> { emas, perak, perunggu }
        
        for (const row of dataRows) {
            if (!row || row.length === 0) continue;
            const tahunRaw = row[colTahun];
            const tahun = parseInt(tahunRaw, 10);
            if (isNaN(tahun)) continue;
            
            const emas = parseInt(row[colEmas]) || 0;
            const perak = parseInt(row[colPerak]) || 0;
            const perunggu = parseInt(row[colPerunggu]) || 0;
            
            if (!tahunMap.has(tahun)) {
                tahunMap.set(tahun, { emas: 0, perak: 0, perunggu: 0 });
            }
            const agg = tahunMap.get(tahun);
            agg.emas += emas;
            agg.perak += perak;
            agg.perunggu += perunggu;
        }
        
        if (tahunMap.size === 0) return null;
        
        const sortedYears = Array.from(tahunMap.keys()).sort((a, b) => a - b);
        const emasArr = sortedYears.map(y => tahunMap.get(y).emas);
        const perakArr = sortedYears.map(y => tahunMap.get(y).perak);
        const perungguArr = sortedYears.map(y => tahunMap.get(y).perunggu);
        
        return {
            sheetName: sheetName,
            years: sortedYears,
            emas: emasArr,
            perak: perakArr,
            perunggu: perungguArr
        };
    }
    
    // Load SheetJS library
    function loadSheetJSLib() {
        return new Promise((resolve, reject) => {
            if (window.XLSX) {
                resolve(window.XLSX);
                return;
            }
            const script = document.createElement('script');
            script.src = 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';
            script.onload = () => resolve(window.XLSX);
            script.onerror = () => reject(new Error('Gagal memuat library SheetJS'));
            document.head.appendChild(script);
        });
    }
    
    // Memuat spreadsheet dari URL XLSX
    async function loadSpreadsheet() {
        hideError();
        loadingDiv.style.display = 'block';
        sheetPanel.style.display = 'none';
        
        try {
            const XLSX = await loadSheetJSLib();
            
            // Konversi ke URL XLSX untuk mendapatkan semua sheet
            const xlsxUrl = getXlsxUrlFromPublicCsv(SPREADSHEET_URL);
            console.log("Loading XLSX from:", xlsxUrl);
            
            const response = await fetch(xlsxUrl);
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: Gagal mengunduh file. Pastikan spreadsheet dapat diakses publik.`);
            }
            
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false, defval: "" });
            
            const sheets = workbook.SheetNames;
            if (!sheets || sheets.length === 0) {
                throw new Error("Tidak ada sheet dalam spreadsheet.");
            }
            
            const parsedData = {};
            
            for (const sheetName of sheets) {
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                if (rows && rows.length > 1) {
                    const parsed = parseSheetData(rows, sheetName);
                    if (parsed && parsed.years && parsed.years.length > 0) {
                        parsedData[sheetName] = parsed;
                    } else {
                        console.warn(`Sheet "${sheetName}" tidak memiliki data medali yang valid (minimal Tahun & medali).`);
                    }
                }
            }
            
            const sheetCount = Object.keys(parsedData).length;
            if (sheetCount === 0) {
                throw new Error("Tidak ada sheet yang berisi data medali. Pastikan kolom: Tahun, Emas, Perak, Perunggu.");
            }
            
            allSheetsData = parsedData;
            sheetNames = Object.keys(parsedData);
            
            // Render UI
            renderSheetButtons();
            if (sheetNames.length > 0) {
                renderChart(sheetNames[0]);
                sheetPanel.style.display = 'block';
            }
            
            loadingDiv.style.display = 'none';
            
        } catch (err) {
            console.error(err);
            loadingDiv.style.display = 'none';
            showError(err.message + " Periksa kembali link spreadsheet dan pastikan akses publik (File → Bagikan → 'Siapa saja dengan tautan dapat melihat').");
        }
    }
    
    function renderSheetButtons() {
        sheetButtonsDiv.innerHTML = '';
        sheetNames.forEach((sheet, idx) => {
            const btn = document.createElement('button');
            btn.innerText = sheet;
            btn.classList.add('sheet-btn');
            if (idx === 0) btn.classList.add('active');
            btn.addEventListener('click', () => {
                // Update active class
                document.querySelectorAll('.sheet-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                renderChart(sheet);
            });
            sheetButtonsDiv.appendChild(btn);
        });
    }
    
    function renderChart(sheetName) {
        const data = allSheetsData[sheetName];
        if (!data) return;
        
        const ctx = canvas.getContext('2d');
        if (currentChart) {
            currentChart.destroy();
        }
        
        currentChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: data.years.map(y => y.toString()),
                datasets: [
                    {
                        label: '🥇 Emas',
                        data: data.emas,
                        backgroundColor: 'rgba(255, 193, 7, 0.75)',
                        borderColor: '#f59e0b',
                        borderWidth: 2,
                        borderRadius: 8,
                        barPercentage: 0.65,
                        categoryPercentage: 0.8
                    },
                    {
                        label: '🥈 Perak',
                        data: data.perak,
                        backgroundColor: 'rgba(156, 163, 175, 0.75)',
                        borderColor: '#6b7280',
                        borderWidth: 2,
                        borderRadius: 8,
                        barPercentage: 0.65,
                        categoryPercentage: 0.8
                    },
                    {
                        label: '🥉 Perunggu',
                        data: data.perunggu,
                        backgroundColor: 'rgba(217, 119, 6, 0.7)',
                        borderColor: '#b45309',
                        borderWidth: 2,
                        borderRadius: 8,
                        barPercentage: 0.65,
                        categoryPercentage: 0.8
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: { font: { size: 13, weight: 'bold' }, usePointStyle: true, boxWidth: 12 }
                    },
                    title: {
                        display: true,
                        text: `🏅 Perolehan Medali - ${sheetName}`,
                        font: { size: 18, weight: 'bold' },
                        padding: { bottom: 20 }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        titleFont: { size: 13 },
                        bodyFont: { size: 12 }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: { display: true, text: 'Jumlah Medali', font: { size: 12, weight: 'bold' } },
                        grid: { color: '#e2e8f0', drawBorder: true },
                        ticks: { stepSize: 1, precision: 0 }
                    },
                    x: {
                        title: { display: true, text: 'Tahun', font: { size: 12, weight: 'bold' } },
                        grid: { display: false }
                    }
                },
                animation: {
                    duration: 800,
                    easing: 'easeOutQuart'
                }
            }
        });
    }
    
    // Jalankan otomatis saat halaman dimuat
    window.addEventListener('DOMContentLoaded', () => {
        loadSpreadsheet();
    });