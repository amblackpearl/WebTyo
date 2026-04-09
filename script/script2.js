// ========== KONFIGURASI LANGSUNG ==========
    // Link spreadsheet yang diberikan (format publik)
    const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1SVG5tqs6TUktTN_7wUXDpUYGr0eXMPo63eyktQA2eWg/edit?gid=0#gid=0";
    
    // Tampilkan URL di info bar
    document.getElementById('sourceUrlDisplay').innerText = SPREADSHEET_URL;
    
    // Global variables
    let currentChart = null;
    let allSheetsData = {};      // Key = nama sheet (kategori), Value = { years, emas, perak, perunggu }
    let sheetNamesList = [];
    
    const loadingDiv = document.getElementById('loadingIndicator');
    const errorDiv = document.getElementById('errorMsg');
    const successDiv = document.getElementById('successMsg');
    const sheetPanel = document.getElementById('sheetPanel');
    const sheetButtonsDiv = document.getElementById('sheetButtons');
    const canvas = document.getElementById('medalChart');
    const statsInfo = document.getElementById('statsInfo');
    
    function showError(msg) {
        errorDiv.style.display = 'block';
        errorDiv.innerHTML = `<strong>❌ Error:</strong> ${msg}`;
        loadingDiv.style.display = 'none';
        sheetPanel.style.display = 'none';
        successDiv.style.display = 'none';
    }
    
    function showSuccess(msg) {
        successDiv.style.display = 'block';
        successDiv.innerHTML = `<strong>✅ ${msg}</strong>`;
        setTimeout(() => {
            successDiv.style.display = 'none';
        }, 4000);
    }
    
    function hideError() {
        errorDiv.style.display = 'none';
    }
    
    // Fungsi parsing data dari worksheet (array of array)
    function parseSheetData(worksheetRows, sheetName) {
        if (!worksheetRows || worksheetRows.length < 2) {
            console.warn(`Sheet "${sheetName}" kosong atau tidak cukup baris`);
            return null;
        }
        
        // Cari baris header - mencari baris yang mengandung kata tahun/year dan emas/gold
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(worksheetRows.length, 15); i++) {
            const row = worksheetRows[i];
            if (!row || row.length === 0) continue;
            const rowStr = row.join(' ').toLowerCase();
            if ((rowStr.includes('tahun') || rowStr.includes('year')) && 
                (rowStr.includes('emas') || rowStr.includes('gold') || rowStr.includes('perak') || rowStr.includes('silver'))) {
                headerRowIndex = i;
                break;
            }
        }
        
        if (headerRowIndex === -1) {
            // Jika tidak ketemu, coba asumsikan baris pertama sebagai header
            headerRowIndex = 0;
            console.log(`Sheet "${sheetName}": Menggunakan baris pertama sebagai header`);
        }
        
        const headers = worksheetRows[headerRowIndex].map(cell => (cell || '').toString().toLowerCase().trim());
        
        // Cari indeks kolom dengan berbagai kemungkinan nama
        let colNama = headers.findIndex(h => h.includes('nama') || h === 'name' || h.includes('negara') || h.includes('kontingen') || h.includes('atlet'));
        let colTahun = headers.findIndex(h => h.includes('tahun') || h === 'year');
        let colEmas = headers.findIndex(h => h.includes('emas') || h === 'gold');
        let colPerak = headers.findIndex(h => h.includes('perak') || h === 'silver');
        let colPerunggu = headers.findIndex(h => h.includes('perunggu') || h === 'bronze');
        
        // Fallback berdasarkan posisi umum jika tidak ketemu
        if (colTahun === -1 && headers.length > 1) colTahun = 1;
        if (colEmas === -1 && headers.length > 2) colEmas = 2;
        if (colPerak === -1 && headers.length > 3) colPerak = 3;
        if (colPerunggu === -1 && headers.length > 4) colPerunggu = 4;
        if (colNama === -1) colNama = 0;
        
        console.log(`Sheet "${sheetName}" - Kolom: Tahun=${colTahun}, Emas=${colEmas}, Perak=${colPerak}, Perunggu=${colPerunggu}`);
        
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
        
        if (tahunMap.size === 0) {
            console.warn(`Sheet "${sheetName}" tidak memiliki data numerik tahun yang valid`);
            return null;
        }
        
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
    
    // Konversi URL publik ke XLSX untuk mendapatkan semua sheet
    function getXlsxUrlFromPublicCsv(csvUrl) {
        const match = csvUrl.match(/\/e\/([a-zA-Z0-9_-]+)/);
        if (match) {
            const docId = match[1];
            return `https://docs.google.com/spreadsheets/d/e/${docId}/export?format=xlsx`;
        }
        return csvUrl.replace('pub?output=csv', 'export?format=xlsx');
    }
    
    // Memuat spreadsheet dari URL XLSX
    async function loadSpreadsheet() {
        hideError();
        loadingDiv.style.display = 'block';
        sheetPanel.style.display = 'none';
        
        try {
            const XLSX = await loadSheetJSLib();
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
            
            console.log(`Ditemukan ${sheets.length} sheet:`, sheets);
            
            const parsedData = {};
            let successCount = 0;
            
            for (const sheetName of sheets) {
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                if (rows && rows.length > 1) {
                    const parsed = parseSheetData(rows, sheetName);
                    if (parsed && parsed.years && parsed.years.length > 0) {
                        parsedData[sheetName] = parsed;
                        successCount++;
                        console.log(`✅ Sheet "${sheetName}" berhasil diproses, tahun: ${parsed.years.join(', ')}`);
                    } else {
                        console.warn(`⚠️ Sheet "${sheetName}" tidak memiliki data medali yang valid`);
                    }
                } else {
                    console.warn(`⚠️ Sheet "${sheetName}" kosong`);
                }
            }
            
            if (successCount === 0) {
                throw new Error("Tidak ada sheet yang berisi data medali yang valid. Pastikan setiap sheet memiliki kolom: Tahun, Emas, Perak, Perunggu.");
            }
            
            allSheetsData = parsedData;
            sheetNamesList = Object.keys(parsedData);
            
            // Update UI: tampilkan semua kategori (nama sheet)
            renderSheetButtons();
            if (sheetNamesList.length > 0) {
                renderChart(sheetNamesList[0]);
                sheetPanel.style.display = 'block';
                showSuccess(`Berhasil memuat ${sheetNamesList.length} kategori (sheet): ${sheetNamesList.join(', ')}`);
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
        sheetNamesList.forEach((sheetName, idx) => {
            const btn = document.createElement('button');
            // TAMPILKAN NAMA SHEET ASLI sebagai teks kategori
            btn.innerText = sheetName;
            btn.classList.add('sheet-btn');
            if (idx === 0) btn.classList.add('active');
            btn.title = `Klik untuk melihat grafik medali ${sheetName}`;
            btn.addEventListener('click', () => {
                // Update active class pada semua tombol
                document.querySelectorAll('.sheet-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                renderChart(sheetName);
                // Update info stats
                const data = allSheetsData[sheetName];
                if (data) {
                    statsInfo.innerHTML = `📊 Menampilkan data <strong>${sheetName}</strong> | Tahun: ${data.years.join(' → ')} | Total Medali: ${data.emas.reduce((a,b)=>a+b,0) + data.perak.reduce((a,b)=>a+b,0) + data.perunggu.reduce((a,b)=>a+b,0)}`;
                }
            });
            sheetButtonsDiv.appendChild(btn);
        });
        
        // Update stats untuk sheet pertama
        if (sheetNamesList.length > 0) {
            const firstData = allSheetsData[sheetNamesList[0]];
            if (firstData) {
                statsInfo.innerHTML = `📊 Menampilkan data <strong>${sheetNamesList[0]}</strong> | Tahun: ${firstData.years.join(' → ')} | Total Medali: ${firstData.emas.reduce((a,b)=>a+b,0) + firstData.perak.reduce((a,b)=>a+b,0) + firstData.perunggu.reduce((a,b)=>a+b,0)}`;
            }
        }
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
                        backgroundColor: 'rgba(251, 191, 36, 0.85)',
                        borderColor: '#f59e0b',
                        borderWidth: 2,
                        borderRadius: 8,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8
                    },
                    {
                        label: '🥈 Perak',
                        data: data.perak,
                        backgroundColor: 'rgba(156, 163, 175, 0.85)',
                        borderColor: '#6b7280',
                        borderWidth: 2,
                        borderRadius: 8,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8
                    },
                    {
                        label: '🥉 Perunggu',
                        data: data.perunggu,
                        backgroundColor: 'rgba(217, 119, 6, 0.8)',
                        borderColor: '#b45309',
                        borderWidth: 2,
                        borderRadius: 8,
                        barPercentage: 0.7,
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
                        padding: { bottom: 20, top: 10 },
                        color: '#1e293b'
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        titleFont: { size: 13, weight: 'bold' },
                        bodyFont: { size: 12 },
                        callbacks: {
                            label: function(context) {
                                return `${context.dataset.label}: ${context.raw} medali`;
                            }
                        }
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
                    duration: 700,
                    easing: 'easeOutQuart'
                }
            }
        });
    }
    
    // Jalankan otomatis saat halaman dimuat
    window.addEventListener('DOMContentLoaded', () => {
        loadSpreadsheet();
    });