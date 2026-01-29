const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(express.static('public'));

// File upload config
const storage = multer.diskStorage({
    destination: './uploads/',
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage });

// Data files
const dataDir = './data';
const productionFile = path.join(dataDir, 'production_data.json');
const turnaroundFile = path.join(dataDir, 'turnaround_data.json');
const bypassFile = path.join(dataDir, 'bypass_data.json');
const usageFile = path.join(dataDir, 'usage_data.json');
const productUsageFile = path.join(dataDir, 'product_usage.json');
const productWastageFile = path.join(dataDir, 'product_wastage.json');
const detailedWastageFile = path.join(dataDir, 'detailed_wastage.json');
const stockDosesFile = path.join(dataDir, 'stock_doses.json');

// Load data
function loadData(file) {
    if (fs.existsSync(file)) {
        return JSON.parse(fs.readFileSync(file, 'utf8'));
    }
    return [];
}

function saveData(file, data) {
    fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

let productionData = loadData(productionFile);
let turnaroundData = loadData(turnaroundFile);
let bypassData = loadData(bypassFile);
let usageData = loadData(usageFile);
let productUsage = loadData(productUsageFile);
let productWastage = loadData(productWastageFile);
let detailedWastage = loadData(detailedWastageFile);
let stockDoses = loadData(stockDosesFile);

// Utility functions
function excelDateToISO(serial) {
    const utc_days = Math.floor(serial - 25569);
    const date = new Date(utc_days * 86400 * 1000);
    return date.toISOString().split('T')[0];
}

function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
}

// ============ API ENDPOINTS ============

// Summary stats
app.get('/api/summary', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    // Production summary
    const recentProduction = productionData.filter(d => new Date(d.date) >= cutoff);
    const totalDoses = recentProduction.reduce((sum, d) => sum + d.total_doses, 0);

    // Turnaround summary
    const recentTurnaround = turnaroundData.filter(d => new Date(d.date) >= cutoff);
    const avgTurnaround = recentTurnaround.length > 0
        ? recentTurnaround.reduce((sum, d) => sum + d.avg_turnaround, 0) / recentTurnaround.length
        : 0;

    // Bypass summary
    const totalBypasses = bypassData.reduce((sum, d) => sum + d.total_bypasses, 0);
    const bypassRate = totalDoses > 0 ? (totalBypasses * 100 / totalDoses).toFixed(2) : 0;

    res.json({
        total_doses: totalDoses,
        days_tracked: recentProduction.length,
        avg_turnaround_minutes: parseFloat(avgTurnaround.toFixed(1)),
        total_bypasses: totalBypasses,
        bypass_rate: parseFloat(bypassRate),
        data_range: {
            production: productionData.length > 0 ? {
                start: productionData[productionData.length - 1]?.date,
                end: productionData[0]?.date
            } : null
        }
    });
});

// Production trends (daily doses)
app.get('/api/production/daily', (req, res) => {
    const { days = 30, grouping = 'daily' } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    let filtered = productionData.filter(d => new Date(d.date) >= cutoff);

    if (grouping === 'weekly') {
        const byWeek = {};
        filtered.forEach(d => {
            const date = new Date(d.date + 'T00:00:00');
            const day = date.getDay();
            const diff = date.getDate() - day + (day === 0 ? -6 : 1);
            const monday = new Date(date);
            monday.setDate(diff);
            const weekKey = 'Week of ' + monday.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

            if (!byWeek[weekKey]) byWeek[weekKey] = { total_doses: 0, sortKey: d.date, hourly: {} };
            byWeek[weekKey].total_doses += d.total_doses;
            if (d.date < byWeek[weekKey].sortKey) byWeek[weekKey].sortKey = d.date;

            // Aggregate hourly
            if (d.hourly) {
                Object.entries(d.hourly).forEach(([hour, count]) => {
                    byWeek[weekKey].hourly[hour] = (byWeek[weekKey].hourly[hour] || 0) + count;
                });
            }
        });

        filtered = Object.entries(byWeek).map(([date, data]) => ({
            date,
            total_doses: data.total_doses,
            sortKey: data.sortKey,
            hourly: data.hourly
        })).sort((a, b) => a.sortKey.localeCompare(b.sortKey));
    }

    res.json(filtered);
});

// Hourly production patterns
app.get('/api/production/hourly', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = productionData.filter(d => new Date(d.date) >= cutoff);

    // Aggregate by hour
    const byHour = {};
    for (let h = 0; h < 24; h++) {
        byHour[h] = { hour: h, total_doses: 0, days_count: 0 };
    }

    filtered.forEach(d => {
        if (d.hourly) {
            Object.entries(d.hourly).forEach(([hour, count]) => {
                const h = parseInt(hour);
                if (byHour[h] !== undefined) {
                    byHour[h].total_doses += count;
                    byHour[h].days_count++;
                }
            });
        }
    });

    const result = Object.values(byHour).map(h => ({
        hour: h.hour,
        hour_label: formatHour(h.hour),
        total_doses: h.total_doses,
        avg_doses: h.days_count > 0 ? Math.round(h.total_doses / h.days_count) : 0
    }));

    res.json(result);
});

// Turnaround time analysis
app.get('/api/turnaround/daily', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = turnaroundData.filter(d => new Date(d.date) >= cutoff);
    res.json(filtered.sort((a, b) => a.date.localeCompare(b.date)));
});

// Turnaround by priority
app.get('/api/turnaround/by-priority', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = turnaroundData.filter(d => new Date(d.date) >= cutoff);

    // Aggregate by priority
    const byPriority = {};
    filtered.forEach(d => {
        if (d.by_priority) {
            Object.entries(d.by_priority).forEach(([priority, stats]) => {
                if (!byPriority[priority]) {
                    byPriority[priority] = { count: 0, total_time: 0 };
                }
                byPriority[priority].count += stats.count;
                byPriority[priority].total_time += stats.total_time;
            });
        }
    });

    const result = Object.entries(byPriority)
        .map(([priority, stats]) => ({
            priority: parseInt(priority),
            priority_label: getPriorityLabel(parseInt(priority)),
            count: stats.count,
            avg_turnaround: stats.count > 0 ? parseFloat((stats.total_time / stats.count).toFixed(1)) : 0
        }))
        .sort((a, b) => a.priority - b.priority);

    res.json(result);
});

// Turnaround by workstation
app.get('/api/turnaround/by-workstation', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = turnaroundData.filter(d => new Date(d.date) >= cutoff);

    // Aggregate by workstation
    const byWorkstation = {};
    filtered.forEach(d => {
        if (d.by_workstation) {
            Object.entries(d.by_workstation).forEach(([ws, stats]) => {
                if (!byWorkstation[ws]) {
                    byWorkstation[ws] = { count: 0, total_time: 0 };
                }
                byWorkstation[ws].count += stats.count;
                byWorkstation[ws].total_time += stats.total_time;
            });
        }
    });

    const result = Object.entries(byWorkstation)
        .map(([workstation, stats]) => ({
            workstation,
            count: stats.count,
            avg_turnaround: stats.count > 0 ? parseFloat((stats.total_time / stats.count).toFixed(1)) : 0
        }))
        .sort((a, b) => b.count - a.count);

    res.json(result);
});

// Top drugs by volume
app.get('/api/turnaround/top-drugs', (req, res) => {
    const { days = 30, limit = 20 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = turnaroundData.filter(d => new Date(d.date) >= cutoff);

    // Aggregate by drug
    const byDrug = {};
    filtered.forEach(d => {
        if (d.by_drug) {
            Object.entries(d.by_drug).forEach(([drug, stats]) => {
                if (!byDrug[drug]) {
                    byDrug[drug] = { count: 0, total_time: 0 };
                }
                byDrug[drug].count += stats.count;
                byDrug[drug].total_time += stats.total_time;
            });
        }
    });

    const result = Object.entries(byDrug)
        .map(([drug, stats]) => ({
            drug,
            count: stats.count,
            avg_turnaround: stats.count > 0 ? parseFloat((stats.total_time / stats.count).toFixed(1)) : 0
        }))
        .sort((a, b) => b.count - a.count)
        .slice(0, parseInt(limit));

    res.json(result);
});

// Bypass analysis
app.get('/api/bypass/summary', (req, res) => {
    const byLocation = {};
    bypassData.forEach(d => {
        if (!byLocation[d.location]) {
            byLocation[d.location] = { total_bypasses: 0, hourly: {} };
        }
        byLocation[d.location].total_bypasses += d.total_bypasses;

        if (d.hourly) {
            Object.entries(d.hourly).forEach(([hour, count]) => {
                byLocation[d.location].hourly[hour] = (byLocation[d.location].hourly[hour] || 0) + count;
            });
        }
    });

    const result = Object.entries(byLocation)
        .map(([location, data]) => ({
            location,
            total_bypasses: data.total_bypasses,
            hourly: data.hourly
        }))
        .sort((a, b) => b.total_bypasses - a.total_bypasses);

    res.json(result);
});

// Bypass hourly pattern
app.get('/api/bypass/hourly', (req, res) => {
    const byHour = {};
    for (let h = 0; h < 24; h++) {
        byHour[h] = 0;
    }

    bypassData.forEach(d => {
        if (d.hourly) {
            Object.entries(d.hourly).forEach(([hour, count]) => {
                const h = parseInt(hour);
                if (byHour[h] !== undefined) {
                    byHour[h] += count;
                }
            });
        }
    });

    const result = Object.entries(byHour).map(([hour, count]) => ({
        hour: parseInt(hour),
        hour_label: formatHour(parseInt(hour)),
        bypasses: count
    }));

    res.json(result);
});

// ============ UPLOAD ENDPOINTS ============

// Upload Production Stats (Dose Order Prep Statistics by Date)
app.post('/api/upload/production', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const results = { added: 0, updated: 0, errors: [] };

        // Find header row
        const headerRow = rawData[0];
        if (!headerRow || headerRow[0] !== 'EntryDate') {
            return res.status(400).json({ error: 'Invalid file format - expected EntryDate in first column' });
        }

        // Parse hour columns
        const hourMap = {};
        for (let i = 1; i < headerRow.length; i++) {
            const hourStr = headerRow[i];
            if (typeof hourStr === 'string' && hourStr.includes(':')) {
                const hour = parseInt(hourStr.split(':')[0]);
                hourMap[i] = hour;
            }
        }

        // Process data rows
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || !row[0]) continue;

            const dateSerial = row[0];
            const date = excelDateToISO(dateSerial);

            const hourly = {};
            let totalDoses = 0;

            Object.entries(hourMap).forEach(([colIdx, hour]) => {
                const count = parseInt(row[parseInt(colIdx)]) || 0;
                hourly[hour] = count;
                totalDoses += count;
            });

            const existingIdx = productionData.findIndex(d => d.date === date);
            const record = {
                id: existingIdx >= 0 ? productionData[existingIdx].id : generateId(),
                date,
                total_doses: totalDoses,
                hourly,
                uploaded_at: new Date().toISOString()
            };

            if (existingIdx >= 0) {
                productionData[existingIdx] = record;
                results.updated++;
            } else {
                productionData.push(record);
                results.added++;
            }
        }

        productionData.sort((a, b) => b.date.localeCompare(a.date));
        saveData(productionFile, productionData);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Upload Turnaround Report
app.post('/api/upload/turnaround', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const results = { added: 0, updated: 0, total_doses: 0, errors: [] };

        // Find columns
        const headerRow = rawData[0];
        const colIdx = {
            doseId: headerRow.indexOf('DoseID'),
            priority: headerRow.indexOf('Priority'),
            patientLocation: headerRow.indexOf('PatientLocation'),
            facility: headerRow.indexOf('Facility'),
            hazmat: headerRow.indexOf('HazMat'),
            doseDesc: headerRow.indexOf('DoseDescription'),
            status: headerRow.indexOf('DoseStatus'),
            totalTimeline: headerRow.indexOf('Total Dose Timeline - Start to Sort (m)'),
            totalPrep: headerRow.indexOf('Total Prep Time (m)'),
            workstation: headerRow.indexOf('WorkStation'),
            queueToPrint: headerRow.indexOf('Queued to Printed Time (m)'),
            queueToSort: headerRow.indexOf('Queued to Sorted Time (m)')
        };

        // We need to extract date from the data - use queueToSort time to estimate date
        // Group by date
        const byDate = {};

        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || !row[colIdx.doseId]) continue;

            // Extract drug name (first part before dose info)
            const doseDesc = row[colIdx.doseDesc] || '';
            const drugName = doseDesc.split(' ')[0];

            const totalTimeline = parseFloat(row[colIdx.totalTimeline]) || 0;
            const priority = parseInt(row[colIdx.priority]) || 0;
            const workstation = row[colIdx.workstation] || 'Unknown';
            const status = row[colIdx.status] || '';

            // Skip non-sorted doses
            if (status !== 'Sorted') continue;

            // For now, we'll use today's date minus row index as a proxy
            // In reality, we'd want the actual date from the report parameters
            // Let's group all into a single batch for now and use upload date
            const date = new Date().toISOString().split('T')[0];

            if (!byDate[date]) {
                byDate[date] = {
                    count: 0,
                    total_time: 0,
                    by_priority: {},
                    by_workstation: {},
                    by_drug: {}
                };
            }

            byDate[date].count++;
            byDate[date].total_time += totalTimeline;
            results.total_doses++;

            // By priority
            if (!byDate[date].by_priority[priority]) {
                byDate[date].by_priority[priority] = { count: 0, total_time: 0 };
            }
            byDate[date].by_priority[priority].count++;
            byDate[date].by_priority[priority].total_time += totalTimeline;

            // By workstation
            if (!byDate[date].by_workstation[workstation]) {
                byDate[date].by_workstation[workstation] = { count: 0, total_time: 0 };
            }
            byDate[date].by_workstation[workstation].count++;
            byDate[date].by_workstation[workstation].total_time += totalTimeline;

            // By drug
            if (drugName && !byDate[date].by_drug[drugName]) {
                byDate[date].by_drug[drugName] = { count: 0, total_time: 0 };
            }
            if (drugName) {
                byDate[date].by_drug[drugName].count++;
                byDate[date].by_drug[drugName].total_time += totalTimeline;
            }
        }

        // Save aggregated data
        Object.entries(byDate).forEach(([date, data]) => {
            const existingIdx = turnaroundData.findIndex(d => d.date === date);
            const record = {
                id: existingIdx >= 0 ? turnaroundData[existingIdx].id : generateId(),
                date,
                total_doses: data.count,
                avg_turnaround: data.count > 0 ? parseFloat((data.total_time / data.count).toFixed(1)) : 0,
                by_priority: data.by_priority,
                by_workstation: data.by_workstation,
                by_drug: data.by_drug,
                uploaded_at: new Date().toISOString()
            };

            if (existingIdx >= 0) {
                turnaroundData[existingIdx] = record;
                results.updated++;
            } else {
                turnaroundData.push(record);
                results.added++;
            }
        });

        turnaroundData.sort((a, b) => b.date.localeCompare(a.date));
        saveData(turnaroundFile, turnaroundData);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Upload Bypass Stats
app.post('/api/upload/bypass', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const results = { added: 0, updated: 0, errors: [] };

        // Find header row
        const headerRow = rawData[0];
        if (!headerRow || headerRow[0] !== 'Location') {
            return res.status(400).json({ error: 'Invalid file format - expected Location in first column' });
        }

        // Parse hour columns
        const hourMap = {};
        for (let i = 1; i < headerRow.length; i++) {
            const hourStr = headerRow[i];
            if (typeof hourStr === 'string' && hourStr.includes(':')) {
                const hour = parseInt(hourStr.split(':')[0]);
                hourMap[i] = hour;
            }
        }

        // Process data rows
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || !row[0]) continue;

            const location = row[0];
            const hourly = {};
            let totalBypasses = 0;

            Object.entries(hourMap).forEach(([colIdx, hour]) => {
                const count = parseInt(row[parseInt(colIdx)]) || 0;
                hourly[hour] = count;
                totalBypasses += count;
            });

            const existingIdx = bypassData.findIndex(d => d.location === location);
            const record = {
                id: existingIdx >= 0 ? bypassData[existingIdx].id : generateId(),
                location,
                total_bypasses: totalBypasses,
                hourly,
                uploaded_at: new Date().toISOString()
            };

            if (existingIdx >= 0) {
                bypassData[existingIdx] = record;
                results.updated++;
            } else {
                bypassData.push(record);
                results.added++;
            }
        }

        saveData(bypassFile, bypassData);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Upload Usage/Wastage Report
app.post('/api/upload/usage', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const results = { added: 0, updated: 0, total_records: 0, total_waste_ml: 0, errors: [] };

        // Find columns
        const headerRow = rawData[0];
        const colIdx = {
            patientLocation: headerRow.indexOf('Patient Location'),
            prepTime: headerRow.indexOf('Dose Preparation Time'),
            doseDesc: headerRow.indexOf('Dose Description'),
            doseId: headerRow.indexOf('Dose ID'),
            productName: headerRow.indexOf('Product Name'),
            productNDC: headerRow.indexOf('Product NDC'),
            totalVolume: headerRow.indexOf('Product Total Volume'),
            unusedVolume: headerRow.indexOf('Product Unused Volume'),
            multiDose: headerRow.indexOf('Multi-Dose Product')
        };

        // Group by date and product
        const byDate = {};
        const byProduct = {};

        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || !row[colIdx.productName]) continue;

            results.total_records++;

            const prepTimeSerial = row[colIdx.prepTime];
            const date = prepTimeSerial ? excelDateToISO(prepTimeSerial) : null;
            if (!date) continue;

            const productName = row[colIdx.productName] || 'Unknown';
            const productNDC = row[colIdx.productNDC] || '';
            const totalVolume = parseFloat(row[colIdx.totalVolume]) || 0;
            const unusedVolume = parseFloat(row[colIdx.unusedVolume]) || 0;
            const wasteVolume = unusedVolume;
            const usedVolume = totalVolume - unusedVolume;
            const isMultiDose = row[colIdx.multiDose] === 'Yes';
            const location = row[colIdx.patientLocation] || 'Unknown';

            results.total_waste_ml += wasteVolume;

            // Aggregate by date
            if (!byDate[date]) {
                byDate[date] = {
                    total_volume: 0,
                    used_volume: 0,
                    waste_volume: 0,
                    product_count: 0,
                    by_product: {},
                    by_location: {}
                };
            }
            byDate[date].total_volume += totalVolume;
            byDate[date].used_volume += usedVolume;
            byDate[date].waste_volume += wasteVolume;
            byDate[date].product_count++;

            // By product within date
            if (!byDate[date].by_product[productName]) {
                byDate[date].by_product[productName] = { used: 0, waste: 0, count: 0, ndc: productNDC };
            }
            byDate[date].by_product[productName].used += usedVolume;
            byDate[date].by_product[productName].waste += wasteVolume;
            byDate[date].by_product[productName].count++;

            // By location within date
            const locKey = location.split('-')[0]; // Simplify location
            if (!byDate[date].by_location[locKey]) {
                byDate[date].by_location[locKey] = { used: 0, waste: 0, count: 0 };
            }
            byDate[date].by_location[locKey].used += usedVolume;
            byDate[date].by_location[locKey].waste += wasteVolume;
            byDate[date].by_location[locKey].count++;

            // Global product tracking
            if (!byProduct[productName]) {
                byProduct[productName] = { used: 0, waste: 0, count: 0, ndc: productNDC };
            }
            byProduct[productName].used += usedVolume;
            byProduct[productName].waste += wasteVolume;
            byProduct[productName].count++;
        }

        // Save daily usage data
        Object.entries(byDate).forEach(([date, data]) => {
            const existingIdx = usageData.findIndex(d => d.date === date);
            const record = {
                id: existingIdx >= 0 ? usageData[existingIdx].id : generateId(),
                date,
                total_volume: parseFloat(data.total_volume.toFixed(2)),
                used_volume: parseFloat(data.used_volume.toFixed(2)),
                waste_volume: parseFloat(data.waste_volume.toFixed(2)),
                waste_percent: data.total_volume > 0 ? parseFloat((data.waste_volume * 100 / data.total_volume).toFixed(1)) : 0,
                product_count: data.product_count,
                by_product: data.by_product,
                by_location: data.by_location,
                uploaded_at: new Date().toISOString()
            };

            if (existingIdx >= 0) {
                usageData[existingIdx] = record;
                results.updated++;
            } else {
                usageData.push(record);
                results.added++;
            }
        });

        usageData.sort((a, b) => b.date.localeCompare(a.date));
        saveData(usageFile, usageData);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Usage/Wastage API endpoints
app.get('/api/usage/daily', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = usageData.filter(d => new Date(d.date) >= cutoff);
    res.json(filtered.sort((a, b) => a.date.localeCompare(b.date)));
});

app.get('/api/usage/top-waste', (req, res) => {
    const { days = 30, limit = 20 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = usageData.filter(d => new Date(d.date) >= cutoff);

    // Aggregate by product
    const byProduct = {};
    filtered.forEach(d => {
        if (d.by_product) {
            Object.entries(d.by_product).forEach(([product, stats]) => {
                if (!byProduct[product]) {
                    byProduct[product] = { used: 0, waste: 0, count: 0, ndc: stats.ndc };
                }
                byProduct[product].used += stats.used;
                byProduct[product].waste += stats.waste;
                byProduct[product].count += stats.count;
            });
        }
    });

    const result = Object.entries(byProduct)
        .map(([product, stats]) => ({
            product,
            ndc: stats.ndc,
            used_ml: parseFloat(stats.used.toFixed(1)),
            waste_ml: parseFloat(stats.waste.toFixed(1)),
            total_ml: parseFloat((stats.used + stats.waste).toFixed(1)),
            waste_percent: (stats.used + stats.waste) > 0 ? parseFloat((stats.waste * 100 / (stats.used + stats.waste)).toFixed(1)) : 0,
            count: stats.count
        }))
        .sort((a, b) => b.waste_ml - a.waste_ml)
        .slice(0, parseInt(limit));

    res.json(result);
});

app.get('/api/usage/by-location', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = usageData.filter(d => new Date(d.date) >= cutoff);

    // Aggregate by location
    const byLocation = {};
    filtered.forEach(d => {
        if (d.by_location) {
            Object.entries(d.by_location).forEach(([location, stats]) => {
                if (!byLocation[location]) {
                    byLocation[location] = { used: 0, waste: 0, count: 0 };
                }
                byLocation[location].used += stats.used;
                byLocation[location].waste += stats.waste;
                byLocation[location].count += stats.count;
            });
        }
    });

    const result = Object.entries(byLocation)
        .map(([location, stats]) => ({
            location,
            used_ml: parseFloat(stats.used.toFixed(1)),
            waste_ml: parseFloat(stats.waste.toFixed(1)),
            waste_percent: (stats.used + stats.waste) > 0 ? parseFloat((stats.waste * 100 / (stats.used + stats.waste)).toFixed(1)) : 0,
            count: stats.count
        }))
        .sort((a, b) => b.waste_ml - a.waste_ml);

    res.json(result);
});

app.get('/api/usage/summary', (req, res) => {
    const { days = 30 } = req.query;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const filtered = usageData.filter(d => new Date(d.date) >= cutoff);

    const totalVolume = filtered.reduce((sum, d) => sum + d.total_volume, 0);
    const usedVolume = filtered.reduce((sum, d) => sum + d.used_volume, 0);
    const wasteVolume = filtered.reduce((sum, d) => sum + d.waste_volume, 0);
    const productCount = filtered.reduce((sum, d) => sum + d.product_count, 0);

    res.json({
        total_volume_ml: parseFloat(totalVolume.toFixed(1)),
        used_volume_ml: parseFloat(usedVolume.toFixed(1)),
        waste_volume_ml: parseFloat(wasteVolume.toFixed(1)),
        waste_percent: totalVolume > 0 ? parseFloat((wasteVolume * 100 / totalVolume).toFixed(1)) : 0,
        product_uses: productCount,
        days_tracked: filtered.length
    });
});

// ============ AGGREGATE PRODUCT USAGE/WASTAGE ============

// Upload aggregate Product Usage report (by location)
app.post('/api/upload/product-usage', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet);

        const results = { products: 0, locations: 0, total_product_uses: 0, total_doses: 0 };

        // Aggregate by product and location
        const byProduct = {};
        const byLocation = {};

        rawData.forEach(row => {
            const name = row.Name || row.name || '';
            const ndc = row.NDCcode || row.NDC || '';
            const location = row['Location Name'] || row.Location || 'Unknown';
            const productCount = parseInt(row.ProductCount) || 0;
            const doseCount = parseInt(row.DoseCount) || 0;

            if (!name) return;

            // By product
            if (!byProduct[name]) {
                byProduct[name] = { ndc, productCount: 0, doseCount: 0, locations: {} };
            }
            byProduct[name].productCount += productCount;
            byProduct[name].doseCount += doseCount;
            byProduct[name].locations[location] = (byProduct[name].locations[location] || 0) + productCount;

            // By location
            if (!byLocation[location]) {
                byLocation[location] = { productCount: 0, doseCount: 0, products: {} };
            }
            byLocation[location].productCount += productCount;
            byLocation[location].doseCount += doseCount;
            byLocation[location].products[name] = (byLocation[location].products[name] || 0) + productCount;

            results.total_product_uses += productCount;
            results.total_doses += doseCount;
        });

        results.products = Object.keys(byProduct).length;
        results.locations = Object.keys(byLocation).length;

        // Save data
        productUsage = {
            by_product: byProduct,
            by_location: byLocation,
            summary: {
                total_products: results.products,
                total_locations: results.locations,
                total_product_uses: results.total_product_uses,
                total_doses: results.total_doses
            },
            uploaded_at: new Date().toISOString()
        };
        saveData(productUsageFile, productUsage);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Upload aggregate Product Wastage report
app.post('/api/upload/product-wastage', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet);

        const results = { products: 0, total_waste_ml: 0, total_waste_dollars: 0 };

        const products = [];
        const byType = {};

        rawData.forEach(row => {
            const name = row.Name || row.name || '';
            const ndc = row.NDCcode || row.NDC || '';
            const type = row.Type || 'Unknown';
            const productSize = parseFloat(row['Product Size(ml)']) || 0;
            const productCost = parseFloat(row['Product Cost(Dollar Amount)']) || 0;
            const partialCount = parseInt(row['Partial Products Count']) || 0;
            const wasteML = parseFloat(row['Total Wastage(mL)']) || 0;
            const wasteDollars = parseFloat(row['Total Wastage(Dollar Amount)']) || 0;

            if (!name) return;

            products.push({
                name,
                ndc,
                type,
                product_size_ml: productSize,
                product_cost: productCost,
                partial_count: partialCount,
                waste_ml: parseFloat(wasteML.toFixed(2)),
                waste_dollars: parseFloat(wasteDollars.toFixed(2)),
                waste_percent: productSize > 0 && partialCount > 0
                    ? parseFloat((wasteML * 100 / (productSize * partialCount)).toFixed(1))
                    : 0
            });

            // By type
            if (!byType[type]) {
                byType[type] = { count: 0, waste_ml: 0, waste_dollars: 0 };
            }
            byType[type].count++;
            byType[type].waste_ml += wasteML;
            byType[type].waste_dollars += wasteDollars;

            results.total_waste_ml += wasteML;
            results.total_waste_dollars += wasteDollars;
        });

        results.products = products.length;

        // Sort by waste volume
        products.sort((a, b) => b.waste_ml - a.waste_ml);

        // Save data
        productWastage = {
            products,
            by_type: byType,
            summary: {
                total_products: results.products,
                total_waste_ml: parseFloat(results.total_waste_ml.toFixed(1)),
                total_waste_dollars: parseFloat(results.total_waste_dollars.toFixed(2)),
                total_partial_products: products.reduce((sum, p) => sum + p.partial_count, 0)
            },
            uploaded_at: new Date().toISOString()
        };
        saveData(productWastageFile, productWastage);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// API endpoints for aggregate data
app.get('/api/product-usage/summary', (req, res) => {
    if (!productUsage.summary) {
        return res.json({ total_products: 0, total_locations: 0, total_product_uses: 0, total_doses: 0 });
    }
    res.json(productUsage.summary);
});

app.get('/api/product-usage/by-location', (req, res) => {
    if (!productUsage.by_location) {
        return res.json([]);
    }
    const result = Object.entries(productUsage.by_location)
        .filter(([loc]) => loc && loc !== 'undefined')
        .map(([location, data]) => ({
            location,
            product_count: data.productCount,
            dose_count: data.doseCount,
            unique_products: Object.keys(data.products).length
        }))
        .sort((a, b) => b.product_count - a.product_count);
    res.json(result);
});

app.get('/api/product-usage/top-products', (req, res) => {
    const { limit = 20 } = req.query;
    if (!productUsage.by_product) {
        return res.json([]);
    }
    const result = Object.entries(productUsage.by_product)
        .map(([name, data]) => ({
            name,
            ndc: data.ndc,
            product_count: data.productCount,
            dose_count: data.doseCount,
            locations_used: Object.keys(data.locations).length
        }))
        .sort((a, b) => b.product_count - a.product_count)
        .slice(0, parseInt(limit));
    res.json(result);
});

app.get('/api/product-wastage/summary', (req, res) => {
    if (!productWastage.summary) {
        return res.json({ total_products: 0, total_waste_ml: 0, total_waste_dollars: 0, total_partial_products: 0 });
    }
    res.json(productWastage.summary);
});

app.get('/api/product-wastage/top-waste', (req, res) => {
    const { limit = 20, sortBy = 'waste_ml' } = req.query;
    if (!productWastage.products) {
        return res.json([]);
    }
    const sorted = [...productWastage.products].sort((a, b) => {
        if (sortBy === 'waste_dollars') return b.waste_dollars - a.waste_dollars;
        if (sortBy === 'waste_percent') return b.waste_percent - a.waste_percent;
        return b.waste_ml - a.waste_ml;
    });
    res.json(sorted.slice(0, parseInt(limit)));
});

app.get('/api/product-wastage/by-type', (req, res) => {
    if (!productWastage.by_type) {
        return res.json([]);
    }
    const result = Object.entries(productWastage.by_type)
        .map(([type, data]) => ({
            type,
            product_count: data.count,
            waste_ml: parseFloat(data.waste_ml.toFixed(1)),
            waste_dollars: parseFloat(data.waste_dollars.toFixed(2))
        }))
        .sort((a, b) => b.waste_ml - a.waste_ml);
    res.json(result);
});

// ============ DETAILED WASTAGE (with dates) ============

// Upload detailed usage/wastage report (has dates, no patient names)
app.post('/api/upload/detailed-wastage', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet);

        const results = { records: 0, dates: 0, total_waste_ml: 0 };

        // Group by date
        const byDate = {};
        const byProduct = {};
        const byLocation = {};

        rawData.forEach(row => {
            const prepTimeSerial = row['Dose Preparation Time'];
            if (!prepTimeSerial) return;

            const date = excelDateToISO(prepTimeSerial);
            const location = row['Patient Location'] || 'Unknown';
            const productName = row['Product Name'] || '';
            const productNDC = row['Product NDC'] || '';
            const totalVolume = parseFloat(row['Product Total Volume']) || 0;
            const unusedVolume = parseFloat(row['Product Unused Volume']) || 0;
            const isMultiDose = row['Multi-Dose Product'] === 'Yes';
            const doseDesc = row['Dose Description'] || '';
            const budTime = row['Product BUD Time'] ? excelDateToISO(row['Product BUD Time']) : null;

            // Determine if this is a stock solution
            const isStock = doseDesc.toLowerCase().includes('stock') ||
                           productName.toLowerCase().includes('stock');

            const wasteVolume = unusedVolume;
            const usedVolume = totalVolume - unusedVolume;

            results.records++;
            results.total_waste_ml += wasteVolume;

            // By date
            if (!byDate[date]) {
                byDate[date] = {
                    total_volume: 0,
                    used_volume: 0,
                    waste_volume: 0,
                    product_count: 0,
                    multi_dose_count: 0,
                    stock_count: 0
                };
            }
            byDate[date].total_volume += totalVolume;
            byDate[date].used_volume += usedVolume;
            byDate[date].waste_volume += wasteVolume;
            byDate[date].product_count++;
            if (isMultiDose) byDate[date].multi_dose_count++;
            if (isStock) byDate[date].stock_count++;

            // By product
            if (!byProduct[productName]) {
                byProduct[productName] = {
                    ndc: productNDC,
                    total_volume: 0,
                    used_volume: 0,
                    waste_volume: 0,
                    count: 0,
                    is_multi_dose: isMultiDose,
                    is_stock: isStock
                };
            }
            byProduct[productName].total_volume += totalVolume;
            byProduct[productName].used_volume += usedVolume;
            byProduct[productName].waste_volume += wasteVolume;
            byProduct[productName].count++;

            // By location
            const locKey = location.split('-')[0];
            if (!byLocation[locKey]) {
                byLocation[locKey] = { total_volume: 0, used_volume: 0, waste_volume: 0, count: 0 };
            }
            byLocation[locKey].total_volume += totalVolume;
            byLocation[locKey].used_volume += usedVolume;
            byLocation[locKey].waste_volume += wasteVolume;
            byLocation[locKey].count++;
        });

        results.dates = Object.keys(byDate).length;

        // Save data
        detailedWastage = {
            by_date: byDate,
            by_product: byProduct,
            by_location: byLocation,
            summary: {
                total_records: results.records,
                total_dates: results.dates,
                total_waste_ml: parseFloat(results.total_waste_ml.toFixed(1)),
                date_range: {
                    start: Object.keys(byDate).sort()[0],
                    end: Object.keys(byDate).sort().pop()
                }
            },
            uploaded_at: new Date().toISOString()
        };
        saveData(detailedWastageFile, detailedWastage);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Detailed wastage API endpoints
app.get('/api/detailed-wastage/summary', (req, res) => {
    if (!detailedWastage.summary) {
        return res.json({ total_records: 0, total_dates: 0, total_waste_ml: 0 });
    }
    res.json(detailedWastage.summary);
});

app.get('/api/detailed-wastage/daily', (req, res) => {
    const { days = 30 } = req.query;
    if (!detailedWastage.by_date) {
        return res.json([]);
    }

    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - parseInt(days));

    const result = Object.entries(detailedWastage.by_date)
        .filter(([date]) => new Date(date) >= cutoff)
        .map(([date, data]) => ({
            date,
            total_volume: parseFloat(data.total_volume.toFixed(1)),
            used_volume: parseFloat(data.used_volume.toFixed(1)),
            waste_volume: parseFloat(data.waste_volume.toFixed(1)),
            waste_percent: data.total_volume > 0
                ? parseFloat((data.waste_volume * 100 / data.total_volume).toFixed(1))
                : 0,
            product_count: data.product_count,
            multi_dose_count: data.multi_dose_count,
            stock_count: data.stock_count
        }))
        .sort((a, b) => a.date.localeCompare(b.date));

    res.json(result);
});

app.get('/api/detailed-wastage/by-product', (req, res) => {
    const { limit = 25, sortBy = 'waste_volume' } = req.query;
    if (!detailedWastage.by_product) {
        return res.json([]);
    }

    const result = Object.entries(detailedWastage.by_product)
        .map(([name, data]) => ({
            name,
            ndc: data.ndc,
            total_volume: parseFloat(data.total_volume.toFixed(1)),
            used_volume: parseFloat(data.used_volume.toFixed(1)),
            waste_volume: parseFloat(data.waste_volume.toFixed(1)),
            waste_percent: data.total_volume > 0
                ? parseFloat((data.waste_volume * 100 / data.total_volume).toFixed(1))
                : 0,
            count: data.count,
            is_multi_dose: data.is_multi_dose,
            is_stock: data.is_stock
        }))
        .sort((a, b) => {
            if (sortBy === 'waste_percent') return b.waste_percent - a.waste_percent;
            if (sortBy === 'count') return b.count - a.count;
            return b.waste_volume - a.waste_volume;
        })
        .slice(0, parseInt(limit));

    res.json(result);
});

app.get('/api/detailed-wastage/by-location', (req, res) => {
    if (!detailedWastage.by_location) {
        return res.json([]);
    }

    const result = Object.entries(detailedWastage.by_location)
        .map(([location, data]) => ({
            location,
            total_volume: parseFloat(data.total_volume.toFixed(1)),
            used_volume: parseFloat(data.used_volume.toFixed(1)),
            waste_volume: parseFloat(data.waste_volume.toFixed(1)),
            waste_percent: data.total_volume > 0
                ? parseFloat((data.waste_volume * 100 / data.total_volume).toFixed(1))
                : 0,
            count: data.count
        }))
        .sort((a, b) => b.waste_volume - a.waste_volume);

    res.json(result);
});

// ============ STOCK DOSES ============

// Upload Completed Stock and Dilution Doses report
app.post('/api/upload/stock-doses', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet);

        const results = { stocks: 0, dilutions: 0, total_doses: 0 };

        const stocks = [];
        const dilutions = [];

        rawData.forEach(row => {
            const dose = row.Dose || row.dose || '';
            const total = parseInt(row.Total) || 0;

            if (!dose) return;

            results.total_doses += total;

            // Parse the dose string to extract details
            // Format: "DILUTION: BAXA - ASCORBIC ACID 100MG/ML DILUTION IN SW - 25ML"
            // or "STOCK: ..."
            const isDilution = dose.toUpperCase().startsWith('DILUTION:');
            const isStock = dose.toUpperCase().startsWith('STOCK:');

            // Extract drug name and size
            let drugName = dose;
            let size = '';
            const sizeMatch = dose.match(/(\d+\.?\d*)\s*(ML|MG|MCG|UNITS?)/i);
            if (sizeMatch) {
                size = sizeMatch[0];
            }

            // Clean up drug name
            drugName = dose
                .replace(/^(DILUTION|STOCK):\s*/i, '')
                .replace(/BAXA\s*-\s*/i, '')
                .trim();

            const record = {
                name: drugName,
                original: dose,
                total,
                size,
                type: isDilution ? 'Dilution' : isStock ? 'Stock' : 'Other'
            };

            if (isDilution) {
                dilutions.push(record);
                results.dilutions++;
            } else {
                stocks.push(record);
                results.stocks++;
            }
        });

        // Sort by total
        stocks.sort((a, b) => b.total - a.total);
        dilutions.sort((a, b) => b.total - a.total);

        // Save data
        stockDoses = {
            stocks,
            dilutions,
            all: [...stocks, ...dilutions].sort((a, b) => b.total - a.total),
            summary: {
                total_stock_types: results.stocks,
                total_dilution_types: results.dilutions,
                total_doses_made: results.total_doses
            },
            uploaded_at: new Date().toISOString()
        };
        saveData(stockDosesFile, stockDoses);
        fs.unlinkSync(req.file.path);

        res.json({ success: true, results });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Stock doses API endpoints
app.get('/api/stock-doses/summary', (req, res) => {
    if (!stockDoses.summary) {
        return res.json({ total_stock_types: 0, total_dilution_types: 0, total_doses_made: 0 });
    }
    res.json(stockDoses.summary);
});

app.get('/api/stock-doses/all', (req, res) => {
    const { limit = 50, type = 'all' } = req.query;
    if (!stockDoses.all) {
        return res.json([]);
    }

    let data = stockDoses.all;
    if (type === 'stock') data = stockDoses.stocks || [];
    if (type === 'dilution') data = stockDoses.dilutions || [];

    res.json(data.slice(0, parseInt(limit)));
});

app.get('/api/stock-doses/top', (req, res) => {
    const { limit = 20 } = req.query;
    if (!stockDoses.all) {
        return res.json([]);
    }
    res.json(stockDoses.all.slice(0, parseInt(limit)));
});

// Clear all data
app.post('/api/clear', (req, res) => {
    productionData = [];
    turnaroundData = [];
    bypassData = [];
    usageData = [];

    saveData(productionFile, productionData);
    saveData(turnaroundFile, turnaroundData);
    saveData(bypassFile, bypassData);
    saveData(usageFile, usageData);

    res.json({ success: true, message: 'All data cleared' });
});

// Helper functions
function formatHour(hour) {
    if (hour === 0) return '12 AM';
    if (hour === 12) return '12 PM';
    return hour < 12 ? `${hour} AM` : `${hour - 12} PM`;
}

function getPriorityLabel(priority) {
    const labels = {
        10: 'STAT',
        20: 'ASAP',
        30: 'Urgent',
        40: 'Routine',
        45: 'Standard',
        50: 'Low',
        55: 'Batch'
    };
    return labels[priority] || `Priority ${priority}`;
}

// Start server
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`
============================================================
  DoseEdge Analytics Dashboard
============================================================

  Dashboard is running at: http://localhost:${PORT}

============================================================
`);
});
