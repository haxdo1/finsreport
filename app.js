let dashboardData = [];
const charts = { trend: null, category: null, division: null, personnel: null };

const ExcelInput = document.getElementById('excelInput');
const YearSelector = document.getElementById('yearSelector');
const MonthSelector = document.getElementById('monthSelector');
const TypeSelector = document.getElementById('typeSelector');
const DataSourceSelector = document.getElementById('dataSourceSelector');
const SearchInput = document.getElementById('searchInput'); // New
const UploadBtnLabel = document.getElementById('uploadBtnLabel');

// --- Core Event Listeners ---
ExcelInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const raw = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        processData(raw);
    };
    reader.readAsArrayBuffer(file);
});

DataSourceSelector.addEventListener('change', (e) => {
    const value = e.target.value;
    if (value === 'upload') {
        UploadBtnLabel.classList.remove('hidden');
    } else if (value === 'combined') {
        UploadBtnLabel.classList.add('hidden');
        loadCombinedData();
    } else {
        UploadBtnLabel.classList.add('hidden');
        loadLocalFile(value);
    }
});

function loadCombinedData() {
    document.getElementById('insightsText').innerText = "Loading combined data...";

    Promise.all([
        fetch('Proposal.xlsx').then(res => { if (!res.ok) throw new Error("Proposal.xlsx missing"); return res.arrayBuffer(); }),
        fetch('Reimbursement.xlsx').then(res => { if (!res.ok) throw new Error("Reimbursement.xlsx missing"); return res.arrayBuffer(); })
    ]).then(([propData, reimbData]) => {
        const wb1 = XLSX.read(propData, { type: 'array' });
        const raw1 = XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]]);

        const wb2 = XLSX.read(reimbData, { type: 'array' });
        const raw2 = XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]);

        const combined = [...raw1, ...raw2];
        processData(combined);
    }).catch(err => {
        console.error("Error loading combined files:", err);
        document.getElementById('insightsText').innerText = "Failed to load combined data.";
    });
}

function loadLocalFile(url) {
    // Show loading state (optional, but good for UX)
    const originalText = document.getElementById('insightsText').innerText;
    document.getElementById('insightsText').innerText = "Loading data...";

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            if (!workbook.SheetNames.length) throw new Error("Excel file is empty or has no sheets.");

            const firstSheetName = workbook.SheetNames[0];
            const raw = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);

            if (!raw || raw.length === 0) throw new Error("Sheet is empty.");

            processData(raw);
        })
        .catch(err => {
            console.error("Error loading local file:", err);
            alert(`Error loading file: ${err.message}\n\nNote: If you are opening index.html directly (file://), you MUST use a local server (like XAMPP, Live Server, or 'python -m http.server'). content fetch detection blocked by browser security.`);
            document.getElementById('insightsText').innerText = "Failed to load data.";
        });
}

function processData(rawData) {
    dashboardData = rawData.map(row => {
        const cleanRow = {};
        Object.keys(row).forEach(key => cleanRow[key.trim()] = row[key]);
        const date = parseDate(cleanRow['Date & Time'] || cleanRow['Date']);
        return {
            date, year: date?.getFullYear(), month: date?.getMonth(),
            issuedBy: cleanRow['Issued By'] || 'Unknown',
            action: (cleanRow['Action'] || 'Unknown').trim(),
            balance: parseCurrency(cleanRow['Balance']),
            reason: cleanRow['Reason'] || '',
            category: categorize(cleanRow['Reason'] || ''),
            division: extractDiv(cleanRow['Reason'] || '')
        };
    }).filter(d => d.year != null);

    // Sort simply for processing (though renderTable does its own sort)
    // We need time-based sort for the pairing logic to work (Mistake -> Correction)
    dashboardData.sort((a, b) => {
        const da = parseDate(a.date);
        const db = parseDate(b.date);
        return da - db;
    });

    excludeCorrectionPairs(dashboardData);

    renderSelectors();
    updateUI();
}

function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date((val - 25569) * 86400 * 1000); // Excel serial date

    // Ensure val is a string before splitting
    const str = String(val).trim();
    if (!str) return null;

    // Try standard date constructor first
    const d = new Date(str);
    if (!isNaN(d.getTime()) && str.includes('-')) return d; // Prefer standard if ISO format

    // Custom parsing for DD/MM/YY HH:MM (e.g., 10/2/26 22:18)
    const parts = str.split(' ');
    const datePart = parts[0];
    const timePart = parts[1] || '00:00';

    if (datePart.includes('/')) {
        const dParts = datePart.split('/');
        const tParts = timePart.split(':');

        if (dParts.length === 3) {
            let day = parseInt(dParts[0]);
            let month = parseInt(dParts[1]) - 1;
            let year = parseInt(dParts[2]);
            if (year < 100) year += 2000;

            return new Date(year, month, day, parseInt(tParts[0] || 0), parseInt(tParts[1] || 0));
        }
    }

    return d; // Fallback to standard date
}

function parseCurrency(val) {
    if (typeof val === 'number') return val;
    return parseFloat(val?.toString().replace(/[^0-9.-]+/g, "") || 0);
}

function categorize(r) {
    r = r.toLowerCase();
    if (r.includes('salary')) return 'Salaries';
    if (r.includes('proposal')) return 'Proposals';
    if (r.includes('staff of the month')) return 'Bonuses';
    if (r.includes('reimbursement')) return 'Reimbursements';
    if (r.includes('bank') || r.includes('transfer')) return 'Internal Transfers';
    if (r.includes('severance')) return 'Severance Pay';
    if (['miss', '-', 'miss '].includes(r.trim())) return 'Correction/Misc';
    return 'Miscellaneous';
}

function extractDiv(r) {
    r = r.replace(/\[PROPOSAL\]/gi, '').trim().toLowerCase();
    if (r.includes('bnmc')) return 'BNMC';
    if (r.includes('ens')) return 'EnS';
    if (r.includes('finan')) return 'Financial Regulator';
    if (r.includes('cc ')) return 'CC';
    if (r.includes('hrd') || r.includes('re-training')) return 'HRD';
    if (r.includes('reimbursement')) return 'Reimbursement Team';
    if (r.includes('bank') || r.includes('transfer')) return 'Internal Transfer';
    const m = r.match(/ - (.*?) (salary|staff)/i);
    return m ? m[1].trim().toUpperCase() : 'Internal/Misc';
}

function renderSelectors() {
    const years = [...new Set(dashboardData.map(d => d.year))].sort((a, b) => b - a);
    YearSelector.innerHTML = '<option value="all">All Time</option>' + years.map(y => `<option value="${y}">${y}</option>`).join('');
    const mNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    MonthSelector.innerHTML = '<option value="all">All Months</option>' + mNames.map((n, i) => `<option value="${i}">${n}</option>`).join('');

    YearSelector.onchange = MonthSelector.onchange = TypeSelector.onchange = SearchInput.oninput = updateUI; // Add Search listener

    ['emptyState', 'insightsSection', 'statsSection', 'closingSection', 'chartsSection', 'filterSection', 'reportSection'].forEach(id => {
        document.getElementById(id).classList.toggle('hidden', id === 'emptyState');
    });
}

function isCorrection(r) {
    if (!r) return false;
    // Match whole words only to avoid false positives (e.g., 'contest' shouldn't match 'test')
    const keywords = ['correction', 'mistake', 'error', 'double', 'wrong', 'test', 'refund', 'return', 'void'];
    const regex = new RegExp(`\\b(${keywords.join('|')})\\b`, 'i');

    if (regex.test(r)) return true;
    return ['miss', '-', 'miss '].includes(r.trim().toLowerCase());
    if (regex.test(r)) return true;
    return ['miss', '-', 'miss '].includes(r.trim().toLowerCase());
}

function excludeCorrectionPairs(data) {
    // Reset exclusion state
    data.forEach(d => d.isExcluded = false);

    // Iterate to find corrections and pair them with previous mistakes
    for (let i = 0; i < data.length; i++) {
        const current = data[i];

        // If this is identified as a correction (e.g. "miss", "error", "void")
        if (isCorrection(current.reason) && !current.isExcluded) {

            // Look backwards for the matching "Mistake"
            for (let j = i - 1; j >= 0; j--) {
                const candidate = data[j];

                // Skip if already excluded or too far back (optional optimization, but we'll scan all for now)
                if (candidate.isExcluded) continue;

                // Match Criteria:
                // 1. Same Person
                // 2. Same Amount
                // 3. Opposite Action (Deposit vs Withdraw)
                const isSamePerson = candidate.issuedBy === current.issuedBy;
                const isSameAmount = Math.abs(candidate.balance - current.balance) < 0.01;
                const isOppositeAction = candidate.action.toLowerCase() !== current.action.toLowerCase();

                if (isSamePerson && isSameAmount && isOppositeAction) {
                    // FOUND IT!
                    // Mark both as excluded
                    current.isExcluded = true;
                    candidate.isExcluded = true;

                    // We found the pair, stop looking for this specific correction
                    break;
                }
            }

            // If no pair found, we still exclude the correction itself (single-sided fix? or just a note)
            // The previous logic did this, so we keep it implies "this transaction is a correction so don't count it"
            current.isExcluded = true;
        }
    }
}

function updateUI() {
    const year = YearSelector.value === 'all' ? 'all' : parseInt(YearSelector.value);
    const month = MonthSelector.value === 'all' ? 'all' : parseInt(MonthSelector.value);
    const type = TypeSelector.value;
    const searchQuery = SearchInput.value.toLowerCase(); // Get search query

    let filtered = dashboardData;
    if (year !== 'all') filtered = filtered.filter(d => d.year === year);
    if (month !== 'all') filtered = filtered.filter(d => d.month === month);
    if (type !== 'all') filtered = filtered.filter(d => d.action.toLowerCase() === type);

    // Apply Search Filter
    if (searchQuery) {
        filtered = filtered.filter(d =>
            d.reason.toLowerCase().includes(searchQuery) ||
            d.issuedBy.toLowerCase().includes(searchQuery)
        );
    }

    // Exclude corrections/errors from all stats and views
    // Exclude corrections/errors from all stats and views
    // Now uses the smart pairing flag 'isExcluded'
    filtered = filtered.filter(d => !d.isExcluded);

    const dep = filtered.filter(d => d.action.toLowerCase() === 'deposit').reduce((s, d) => s + d.balance, 0);
    const wit = filtered.filter(d => d.action.toLowerCase() === 'withdraw').reduce((s, d) => s + d.balance, 0);
    const net = dep - wit;

    document.getElementById('totalDeposit').textContent = fmt(dep);
    document.getElementById('totalWithdraw').textContent = fmt(wit);
    document.getElementById('netBalance').textContent = fmt(net);
    document.getElementById('netBalance').style.color = net >= 0 ? '#10b981' : '#ef4444';
    document.getElementById('transCount').textContent = `${filtered.length} Records`;

    renderClosingSummary(net, year, month);
    renderTable(filtered);
    renderCharts(filtered);
    renderInsights(filtered);
}

function renderClosingSummary(net, year, month) {
    const section = document.getElementById('closingSection');
    const iconContainer = document.getElementById('closingIconStatus');
    const icon = document.getElementById('closingIcon');
    const statusText = document.getElementById('closingStatusText');
    const instruction = document.getElementById('closingInstruction');
    const instructionText = document.getElementById('closingInstructionText');
    const copyBtn = document.getElementById('copyClosingInstructionBtn');
    const balancedBadge = document.getElementById('balancedBadge');

    section.classList.remove('hidden');

    if (Math.abs(net) < 0.01) {
        // Balanced
        iconContainer.className = 'w-14 h-14 rounded-2xl bg-green-500/20 flex items-center justify-center text-green-500';
        icon.setAttribute('data-lucide', 'check-circle-2');
        statusText.innerHTML = '<span class="text-green-400 font-bold">Laporan Seimbang</span> • Periode sudah ditutup (Nol).';
        instruction.classList.add('hidden');
        copyBtn.classList.add('hidden');
        balancedBadge.classList.remove('hidden');
    } else {
        // Unbalanced / Open
        const isDeficit = net < 0;
        const actionNeeded = isDeficit ? 'DEPOSIT' : 'WITHDRAW';
        const amtNeeded = fmt(Math.abs(net));
        const statusLabel = isDeficit ? 'KEKURANGAN (Deficit)' : 'SISA SALDO (Surplus)';
        const statusColor = isDeficit ? 'text-red-400' : 'text-accent-blue';
        const iconName = isDeficit ? 'alert-circle' : 'wallet';

        let periodName = "";
        if (year === 'all') {
            periodName = "All Time";
        } else if (month === 'all') {
            periodName = `Yearly ${year}`;
        } else {
            const monthText = MonthSelector.options[MonthSelector.selectedIndex].text;
            periodName = `${monthText} ${year}`;
        }

        iconContainer.className = `w-14 h-14 rounded-2xl bg-white/5 flex items-center justify-center ${statusColor}`;
        icon.setAttribute('data-lucide', iconName);
        statusText.innerHTML = `<span class="${statusColor} font-bold">${statusLabel}</span> • ${fmt(net)}`;

        instruction.classList.remove('hidden');
        instructionText.textContent = `Untuk menolkan saldo, catat transaksi ${actionNeeded} sebesar ${amtNeeded} dengan alasan: "[CLOSE] Tutup Buku ${periodName}".`;

        copyBtn.classList.remove('hidden');
        balancedBadge.classList.add('hidden');

        copyBtn.onclick = () => {
            const text = `[CLOSING STATEMENT] ${periodName}\nNominal: ${amtNeeded}\nTindakan: ${actionNeeded}\nAlasan: [CLOSE] Tutup Buku ${periodName}`;
            navigator.clipboard.writeText(text);
            const originalHTML = copyBtn.innerHTML;
            copyBtn.innerHTML = '<i data-lucide="check" class="w-4 h-4"></i> Berhasil Disalin';
            lucide.createIcons();
            setTimeout(() => { copyBtn.innerHTML = originalHTML; lucide.createIcons(); }, 2000);
        };
    }
    lucide.createIcons();
}

function renderTable(data) {
    // Sort oldest to newest for accurate running balance
    data.sort((a, b) => {
        const da = parseDate(a.date);
        const db = parseDate(b.date);
        return da - db;
    });

    let runningBalance = 0;

    document.getElementById('reportTableBody').innerHTML = data.map(r => {
        const isDeposit = r.action.toLowerCase() === 'deposit';
        const debit = isDeposit ? r.balance : 0;
        const credit = !isDeposit ? r.balance : 0;

        if (isDeposit) runningBalance += r.balance;
        else runningBalance -= r.balance;

        return `
        <tr class="hover:bg-white/5 transition-colors group">
            <td class="px-6 py-4 font-mono text-xs text-slate-400 whitespace-nowrap">
                ${r.date instanceof Date ? r.date.toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : r.date}
            </td>
            <td class="px-6 py-4">
                <div class="flex items-center gap-2">
                    <div class="w-6 h-6 rounded-full bg-white/5 flex items-center justify-center text-[10px] font-bold text-slate-400 group-hover:bg-accent-blue group-hover:text-white transition-colors">
                        ${r.issuedBy.charAt(0)}
                    </div>
                    <span class="font-medium text-slate-200">${r.issuedBy}</span>
                </div>
            </td>
            <td class="px-6 py-4 text-right font-mono text-green-400">${debit > 0 ? fmt(debit) : '-'}</td>
            <td class="px-6 py-4 text-right font-mono text-red-400">${credit > 0 ? fmt(credit) : '-'}</td>
            <td class="px-6 py-4 text-right font-bold text-white font-mono bg-white/2">${fmt(runningBalance)}</td>
            <td class="px-6 py-4 text-slate-400 text-xs break-words" title="${r.reason}">
                ${extractDiv(r.reason) !== 'Internal/Misc' ? `<span class="px-2 py-0.5 rounded-md bg-white/5 text-white/50 text-[10px] uppercase font-bold mr-2 border border-white/5 align-middle">${extractDiv(r.reason)}</span>` : ''}
                <span class="align-middle leading-relaxed">${r.reason}</span>
            </td>
        </tr>
    `}).join('');
}

function renderCharts(data) {
    const isAll = YearSelector.value === 'all';
    const withdrawals = data.filter(d => d.action.toLowerCase() === 'withdraw');

    renderTrend(data, isAll);
    renderDoughnut('categoryChart', 'category', withdrawals, 'category'); // Use withdrawals only
    renderBar('divisionChart', 'division', withdrawals, 'division', 'rgba(236, 72, 153, 0.4)', '#ec4899'); // Spending only
    renderBar('personnelChart', 'personnel', data, 'issuedBy', 'rgba(59, 130, 246, 0.5)', '#3b82f6'); // Operations count (All)
}

function renderTrend(data, isAll) {
    const ctx = document.getElementById('trendChart').getContext('2d');
    let labels, dData, wData;
    if (isAll) {
        labels = [...new Set(dashboardData.map(d => d.year))].sort();
        dData = labels.map(y => data.filter(d => d.year === y && d.action.toLowerCase() === 'deposit').reduce((s, d) => s + d.balance, 0));
        wData = labels.map(y => data.filter(d => d.year === y && d.action.toLowerCase() === 'withdraw').reduce((s, d) => s + d.balance, 0));
    } else {
        labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        dData = labels.map((_, i) => data.filter(d => d.month === i && d.action.toLowerCase() === 'deposit').reduce((s, d) => s + d.balance, 0));
        wData = labels.map((_, i) => data.filter(d => d.month === i && d.action.toLowerCase() === 'withdraw').reduce((s, d) => s + d.balance, 0));
    }
    if (charts.trend) charts.trend.destroy();
    charts.trend = new Chart(ctx, {
        type: 'line',
        data: {
            labels, datasets: [
                { label: 'Deposit', data: dData, borderColor: '#10b981', tension: 0.4 },
                { label: 'Withdraw', data: wData, borderColor: '#ef4444', tension: 0.4 }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { x: { grid: { display: false } }, y: { grid: { color: 'rgba(255,255,255,0.05)' } } } }
    });
}

function renderDoughnut(canvasId, chartKey, data, key) {
    const counts = {};
    data.filter(d => d.action.toLowerCase() === 'withdraw').forEach(d => counts[d[key]] = (counts[d[key]] || 0) + d.balance);
    const ctx = document.getElementById(canvasId).getContext('2d');
    if (charts[chartKey]) charts[chartKey].destroy();
    charts[chartKey] = new Chart(ctx, {
        type: 'doughnut',
        data: { labels: Object.keys(counts), datasets: [{ data: Object.values(counts), backgroundColor: ['#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b'], borderWidth: 0 }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 10 } } } }, cutout: '70%' }
    });
}

function renderBar(canvasId, chartKey, data, key, bg, border) {
    const counts = {};
    data.forEach(d => counts[d[key]] = (counts[d[key]] || 0) + (canvasId === 'personnelChart' ? 1 : d.balance));
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const ctx = document.getElementById(canvasId).getContext('2d');
    if (charts[chartKey]) charts[chartKey].destroy();
    charts[chartKey] = new Chart(ctx, {
        type: 'bar',
        data: { labels: sorted.map(s => s[0]), datasets: [{ data: sorted.map(s => s[1]), backgroundColor: bg, borderColor: border, borderWidth: 1, borderRadius: 4 }] },
        options: { indexAxis: chartKey === 'division' ? 'y' : 'x', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
}

function renderInsights(data) {
    if (data.length === 0) return;
    const withdraws = data.filter(d => d.action.toLowerCase() === 'withdraw');
    const catCounts = {}; withdraws.forEach(d => catCounts[d.category] = (catCounts[d.category] || 0) + d.balance);
    const topCat = Object.entries(catCounts).sort((a, b) => b[1] - a[1])[0] || ['None', 0];
    const topPerson = Object.entries(data.reduce((a, c) => (a[c.issuedBy] = (a[c.issuedBy] || 0) + 1, a), {})).sort((a, b) => b[1] - a[1])[0];

    document.getElementById('insightsText').innerHTML = `
        Highest Spend: <span class="text-white font-bold">${topCat[0]}</span> (${fmt(topCat[1])}) • 
        Most Active: <span class="text-white font-bold">${topPerson?.[0]}</span> (${topPerson?.[1]} ops) • 
        Net: <span class="text-white font-bold">${fmt(data.filter(d => d.action.toLowerCase() === 'deposit').reduce((s, d) => s + d.balance, 0) - data.filter(d => d.action.toLowerCase() === 'withdraw').reduce((s, d) => s + d.balance, 0))}</span>
    `;
}

// --- Global Formatting ---
const fmt = v => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(v);

function downloadReport() {
    // 1. Get current filters
    const year = YearSelector.value === 'all' ? 'all' : parseInt(YearSelector.value);
    const month = MonthSelector.value === 'all' ? 'all' : parseInt(MonthSelector.value);
    const type = TypeSelector.value;
    const searchQuery = SearchInput.value.toLowerCase();

    // 2. Filter data (Same logic as updateUI)
    let filtered = dashboardData;
    if (year !== 'all') filtered = filtered.filter(d => d.year === year);
    if (month !== 'all') filtered = filtered.filter(d => d.month === month);
    if (type !== 'all') filtered = filtered.filter(d => d.action.toLowerCase() === type);

    if (searchQuery) {
        filtered = filtered.filter(d =>
            d.reason.toLowerCase().includes(searchQuery) ||
            d.issuedBy.toLowerCase().includes(searchQuery)
        );
    }

    // Exclude corrections
    filtered = filtered.filter(d => !d.isExcluded);

    // 3. Prepare Title and Filename helpers
    const typeLabel = (TypeSelector.options[TypeSelector.selectedIndex].text.split(' ')[0] || "REPORT").toUpperCase();
    const monthLabel = month === 'all' ? "FULL YEAR" : MonthSelector.options[MonthSelector.selectedIndex].text.toUpperCase();
    const yearLabel = year === 'all' ? "ALL TIME" : year;
    const title = `${typeLabel === 'ALL' ? 'FINANCIAL REPORT' : typeLabel} (${monthLabel} ${yearLabel})`;

    // 4. Calculate Totals
    const totalDebit = filtered.reduce((sum, d) => sum + (d.action.toLowerCase() === 'deposit' ? d.balance : 0), 0);
    const totalCredit = filtered.reduce((sum, d) => sum + (d.action.toLowerCase() === 'withdraw' ? d.balance : 0), 0);

    // 5. Build Data Array
    const data = [
        [title, "", "", "", "", ""], // Row 1: Title (Merged A1:F1)
        ["TANGGAL", "ISSUED BY", "NAMA PROPOSAL", "DEBIT (MASUK)", "KREDIT (KELUAR)", "SALDO (BALANCE)"], // Row 2: Headers
    ];

    let runningBalance = 0;

    // Add Data Rows
    filtered.forEach(d => {
        const dateStr = d.date instanceof Date
            ? d.date.toLocaleDateString('en-GB') // DD/MM/YYYY
            : d.date;

        const isDeposit = d.action.toLowerCase() === 'deposit';
        const debit = isDeposit ? d.balance : 0;
        const credit = !isDeposit ? d.balance : 0;

        if (isDeposit) runningBalance += d.balance;
        else runningBalance -= d.balance;

        data.push([
            dateStr,
            d.issuedBy, // Issuer
            d.reason,   // NAMA PROPOSAL
            debit,      // DEBIT
            credit,     // KREDIT
            runningBalance // SALDO
        ]);
    });

    // Add Footer Row
    // Total should be net balance? Or just separate totals?
    // Usually "Saldo" column bottom is the final balance.
    data.push(["", "", "TOTAL SELURUH", totalDebit, totalCredit, runningBalance]);

    // 6. Generate Sheet
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Merge Title Cell (A1:F1)
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });

    // Define Styles
    const borderStyle = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
    };

    const headerStyle = {
        font: { bold: true, sz: 12 },
        alignment: { horizontal: "center", vertical: "center" },
        border: borderStyle,
        fill: { fgColor: { rgb: "EFEFEF" } }
    };

    const titleStyle = {
        font: { bold: true, sz: 14 },
        alignment: { horizontal: "center", vertical: "center" },
        border: borderStyle
    };

    const cellStyle = {
        border: borderStyle,
        alignment: { vertical: "center" }
    };

    const currencyStyle = {
        border: borderStyle,
        alignment: { vertical: "center" },
        numFmt: "$#,##0"
    };

    const boldCurrencyStyle = {
        font: { bold: true },
        border: borderStyle,
        alignment: { vertical: "center" },
        numFmt: "$#,##0"
    };

    const boldCellStyle = {
        font: { bold: true },
        border: borderStyle,
        alignment: { vertical: "center" }
    };


    // Apply Styles to All Cells
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellRef]) continue;

            // 1. Title Row (Row 0)
            if (R === 0) {
                ws[cellRef].s = titleStyle;
            }
            // 2. Header Row (Row 1)
            else if (R === 1) {
                ws[cellRef].s = headerStyle;
            }
            // 3. Data Rows
            else {
                // Last Row (Total) Check
                const isLastRow = R === range.e.r;

                if (isLastRow) {
                    // Footer styling
                    if (C >= 3) ws[cellRef].s = boldCurrencyStyle; // Debit, Credit, Saldo Totals
                    else ws[cellRef].s = boldCellStyle; // Label
                } else {
                    // Normal Data
                    if (C >= 3) {
                        ws[cellRef].s = currencyStyle; // Debit, Credit, Saldo
                    }
                    else ws[cellRef].s = cellStyle;
                }
            }
        }
    }

    // Basic Column Widths
    ws['!cols'] = [
        { wch: 15 }, // Date
        { wch: 20 }, // Issuer
        { wch: 50 }, // Reason
        { wch: 20 }, // Debit
        { wch: 20 }, // Credit
        { wch: 20 }  // Saldo
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");

    // 7. Download
    const safeTitle = title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    XLSX.writeFile(wb, `${safeTitle}.xlsx`);
}

