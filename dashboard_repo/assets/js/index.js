const toggleBtn = document.getElementById('toggleSidebar');
const sidebar = document.getElementById('sidebar');
const icon = toggleBtn.querySelector('i');
const sidebarLogo = document.getElementById('sidebarLogo');

toggleBtn.addEventListener('click', () => {
  sidebar.classList.toggle('collapsed');
  icon.classList.toggle('bi-chevron-left');
  icon.classList.toggle('bi-chevron-right');

  sidebarLogo.src = sidebar.classList.contains('collapsed') 
    ? 'assets/img/ud_logo_short.png' 
    : 'assets/img/ud_logo_long.png';
});

// --- Date parsing & quarter calculation ---
function parseDate(value) {
  if (!value) return null;
  if (typeof value === 'number') {
    const epoch = new Date(1899, 11, 30);
    epoch.setDate(epoch.getDate() + value);
    return epoch;
  }
  const d = new Date(value);
  return isNaN(d) ? null : d;
}

function getQuarter(date) {
  return Math.floor(date.getMonth() / 3) + 1;
}

function isInQuarter(dateValue, year, quarter) {
  const date = parseDate(dateValue);
  if (!date) return false;
  return date.getFullYear() === year && getQuarter(date) === quarter;
}

// --- Color shade utilities ---
function hexToHSL(H) {
  let r = 0, g = 0, b = 0;
  if (H.length === 4) {
    r = "0x" + H[1] + H[1];
    g = "0x" + H[2] + H[2];
    b = "0x" + H[3] + H[3];
  } else if (H.length === 7) {
    r = "0x" + H[1] + H[2];
    g = "0x" + H[3] + H[4];
    b = "0x" + H[5] + H[6];
  }
  r /= 255; g /= 255; b /= 255;
  const cmin = Math.min(r, g, b), cmax = Math.max(r, g, b), delta = cmax - cmin;
  let h = 0, s = 0, l = 0;

  if (delta === 0) h = 0;
  else if (cmax === r) h = ((g - b) / delta) % 6;
  else if (cmax === g) h = (b - r) / delta + 2;
  else h = (r - g) / delta + 4;
  h = Math.round(h * 60);
  if (h < 0) h += 360;

  l = (cmax + cmin) / 2;
  s = delta === 0 ? 0 : delta / (1 - Math.abs(2 * l - 1));
  s = +(s * 100).toFixed(1);
  l = +(l * 100).toFixed(1);
  return { h, s, l };
}

function generateShades(values, baseHex) {
  const baseHSL = hexToHSL(baseHex);
  const maxVal = Math.max(...values);
  const minVal = Math.min(...values);
  return values.map(v => {
    const lightness = baseHSL.l + (maxVal - v) / (maxVal - minVal + 0.0001) * 40;
    return `hsl(${baseHSL.h}, ${baseHSL.s}%, ${lightness}%)`;
  });
}

window.addEventListener('DOMContentLoaded', () => {
  // --- SharePoint Excel URL ---
  const excelPath = 'https://uniondigitalph.sharepoint.com/sites/DataDashboardsRepository/_layouts/15/download.aspx?SourceUrl=/sites/DataDashboardsRepository/Shared%20Documents/Dashboard%20Inventory.xlsx';

  fetch(excelPath)
    .then(res => res.arrayBuffer())
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      // --- Scorecards ---
      const uniqueLinks = new Set(jsonData.map(r => r['Report Link']));
      document.getElementById('totalDashboards').innerText = uniqueLinks.size;

      const now = new Date();
      const currentQuarter = getQuarter(now);
      const currentYear = now.getFullYear();
      const prevQuarterDate = new Date(now);
      prevQuarterDate.setMonth(now.getMonth() - 3);
      const prevQuarter = getQuarter(prevQuarterDate);
      const prevYear = prevQuarterDate.getFullYear();

      const currentQuarterCount = new Set(
        jsonData.filter(r => isInQuarter(r['Created_at'], currentYear, currentQuarter))
                .map(r => r['Report Link'])
      ).size;
      document.getElementById('thisQuarter').innerText = currentQuarterCount;

      const prevQuarterCount = new Set(
        jsonData.filter(r => isInQuarter(r['Created_at'], prevYear, prevQuarter))
                .map(r => r['Report Link'])
      ).size;

      let percent = 0;
      if (prevQuarterCount > 0) {
        percent = ((currentQuarterCount - prevQuarterCount)/prevQuarterCount)*100;
      } else if (currentQuarterCount > 0) {
        percent = 100;
      }
      const percentElem = document.getElementById('percentChange');
      percentElem.innerText = `${percent.toFixed(1)}%`;
      percentElem.style.color = percent >= 0 ? 'lightgreen' : 'red';

      // --- Tribe counts & top tribe ---
      const tribeCounts = {};
      jsonData.forEach(r => {
        const tribe = r['Tribe Owner'] || 'Unknown';
        tribeCounts[tribe] = (tribeCounts[tribe] || 0) + 1;
      });

      let topTribe = '-';
      let maxCount = 0;
      Object.entries(tribeCounts).forEach(([tribe, count]) => {
        if (count > maxCount) {
          maxCount = count;
          topTribe = tribe;
        }
      });
      document.getElementById('topTribe').innerText = topTribe;

      // --- Monthly Line Chart ---
      const monthlyCounts = {};
      jsonData.forEach(r => {
        const date = parseDate(r['Created_at']);
        if (!date) return;
        const monthKey = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}`;
        if (!monthlyCounts[monthKey]) monthlyCounts[monthKey] = new Set();
        monthlyCounts[monthKey].add(r['Report Link']);
      });

      const sortedMonths = Object.keys(monthlyCounts).sort();
      const monthlyValues = sortedMonths.map(m => monthlyCounts[m].size);

      const ctxMonthly = document.getElementById('monthlyChart').getContext('2d');
      new Chart(ctxMonthly, {
        type: 'line',
        data: {
          labels: sortedMonths,
          datasets: [{
            label: 'Published Dashboards',
            data: monthlyValues,
            borderColor: '#573789',
            backgroundColor: 'rgba(87,55,137,0.1)',
            tension: 0.3,
            fill: true,
            pointRadius: 5,
            pointBackgroundColor: '#573789'
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { display: false } },
          scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } },
          maintainAspectRatio: false
        }
      });

      // --- Tribe Bar & Pie Chart (sorted descending) ---
      let sortedTribes = Object.entries(tribeCounts)
        .sort((a,b) => b[1] - a[1]);

      const sortedTribeLabels = sortedTribes.map(t => t[0]);
      const sortedTribeValues = sortedTribes.map(t => t[1]);
      const sortedBarColors = generateShades(sortedTribeValues, '#573789');

      // --- Tribe Horizontal Bar Chart ---
      const ctxBar = document.getElementById('tribeBarChart').getContext('2d');
      new Chart(ctxBar, {
        type: 'bar',
        data: {
          labels: sortedTribeLabels,
          datasets: [{
            label: 'Number of Dashboards',
            data: sortedTribeValues,
            backgroundColor: sortedBarColors
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          plugins: { legend: { display: false } },
          scales: { x: { beginAtZero: true, ticks: { stepSize: 1 } } },
          maintainAspectRatio: false
        }
      });

      // --- Tribe Pie Chart ---
      const ctxPie = document.getElementById('tribePieChart').getContext('2d');
      const pieValues = sortedTribeValues.map(v => ((v / sortedTribeValues.reduce((a, b) => a + b, 0)) * 100).toFixed(1));

      new Chart(ctxPie, {
        type: 'pie',
        data: {
          labels: sortedTribeLabels,
          datasets: [{
            data: pieValues,
            backgroundColor: sortedBarColors,
            borderColor: '#fff',
            borderWidth: 1
          }]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { position: 'bottom' },
            tooltip: {
              callbacks: {
                label: function (context) {
                  const label = context.label || '';
                  const value = context.parsed || 0;
                  return `${label}: ${value}%`;
                }
              }
            }
          },
          maintainAspectRatio: false
        }
      });

    })
    .catch(err => console.error("Error loading Excel file:", err));
});
