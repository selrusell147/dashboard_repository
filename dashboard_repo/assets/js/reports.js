// reports.js

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

// --- Excel date parsing ---
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

// --- Format date to YYYY-MM-DD ---
function formatDate(date) {
  if (!date) return '-';
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

window.addEventListener('DOMContentLoaded', () => {
  const excelPath = 'https://uniondigitalph.sharepoint.com/sites/DataDashboardsRepository/_layouts/15/download.aspx?SourceUrl=/sites/DataDashboardsRepository/Shared%20Documents/Dashboard%20Inventory.xlsx';

  fetch(excelPath)
    .then(res => res.arrayBuffer())
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      const tribes = [
        'Audit','Collections','Commercial and Revenue','Compliance','Finance',
        'Human Resource','Information Security','Legal','Marketing','Operations',
        'Product','Risk Management','Technology','Treasury'
      ];

      const reportsContainer = document.getElementById('reportsContainer');
      reportsContainer.innerHTML = '';

      tribes.forEach((tribe, index) => {
        const tribeReports = jsonData
          .filter(r => (r['Tribe Owner'] || 'Unknown') === tribe)
          .sort((a, b) => {
            const dateA = parseDate(a['Created_at']);
            const dateB = parseDate(b['Created_at']);
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1;
            if (!dateB) return -1;
            return dateA - dateB;
          });

        const tribeCard = document.createElement('div');
        tribeCard.classList.add('card', 'mb-3');

        const tribeHeader = document.createElement('div');
        tribeHeader.classList.add('card-header', 'd-flex', 'justify-content-between', 'align-items-center');
        tribeHeader.style.cursor = 'pointer';
        tribeHeader.setAttribute('data-bs-toggle', 'collapse');
        tribeHeader.setAttribute('data-bs-target', `#collapseTribe${index}`);
        tribeHeader.setAttribute('aria-expanded', 'false');
        tribeHeader.setAttribute('aria-controls', `collapseTribe${index}`);
        tribeHeader.innerHTML = `
          <strong>${tribe}</strong>
          <span class="badge bg-secondary">${tribeReports.length}</span>
        `;

        const tribeBody = document.createElement('div');
        tribeBody.classList.add('collapse');
        tribeBody.id = `collapseTribe${index}`;

        const cardBody = document.createElement('div');
        cardBody.classList.add('card-body', 'p-2');

        const table = document.createElement('table');
        table.classList.add('tribe-table', 'table', 'table-sm', 'table-striped', 'mb-0');

        const thead = document.createElement('thead');
        thead.innerHTML = `
          <tr>
            <th>Created At</th>
            <th>Report Name</th>
            <th>Description</th>
            <th>Report Link</th>
          </tr>
        `;
        table.appendChild(thead);

        const tbody = document.createElement('tbody');

        if (tribeReports.length === 0) {
          const tr = document.createElement('tr');
          tr.innerHTML = `<td colspan="4"><em>No dashboards available</em></td>`;
          tbody.appendChild(tr);
        } else {
          tribeReports.forEach(r => {
            const tr = document.createElement('tr');
            const date = formatDate(parseDate(r['Created_at']));
            const reportName = r['Report Name'] || '-';
            const desc = r['Description'] || '-';
            const link = r['Report Link']
              ? `<a href="${r['Report Link']}" target="_blank">Click to view dashboard</a>`
              : '-';
            tr.innerHTML = `
              <td>${date}</td>
              <td>${reportName}</td>
              <td>${desc}</td>
              <td>${link}</td>
            `;
            tbody.appendChild(tr);
          });
        }

        table.appendChild(tbody);
        cardBody.appendChild(table);
        tribeBody.appendChild(cardBody);
        tribeCard.appendChild(tribeHeader);
        tribeCard.appendChild(tribeBody);
        reportsContainer.appendChild(tribeCard);
      });

      // --- Search Setup ---
      const searchInput = document.getElementById('searchInput');
      const suggestionsList = document.getElementById('suggestionsList');
      const reportNames = [...new Set(jsonData.map(r => r['Report Name']).filter(Boolean))];
      let activeSuggestionIndex = -1;

      function showSuggestions(value) {
        const filtered = reportNames
          .filter(name => name.toLowerCase().includes(value.toLowerCase()))
          .slice(0, 10);

        suggestionsList.innerHTML = '';
        activeSuggestionIndex = -1;

        if (filtered.length === 0) {
          suggestionsList.style.display = 'none';
          return;
        }

        filtered.forEach(name => {
          const li = document.createElement('li');
          li.classList.add('list-group-item', 'list-group-item-action');
          li.textContent = name;
          li.addEventListener('click', () => {
            searchInput.value = name;
            suggestionsList.style.display = 'none';
            filterReports(name);
          });
          suggestionsList.appendChild(li);
        });

        suggestionsList.style.display = 'block';
      }

      function updateActiveSuggestion() {
        const items = suggestionsList.querySelectorAll('li');
        items.forEach((item, i) => {
          item.classList.toggle('active', i === activeSuggestionIndex);
        });
      }

      searchInput.addEventListener('input', e => {
        const value = e.target.value.trim();
        if (value.length === 0) {
          suggestionsList.style.display = 'none';
          resetReports();
          return;
        }
        showSuggestions(value);
      });

      searchInput.addEventListener('keydown', e => {
        const items = suggestionsList.querySelectorAll('li');
        if (suggestionsList.style.display === 'none' || items.length === 0) return;

        if (e.key === 'ArrowDown') {
          e.preventDefault();
          activeSuggestionIndex = (activeSuggestionIndex + 1) % items.length;
          updateActiveSuggestion();
        } else if (e.key === 'ArrowUp') {
          e.preventDefault();
          activeSuggestionIndex = (activeSuggestionIndex - 1 + items.length) % items.length;
          updateActiveSuggestion();
        } else if (e.key === 'Enter') {
          e.preventDefault();
          if (activeSuggestionIndex >= 0 && items[activeSuggestionIndex]) {
            const selected = items[activeSuggestionIndex].textContent;
            searchInput.value = selected;
            suggestionsList.style.display = 'none';
            filterReports(selected);
          } else {
            const value = e.target.value.trim();
            suggestionsList.style.display = 'none';
            filterReports(value);
          }
        } else if (e.key === 'Escape') {
          suggestionsList.style.display = 'none';
        }
      });

      function filterReports(query) {
        const cards = document.querySelectorAll('.card');
        cards.forEach(card => {
          const rows = card.querySelectorAll('tbody tr');
          let tribeHasMatch = false;

          rows.forEach(row => {
            const nameCell = row.children[1];
            if (!nameCell) return;
            const matches = nameCell.textContent.toLowerCase().includes(query.toLowerCase());
            row.style.display = matches ? '' : 'none';
            if (matches) tribeHasMatch = true;
          });

          card.style.display = tribeHasMatch ? '' : 'none';
        });
      }

      function resetReports() {
        const cards = document.querySelectorAll('.card');
        cards.forEach(card => {
          card.style.display = '';
          card.querySelectorAll('tbody tr').forEach(row => row.style.display = '');
        });
      }

    })
    .catch(err => console.error("Error loading Excel file:", err));
});
