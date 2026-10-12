/**
 * FENNEX CUSTOMER HUB — FRONTEND APPLICATION LOGIC
 * Manages CRUD operations via .NET 10 Web API endpoints, responsive UI state,
 * live search/filtering, modal interactions, and fallback demo preview.
 */

// Application State
const state = {
    customers: [],
    filteredCustomers: [],
    isDemoMode: false,
    activeView: 'table', // 'table' | 'cards'
    activeFilter: 'all',
    activeSort: 'id-asc',
    searchQuery: '',
    selectedCustomerId: null
};

// Initial Mock / Sample Dataset for instant preview if Azure SQL connection is unavailable
const sampleCustomers = [
    {
        id: 1,
        sharePointId: 101,
        name: "Transocean Ltd",
        address: "Turmstrasse 30, 6312 Steinhausen, Switzerland",
        numberOfRigs: 38,
        numberOfJackUps: 0,
        numberOfModus: 28,
        siteUrls: "https://transocean.sharepoint.com/fleet, https://deepwater.com/ops",
        spLastModified: "2026-09-15T14:30:00Z",
        syncedAt: "2026-10-10T08:00:00Z",
        createdAt: "2026-08-01T10:00:00Z"
    },
    {
        id: 2,
        sharePointId: 102,
        name: "Valaris Limited",
        address: "110 Cannon Street, London, EC4N 6EU, United Kingdom",
        numberOfRigs: 52,
        numberOfJackUps: 35,
        numberOfModus: 11,
        siteUrls: "https://valaris.sharepoint.com/operations",
        spLastModified: "2026-09-28T11:15:00Z",
        syncedAt: "2026-10-10T08:00:00Z",
        createdAt: "2026-08-01T10:00:00Z"
    },
    {
        id: 3,
        sharePointId: 103,
        name: "Noble Corporation",
        address: "13135 Dairy Ashford, Suite 800, Sugar Land, TX 77478",
        numberOfRigs: 31,
        numberOfJackUps: 19,
        numberOfModus: 12,
        siteUrls: "https://noblecorp.sharepoint.com/rigs",
        spLastModified: "2026-10-02T16:45:00Z",
        syncedAt: "2026-10-10T08:00:00Z",
        createdAt: "2026-08-05T12:00:00Z"
    },
    {
        id: 4,
        sharePointId: 104,
        name: "Seadrill Limited",
        address: "Park Place, 55 Par-la-Ville Road, Hamilton, Bermuda",
        numberOfRigs: 18,
        numberOfJackUps: 10,
        numberOfModus: 8,
        siteUrls: "https://seadrill.sharepoint.com/sites/drilling",
        spLastModified: "2026-10-05T09:20:00Z",
        syncedAt: "2026-10-10T08:00:00Z",
        createdAt: "2026-08-10T09:30:00Z"
    }
];

// DOM Elements Cache
const elements = {
    // Buttons & Toggles
    refreshBtn: document.getElementById('refreshDataBtn'),
    openCreateModalBtn: document.getElementById('openCreateModalBtn'),
    toggleDemoModeBtn: document.getElementById('toggleDemoModeBtn'),
    demoModeBtnText: document.getElementById('demoModeBtnText'),
    connectionStatusText: document.getElementById('connectionStatusText'),
    envTag: document.querySelector('.env-tag'),
    tableViewBtn: document.getElementById('tableViewBtn'),
    cardsViewBtn: document.getElementById('cardsViewBtn'),
    exportExcelBtn: document.getElementById('exportExcelBtn'),
    exportWordBtn: document.getElementById('exportWordBtn'),
    detailsExportWordBtn: document.getElementById('detailsExportWordBtn'),
    searchInput: document.getElementById('searchInput'),
    clearSearchBtn: document.getElementById('clearSearchBtn'),
    filterAssetType: document.getElementById('filterAssetType'),
    sortBySelect: document.getElementById('sortBySelect'),

    // Containers & States
    loadingState: document.getElementById('loadingState'),
    emptyState: document.getElementById('emptyState'),
    emptyAddBtn: document.getElementById('emptyAddBtn'),
    tableViewContainer: document.getElementById('tableViewContainer'),
    cardsViewContainer: document.getElementById('cardsViewContainer'),
    customersTableBody: document.getElementById('customersTableBody'),
    recordsCountText: document.getElementById('recordsCountText'),
    dataTimestamp: document.getElementById('dataTimestamp'),
    systemAlertBanner: document.getElementById('systemAlertBanner'),
    enableDemoDataAlertBtn: document.getElementById('enableDemoDataAlertBtn'),
    dismissAlertBtn: document.getElementById('dismissAlertBtn'),
    toastContainer: document.getElementById('toastContainer'),

    // Metrics
    statTotalCustomers: document.getElementById('statTotalCustomers'),
    statTotalRigs: document.getElementById('statTotalRigs'),
    statTotalJackUps: document.getElementById('statTotalJackUps'),
    statTotalModus: document.getElementById('statTotalModus'),

    // Modals
    createModal: document.getElementById('createModal'),
    createForm: document.getElementById('createCustomerForm'),
    editModal: document.getElementById('editModal'),
    editForm: document.getElementById('editCustomerForm'),
    detailsModal: document.getElementById('detailsModal'),
    detailsContent: document.getElementById('detailsContent'),
    detailsCustomerName: document.getElementById('detailsCustomerName'),
    detailsEditBtn: document.getElementById('detailsEditBtn'),
    deleteModal: document.getElementById('deleteModal'),
    deleteCustomerName: document.getElementById('deleteCustomerName'),
    deleteCustomerId: document.getElementById('deleteCustomerId'),
    confirmDeleteBtn: document.getElementById('confirmDeleteBtn')
};

// -----------------------------------------------------------------------------
// Initialization
// -----------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    loadCustomers();
});

function setupEventListeners() {
    // Data Loading & Controls
    elements.refreshBtn.addEventListener('click', () => loadCustomers(true));
    elements.openCreateModalBtn.addEventListener('click', () => openModal('createModal'));
    elements.emptyAddBtn.addEventListener('click', () => openModal('createModal'));

    // Export Controls (Task 7 & Task 8)
    if (elements.exportExcelBtn) {
        elements.exportExcelBtn.addEventListener('click', triggerExcelExport);
    }
    if (elements.exportWordBtn) {
        elements.exportWordBtn.addEventListener('click', triggerWordExport);
    }
    if (elements.detailsExportWordBtn) {
        elements.detailsExportWordBtn.addEventListener('click', () => {
            if (state.selectedCustomerId) {
                exportCustomerWord(state.selectedCustomerId);
            }
        });
    }

    // Demo Mode Toggles
    elements.toggleDemoModeBtn.addEventListener('click', toggleDemoMode);
    elements.enableDemoDataAlertBtn.addEventListener('click', () => {
        elements.systemAlertBanner.classList.add('hidden');
        enableDemoMode();
    });
    elements.dismissAlertBtn.addEventListener('click', () => {
        elements.systemAlertBanner.classList.add('hidden');
    });

    // Search & Filter Listeners
    elements.searchInput.addEventListener('input', (e) => {
        state.searchQuery = e.target.value.trim().toLowerCase();
        elements.clearSearchBtn.classList.toggle('hidden', state.searchQuery === '');
        applyFiltersAndRender();
    });

    elements.clearSearchBtn.addEventListener('click', () => {
        elements.searchInput.value = '';
        state.searchQuery = '';
        elements.clearSearchBtn.classList.add('hidden');
        applyFiltersAndRender();
    });

    elements.filterAssetType.addEventListener('change', (e) => {
        state.activeFilter = e.target.value;
        applyFiltersAndRender();
    });

    elements.sortBySelect.addEventListener('change', (e) => {
        state.activeSort = e.target.value;
        applyFiltersAndRender();
    });

    // View Switcher
    elements.tableViewBtn.addEventListener('click', () => switchView('table'));
    elements.cardsViewBtn.addEventListener('click', () => switchView('cards'));

    // Modal Close Triggers
    document.querySelectorAll('[data-close]').forEach(btn => {
        btn.addEventListener('click', () => closeModal(btn.getAttribute('data-close')));
    });

    // Modal backdrop click
    document.querySelectorAll('.modal-overlay').forEach(modal => {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal(modal.id);
        });
    });

    // Form Submissions
    elements.createForm.addEventListener('submit', handleCreateCustomer);
    elements.editForm.addEventListener('submit', handleEditCustomer);
    elements.confirmDeleteBtn.addEventListener('click', handleConfirmDelete);

    // Details Edit Trigger
    elements.detailsEditBtn.addEventListener('click', () => {
        closeModal('detailsModal');
        if (state.selectedCustomerId) {
            openEditModal(state.selectedCustomerId);
        }
    });
}

// -----------------------------------------------------------------------------
// API CRUD Operations
// -----------------------------------------------------------------------------

/**
 * Loads customer records from GET /api/customers or demo data
 */
async function loadCustomers(isManualRefresh = false) {
    if (state.isDemoMode) {
        renderDemoData();
        if (isManualRefresh) showToast('Sample demo data refreshed', 'info');
        return;
    }

    setLoading(true);

    try {
        const response = await fetch('/api/customers');

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ error: 'Server returned error ' + response.status }));
            throw new Error(errorData.error || `HTTP ${response.status}`);
        }

        const data = await response.json();
        state.customers = Array.isArray(data) ? data : [];
        elements.systemAlertBanner.classList.add('hidden');
        
        applyFiltersAndRender();
        updateMetrics();
        updateTimestamp();

        if (isManualRefresh) {
            showToast(`Loaded ${state.customers.length} customer records from Azure SQL`, 'success');
        }
    } catch (err) {
        console.warn('API Error loading customers:', err);
        // Show informative alert about Azure SQL firewall / connection
        elements.systemAlertBanner.classList.remove('hidden');
        document.getElementById('alertMessage').textContent =
            `Could not query Azure SQL: ${err.message}. If Azure SQL Firewall blocks your current IP, whitelist it in the Azure Portal or click below to explore in Demo Mode.`;

        showToast('Database connection unavailable: ' + err.message, 'error');
        
        // If initial load fails and list is empty, show empty state or prompt demo mode
        if (state.customers.length === 0) {
            state.filteredCustomers = [];
            renderEmptyState();
        }
    } finally {
        setLoading(false);
    }
}

/**
 * Handles POST /api/customers
 */
async function handleCreateCustomer(e) {
    e.preventDefault();

    const form = elements.createForm;
    const name = form.name.value.trim();
    const sharePointId = parseInt(form.sharePointId.value, 10);
    const address = form.address.value.trim() || null;
    const numberOfRigs = form.numberOfRigs.value ? parseInt(form.numberOfRigs.value, 10) : 0;
    const numberOfJackUps = form.numberOfJackUps.value ? parseInt(form.numberOfJackUps.value, 10) : 0;
    const numberOfModus = form.numberOfModus.value ? parseInt(form.numberOfModus.value, 10) : 0;
    const siteUrls = form.siteUrls.value.trim() || null;
    const spLastModified = form.spLastModified.value ? new Date(form.spLastModified.value).toISOString() : null;

    // Validation
    let isValid = true;
    clearFieldErrors();

    if (!name) {
        showFieldError('createName', 'Customer Name is required.');
        isValid = false;
    }

    if (isNaN(sharePointId) || sharePointId <= 0) {
        showFieldError('createSharePointId', 'A valid positive SharePoint ID is required.');
        isValid = false;
    }

    if (!isValid) return;

    const payload = {
        sharePointId,
        name,
        address,
        numberOfRigs,
        numberOfJackUps,
        numberOfModus,
        siteUrls,
        spLastModified
    };

    if (state.isDemoMode) {
        // Mock creation
        const newCustomer = {
            id: Math.max(...state.customers.map(c => c.id), 0) + 1,
            ...payload,
            syncedAt: new Date().toISOString(),
            createdAt: new Date().toISOString()
        };
        state.customers.unshift(newCustomer);
        closeModal('createModal');
        form.reset();
        applyFiltersAndRender();
        updateMetrics();
        showToast(`Created customer "${name}" (Demo Mode)`, 'success');
        return;
    }

    const submitBtn = document.getElementById('submitCreateBtn');
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Saving...';

    try {
        const response = await fetch('/api/customers', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (response.status === 201) {
            const created = await response.json();
            showToast(`Customer "${created.name}" created successfully (SQL ID: ${created.id})`, 'success');
            closeModal('createModal');
            form.reset();
            await loadCustomers();
        } else if (response.status === 409) {
            const err = await response.json();
            showToast(err.error || 'Duplicate SharePoint ID.', 'warning');
            showFieldError('createSharePointId', err.error || 'SharePointId already exists in database.');
        } else {
            const err = await response.json().catch(() => ({}));
            showToast(err.error || 'Failed to create customer.', 'error');
        }
    } catch (err) {
        showToast('Network error: ' + err.message, 'error');
    } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<i class="fa-solid fa-check"></i> <span>Save Customer</span>';
    }
}

/**
 * Handles PUT /api/customers/{id}
 */
async function handleEditCustomer(e) {
    e.preventDefault();

    const form = elements.editForm;
    const id = parseInt(form.id.value, 10);
    const name = form.name.value.trim();
    const address = form.address.value.trim() || null;
    const numberOfRigs = form.numberOfRigs.value ? parseInt(form.numberOfRigs.value, 10) : null;
    const numberOfJackUps = form.numberOfJackUps.value ? parseInt(form.numberOfJackUps.value, 10) : null;
    const numberOfModus = form.numberOfModus.value ? parseInt(form.numberOfModus.value, 10) : null;
    const siteUrls = form.siteUrls.value.trim() || null;
    const spLastModified = form.spLastModified.value ? new Date(form.spLastModified.value).toISOString() : null;

    if (!name) {
        showFieldError('editName', 'Name is required.');
        return;
    }

    const payload = {
        name,
        address,
        numberOfRigs,
        numberOfJackUps,
        numberOfModus,
        siteUrls,
        spLastModified
    };

    if (state.isDemoMode) {
        const idx = state.customers.findIndex(c => c.id === id);
        if (idx !== -1) {
            state.customers[idx] = {
                ...state.customers[idx],
                ...payload,
                syncedAt: new Date().toISOString()
            };
            closeModal('editModal');
            applyFiltersAndRender();
            updateMetrics();
            showToast(`Updated "${name}" (Demo Mode)`, 'success');
        }
        return;
    }

    const submitBtn = document.getElementById('submitEditBtn');
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Updating...';

    try {
        const response = await fetch(`/api/customers/${id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            const updated = await response.json();
            showToast(`Updated customer "${updated.name}" successfully`, 'success');
            closeModal('editModal');
            await loadCustomers();
        } else if (response.status === 404) {
            showToast(`Customer with ID ${id} not found.`, 'error');
        } else {
            const err = await response.json().catch(() => ({}));
            showToast(err.error || 'Failed to update customer.', 'error');
        }
    } catch (err) {
        showToast('Network error: ' + err.message, 'error');
    } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> <span>Update Customer</span>';
    }
}

/**
 * Handles DELETE /api/customers/{id}
 */
async function handleConfirmDelete() {
    const id = state.selectedCustomerId;
    if (!id) return;

    if (state.isDemoMode) {
        state.customers = state.customers.filter(c => c.id !== id);
        closeModal('deleteModal');
        applyFiltersAndRender();
        updateMetrics();
        showToast(`Deleted customer ID ${id} (Demo Mode)`, 'success');
        return;
    }

    const deleteBtn = elements.confirmDeleteBtn;
    deleteBtn.disabled = true;
    deleteBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Deleting...';

    try {
        const response = await fetch(`/api/customers/${id}`, {
            method: 'DELETE'
        });

        if (response.status === 204) {
            showToast(`Customer ID ${id} deleted permanently`, 'success');
            closeModal('deleteModal');
            await loadCustomers();
        } else if (response.status === 404) {
            showToast(`Customer ID ${id} not found.`, 'error');
        } else {
            const err = await response.json().catch(() => ({}));
            showToast(err.error || 'Failed to delete customer.', 'error');
        }
    } catch (err) {
        showToast('Network error: ' + err.message, 'error');
    } finally {
        deleteBtn.disabled = false;
        deleteBtn.innerHTML = '<i class="fa-solid fa-trash"></i> <span>Delete Permanently</span>';
    }
}

// -----------------------------------------------------------------------------
// UI Filters, Search, and Sorting
// -----------------------------------------------------------------------------
function applyFiltersAndRender() {
    let result = [...state.customers];

    // Search query
    if (state.searchQuery) {
        const q = state.searchQuery;
        result = result.filter(c => {
            const name = (c.name || '').toLowerCase();
            const addr = (c.address || '').toLowerCase();
            const spId = String(c.sharePointId || '');
            const id = String(c.id || '');
            const urls = (c.siteUrls || '').toLowerCase();
            return name.includes(q) || addr.includes(q) || spId.includes(q) || id.includes(q) || urls.includes(q);
        });
    }

    // Asset filter
    if (state.activeFilter === 'has-rigs') {
        result = result.filter(c => (c.numberOfRigs || 0) > 0);
    } else if (state.activeFilter === 'has-jackups') {
        result = result.filter(c => (c.numberOfJackUps || 0) > 0);
    } else if (state.activeFilter === 'has-modus') {
        result = result.filter(c => (c.numberOfModus || 0) > 0);
    }

    // Sorting
    switch (state.activeSort) {
        case 'id-asc':
            result.sort((a, b) => a.id - b.id);
            break;
        case 'id-desc':
            result.sort((a, b) => b.id - a.id);
            break;
        case 'name-asc':
            result.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            break;
        case 'name-desc':
            result.sort((a, b) => (b.name || '').localeCompare(a.name || ''));
            break;
        case 'rigs-desc':
            result.sort((a, b) => (b.numberOfRigs || 0) - (a.numberOfRigs || 0));
            break;
        case 'recent-sync':
            result.sort((a, b) => new Date(b.syncedAt || 0) - new Date(a.syncedAt || 0));
            break;
    }

    state.filteredCustomers = result;

    elements.recordsCountText.textContent = `Showing ${result.length} of ${state.customers.length} customers`;

    if (result.length === 0) {
        renderEmptyState();
    } else {
        elements.emptyState.classList.add('hidden');
        if (state.activeView === 'table') {
            renderTable(result);
            elements.tableViewContainer.classList.remove('hidden');
            elements.cardsViewContainer.classList.add('hidden');
        } else {
            renderCards(result);
            elements.cardsViewContainer.classList.remove('hidden');
            elements.tableViewContainer.classList.add('hidden');
        }
    }
}

// -----------------------------------------------------------------------------
// Rendering: Table & Cards
// -----------------------------------------------------------------------------
function renderTable(customers) {
    const tbody = elements.customersTableBody;
    tbody.innerHTML = '';

    customers.forEach(customer => {
        const tr = document.createElement('tr');

        const rigs = customer.numberOfRigs ?? 0;
        const jackUps = customer.numberOfJackUps ?? 0;
        const modus = customer.numberOfModus ?? 0;

        const spDate = customer.spLastModified ? formatDate(customer.spLastModified) : '<span class="text-faint">—</span>';
        const syncDate = customer.syncedAt ? formatDate(customer.syncedAt) : '<span class="text-faint">—</span>';

        const siteUrlPills = renderUrlPills(customer.siteUrls);

        tr.innerHTML = `
            <td><span class="id-badge">#${customer.id}</span></td>
            <td><span class="sp-badge">SP-${customer.sharePointId}</span></td>
            <td>
                <div class="customer-cell">
                    <span class="customer-name">${escapeHtml(customer.name || 'Unnamed')}</span>
                    ${customer.address ? `<span class="customer-address"><i class="fa-solid fa-location-dot"></i>${escapeHtml(customer.address)}</span>` : ''}
                </div>
            </td>
            <td>
                ${siteUrlPills || '<span class="text-faint">No links recorded</span>'}
            </td>
            <td style="text-align: center;">
                <div class="rig-badges">
                    <span class="badge-rig rigs" title="Oil Rigs"><i class="fa-solid fa-tower-broadcast"></i> ${rigs} Rigs</span>
                    <span class="badge-rig jackups" title="Jack-Up Units"><i class="fa-solid fa-anchor"></i> ${jackUps} JU</span>
                    <span class="badge-rig modus" title="MODU Vessels"><i class="fa-solid fa-ship"></i> ${modus} MODU</span>
                </div>
            </td>
            <td>${spDate}</td>
            <td>${syncDate}</td>
            <td style="text-align: right;">
                <div class="action-buttons">
                    <button class="btn-icon view" title="View Details" onclick="openDetailsModal(${customer.id})">
                        <i class="fa-solid fa-eye"></i>
                    </button>
                    <button class="btn-icon export-word" title="Export Profile to Word (.docx)" onclick="exportCustomerWord(${customer.id})">
                        <i class="fa-solid fa-file-word"></i>
                    </button>
                    <button class="btn-icon edit" title="Edit Customer" onclick="openEditModal(${customer.id})">
                        <i class="fa-solid fa-pen-to-square"></i>
                    </button>
                    <button class="btn-icon delete" title="Delete Customer" onclick="openDeleteModal(${customer.id}, '${escapeHtml(customer.name || 'Customer')}')">
                        <i class="fa-solid fa-trash"></i>
                    </button>
                </div>
            </td>
        `;

        tbody.appendChild(tr);
    });
}

function renderCards(customers) {
    const container = elements.cardsViewContainer;
    container.innerHTML = '';

    customers.forEach(customer => {
        const card = document.createElement('div');
        card.className = 'customer-card glass-panel';

        const initials = (customer.name || 'C')
            .split(' ')
            .map(w => w[0])
            .slice(0, 2)
            .join('')
            .toUpperCase();

        const rigs = customer.numberOfRigs ?? 0;
        const jackUps = customer.numberOfJackUps ?? 0;
        const modus = customer.numberOfModus ?? 0;

        card.innerHTML = `
            <div class="card-top">
                <div class="card-avatar-wrap">
                    <div class="company-avatar">${initials}</div>
                    <div class="card-title-group">
                        <h4>${escapeHtml(customer.name || 'Unnamed')}</h4>
                        <span class="card-sp-id">ID: #${customer.id} &bull; SharePoint ID: ${customer.sharePointId}</span>
                    </div>
                </div>
                <div class="action-buttons">
                    <button class="btn-icon view" title="View" onclick="openDetailsModal(${customer.id})">
                        <i class="fa-solid fa-eye"></i>
                    </button>
                    <button class="btn-icon export-word" title="Export to Word (.docx)" onclick="exportCustomerWord(${customer.id})">
                        <i class="fa-solid fa-file-word"></i>
                    </button>
                    <button class="btn-icon edit" title="Edit" onclick="openEditModal(${customer.id})">
                        <i class="fa-solid fa-pen-to-square"></i>
                    </button>
                    <button class="btn-icon delete" title="Delete" onclick="openDeleteModal(${customer.id}, '${escapeHtml(customer.name || 'Customer')}')">
                        <i class="fa-solid fa-trash"></i>
                    </button>
                </div>
            </div>

            <div class="card-body">
                ${customer.address ? `
                    <p style="font-size: 0.82rem; color: var(--text-muted); margin-bottom: 0.85rem;">
                        <i class="fa-solid fa-location-dot" style="color: var(--primary); margin-right: 0.35rem;"></i>
                        ${escapeHtml(customer.address)}
                    </p>
                ` : ''}

                <div class="card-rig-metrics">
                    <div>
                        <div class="mini-stat-value" style="color: var(--primary);">${rigs}</div>
                        <div class="mini-stat-label">Rigs</div>
                    </div>
                    <div>
                        <div class="mini-stat-value" style="color: var(--accent-amber);">${jackUps}</div>
                        <div class="mini-stat-label">Jack-Ups</div>
                    </div>
                    <div>
                        <div class="mini-stat-value" style="color: var(--accent-purple);">${modus}</div>
                        <div class="mini-stat-label">MODUs</div>
                    </div>
                </div>

                <div style="margin-top: 0.85rem;">
                    ${renderUrlPills(customer.siteUrls) || ''}
                </div>
            </div>

            <div class="card-footer-meta">
                <span><i class="fa-solid fa-clock-rotate-left"></i> Synced: ${customer.syncedAt ? formatRelativeTime(customer.syncedAt) : 'Never'}</span>
                <span class="id-badge">dbo.Customers</span>
            </div>
        `;

        container.appendChild(card);
    });
}

function renderUrlPills(siteUrls) {
    if (!siteUrls) return '';
    return siteUrls
        .split(',')
        .map(url => url.trim())
        .filter(url => url.length > 0)
        .map(url => {
            const cleanUrl = url.startsWith('http') ? url : `https://${url}`;
            return `<a href="${escapeHtml(cleanUrl)}" target="_blank" rel="noopener noreferrer" class="site-link-pill" title="${escapeHtml(url)}">
                <i class="fa-solid fa-arrow-up-right-from-square"></i> ${escapeHtml(url.replace(/^https?:\/\//, ''))}
            </a>`;
        })
        .join('');
}

// -----------------------------------------------------------------------------
// Modals & User Interaction Handlers
// -----------------------------------------------------------------------------
window.openDetailsModal = function(id) {
    state.selectedCustomerId = id;
    const customer = state.customers.find(c => c.id === id);
    if (!customer) return;

    elements.detailsCustomerName.textContent = customer.name || `Customer #${customer.id}`;
    
    elements.detailsContent.innerHTML = `
        <div class="detail-item">
            <div class="detail-label">Database Primary Key (Id)</div>
            <div class="detail-val"><span class="id-badge">#${customer.id}</span></div>
        </div>
        <div class="detail-item">
            <div class="detail-label">SharePoint List Item ID</div>
            <div class="detail-val"><span class="sp-badge">SP-${customer.sharePointId}</span></div>
        </div>
        <div class="detail-item full-width">
            <div class="detail-label">Company / Account Name</div>
            <div class="detail-val" style="font-weight: 700; color: #fff; font-size: 1.1rem;">${escapeHtml(customer.name || '—')}</div>
        </div>
        <div class="detail-item full-width">
            <div class="detail-label">Headquarters / Operational Address</div>
            <div class="detail-val">${escapeHtml(customer.address || 'Not specified')}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Oil Rigs Count</div>
            <div class="detail-val" style="color: var(--primary); font-weight: 700;">${customer.numberOfRigs ?? '0'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Jack-Up Units</div>
            <div class="detail-val" style="color: var(--accent-amber); font-weight: 700;">${customer.numberOfJackUps ?? '0'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">MODU Vessels</div>
            <div class="detail-val" style="color: var(--accent-purple); font-weight: 700;">${customer.numberOfModus ?? '0'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Total Fleet Units</div>
            <div class="detail-val" style="font-weight: 700;">${(customer.numberOfRigs || 0) + (customer.numberOfJackUps || 0) + (customer.numberOfModus || 0)}</div>
        </div>
        <div class="detail-item full-width">
            <div class="detail-label">Associated Site URLs</div>
            <div class="detail-val">${renderUrlPills(customer.siteUrls) || 'No URLs provided'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">SharePoint Last Modified</div>
            <div class="detail-val">${customer.spLastModified ? formatDate(customer.spLastModified) : '—'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Last Synced to Azure SQL</div>
            <div class="detail-val">${customer.syncedAt ? formatDate(customer.syncedAt) : '—'}</div>
        </div>
        <div class="detail-item full-width">
            <div class="detail-label">SQL Record Created At</div>
            <div class="detail-val">${customer.createdAt ? formatDate(customer.createdAt) : '—'}</div>
        </div>
    `;

    openModal('detailsModal');
};

window.openEditModal = function(id) {
    state.selectedCustomerId = id;
    const customer = state.customers.find(c => c.id === id);
    if (!customer) return;

    const form = elements.editForm;
    form.id.value = customer.id;
    form.name.value = customer.name || '';
    form.address.value = customer.address || '';
    form.numberOfRigs.value = customer.numberOfRigs ?? '';
    form.numberOfJackUps.value = customer.numberOfJackUps ?? '';
    form.numberOfModus.value = customer.numberOfModus ?? '';
    form.siteUrls.value = customer.siteUrls || '';
    
    if (customer.spLastModified) {
        // Format to YYYY-MM-DDTHH:mm
        const d = new Date(customer.spLastModified);
        const pad = (n) => String(n).padStart(2, '0');
        form.spLastModified.value = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    } else {
        form.spLastModified.value = '';
    }

    document.getElementById('editModalSubtitle').textContent = `Editing #${customer.id} — SharePoint ID: ${customer.sharePointId}`;
    clearFieldErrors();
    openModal('editModal');
};

window.openDeleteModal = function(id, name) {
    state.selectedCustomerId = id;
    elements.deleteCustomerId.textContent = id;
    elements.deleteCustomerName.textContent = name;
    openModal('deleteModal');
};

function openModal(id) {
    const modal = document.getElementById(id);
    if (modal) modal.classList.remove('hidden');
}

function closeModal(id) {
    const modal = document.getElementById(id);
    if (modal) modal.classList.add('hidden');
}

function switchView(viewType) {
    state.activeView = viewType;
    elements.tableViewBtn.classList.toggle('active', viewType === 'table');
    elements.cardsViewBtn.classList.toggle('active', viewType === 'cards');
    applyFiltersAndRender();
}

// -----------------------------------------------------------------------------
// Demo Mode Support
// -----------------------------------------------------------------------------
function toggleDemoMode() {
    if (state.isDemoMode) {
        disableDemoMode();
    } else {
        enableDemoMode();
    }
}

function enableDemoMode() {
    state.isDemoMode = true;
    elements.demoModeBtnText.textContent = 'Demo Mode: ON';
    elements.toggleDemoModeBtn.classList.add('btn-primary');
    elements.toggleDemoModeBtn.classList.remove('btn-ghost');
    elements.connectionStatusText.textContent = 'Demo Sandbox';
    elements.envTag.classList.add('demo');

    renderDemoData();
    showToast('Loaded interactive Demo Mode with sample offshore fleets', 'info');
}

function disableDemoMode() {
    state.isDemoMode = false;
    elements.demoModeBtnText.textContent = 'Demo Mode: Off';
    elements.toggleDemoModeBtn.classList.remove('btn-primary');
    elements.toggleDemoModeBtn.classList.add('btn-ghost');
    elements.connectionStatusText.textContent = 'Live API Mode';
    elements.envTag.classList.remove('demo');

    loadCustomers(true);
}

function renderDemoData() {
    state.customers = [...sampleCustomers];
    applyFiltersAndRender();
    updateMetrics();
    updateTimestamp();
}

// -----------------------------------------------------------------------------
// Statistics & Metrics
// -----------------------------------------------------------------------------
function updateMetrics() {
    const list = state.customers;
    const totalCustomers = list.length;
    const totalRigs = list.reduce((sum, c) => sum + (c.numberOfRigs || 0), 0);
    const totalJackUps = list.reduce((sum, c) => sum + (c.numberOfJackUps || 0), 0);
    const totalModus = list.reduce((sum, c) => sum + (c.numberOfModus || 0), 0);

    animateCounter(elements.statTotalCustomers, totalCustomers);
    animateCounter(elements.statTotalRigs, totalRigs);
    animateCounter(elements.statTotalJackUps, totalJackUps);
    animateCounter(elements.statTotalModus, totalModus);
}

function animateCounter(element, targetValue) {
    const startValue = parseInt(element.textContent, 10) || 0;
    if (startValue === targetValue) {
        element.textContent = targetValue;
        return;
    }

    const duration = 600;
    const startTime = performance.now();

    function update(time) {
        const elapsed = time - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const current = Math.floor(startValue + (targetValue - startValue) * progress);
        element.textContent = current;

        if (progress < 1) {
            requestAnimationFrame(update);
        } else {
            element.textContent = targetValue;
        }
    }

    requestAnimationFrame(update);
}

function updateTimestamp() {
    const now = new Date();
    elements.dataTimestamp.textContent = `Last refreshed: ${now.toLocaleTimeString()}`;
}

// -----------------------------------------------------------------------------
// Toast Notifications
// -----------------------------------------------------------------------------
function showToast(message, type = 'info') {
    const container = elements.toastContainer;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let icon = 'fa-info-circle';
    if (type === 'success') icon = 'fa-circle-check';
    if (type === 'error') icon = 'fa-circle-exclamation';
    if (type === 'warning') icon = 'fa-triangle-exclamation';

    toast.innerHTML = `
        <i class="fa-solid ${icon}"></i>
        <span>${escapeHtml(message)}</span>
    `;

    container.appendChild(toast);

    setTimeout(() => {
        toast.style.animation = 'toastIn 0.3s cubic-bezier(0.16, 1, 0.3, 1) reverse forwards';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// -----------------------------------------------------------------------------
// Utility Functions
// -----------------------------------------------------------------------------
function setLoading(isLoading) {
    elements.loadingState.classList.toggle('hidden', !isLoading);
    if (isLoading) {
        elements.tableViewContainer.classList.add('hidden');
        elements.cardsViewContainer.classList.add('hidden');
        elements.emptyState.classList.add('hidden');
    }
}

function renderEmptyState() {
    elements.tableViewContainer.classList.add('hidden');
    elements.cardsViewContainer.classList.add('hidden');
    elements.emptyState.classList.remove('hidden');
}

function showFieldError(fieldId, message) {
    const errEl = document.getElementById(`error_${fieldId}`);
    if (errEl) errEl.textContent = message;
}

function clearFieldErrors() {
    document.querySelectorAll('.field-error').forEach(el => el.textContent = '');
}

function formatDate(dateStr) {
    try {
        const d = new Date(dateStr);
        return d.toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return dateStr;
    }
}

function formatRelativeTime(dateStr) {
    try {
        const diffMs = new Date() - new Date(dateStr);
        const mins = Math.floor(diffMs / 60000);
        if (mins < 1) return 'Just now';
        if (mins < 60) return `${mins}m ago`;
        const hours = Math.floor(mins / 60);
        if (hours < 24) return `${hours}h ago`;
        return `${Math.floor(hours / 24)}d ago`;
    } catch {
        return 'Recently';
    }
}

function escapeHtml(str) {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

// -----------------------------------------------------------------------------
// OpenXML Export Handlers (Task 7: Excel, Task 8: Word)
// -----------------------------------------------------------------------------
async function triggerExcelExport() {
    showToast('Generating Excel (.xlsx) spreadsheet via OpenXML...', 'info');

    try {
        const response = await fetch('/api/customers/export/excel');
        if (!response.ok) {
            const errJson = await response.json().catch(() => ({ error: 'Export failed with status ' + response.status }));
            throw new Error(errJson.error || 'Server error exporting to Excel');
        }

        const blob = await response.blob();
        downloadBlob(blob, `Fennex_Customers_Export_${getTimestampString()}.xlsx`);
        showToast('Excel workbook exported successfully!', 'success');
    } catch (err) {
        console.warn('Excel export error:', err);
        showToast('Excel Export Notice: ' + err.message, 'warning');
    }
}

async function triggerWordExport() {
    showToast('Generating Word (.docx) fleet report via OpenXML...', 'info');

    try {
        const response = await fetch('/api/customers/export/word');
        if (!response.ok) {
            const errJson = await response.json().catch(() => ({ error: 'Export failed with status ' + response.status }));
            throw new Error(errJson.error || 'Server error exporting to Word');
        }

        const blob = await response.blob();
        downloadBlob(blob, `Fennex_Fleet_Report_${getTimestampString()}.docx`);
        showToast('Word fleet report exported successfully!', 'success');
    } catch (err) {
        console.warn('Word export error:', err);
        showToast('Word Export Notice: ' + err.message, 'warning');
    }
}

window.exportCustomerWord = async function(id) {
    showToast(`Generating Word document for Customer #${id}...`, 'info');

    try {
        const response = await fetch(`/api/customers/${id}/export/word`);
        if (!response.ok) {
            const errJson = await response.json().catch(() => ({ error: 'Export failed with status ' + response.status }));
            throw new Error(errJson.error || 'Server error exporting customer profile');
        }

        const customer = state.customers.find(c => c.id === id);
        const nameSlug = (customer?.name || 'Customer').replace(/[^a-zA-Z0-9]/g, '_');
        const blob = await response.blob();
        downloadBlob(blob, `Fennex_Customer_${nameSlug}_${id}.docx`);
        showToast(`Word document exported for Customer #${id}!`, 'success');
    } catch (err) {
        console.warn('Single Word export error:', err);
        showToast('Profile Export Notice: ' + err.message, 'warning');
    }
};

function downloadBlob(blob, fileName) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 200);
}

function getTimestampString() {
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
}
