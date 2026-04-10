const searchInput = document.getElementById('search-input');
const filterStatus = document.getElementById('filter-status');
const filterDate = document.getElementById('filter-date');
const filterTime = document.getElementById('filter-time');
const filterType = document.getElementById('filter-type');
const filterPlace = document.getElementById('filter-place');
const table = document.getElementById('orders-table');
const codeTableBody = document.querySelector('.code-generator-block tbody');
const kpiValues = document.querySelectorAll('.kpi strong');
const generateCodeBtn = document.getElementById('generate-code-btn');
const copyCodeBtn = document.getElementById('copy-code-btn');
const generatedCode = document.getElementById('generated-code');
let refreshTimer = null;
let inFlightRefresh = false;
let lastUserActivityAt = Date.now();

const REFRESH_FAST_MS = 15000;
const IDLE_AFTER_MS = 120000;

function normalize(value) {
    return (value || '').toString().toLowerCase().trim();
}

function matchFilter(value, expected) {
    if (!expected) return true;
    return normalize(value) === normalize(expected);
}

function applyFilters() {
    if (!table) return;

    const rows = table.querySelectorAll('tbody tr');
    const search = normalize(searchInput?.value);

    rows.forEach((row) => {
        const code = row.dataset.code || '';
        const name = row.dataset.name || '';
        const lastname = row.dataset.lastname || '';
        const phone = row.dataset.phone || '';
        const status = row.dataset.status || '';
        const date = row.dataset.date || '';
        const time = row.dataset.time || '';
        const type = row.dataset.type || '';
        const place = row.dataset.place || '';

        const searchable = normalize(`${code} ${name} ${lastname} ${phone}`);
        const passSearch = !search || searchable.includes(search);

        const passStatus = matchFilter(status, filterStatus?.value);
        const passDate = matchFilter(date, filterDate?.value);
        const passTime = matchFilter(time, filterTime?.value);
        const passType = matchFilter(type, filterType?.value);
        const passPlace = !filterPlace?.value || normalize(place).includes(normalize(filterPlace.value));

        row.style.display = passSearch && passStatus && passDate && passTime && passType && passPlace ? '' : 'none';
    });
}

function renderOrders(orders) {
    if (!table) return;

    const estados = Array.from(filterStatus?.options || []).map((option) => option.value).filter(Boolean);
    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    tbody.innerHTML = orders.map((order) => {
        const place = order.lugar_entrega || order.ciudad_destino || '';
        const selectedStatus = order.estado || '';
        const optionsHtml = estados.map((estado) => `<option value="${estado}" ${estado === selectedStatus ? 'selected' : ''}>${estado}</option>`).join('');
        const whatsappCell = order.whatsapp_url ? `<a class="btn btn-light" href="${order.whatsapp_url}" target="_blank" rel="noopener noreferrer">WhatsApp</a>` : '-';

        return `
            <tr
                data-id="${order.id || ''}"
                data-code="${order.codigo_producto || ''}"
                data-name="${order.nombre || ''}"
                data-lastname="${order.apellido || ''}"
                data-phone="${order.telefono || ''}"
                data-status="${order.estado || ''}"
                data-date="${(order.fecha_registro || '').toString().slice(0, 10)}"
                data-time="${order.hora_entrega || ''}"
                data-type="${order.tipo_entrega || ''}"
                data-place="${place}"
            >
                <td>${order.codigo_producto || ''}</td>
                <td>${order.pedido_grupo_id || ''}</td>
                <td>${order.nombre || ''}</td>
                <td>${order.apellido || '-'}</td>
                <td>${order.telefono || ''}</td>
                <td>${order.tipo_entrega || ''}</td>
                <td>${place || '-'}</td>
                <td>${order.fecha_entrega || '-'}</td>
                <td>${order.hora_entrega || '-'}</td>
                <td>${order.costo_total_pedido || '-'}</td>
                <td><select class="status-select">${optionsHtml}</select></td>
                <td>${whatsappCell}</td>
                <td>${order.fecha_registro || ''}</td>
            </tr>
        `;
    }).join('');

    table.querySelectorAll('.status-select').forEach((selectEl) => {
        selectEl.addEventListener('change', (event) => {
            const row = event.target.closest('tr');
            if (!row) return;
            const orderId = row.dataset.id;
            updateStatus(orderId, event.target.value, event.target);
        });
    });

    applyFilters();
}

function updateMetrics(metrics) {
    const values = [
        metrics.pending,
        metrics.delivered,
        metrics.no_entregado,
        metrics.today,
        metrics.personal,
        metrics.envio,
    ];

    kpiValues.forEach((element, index) => {
        if (values[index] !== undefined) {
            element.textContent = values[index];
        }
    });
}

function updateRecentCodes(recentCodes) {
    if (!codeTableBody) return;

    codeTableBody.innerHTML = recentCodes.map((item) => `
        <tr>
            <td>${item.codigo || ''}</td>
            <td>${item.estado_codigo || ''}</td>
            <td>${item.fecha_creacion || '-'}</td>
            <td>${item.fecha_uso || '-'}</td>
            <td>${item.pedido_grupo_id || '-'}</td>
        </tr>
    `).join('');
}

async function refreshDashboard() {
    if (inFlightRefresh) return;

    try {
        inFlightRefresh = true;
        const response = await fetch('/admin/api/orders', { headers: { 'X-Requested-With': 'fetch' } });
        const data = await response.json();
        if (!response.ok || !data.ok) return;

        if (Array.isArray(data.orders)) {
            renderOrders(data.orders);
        }

        if (data.metrics) {
            updateMetrics(data.metrics);
        }

        if (Array.isArray(data.recent_codes)) {
            updateRecentCodes(data.recent_codes);
        }
    } catch (error) {
        console.warn('No se pudo actualizar el panel automaticamente', error);
    } finally {
        inFlightRefresh = false;
    }
}

function clearRefreshTimer() {
    if (refreshTimer) {
        clearTimeout(refreshTimer);
        refreshTimer = null;
    }
}

function getNextRefreshDelay() {
    if (document.visibilityState !== 'visible') {
        return null;
    }

    const isIdle = Date.now() - lastUserActivityAt > IDLE_AFTER_MS;
    if (isIdle) {
        return null;
    }

    return REFRESH_FAST_MS;
}

function scheduleRefresh() {
    clearRefreshTimer();

    const delay = getNextRefreshDelay();
    if (delay === null) return;

    refreshTimer = setTimeout(async () => {
        await refreshDashboard();
        scheduleRefresh();
    }, delay);
}

async function refreshNowAndReschedule() {
    if (document.visibilityState !== 'visible') {
        clearRefreshTimer();
        return;
    }

    await refreshDashboard();
    scheduleRefresh();
}

function registerUserActivity() {
    const wasIdle = Date.now() - lastUserActivityAt > IDLE_AFTER_MS;
    lastUserActivityAt = Date.now();

    if (wasIdle && !refreshTimer) {
        refreshNowAndReschedule();
    }
}

async function updateStatus(orderId, status, selectElement) {
    try {
        const response = await fetch('/admin/status', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ order_id: orderId, status }),
        });

        const data = await response.json();
        if (!response.ok || !data.ok) {
            throw new Error(data.message || 'No se pudo actualizar');
        }

        const row = selectElement.closest('tr');
        if (row) row.dataset.status = status;
        applyFilters();
        refreshNowAndReschedule();
    } catch (error) {
        alert(error.message || 'Error al actualizar estado');
    }
}

if (table) {
    refreshNowAndReschedule();

    ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'].forEach((eventName) => {
        window.addEventListener(eventName, registerUserActivity, { passive: true });
    });

    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') {
            refreshNowAndReschedule();
        } else {
            clearRefreshTimer();
        }
    });

    window.addEventListener('focus', refreshNowAndReschedule);
    window.addEventListener('blur', clearRefreshTimer);
    window.addEventListener('beforeunload', clearRefreshTimer);
}

[searchInput, filterStatus, filterDate, filterTime, filterType, filterPlace]
    .filter(Boolean)
    .forEach((el) => el.addEventListener('input', applyFilters));

[filterStatus, filterDate, filterTime, filterType, filterPlace]
    .filter(Boolean)
    .forEach((el) => el.addEventListener('change', applyFilters));

applyFilters();

async function generateCode() {
    if (!generateCodeBtn || !generatedCode) return;

    try {
        generateCodeBtn.disabled = true;
        const response = await fetch('/admin/codes/generate', { method: 'POST' });
        const data = await response.json();
        if (!response.ok || !data.ok) {
            throw new Error(data.message || 'No se pudo generar codigo');
        }
        generatedCode.textContent = data.code || '- - -';
        refreshNowAndReschedule();
    } catch (error) {
        alert(error.message || 'Error al generar codigo');
    } finally {
        generateCodeBtn.disabled = false;
    }
}

async function copyGeneratedCode() {
    const code = (generatedCode?.textContent || '').trim();
    if (!code || code === '- - -') {
        alert('Primero genera un codigo');
        return;
    }

    try {
        await navigator.clipboard.writeText(code);
        copyCodeBtn.textContent = 'Copiado';
        setTimeout(() => {
            copyCodeBtn.textContent = 'Copiar codigo';
        }, 1200);
    } catch {
        alert('No se pudo copiar automaticamente');
    }
}

if (generateCodeBtn) {
    generateCodeBtn.addEventListener('click', generateCode);
}

if (copyCodeBtn) {
    copyCodeBtn.addEventListener('click', copyGeneratedCode);
}
