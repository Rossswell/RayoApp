const { ipcRenderer } = require('electron');

// --- ELEMENTS ---
const splashView = document.getElementById('splash-view');
const appView = document.getElementById('app-view');
const progressBar = document.getElementById('progress-bar');
const sidebar = document.getElementById('sidebar');
const sidebarToggle = document.getElementById('sidebar-toggle');
const closeBtn = document.getElementById('close-btn');
const minimizeBtn = document.getElementById('minimize-btn');
const dashboardToggle = document.getElementById('dashboard-toggle');
const groupDashboard = document.getElementById('group-dashboard');
const dashboardRendimientoBtn = document.getElementById('dashboard-rendimiento-btn');
const dashboardGananciasBtn = document.getElementById('dashboard-ganancias-btn');
const flashHighToggle = document.getElementById('flash-high-toggle');
const groupFlashHigh = document.getElementById('group-flash-high');
const empleadosToggle = document.getElementById('empleados-toggle');
const groupEmpleados = document.getElementById('group-empleados');
const addEmployeeBtn = document.getElementById('add-employee-btn');
const listEmployeesBtn = document.getElementById('list-employees-btn');
const addEmployeeForm = document.getElementById('add-employee-form');
const cancelAddBtn = document.getElementById('cancel-add-btn');
const employeesBody = document.getElementById('employees-body');
const employeeSearch = document.getElementById('employee-search');
const refreshListBtn = document.getElementById('refresh-list-btn');
const listLoading = document.getElementById('list-loading');
const listEmpty = document.getElementById('list-empty');

const checkUpdateBtn = document.getElementById('check-update-btn');
const updateBtnLabel = document.getElementById('update-btn-label');

// --- STATE ---
let loadProgress = 0;
let allEmployees = [];
let currentEmployeeInModal = null;

// --- INITIALIZATION ---
document.addEventListener('DOMContentLoaded', () => {
    simulateLoading();
    setupEventListeners();
    initHomeWidgets();
});

// --- CORE FUNCTIONS ---

async function simulateLoading() {
    // Animate to ~95% while real data loads in parallel
    const interval = setInterval(() => {
        if (loadProgress < 95) {
            loadProgress += Math.random() * 2;
            if (loadProgress > 95) loadProgress = 95;
            progressBar.style.width = loadProgress + '%';
        }
    }, 150);

    try {
        // Wait for both employees data and attendance stats
        const [empRes, statsRes] = await Promise.all([
            ipcRenderer.invoke('get-employees'),
            ipcRenderer.invoke('get-attendance-stats')
        ]);

        if (empRes.success) {
            allEmployees = empRes.employees || [];
            rendEmployees = [...allEmployees]; // Pre-populate for Rendimiento
        }
        
        if (statsRes.success) {
            rendAllRecords = statsRes.records || [];
        }
    } catch(e) {
        console.error('Critical initial load failed:', e);
    }

    clearInterval(interval);
    loadProgress = 100;
    progressBar.style.width = '100%';
    setTimeout(finishLoading, 400);
}

function finishLoading() {
    splashView.classList.add('fade-out');
    
    setTimeout(() => {
        splashView.classList.remove('active');
        appView.classList.add('active');
        
        sidebar.classList.add('collapsed');
        setTimeout(() => {
            sidebar.classList.add('animate-icons');
        }, 300);

        setTimeout(() => {
            appView.style.opacity = '1';
        }, 50);
    }, 800);
}

function showPage(pageId) {
    const pages = document.querySelectorAll('.page-view');
    pages.forEach(page => page.classList.remove('active'));

    const activePage = document.getElementById(pageId);
    if (activePage) activePage.classList.add('active');

    // Clear all nav active states
    document.querySelectorAll('.nav-item, .submenu-item').forEach(item => item.classList.remove('active'));

    const navMap = {
        'home-view':           null,
        'dashboard-view':      'dashboard-rendimiento-btn',
        'ganancias-view':      'dashboard-ganancias-btn',
        'add-employee-view':   'add-employee-btn',
        'list-employees-view': 'list-employees-btn',
        'assign-schedule-view':'assign-schedule-btn',
        'nomina-view':         'nomina-btn',
    };
    const btnId = navMap[pageId];
    if (btnId) {
        const btn = document.getElementById(btnId);
        if (btn) btn.classList.add('active');
    }

    if (pageId === 'dashboard-view' && typeof loadDashboardData === 'function') {
        loadDashboardData();
    }
}

// ── HOME PAGE ────────────────────────────────────────────────────────────────

(function initHomePage() {
    const DAYS_ES   = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
    const MONTHS_ES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

    // Clock
    function tickClock() {
        const now  = new Date();
        const h    = String(now.getHours()).padStart(2,'0');
        const m    = String(now.getMinutes()).padStart(2,'0');
        const s    = String(now.getSeconds()).padStart(2,'0');
        const el   = document.getElementById('home-clock');
        const elD  = document.getElementById('home-clock-date');
        const elLbl= document.getElementById('home-day-label');
        if (el)   el.textContent  = `${h}:${m}:${s}`;
        if (elD)  elD.textContent = `${DAYS_ES[now.getDay()]}, ${now.getDate()} de ${MONTHS_ES[now.getMonth()]} ${now.getFullYear()}`;
        if (elLbl && !elLbl._set) { elLbl.textContent = DAYS_ES[now.getDay()]; elLbl._set = true; }
    }
    tickClock();
    setInterval(tickClock, 1000);

    // Calendar
    let calDate = new Date();
    function renderCal() {
        const lbl = document.getElementById('cal-month-label');
        const container = document.getElementById('cal-days');
        if (!lbl || !container) return;
        const today = new Date();
        const y = calDate.getFullYear(), mo = calDate.getMonth();
        lbl.textContent = `${MONTHS_ES[mo]} ${y}`;
        const first = new Date(y, mo, 1);
        let startDow = first.getDay(); // 0=Sun
        startDow = startDow === 0 ? 6 : startDow - 1; // shift to Mon=0
        const daysInMonth = new Date(y, mo + 1, 0).getDate();
        let html = '';
        for (let i = 0; i < startDow; i++) html += `<span class="home-cal-day empty"></span>`;
        for (let d = 1; d <= daysInMonth; d++) {
            const isToday = d === today.getDate() && mo === today.getMonth() && y === today.getFullYear();
            html += `<span class="home-cal-day${isToday ? ' today' : ''}">${d}</span>`;
        }
        container.innerHTML = html;
    }
    renderCal();
    document.getElementById('cal-prev')?.addEventListener('click', () => { calDate.setMonth(calDate.getMonth()-1); renderCal(); });
    document.getElementById('cal-next')?.addEventListener('click', () => { calDate.setMonth(calDate.getMonth()+1); renderCal(); });

    // Dollar rates
    async function loadRates() {
        const list = document.getElementById('rates-list');
        const upd  = document.getElementById('rates-updated');
        if (!list) return;
        try {
            const res  = await fetch('https://ve.dolarapi.com/v1/dolares');
            const data = await res.json();
            const nameMap = { oficial:'BCV Oficial', paralelo:'Paralelo (USDT)', promedio:'Promedio', cripto:'Cripto', coinbase:'Coinbase', yadio:'Yadio' };
            list.innerHTML = data.map(r => {
                const label = nameMap[r.fuente?.toLowerCase()] || r.nombre || r.fuente;
                const val   = r.promedio ? Number(r.promedio).toLocaleString('es-VE', { minimumFractionDigits:2, maximumFractionDigits:2 }) : '—';
                return `<div class="home-rate-row">
                    <span class="home-rate-name">${label}</span>
                    <div style="display:flex;align-items:center;gap:8px;">
                        <span class="home-rate-value">Bs. ${val}</span>
                        <span class="home-rate-badge">VES</span>
                    </div>
                </div>`;
            }).join('');
            if (upd) upd.textContent = 'Actualizado ' + new Date().toLocaleTimeString('es-VE', {hour:'2-digit',minute:'2-digit'});
        } catch {
            if (list) list.innerHTML = '<div class="home-rates-loading">No se pudieron cargar las tasas.</div>';
        }
    }
    loadRates();
    setInterval(loadRates, 5 * 60 * 1000); // refresh every 5 min
})();

// --- EVENT LISTENERS ---

function setupEventListeners() {
    // --- CUSTOM DATE PICKER LOGIC ---
class CalendarPicker {
    constructor(inputId) {
        this.input = document.getElementById(inputId);
        this.currentDate = new Date();
        this.selectedDate = null;
        this.isOpen = false;
        this.view = 'days'; // 'days', 'months', 'years'
        
        this.picker = document.createElement('div');
        this.picker.className = 'calendar-picker';
        document.body.appendChild(this.picker);
        
        this.init();
    }

    init() {
        if (!this.input) return;
        this.input.classList.add('date-trigger');
        this.input.readOnly = true;

        this.input.addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggle();
        });

        document.addEventListener('click', (e) => {
            if (this.isOpen && !this.picker.contains(e.target) && e.target !== this.input) {
                this.close();
            }
        });

        this.render();
    }

    toggle() {
        if (this.isOpen) this.close();
        else this.open();
    }

    open() {
        const rect = this.input.getBoundingClientRect();
        this.picker.style.top = `${rect.bottom + window.scrollY + 5}px`;
        this.picker.style.left = `${rect.left + window.scrollX}px`;
        this.view = 'days';
        this.render();
        this.picker.classList.add('active');
        this.isOpen = true;
    }

    close() {
        this.picker.classList.remove('active');
        this.isOpen = false;
    }

    switchView(newView) {
        this.view = newView;
        this.render();
    }

    render() {
        const year = this.currentDate.getFullYear();
        const monthName = new Intl.DateTimeFormat('es-ES', { month: 'long' }).format(this.currentDate);

        this.picker.innerHTML = `
            <div class="calendar-header">
                <div class="calendar-title">
                    <span class="month-label">${monthName.charAt(0).toUpperCase() + monthName.slice(1)}</span>
                    <span class="year-label">${year}</span>
                </div>
                <div class="calendar-nav">
                    <button class="nav-btn prev-btn">‹</button>
                    <button class="nav-btn next-btn">›</button>
                </div>
            </div>
            <div class="calendar-grid grid-${this.view}">
                ${this.renderViewContent()}
            </div>
        `;

        this.setupHeaderEvents();
        this.setupNavEvents();
        this.setupGridEvents();
    }

    renderViewContent() {
        const year = this.currentDate.getFullYear();
        const month = this.currentDate.getMonth();

        if (this.view === 'days') {
            const firstDay = new Date(year, month, 1).getDay();
            const daysInMonth = new Date(year, month + 1, 0).getDate();
            const prevMonthDays = new Date(year, month, 0).getDate();
            
            let html = `
                <div class="calendar-day-name">Lu</div><div class="calendar-day-name">Ma</div>
                <div class="calendar-day-name">Mi</div><div class="calendar-day-name">Ju</div>
                <div class="calendar-day-name">Vi</div><div class="calendar-day-name">Sa</div>
                <div class="calendar-day-name">Do</div>
            `;
            
            // Monday start adjustment
            let startingDay = firstDay === 0 ? 6 : firstDay - 1;
            for (let i = startingDay; i > 0; i--) {
                html += `<div class="calendar-item calendar-day other-month">${prevMonthDays - i + 1}</div>`;
            }

            for (let d = 1; d <= daysInMonth; d++) {
                const isToday = new Date().toDateString() === new Date(year, month, d).toDateString();
                const isSelected = this.selectedDate?.toDateString() === new Date(year, month, d).toDateString();
                html += `<div class="calendar-item calendar-day ${isToday ? 'today' : ''} ${isSelected ? 'selected' : ''}">${d}</div>`;
            }
            return html;
        }

        if (this.view === 'months') {
            const months = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
            return months.map((m, i) => {
                const isSelected = this.currentDate.getMonth() === i;
                return `<div class="calendar-item calendar-month ${isSelected ? 'selected' : ''}" data-month="${i}">${m}</div>`;
            }).join('');
        }

        if (this.view === 'years') {
            const startYear = year - (year % 16);
            let html = '';
            for (let i = 0; i < 16; i++) {
                const y = startYear + i;
                const isSelected = year === y;
                html += `<div class="calendar-item calendar-year ${isSelected ? 'selected' : ''}" data-year="${y}">${y}</div>`;
            }
            return html;
        }
    }

    setupHeaderEvents() {
        const mLabel = this.picker.querySelector('.month-label');
        const yLabel = this.picker.querySelector('.year-label');

        mLabel.onclick = (e) => { e.stopPropagation(); this.switchView('months'); };
        yLabel.onclick = (e) => { e.stopPropagation(); this.switchView('years'); };
    }

    setupNavEvents() {
        const prevBtn = this.picker.querySelector('.prev-btn');
        const nextBtn = this.picker.querySelector('.next-btn');

        prevBtn.onclick = (e) => {
            e.stopPropagation();
            if (this.view === 'days') this.currentDate.setMonth(this.currentDate.getMonth() - 1);
            else if (this.view === 'years') this.currentDate.setFullYear(this.currentDate.getFullYear() - 16);
            this.render();
        };

        nextBtn.onclick = (e) => {
            e.stopPropagation();
            if (this.view === 'days') this.currentDate.setMonth(this.currentDate.getMonth() + 1);
            else if (this.view === 'years') this.currentDate.setFullYear(this.currentDate.getFullYear() + 16);
            this.render();
        };
    }

    setupGridEvents() {
        this.picker.querySelectorAll('.calendar-item').forEach(item => {
            item.onclick = (e) => {
                e.stopPropagation();
                if (this.view === 'days' && !item.classList.contains('other-month')) {
                    const day = parseInt(item.innerText);
                    this.selectDate(new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), day));
                } else if (this.view === 'months') {
                    this.currentDate.setMonth(parseInt(item.dataset.month));
                    this.switchView('days');
                } else if (this.view === 'years') {
                    this.currentDate.setFullYear(parseInt(item.dataset.year));
                    this.switchView('months');
                }
            };
        });
    }

    selectDate(date) {
        this.selectedDate = date;
        const d = String(date.getDate()).padStart(2, '0');
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const y = date.getFullYear();
        this.input.value = `${d}/${m}/${y}`;
        this.close();
    }
}

// Sidebar Toggle
    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
            if (sidebar.classList.contains('collapsed')) {
                document.querySelectorAll('.nav-group').forEach(group => group.classList.remove('open'));
            }
        });
    }

    // Company Group Toggle
    if (flashHighToggle && groupFlashHigh) {
        flashHighToggle.addEventListener('click', () => {
            if (sidebar.classList.contains('collapsed')) {
                sidebar.classList.remove('collapsed');
                setTimeout(() => groupFlashHigh.classList.add('open'), 200);
            } else {
                groupFlashHigh.classList.toggle('open');
            }
            showPage('home-view');
        });
    }

    // Submenu Toggle
    if (empleadosToggle && groupEmpleados) {
        empleadosToggle.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent company toggle from triggering
            if (sidebar.classList.contains('collapsed')) {
                sidebar.classList.remove('collapsed');
                setTimeout(() => groupEmpleados.classList.add('open'), 200);
            } else {
                groupEmpleados.classList.toggle('open');
            }
        });
    }

    // Dashboard Submenu Toggle
    if (dashboardToggle && groupDashboard) {
        dashboardToggle.addEventListener('click', (e) => {
            e.stopPropagation();
            if (sidebar.classList.contains('collapsed')) {
                sidebar.classList.remove('collapsed');
                setTimeout(() => groupDashboard.classList.add('open'), 200);
            } else {
                groupDashboard.classList.toggle('open');
            }
        });
    }

    // Navigation
    if (dashboardRendimientoBtn) {
        dashboardRendimientoBtn.addEventListener('click', (e) => {
            e.preventDefault();
            showPage('dashboard-view');
        });
    }

    if (dashboardGananciasBtn) {
        dashboardGananciasBtn.addEventListener('click', (e) => {
            e.preventDefault();
            showPage('ganancias-view');
        });
    }

    if (addEmployeeBtn) {
        addEmployeeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            showPage('add-employee-view');
            
            // Initialize pickers if not already initialized
            if (!window.startDatePicker) window.startDatePicker = new CalendarPicker('emp-inicio');
            if (!window.birthdayPicker) window.birthdayPicker = new CalendarPicker('emp-cumple');
        });
    }

    // Window Controls
    if (closeBtn) closeBtn.addEventListener('click', () => ipcRenderer.send('app-close'));
    if (minimizeBtn) minimizeBtn.addEventListener('click', () => ipcRenderer.send('app-minimize'));

    // Form Handling
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const submitBtn = addEmployeeForm.querySelector('button[type="submit"]');
            const originalBtnText = submitBtn.innerText;
            
            // Disable form during submission
            submitBtn.disabled = true;
            submitBtn.innerText = 'Guardando...';
            submitBtn.classList.add('loading');

            const formData = new FormData(addEmployeeForm);
            const employeeData = Object.fromEntries(formData.entries());
            
            try {
                const result = await ipcRenderer.invoke('add-employee-to-sheet', employeeData);
                
                if (result.success) {
                    alert('✅ Empleado registrado exitosamente en Google Sheets.');
                    addEmployeeForm.reset();
                    showPage('dashboard-view');
                } else {
                    alert('❌ Error al registrar: ' + result.error);
                }
            } catch (error) {
                console.error('IPC Error:', error);
                alert('❌ Error de conexión al intentar guardar.');
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerText = originalBtnText;
                submitBtn.classList.remove('loading');
            }
        });
    }

if (cancelAddBtn) {
        cancelAddBtn.addEventListener('click', () => {
            addEmployeeForm.reset();
            showPage('dashboard-view');
        });
    }

    // --- EMPLOYEES LIST LOGIC ---

    if (listEmployeesBtn) {
        listEmployeesBtn.addEventListener('click', (e) => {
            e.preventDefault();
            showPage('list-employees-view');
            loadEmployees();
        });
    }

    if (refreshListBtn) {
        refreshListBtn.addEventListener('click', () => {
            loadEmployees();
        });
    }

    if (employeeSearch) {
        employeeSearch.addEventListener('input', (e) => {
            renderEmployees(e.target.value);
        });
    }

    async function loadEmployees() {
        if (listLoading) listLoading.style.display = 'block';
        if (listEmpty) listEmpty.style.display = 'none';
        if (employeesBody) employeesBody.innerHTML = '';
        if (refreshListBtn) refreshListBtn.classList.add('loading');

        try {
            const result = await ipcRenderer.invoke('get-employees');
            if (result.success) {
                allEmployees = result.employees;
                renderEmployees();
            } else {
                alert('Error al obtener lista: ' + result.error);
            }
        } catch (error) {
            console.error('Fetch Error:', error);
            alert('Error de conexión con el backend.');
        } finally {
            if (listLoading) listLoading.style.display = 'none';
            if (refreshListBtn) refreshListBtn.classList.remove('loading');
        }
    }

    // Avatar color palette
    const AVATAR_COLORS = [
        { bg: 'linear-gradient(135deg,#fde68a,#f59e0b)', color: '#78350f' },
        { bg: 'linear-gradient(135deg,#a5f3fc,#0891b2)', color: '#ffffff' },
        { bg: 'linear-gradient(135deg,#bbf7d0,#16a34a)', color: '#ffffff' },
        { bg: 'linear-gradient(135deg,#fecaca,#dc2626)', color: '#ffffff' },
        { bg: 'linear-gradient(135deg,#e9d5ff,#7c3aed)', color: '#ffffff' },
        { bg: 'linear-gradient(135deg,#fed7aa,#ea580c)', color: '#ffffff' },
        { bg: 'linear-gradient(135deg,#fbcfe8,#db2777)', color: '#ffffff' },
    ];

    function getAvatarColor(name) {
        let hash = 0;
        for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
        return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
    }

    function getInitials(name) {
        return name.split(' ').map(n => n[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();
    }

    function renderEmployees(searchTerm = '') {
        if (!employeesBody) return;

        const filtered = allEmployees.filter(emp => {
            const s = searchTerm.toLowerCase();
            return emp.nombre.toLowerCase().includes(s) || 
                   emp.cedula.toString().toLowerCase().includes(s) || 
                   emp.area.toLowerCase().includes(s);
        });

        if (filtered.length === 0) {
            employeesBody.innerHTML = '';
            if (listEmpty) listEmpty.style.display = 'block';
            return;
        }

        if (listEmpty) listEmpty.style.display = 'none';

        employeesBody.innerHTML = filtered.map((emp, idx) => {
            const color = getAvatarColor(emp.nombre);
            const initials = getInitials(emp.nombre);
            return `
            <div class="team-card" data-cedula="${emp.cedula}" style="animation-delay:${idx * 0.05}s">
                <div class="team-card-avatar" style="background:${color.bg};color:${color.color}">${initials}</div>
                <div class="team-card-body">
                    <div class="team-card-name">${emp.nombre}</div>
                    <span class="team-card-area">${emp.area}</span>
                    <div class="team-card-meta">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 13a19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 3.6 2.19h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L7.91 9.91a16 16 0 0 0 6.18 6.18l1.8-1.84a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg>
                        <span>${emp.telefono}</span>
                    </div>
                    <div class="team-card-meta">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>
                        <span>${emp.correo}</span>
                    </div>
                </div>
                <div class="team-card-footer">
                    <span class="team-card-since">Desde ${emp.fecha_inicio}</span>
                    <button class="team-card-view-btn">Ver perfil →</button>
                </div>
            </div>
            `;
        }).join('');

        // Attach click handlers to open profile modal
        employeesBody.querySelectorAll('.team-card').forEach(card => {
            card.addEventListener('click', () => {
                const emp = allEmployees.find(e => e.cedula === card.dataset.cedula);
                if (emp) openProfileModal(emp);
            });
        });
    }

    // --- PROFILE MODAL ---

    const profileModal = document.getElementById('employee-profile-modal');
    const closeProfileModalBtn = document.getElementById('close-profile-modal');

    if (closeProfileModalBtn) {
        closeProfileModalBtn.addEventListener('click', closeProfileModal);
    }
    if (profileModal) {
        profileModal.addEventListener('click', (e) => {
            if (e.target === profileModal) closeProfileModal();
        });
    }

    // Delete Employee Button
    const deleteEmployeeBtn = document.getElementById('delete-employee-btn');
    if (deleteEmployeeBtn) {
        deleteEmployeeBtn.addEventListener('click', async () => {
            if (!currentEmployeeInModal) return;
            
            const confirmDelete = confirm(`¿Está seguro que desea eliminar al empleado "${currentEmployeeInModal.nombre}"?\n\nEsta acción no se puede deshacer.`);
            if (!confirmDelete) return;

            deleteEmployeeBtn.disabled = true;
            deleteEmployeeBtn.innerHTML = '<div class="loader" style="width:16px;height:16px;"></div> Eliminando...';

            try {
                const result = await ipcRenderer.invoke('delete-employee', { cedula: currentEmployeeInModal.cedula });
                
                if (result.success) {
                    alert('✅ Empleado eliminado exitosamente.');
                    closeProfileModal();
                    // Refresh the employee list
                    loadEmployees();
                } else {
                    alert('❌ Error al eliminar: ' + result.error);
                }
            } catch (error) {
                console.error('Delete Error:', error);
                alert('❌ Error de conexión al intentar eliminar.');
            } finally {
                deleteEmployeeBtn.disabled = false;
                deleteEmployeeBtn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg> Eliminar Empleado`;
            }
        });
    }

    function closeProfileModal() {
        if (profileModal) {
            profileModal.classList.remove('active');
            setTimeout(() => { profileModal.style.display = 'none'; }, 300);
        }
    }

    async function openProfileModal(emp) {
        if (!profileModal) return;
        profileModal.style.display = 'flex';
        setTimeout(() => profileModal.classList.add('active'), 10);

        const color = getAvatarColor(emp.nombre);
        const initials = getInitials(emp.nombre);

        // Fill in data
        const avatarEl = document.getElementById('modal-avatar');
        avatarEl.textContent = initials;
        avatarEl.style.background = color.bg;
        avatarEl.style.color = color.color;

        document.getElementById('modal-nombre').textContent = emp.nombre;
        document.getElementById('modal-area').textContent = emp.area;
        document.getElementById('modal-correo').textContent = emp.correo;
        document.getElementById('modal-telefono').textContent = emp.telefono;
        document.getElementById('modal-cedula').textContent = emp.cedula;
        document.getElementById('modal-inicio').textContent = emp.fecha_inicio;
        document.getElementById('modal-cumple').textContent = emp.fecha_cumple;
        document.getElementById('modal-ubicacion').textContent = emp.ubicacion || 'Sin asignar';

        // Store current employee for delete functionality
        currentEmployeeInModal = emp;

        // Load schedule
        const scheduleEl = document.getElementById('modal-schedule');
        scheduleEl.innerHTML = `<div class="modal-schedule-loading"><div class="loader" style="width:18px;height:18px;"></div> Cargando horario...</div>`;

        try {
            const result = await ipcRenderer.invoke('get-employee-schedule', { cedula: emp.cedula });
            if (result && result.success && result.schedule) {
                renderModalSchedule(result.schedule, scheduleEl);
            } else {
                scheduleEl.innerHTML = `<div class="modal-schedule-empty">No hay horario asignado aún.</div>`;
            }
        } catch (e) {
            scheduleEl.innerHTML = `<div class="modal-schedule-empty">No hay horario asignado aún.</div>`;
        }
    }

    // --- EDIT EMPLOYEE MODAL ---

    const editModal = document.getElementById('edit-employee-modal');
    const closeEditModalBtn = document.getElementById('close-edit-modal');
    const cancelEditBtn = document.getElementById('cancel-edit-btn');
    const editEmployeeForm = document.getElementById('edit-employee-form');
    const editEmployeeBtn = document.getElementById('edit-employee-btn');

    if (closeEditModalBtn) {
        closeEditModalBtn.addEventListener('click', closeEditModal);
    }
    if (cancelEditBtn) {
        cancelEditBtn.addEventListener('click', closeEditModal);
    }
    if (editModal) {
        editModal.addEventListener('click', (e) => {
            if (e.target === editModal) closeEditModal();
        });
    }

    // Open edit modal from profile
    if (editEmployeeBtn) {
        editEmployeeBtn.addEventListener('click', () => {
            if (!currentEmployeeInModal) return;
            closeProfileModal();
            openEditModal(currentEmployeeInModal);
        });
    }

    function openEditModal(emp) {
        if (!editModal) return;

        const color = getAvatarColor(emp.nombre);
        const initials = getInitials(emp.nombre);

        const avatarEl = document.getElementById('edit-avatar');
        avatarEl.textContent = initials;
        avatarEl.style.background = color.bg;
        avatarEl.style.color = color.color;

        document.getElementById('edit-area-badge').textContent = emp.area;

        // Fill form fields
        document.getElementById('edit-nombre').value = emp.nombre;
        document.getElementById('edit-cedula').value = emp.cedula;
        document.getElementById('edit-telefono').value = emp.telefono;
        document.getElementById('edit-correo').value = emp.correo;
        document.getElementById('edit-area').value = emp.area;
        document.getElementById('edit-inicio').value = emp.fecha_inicio;
        document.getElementById('edit-cumple').value = emp.fecha_cumple;
        document.getElementById('edit-ubicacion').value = emp.ubicacion || '';
        document.getElementById('edit-original-cedula').value = emp.cedula;

        editModal.style.display = 'flex';
        setTimeout(() => editModal.classList.add('active'), 10);
    }

    function closeEditModal() {
        if (editModal) {
            editModal.classList.remove('active');
            setTimeout(() => { editModal.style.display = 'none'; }, 300);
        }
    }

    // Handle edit form submission
    if (editEmployeeForm) {
        editEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const submitBtn = editEmployeeForm.querySelector('button[type="submit"]');
            const originalBtnText = submitBtn.innerText;

            submitBtn.disabled = true;
            submitBtn.innerText = 'Guardando...';
            submitBtn.classList.add('loading');

            const formData = new FormData(editEmployeeForm);
            const employeeData = Object.fromEntries(formData.entries());
            const originalCedula = employeeData.original_cedula;

            try {
                const result = await ipcRenderer.invoke('update-employee', {
                    originalCedula,
                    employeeData
                });

                if (result.success) {
                    alert('✅ Empleado actualizado exitosamente.');
                    closeEditModal();
                    loadEmployees();
                } else {
                    alert('❌ Error al actualizar: ' + result.error);
                }
            } catch (error) {
                console.error('Update Error:', error);
                alert('❌ Error de conexión al intentar actualizar.');
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerText = originalBtnText;
                submitBtn.classList.remove('loading');
            }
        });
    }

    function renderModalSchedule(schedule, container) {
        const days = ['lunes','martes','miercoles','jueves','viernes','sabado','domingo'];
        const dayNames = { lunes:'Lun', martes:'Mar', miercoles:'Mié', jueves:'Jue', viernes:'Vie', sabado:'Sáb', domingo:'Dom' };

        const html = days.map(day => {
            const d = schedule[day];
            const active = d && d.activo;
            return `
            <div class="schedule-pill ${active ? 'active' : 'inactive'}">
                <span class="schedule-pill-day">${dayNames[day]}</span>
                ${active ? `<span class="schedule-pill-time">${d.entrada} – ${d.salida}</span>` : `<span class="schedule-pill-off">Libre</span>`}
            </div>`;
        }).join('');

        container.innerHTML = html;
    }

    // --- SCHEDULE ASSIGNMENT LOGIC ---

    const assignScheduleBtn = document.getElementById('assign-schedule-btn');
    const scheduleEmployeeSearch = document.getElementById('schedule-employee-search');
    const scheduleEmployeeList = document.getElementById('schedule-employee-list');
    const scheduleEmployeeLoading = document.getElementById('schedule-employee-loading');
    const scheduleEmployeeEmpty = document.getElementById('schedule-employee-empty');
    const selectedEmployeeInfo = document.getElementById('selected-employee-info');
    const noEmployeeSelected = document.getElementById('no-employee-selected');
    const scheduleFormContainer = document.getElementById('schedule-form-container');
    const cancelScheduleBtn = document.getElementById('cancel-schedule-btn');
    const saveScheduleBtn = document.getElementById('save-schedule-btn');

    let selectedEmployeeForSchedule = null;
    let scheduleLoadToken = 0;
    let scheduleEmployees = [];

    // Navigate to Schedule Assignment View
    if (assignScheduleBtn) {
        assignScheduleBtn.addEventListener('click', (e) => {
            e.preventDefault();
            showPage('assign-schedule-view');
            loadScheduleEmployees();
        });
    }

    // Load employees for schedule assignment
    async function loadScheduleEmployees() {
        if (scheduleEmployeeLoading) scheduleEmployeeLoading.style.display = 'block';
        if (scheduleEmployeeEmpty) scheduleEmployeeEmpty.style.display = 'none';
        if (scheduleEmployeeList) scheduleEmployeeList.innerHTML = '';

        try {
            const result = await ipcRenderer.invoke('get-employees');
            if (result.success) {
                scheduleEmployees = result.employees;
                renderScheduleEmployees();
            } else {
                console.error('Error loading employees:', result.error);
            }
        } catch (error) {
            console.error('Fetch Error:', error);
        } finally {
            if (scheduleEmployeeLoading) scheduleEmployeeLoading.style.display = 'none';
        }
    }

    function renderScheduleEmployees(searchTerm = '') {
        if (!scheduleEmployeeList) return;

        const filtered = scheduleEmployees.filter(emp => {
            const s = searchTerm.toLowerCase();
            return emp.nombre.toLowerCase().includes(s) || 
                   emp.cedula.toString().toLowerCase().includes(s) || 
                   emp.area.toLowerCase().includes(s);
        });

        if (filtered.length === 0) {
            scheduleEmployeeList.innerHTML = '';
            if (scheduleEmployeeEmpty) scheduleEmployeeEmpty.style.display = 'block';
            return;
        }

        if (scheduleEmployeeEmpty) scheduleEmployeeEmpty.style.display = 'none';

        scheduleEmployeeList.innerHTML = filtered.map(emp => {
            const initials = emp.nombre.split(' ').map(n => n[0]).slice(0, 2).join('').toUpperCase();
            const isSelected = selectedEmployeeForSchedule?.cedula === emp.cedula;
            return `
                <div class="employee-card ${isSelected ? 'selected' : ''}" data-cedula="${emp.cedula}">
                    <div class="employee-avatar">${initials}</div>
                    <div class="employee-info">
                        <div class="name">${emp.nombre}</div>
                        <div class="area">${emp.area}</div>
                    </div>
                    <div class="check-icon">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                            <polyline points="20 6 9 17 4 12"/>
                        </svg>
                    </div>
                </div>
            `;
        }).join('');

        // Add click handlers
        scheduleEmployeeList.querySelectorAll('.employee-card').forEach(card => {
            card.addEventListener('click', () => {
                const cedula = card.dataset.cedula;
                const employee = scheduleEmployees.find(e => e.cedula === cedula);
                selectEmployeeForSchedule(employee);
            });
        });
    }

    function selectEmployeeForSchedule(employee) {
        selectedEmployeeForSchedule = employee;

        // Update UI
        if (selectedEmployeeInfo) {
            selectedEmployeeInfo.style.display = 'flex';
            selectedEmployeeInfo.querySelector('.employee-name').textContent = employee.nombre;
        }

        if (noEmployeeSelected) noEmployeeSelected.style.display = 'none';
        const scheduleLoading = document.getElementById('schedule-loading');
        if (scheduleLoading) scheduleLoading.style.display = 'flex';
        if (scheduleFormContainer) scheduleFormContainer.style.display = 'none';
        const toleranceInput = document.getElementById('schedule-tolerance-input');
        if (toleranceInput) toleranceInput.style.display = 'flex';

        // Update selected card styling
        if (scheduleEmployeeList) {
            scheduleEmployeeList.querySelectorAll('.employee-card').forEach(card => {
                card.classList.toggle('selected', card.dataset.cedula === employee.cedula);
            });
        }

        // Load existing schedule
        loadExistingSchedule(employee.cedula);
    }

    function schedTo24h(timeStr) {
        if (!timeStr) return '08:00';
        const parts = timeStr.toLowerCase().trim().split(' ');
        let [h, m] = parts[0].split(':').map(Number);
        const ampm = parts[1];
        if (ampm === 'pm' && h < 12) h += 12;
        if (ampm === 'am' && h === 12) h = 0;
        return `${String(h).padStart(2, '0')}:${String(m || 0).padStart(2, '0')}`;
    }

    async function loadExistingSchedule(cedula) {
        const token = ++scheduleLoadToken;
        const loadingEl = document.getElementById('schedule-loading');
        const formEl    = document.getElementById('schedule-form-container');

        if (loadingEl) { loadingEl.style.display = 'flex'; }
        if (formEl)    { formEl.style.display    = 'none'; }

        try {
            const result = await ipcRenderer.invoke('get-employee-schedule', { cedula });
            if (token !== scheduleLoadToken) return; // respuesta obsoleta, ignorar

            resetScheduleForm();
            if (result.success && result.schedule) {
                if (result.tolerance) {
                    document.getElementById('schedule-tolerance').value = result.tolerance;
                }
                const days = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
                days.forEach(day => {
                    const dayData = result.schedule[day];
                    const toggle  = document.querySelector(`.day-toggle[data-day="${day}"]`);
                    if (!toggle) return;
                    if (dayData && dayData.activo) {
                        toggle.checked = true;
                        updateDayInputs(day, true);
                        const entryInput = document.querySelector(`.entry-time[data-day="${day}"]`);
                        const exitInput  = document.querySelector(`.exit-time[data-day="${day}"]`);
                        if (entryInput && dayData.entrada) entryInput.value = schedTo24h(dayData.entrada);
                        if (exitInput  && dayData.salida)  exitInput.value  = schedTo24h(dayData.salida);
                    } else {
                        toggle.checked = false;
                        updateDayInputs(day, false);
                    }
                });
            }
        } catch (e) {
            console.error(e);
        } finally {
            if (token === scheduleLoadToken) {
                if (loadingEl) loadingEl.style.display = 'none';
                if (formEl)    formEl.style.display    = 'block';
            }
        }
    }

    function resetScheduleForm() {
        const dayToggles = document.querySelectorAll('.day-toggle');
        dayToggles.forEach(toggle => {
            const day = toggle.dataset.day;
            const isWeekend = day === 'sabado' || day === 'domingo';
            toggle.checked = !isWeekend;
            updateDayInputs(day, !isWeekend);
        });

        // Reset time values
        const timeInputs = document.querySelectorAll('.time-input');
        timeInputs.forEach(input => {
            const day = input.dataset.day;
            if (day === 'sabado' || day === 'domingo') {
                input.value = '08:00';
            } else {
                if (input.classList.contains('entry-time')) {
                    input.value = '08:00';
                } else {
                    input.value = '17:00';
                }
            }
        });
    }

    function updateDayInputs(day, enabled) {
        const dayCard = document.querySelector(`.schedule-day:has([data-day="${day}"])`);
        if (dayCard) {
            const inputs = dayCard.querySelectorAll('.time-input');
            const dayInputs = dayCard.querySelector('.day-inputs');
            
            if (enabled) {
                dayInputs?.classList.remove('disabled');
                inputs.forEach(input => input.disabled = false);
            } else {
                dayInputs?.classList.add('disabled');
                inputs.forEach(input => input.disabled = true);
            }
        }
    }

    // Day toggle handlers
    document.querySelectorAll('.day-toggle').forEach(toggle => {
        toggle.addEventListener('change', (e) => {
            const day = e.target.dataset.day;
            updateDayInputs(day, e.target.checked);
        });
    });

    // Search handler for schedule employees
    if (scheduleEmployeeSearch) {
        scheduleEmployeeSearch.addEventListener('input', (e) => {
            renderScheduleEmployees(e.target.value);
        });
    }

    // Cancel button
    if (cancelScheduleBtn) {
        cancelScheduleBtn.addEventListener('click', () => {
            selectedEmployeeForSchedule = null;
            if (selectedEmployeeInfo) selectedEmployeeInfo.style.display = 'none';
            if (noEmployeeSelected) noEmployeeSelected.style.display = 'flex';
            if (scheduleFormContainer) scheduleFormContainer.style.display = 'none';
            if (scheduleEmployeeList) {
                scheduleEmployeeList.querySelectorAll('.employee-card').forEach(card => {
                    card.classList.remove('selected');
                });
            }
        });
    }

    // Save schedule button
    if (saveScheduleBtn) {
        saveScheduleBtn.addEventListener('click', async () => {
            if (!selectedEmployeeForSchedule) {
                alert('Por favor seleccione un empleado.');
                return;
            }

            let summaryParts = [];
            const days = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
            const dayLabels = { lunes: 'Lunes', martes: 'Martes', miercoles: 'Miércoles', jueves: 'Jueves', viernes: 'Viernes', sabado: 'Sábado', domingo: 'Domingo' };

            days.forEach(day => {
                const toggle = document.querySelector(`.day-toggle[data-day="${day}"]`);
                if (toggle?.checked) {
                    const entryInput = document.querySelector(`.entry-time[data-day="${day}"]`);
                    const exitInput = document.querySelector(`.exit-time[data-day="${day}"]`);
                    
                    const formatTime = (timeStr) => {
                        const [h, m] = timeStr.split(':');
                        const hInt = parseInt(h);
                        const ampm = hInt >= 12 ? 'pm' : 'am';
                        const h12 = hInt % 12 || 12;
                        return `${h12}:${m} ${ampm}`;
                    };

                    const entry = formatTime(entryInput.value);
                    const exit = formatTime(exitInput.value);
                    summaryParts.push(`${dayLabels[day]} (${entry}) - (${exit})`);
                }
            });

            const tolerance = document.getElementById('schedule-tolerance').value;

            const scheduleData = {
                employee: selectedEmployeeForSchedule,
                scheduleSummary: summaryParts.join('; '),
                tolerance: tolerance
            };

            // Disable button during save
            saveScheduleBtn.disabled = true;
            saveScheduleBtn.innerHTML = `
                <div class="loader" style="width: 16px; height: 16px; margin-right: 8px;"></div>
                Guardando...
            `;

            try {
                const result = await ipcRenderer.invoke('save-employee-schedule', scheduleData);
                
                if (result.success) {
                    alert('✅ Horario asignado exitosamente.');
                } else {
                    alert('❌ Error al guardar horario: ' + result.error);
                }
            } catch (error) {
                console.error('IPC Error:', error);
                alert('❌ Error de conexión al intentar guardar.');
            } finally {
                saveScheduleBtn.disabled = false;
                saveScheduleBtn.innerHTML = `
                    <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                        <polyline points="17 21 17 13 7 13 7 21"/>
                        <polyline points="7 3 7 8 15 8"/>
                    </svg>
                    Guardar Horario
                `;
            }
        });
    }
}

// --- RENDIMIENTO EMPLEADOS ---
let rendAllRecords     = [];
let rendEmployees      = [];
let rendSelectedCedula = null;
let rendCurrentWeek    = null;
let rendKpiFilter      = 'semana'; // 'semana' | 'total'
let rendCharts         = {};

function horaToHours(val) {
    if (val === null || val === undefined || val === '') return null;
    const f = parseFloat(val);
    if (!isNaN(f) && f >= 0 && f < 1)  return f * 24;
    if (!isNaN(f) && f >= 1 && f <= 24) return f;
    if (typeof val === 'string' && val.includes(':')) {
        const p = val.split(':');
        return parseInt(p[0]||0) + parseInt(p[1]||0)/60 + parseInt(p[2]||0)/3600;
    }
    return null;
}

function parseWeekNum(semStr) {
    const m = String(semStr||'').match(/(\d+)/);
    return m ? parseInt(m[1]) : null;
}

function parseFechaDate(val) {
    if (!val) return null;
    let d;
    if (val instanceof Date) {
        d = new Date(val.getTime());
    } else if (typeof val === 'string') {
        if (/^\d{4}-\d{2}-\d{2}/.test(val)) {
            const p = val.split('-');
            d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
        } else {
            const p = val.split('/');
            if (p.length === 3) d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
        }
    } else if (typeof val === 'number') {
        // Serial de Excel/Sheets (UTC) -> Local
        const utcDate = new Date(Math.round((val - 25569) * 86400000));
        d = new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate());
    }
    
    if (d && !isNaN(d.getTime())) {
        d.setHours(12, 0, 0, 0); // Forzar mediodía para evitar saltos de día por zona horaria
        return d;
    }
    return null;
}

function normFechaKey(val) {
    const d = parseFechaDate(val);
    if (!d) return String(val);
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function getDateOfISOWeek(w, y) {
    const jan4 = new Date(y, 0, 4);
    const monday = new Date(jan4);
    monday.setDate(jan4.getDate() - (jan4.getDay() + 6) % 7);
    const result = new Date(monday);
    result.setDate(monday.getDate() + (w - 1) * 7);
    result.setHours(12, 0, 0, 0); // Forzar mediodía
    return result;
}

function isoWeekNow() {
    const d = new Date();
    d.setHours(0,0,0,0);
    d.setDate(d.getDate() + 3 - (d.getDay()+6)%7);
    const w1 = new Date(d.getFullYear(), 0, 4);
    return 1 + Math.round(((d - w1) / 86400000 - 3 + (w1.getDay()+6)%7) / 7);
}

function buildDayMap(cedula) {
    const map = {};
    rendAllRecords.filter(r => String(r.cedula).trim() === String(cedula).trim()).forEach(r => {
        const key = normFechaKey(r.fecha);
        if (!map[key]) map[key] = { in: null, out: null, mensajeSistema: '', mensajeSalidaSistema: '', justifEntrada: '', justifSalidaAnticipada: '', justifSalidaTardia: '', weekNum: parseWeekNum(r.semana) };
        const h = horaToHours(r.hora);
        if (r.tipo === 'ENTRADA') {
            map[key].in = h;
            map[key].mensajeSistema = r.mensaje || r.mensajeSistema || '';
            if (r.justifEntrada) map[key].justifEntrada = r.justifEntrada;
        }
        if (r.tipo === 'SALIDA') {
            map[key].out = h;
            map[key].mensajeSalidaSistema = r.mensaje || r.mensajeSistema || '';
            if (r.justifSalidaAnticipada) map[key].justifSalidaAnticipada = r.justifSalidaAnticipada;
            if (r.justifSalidaTardia)     map[key].justifSalidaTardia     = r.justifSalidaTardia;
        }
    });
    return map;
}

async function loadDashboardData() {
    // Show placeholders if no employee selected
    if (!rendSelectedCedula) {
        const avatarEl = document.getElementById('rend-avatar');
        const nameEl = document.getElementById('rend-name');
        const metaEl = document.getElementById('rend-meta');
        if (avatarEl) avatarEl.textContent = '?';
        if (nameEl) nameEl.textContent = 'Ningún empleado seleccionado';
        if (metaEl) metaEl.textContent = 'Haz clic para elegir...';
        
        // Reset KPIs manually for the first time
        const kpis = ['kpi-dias', 'kpi-horas', 'kpi-puntuales', 'kpi-tardanzas-r'];
        kpis.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.textContent = '—';
        });

        // Initialize empty charts
        renderWeeklyChart();
        renderYearlyChart();
        renderDetailTable();
    }

    // Load employees list
    if (rendEmployees.length === 0) {
        try { const r = await ipcRenderer.invoke('get-employees'); if (r.success) rendEmployees = r.employees || []; } catch(e) {}
    }
    // Load all records only if not already loaded
    if (rendAllRecords.length === 0) {
        try {
            const r = await ipcRenderer.invoke('get-attendance-stats');
            if (r.success) rendAllRecords = r.records || [];
        } catch(e) {}
    }
}

function openRendSelector() {
    const modal = document.getElementById('rend-selector-modal');
    if (modal) {
        modal.style.display = 'flex';
        setTimeout(() => modal.classList.add('active'), 10);
        renderRendEmpListModal();
        const searchInput = document.getElementById('rend-emp-search-modal');
        if (searchInput) {
            searchInput.value = '';
            searchInput.focus();
        }
    }
}

function closeRendSelector() {
    const modal = document.getElementById('rend-selector-modal');
    if (modal) {
        modal.classList.remove('active');
        setTimeout(() => { modal.style.display = 'none'; }, 300);
    }
}

function renderRendEmpListModal(filter = '') {
    const list = document.getElementById('rend-emp-list-modal');
    if (!list) return;
    const q = filter.toLowerCase();
    const filtered = rendEmployees.filter(e =>
        e.nombre.toLowerCase().includes(q) || String(e.cedula).includes(q) || (e.area && e.area.toLowerCase().includes(q))
    );
    list.innerHTML = filtered.map(e => {
        const initials = e.nombre.split(' ').filter(Boolean).slice(0,2).map(w=>w[0]).join('').toUpperCase();
        return `<div class="rend-selector-item${e.cedula==rendSelectedCedula?' selected':''}" data-cedula="${e.cedula}">
            <div class="rend-selector-avatar">${initials}</div>
            <div class="rend-selector-info">
                <div class="rend-selector-name">${e.nombre}</div>
                <div class="rend-selector-meta">${e.cedula} · ${e.area||'Sin área'}</div>
            </div>
        </div>`;
    }).join('') || '<div style="font-size:13px;color:var(--text-muted);padding:20px;text-align:center;">No se encontraron empleados</div>';

    list.querySelectorAll('.rend-selector-item').forEach(item => {
        item.addEventListener('click', () => {
            const emp = rendEmployees.find(e => String(e.cedula) === item.dataset.cedula);
            if (emp) selectRendEmployee(emp);
        });
    });
}

function selectRendEmployee(emp) {
    rendSelectedCedula = emp.cedula;
    closeRendSelector();

    const dash = document.getElementById('rend-dashboard');
    if (dash) {
        dash.style.display       = 'flex';
        dash.style.flexDirection = 'column';
        dash.style.gap           = '16px';
    }

    const initials = emp.nombre.split(' ').filter(Boolean).slice(0,2).map(w=>w[0]).join('').toUpperCase();
    document.getElementById('rend-avatar').textContent = initials;
    document.getElementById('rend-name').textContent   = emp.nombre;
    document.getElementById('rend-meta').textContent   = `${emp.cedula} · ${emp.area||'—'}`;

    const dayMap   = buildDayMap(emp.cedula);
    const weekNums = Object.values(dayMap).map(d=>d.weekNum).filter(Boolean);
    rendCurrentWeek = weekNums.length ? Math.max(...weekNums) : isoWeekNow();
    refreshKpis();
    updateRendWeek();
}

function animateValue(id, start, end, duration, suffix = '') {
    const obj = document.getElementById(id);
    if (!obj) return;
    let startTimestamp = null;
    const step = (timestamp) => {
        if (!startTimestamp) startTimestamp = timestamp;
        const progress = Math.min((timestamp - startTimestamp) / duration, 1);
        const current = progress * (end - start) + start;
        obj.textContent = (suffix === 'h' ? current.toFixed(1) : Math.floor(current)) + suffix;
        if (progress < 1) {
            window.requestAnimationFrame(step);
        }
    };
    window.requestAnimationFrame(step);
}

function refreshKpis() {
    if (!rendSelectedCedula) return;
    const dayMap = buildDayMap(rendSelectedCedula);
    
    // Obtenemos el rango de fechas para la semana actual
    const startDate = getDateOfISOWeek(rendCurrentWeek, 2026);
    const endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + 6);

    const workedInTarget = [];
    const entradasInTarget = [];

    Object.entries(dayMap).forEach(([fecha, d]) => {
        const date = parseFechaDate(fecha);
        if (!date) return;
        
        let inRange = true;
        if (rendKpiFilter === 'semana') {
            inRange = (date >= startDate && date <= endDate);
        }

        if (inRange) {
            if (d.in !== null && d.out !== null) workedInTarget.push(d);
            // Si tiene records (ENTRADA), los sumamos para puntuales/tardanzas
            const e = rendAllRecords.filter(r => 
                String(r.cedula).trim() === String(rendSelectedCedula).trim() && 
                r.tipo === 'ENTRADA' && 
                normFechaKey(r.fecha) === normFechaKey(date)
            );
            entradasInTarget.push(...e);
        }
    });

    const totalH = workedInTarget.reduce((s,d) => s + Math.max(0, d.out - d.in), 0);
    const puntuales = entradasInTarget.filter(r => /puntual|tolerada/i.test(r.mensaje||r.mensajeSistema||''));
    const tardanzas = entradasInTarget.filter(r => /tard|anticipada|fuera/i.test(r.mensaje||r.mensajeSistema||''));

    animateValue('kpi-dias', 0, workedInTarget.length, 1000);
    animateValue('kpi-horas', 0, totalH, 1200, 'h');
    animateValue('kpi-puntuales', 0, puntuales.length, 1100);
    animateValue('kpi-tardanzas-r', 0, tardanzas.length, 1300);
}

function updateRendWeek() {
    document.getElementById('rend-week-label').textContent        = `Semana ${rendCurrentWeek}`;
    document.getElementById('rend-detail-week-label').textContent = `Semana ${rendCurrentWeek}`;
    renderWeeklyChart();
    renderYearlyChart();
    renderDetailTable();
}

function renderWeeklyChart() {
    const ctx = document.getElementById('rendWeeklyChart');
    if (!ctx || typeof Chart === 'undefined') return;
    if (rendCharts.weekly) { rendCharts.weekly.destroy(); rendCharts.weekly = null; }
    const dayMap = buildDayMap(rendSelectedCedula);
    const hours  = [0,0,0,0,0,0,0];
    
    // Obtenemos el lunes de la semana seleccionada
    const startDate = getDateOfISOWeek(rendCurrentWeek, 2026);

    for (let i = 0; i < 7; i++) {
        const rowDate = new Date(startDate);
        rowDate.setDate(startDate.getDate() + i);
        const key = normFechaKey(rowDate);
        const d = dayMap[key];
        if (d && d.in !== null && d.out !== null) {
            hours[i] = Math.max(0, d.out - d.in);
        }
    }
    rendCharts.weekly = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'],
            datasets: [{ data: hours, backgroundColor: hours.map(h=>h>0?'#FACC15':'#f1f5f9'), borderRadius: 6, borderSkipped: false }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { callback: v => v+'h' } },
                x: { grid: { display: false } }
            },
            animation: {
                duration: 1500,
                easing: 'easeOutQuart'
            }
        }
    });
}

function renderYearlyChart() {
    const ctx = document.getElementById('rendYearlyChart');
    if (!ctx || typeof Chart === 'undefined') return;
    if (rendCharts.yearly) { rendCharts.yearly.destroy(); rendCharts.yearly = null; }
    const dayMap   = buildDayMap(rendSelectedCedula);
    const weekDays = {};
    Object.values(dayMap).forEach(d => {
        if (d.weekNum && d.in!==null && d.out!==null) weekDays[d.weekNum] = (weekDays[d.weekNum]||0)+1;
    });
    const labels = Array.from({length:52},(_,i)=>i+1);
    rendCharts.yearly = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{ data: labels.map(w=>weekDays[w]||0), backgroundColor: labels.map(w=>w===rendCurrentWeek?'#FACC15':'#e2e8f0'), borderRadius: 3, borderSkipped: false }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { callbacks: { title: c=>`Semana ${c[0].label}`, label: c=>`${c.raw} día(s)` } } },
            scales: {
                y: { beginAtZero: true, max: 7, ticks: { stepSize: 1 }, grid: { color: '#f1f5f9' } },
                x: { grid: { display: false }, ticks: { callback: (_,i)=>(i+1)%4===0?i+1:'', maxRotation: 0 } }
            },
            animation: {
                duration: 2000,
                easing: 'easeOutElastic'
            }
        }
    });
}

function renderDetailTable() {
    const tbody = document.getElementById('rend-table-body');
    if (!tbody) return;
    const dayMap = buildDayMap(rendSelectedCedula);
    
    const badge = msg => {
        if (!msg) return '<span class="rend-badge rend-badge-gray">—</span>';
        if (/puntual/i.test(msg))    return `<span class="rend-badge rend-badge-green">${msg}</span>`;
        if (/tolerada/i.test(msg))   return `<span class="rend-badge rend-badge-green">${msg}</span>`;
        if (/tard/i.test(msg))       return `<span class="rend-badge rend-badge-yellow">${msg}</span>`;
        if (/anticipada/i.test(msg)) return `<span class="rend-badge rend-badge-orange">${msg}</span>`;
        if (/fuera/i.test(msg))      return `<span class="rend-badge rend-badge-red">${msg}</span>`;
        return `<span class="rend-badge rend-badge-gray">${msg}</span>`;
    };
    const fmtH = h => { if (h===null) return '—'; const t=Math.round(h*60); return `${String(Math.floor(t/60)).padStart(2,'0')}:${String(t%60).padStart(2,'0')}`; };
    const DAYS = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
    
    // Obtenemos el lunes de la semana seleccionada
    const startDate = getDateOfISOWeek(rendCurrentWeek, 2026);

    tbody.innerHTML = DAYS.map((name, i) => {
        // Calculamos la fecha exacta para esta fila de la tabla
        const rowDate = new Date(startDate);
        rowDate.setDate(startDate.getDate() + i);
        const key = normFechaKey(rowDate);
        
        const d = dayMap[key];
        const hrs = d && d.in!==null && d.out!==null ? Math.max(0,d.out-d.in).toFixed(1)+'h' : '—';
        const jCell = txt => `<td class="rend-justif-cell" title="${txt||''}">${txt||'—'}</td>`;
        return `<tr>
            <td><strong>${name}</strong></td>
            <td>${d&&d.in!==null?fmtH(d.in):'—'}</td>
            <td>${d&&d.out!==null?fmtH(d.out):'—'}</td>
            <td>${hrs}</td>
            <td>${badge(d?.mensajeSistema)}</td>
            ${jCell(d?.justifEntrada)}
            <td>${badge(d?.mensajeSalidaSistema)}</td>
            ${jCell(d?.justifSalidaAnticipada)}
            ${jCell(d?.justifSalidaTardia)}
        </tr>`;
    }).join('');
}

// Wire up rendimiento search + week nav + modal search
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('rend-emp-search-modal')?.addEventListener('input', e => renderRendEmpListModal(e.target.value));

    // Handle clicking outside modal to close
    document.getElementById('rend-selector-modal')?.addEventListener('click', (e) => {
        if (e.target.id === 'rend-selector-modal') closeRendSelector();
    });

    // Refresh rendimiento data for selected employee
    document.getElementById('rend-refresh-btn')?.addEventListener('click', async () => {
        const btn = document.getElementById('rend-refresh-btn');
        if (!rendSelectedCedula || btn.classList.contains('spinning')) return;
        btn.classList.add('spinning');
        try {
            const r = await ipcRenderer.invoke('get-attendance-stats');
            if (r.success) rendAllRecords = r.records || [];
            refreshKpis();
            updateRendWeek();
        } finally {
            btn.classList.remove('spinning');
        }
    });


    // RayoApp brand → home
    document.getElementById('sidebar-branding')?.addEventListener('click', () => showPage('home-view'));

    const prevWeek = () => { if (rendCurrentWeek > 1)  { rendCurrentWeek--; refreshKpis(); updateRendWeek(); } };
    const nextWeek = () => { if (rendCurrentWeek < 52) { rendCurrentWeek++; refreshKpis(); updateRendWeek(); } };
    document.getElementById('rend-week-prev')?.addEventListener('click', prevWeek);
    document.getElementById('rend-week-next')?.addEventListener('click', nextWeek);
    document.getElementById('rend-detail-week-prev')?.addEventListener('click', prevWeek);
    document.getElementById('rend-detail-week-next')?.addEventListener('click', nextWeek);

    document.querySelectorAll('.rend-filter-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            rendKpiFilter = btn.dataset.filter;
            document.querySelectorAll('.rend-filter-btn').forEach(b => b.classList.toggle('active', b === btn));
            refreshKpis();
        });
    });

    // --- NÓMINA (PAYROLL) LOGIC ---
    const nominaBtn = document.getElementById('nomina-btn');
    const nominaView = document.getElementById('nomina-view');
    const refreshNominaBtn = document.getElementById('refresh-nomina-btn');
    const addNominaBtn = document.getElementById('add-nomina-btn');
    const nominaModal = document.getElementById('add-nomina-modal');
    const closeNominaModalBtn = document.getElementById('close-nomina-modal');
    const cancelNominaBtn = document.getElementById('cancel-nomina-btn');
    const addNominaForm = document.getElementById('add-nomina-form');
    
    // Table Elements
    const nominaBody = document.getElementById('nomina-body');
    const nominaLoading = document.getElementById('nomina-loading');
    const nominaEmpty = document.getElementById('nomina-empty');
    const nominaTableContainer = document.getElementById('nomina-table-container');
    const nominaSearch = document.getElementById('nomina-search');
    const nominaEmpleadoSelect = document.getElementById('nomina-empleado-select');

    let allPayrolls = [];

    // Navigation
    if (nominaBtn) {
        nominaBtn.addEventListener('click', (e) => {
            e.preventDefault();
            showPage('nomina-view');
            loadPayrollData();
        });
    }

    if (refreshNominaBtn) {
        refreshNominaBtn.addEventListener('click', loadPayrollData);
    }

    if (nominaSearch) {
        nominaSearch.addEventListener('input', (e) => renderPayrolls(e.target.value));
    }

    async function loadPayrollData() {
        if (!nominaView.classList.contains('active')) return;
        
        if (nominaLoading) nominaLoading.style.display = 'flex';
        if (nominaEmpty) nominaEmpty.style.display = 'none';
        if (nominaTableContainer) nominaTableContainer.style.display = 'none';
        nominaBody.innerHTML = '';

        try {
            const result = await ipcRenderer.invoke('get-payroll-data');
            if (result.success) {
                allPayrolls = result.records;
                renderPayrolls();
            } else {
                console.error('Error al obtener nómina:', result.error);
                alert('Error al obtener nómina: ' + result.error);
            }
        } catch (error) {
            console.error('IPC Error:', error);
        } finally {
            if (nominaLoading) nominaLoading.style.display = 'none';
        }
    }

    function renderPayrolls(searchTerm = '') {
        const filtered = allPayrolls.filter(p => {
            const s = searchTerm.toLowerCase();
            return p.nombre.toLowerCase().includes(s) || 
                   p.departamento.toLowerCase().includes(s) ||
                   p.cedula.toLowerCase().includes(s);
        });

        if (filtered.length === 0) {
            nominaEmpty.style.display = 'flex';
            nominaTableContainer.style.display = 'none';
            return;
        }

        nominaEmpty.style.display = 'none';
        nominaTableContainer.style.display = 'block';

        nominaBody.innerHTML = filtered.map(p => {
            // formattedValue from Sheets already includes currency symbols if formatted
            const fmt = (val) => val || '—';
            return `
                <tr>
                    <td style="font-weight: 500;">${p.nombre}</td>
                    <td style="color: var(--text-muted); font-size: 13px;">${p.cedula}</td>
                    <td>
                        <span style="background: var(--sidebar-active); padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: 600;">
                            ${p.departamento}
                        </span>
                    </td>
                    <td>${fmt(p.pordia)}</td>
                    <td><span class="nomina-day-badge">${p.dias} Días</span></td>
                    <td class="calc-col">${fmt(p.semana)}</td>
                    <td class="calc-col">${fmt(p.mes)}</td>
                    <td class="calc-col optional-col">${fmt(p.trimestral)}</td>
                    <td class="calc-col optional-col">${fmt(p.semestral)}</td>
                    <td class="calc-col"><span class="nomina-money-badge">${fmt(p.anual)}</span></td>
                </tr>
            `;
        }).join('');
    }

    // Modal Handling
    const openNominaModal = async () => {
        if (!nominaModal) return;
        nominaModal.style.display = 'flex';
        setTimeout(() => nominaModal.classList.add('active'), 10);

        // Reset auto-fill fields
        document.getElementById('nom-nombre').value = '';
        document.getElementById('nom-cedula').value = '';
        document.getElementById('nom-departamento').value = '';

        // Always fetch fresh employees from Sheets for the dropdown
        if (nominaEmpleadoSelect) {
            nominaEmpleadoSelect.innerHTML = '<option value="">Cargando empleados...</option>';
            nominaEmpleadoSelect.disabled = true;
            try {
                const result = await ipcRenderer.invoke('get-employees');
                nominaEmpleadoSelect.innerHTML = '<option value="">-- Seleccione un empleado --</option>';
                if (result.success && result.employees.length > 0) {
                    result.employees.forEach(emp => {
                        const opt = document.createElement('option');
                        opt.value = emp.cedula;
                        opt.dataset.nombre = emp.nombre;
                        opt.dataset.area = emp.area;
                        opt.textContent = emp.nombre;
                        nominaEmpleadoSelect.appendChild(opt);
                    });
                } else {
                    nominaEmpleadoSelect.innerHTML = '<option value="">No hay empleados registrados</option>';
                }
            } catch (e) {
                console.error('Error cargando empleados para nómina:', e);
                nominaEmpleadoSelect.innerHTML = '<option value="">Error al cargar empleados</option>';
            } finally {
                nominaEmpleadoSelect.disabled = false;
            }
        }
    };

    const closeNominaModal = () => {
        if (!nominaModal) return;
        nominaModal.classList.remove('active');
        setTimeout(() => { nominaModal.style.display = 'none'; addNominaForm.reset(); }, 300);
    };

    if (addNominaBtn) addNominaBtn.addEventListener('click', openNominaModal);
    if (closeNominaModalBtn) closeNominaModalBtn.addEventListener('click', closeNominaModal);
    if (cancelNominaBtn) cancelNominaBtn.addEventListener('click', closeNominaModal);

    if (nominaModal) {
        nominaModal.addEventListener('click', (e) => {
            if (e.target === nominaModal) closeNominaModal();
        });
    }

    // Auto-fill on select — reads from the data attributes stored on each <option>
    if (nominaEmpleadoSelect) {
        nominaEmpleadoSelect.addEventListener('change', (e) => {
            const cedula = e.target.value;
            if (!cedula) return;

            const selectedOption = e.target.options[e.target.selectedIndex];
            document.getElementById('nom-nombre').value = selectedOption.dataset.nombre || '';
            document.getElementById('nom-cedula').value = cedula;
            document.getElementById('nom-departamento').value = selectedOption.dataset.area || '';
        });
    }

    if (addNominaForm) {
        addNominaForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const submitBtn = addNominaForm.querySelector('button[type="submit"]');
            const originalBtnText = submitBtn.innerHTML;

            submitBtn.disabled = true;
            submitBtn.innerHTML = '<div class="loader" style="width:16px;height:16px;margin-right:6px;"></div> Guardando...';
            
            const formData = new FormData(addNominaForm);
            const payrollData = Object.fromEntries(formData.entries());

            try {
                const result = await ipcRenderer.invoke('save-payroll-data', payrollData);
                if (result.success) {
                    alert('✅ Registro de nómina guardado y sincronizado con Google Sheets exitosamente.');
                    closeNominaModal();
                    loadPayrollData();
                } else {
                    alert('❌ Error al guardar nómina: ' + result.error);
                }
            } catch (error) {
                console.error('IPC Error:', error);
                alert('❌ Error de conexión al intentar guardar la nómina.');
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerHTML = originalBtnText;
            }
        });
    }
});

// --- AUTO-UPDATER ---

if (checkUpdateBtn) {
    checkUpdateBtn.addEventListener('click', () => {
        checkUpdateBtn.disabled = true;
        checkUpdateBtn.classList.add('spinning');
        updateBtnLabel.textContent = 'Buscando...';
        ipcRenderer.send('check-for-update');
    });
}

ipcRenderer.on('update-not-available', () => {
    if (!checkUpdateBtn) return;
    checkUpdateBtn.disabled = false;
    checkUpdateBtn.classList.remove('spinning');
    checkUpdateBtn.classList.add('up-to-date');
    updateBtnLabel.textContent = 'App actualizada ✓';
    setTimeout(() => {
        checkUpdateBtn.classList.remove('up-to-date');
        updateBtnLabel.textContent = 'Buscar actualización';
    }, 3000);
});

ipcRenderer.on('update-available', (event, version) => {
    if (checkUpdateBtn) {
        checkUpdateBtn.classList.remove('spinning');
        updateBtnLabel.textContent = `Descargando v${version}...`;
    }
    showUpdateBanner(`Nueva versión ${version} disponible. Descargando...`, false);
});

ipcRenderer.on('update-downloaded', () => {
    if (checkUpdateBtn) {
        checkUpdateBtn.disabled = false;
        checkUpdateBtn.classList.remove('spinning');
        updateBtnLabel.textContent = 'Reiniciar para actualizar';
        checkUpdateBtn.onclick = () => ipcRenderer.send('install-update');
    }
    showUpdateBanner('Actualización lista. Reinicia para instalar.', true);
});

function showUpdateBanner(message, showInstallBtn) {
    const existing = document.getElementById('update-banner');
    if (existing) existing.remove();

    const banner = document.createElement('div');
    banner.id = 'update-banner';
    banner.style.cssText = `
        position: fixed; bottom: 24px; right: 24px; z-index: 9999;
        background: #ffffff; color: #1e293b;
        padding: 14px 16px; border-radius: 14px;
        border: 1px solid #e2e8f0; border-left: 4px solid #FACC15;
        box-shadow: 0 8px 24px rgba(0,0,0,0.10);
        display: flex; align-items: center; gap: 12px;
        font-size: 13px; max-width: 300px; font-family: inherit;
        animation: bannerSlideIn 0.3s cubic-bezier(0.4,0,0.2,1);
    `;

    const icon = document.createElement('div');
    icon.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#FACC15" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>`;
    icon.style.flexShrink = '0';
    banner.appendChild(icon);

    const text = document.createElement('span');
    text.textContent = message;
    text.style.cssText = 'flex: 1; line-height: 1.4; font-weight: 500;';
    banner.appendChild(text);

    if (showInstallBtn) {
        const btn = document.createElement('button');
        btn.textContent = 'Reiniciar';
        btn.style.cssText = `
            background: #FACC15; color: #1e293b; border: none;
            padding: 6px 12px; border-radius: 8px; cursor: pointer;
            font-size: 12px; font-weight: 600; white-space: nowrap;
            font-family: inherit;
        `;
        btn.onmouseenter = () => btn.style.background = '#eab308';
        btn.onmouseleave = () => btn.style.background = '#FACC15';
        btn.onclick = () => ipcRenderer.send('install-update');
        banner.appendChild(btn);
    }

    document.body.appendChild(banner);
}
