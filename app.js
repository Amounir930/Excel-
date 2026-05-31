// API Base URL
const API_URL = 'http://127.0.0.1:8000/api';

// State variables
let clientsData = [];
let selectedClient = null;
let currentInspectorTab = 'personal';

// Chart instances
let classificationChart = null;
let financialChart = null;

// Violations Database
const violationsList = [
    { id: 1, name: "تجاوز الإشارة الحمراء", basic: 3000, count: 0 },
    { id: 2, name: "القيادة عكس الاتجاه", basic: 6000, count: 0 },
    { id: 3, name: "تجاوز السرعة بأكثر من 25 كم/س", basic: 900, count: 0 },
    { id: 4, name: "تجاوز السرعة بأقل من 25 كم/س", basic: 500, count: 0 },
    { id: 5, name: "استخدام الجوال أثناء القيادة", basic: 500, count: 0 },
    { id: 6, name: "عدم ربط حزام الأمان", basic: 150, count: 0 },
    { id: 7, name: "القيادة بدون رخصة سارية", basic: 900, count: 0 },
    { id: 8, name: "القيادة تحت تأثير المسكرات/المخدرات", basic: 9000, count: 0 },
    { id: 9, name: "طمس لوحات المركبة", basic: 3000, count: 0 },
    { id: 10, name: "الوقوف في أماكن ذوي الإعاقة", basic: 1000, count: 0 },
    { id: 11, name: "عدم إعطاء أفضلية المرور للمشاة", basic: 500, count: 0 },
    { id: 12, name: "تغيير المسار بشكل مفاجئ", basic: 300, count: 0 },
    { id: 13, name: "عدم استخدام إشارات الانعطاف", basic: 150, count: 0 },
    { id: 14, name: "تأخير تجديد رخصة القيادة", basic: 100, count: 0 },
    { id: 15, name: "عدم تثبيت الحمولة أو تغطيتها", basic: 500, count: 0 }
];

// On Page Load
document.addEventListener("DOMContentLoaded", () => {
    refreshDashboard();
    
    // Smart Search Input Listener
    const searchInput = document.getElementById('client-search-input');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            filterClientsTable(query);
        });
    }
    
    // Auto-calculate Net Salary: Gross Salary - 9% of Basic Salary
    const basicSalInput = document.getElementById('form-basic-sal');
    const grossSalInput = document.getElementById('form-gross-sal');
    const netSalInput = document.getElementById('form-net-sal');

    function calculateNetSalary() {
        const basic = parseFloat(basicSalInput.value);
        const gross = parseFloat(grossSalInput.value);
        
        if (isNaN(basic) || isNaN(gross)) {
            netSalInput.value = '';
            return;
        }
        
        const net = gross - (basic * 0.09);
        netSalInput.value = isNaN(net) || net < 0 ? 0 : Math.round(net);
    }

    if (basicSalInput && grossSalInput && netSalInput) {
        basicSalInput.addEventListener('input', calculateNetSalary);
        grossSalInput.addEventListener('input', calculateNetSalary);
    }
    
    // Close modal when clicking outside of it
    window.addEventListener('click', (e) => {
        const detailModal = document.getElementById('detail-modal');
        if (e.target === detailModal) {
            closeDetailModal();
        }
    });
});

// Format numbers as Saudi Riyals
function formatCurrency(val) {
    return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(val);
}

// Show toast notifications
function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icon = type === 'success' 
        ? '<i class="fa-solid fa-circle-check" style="color: #10B981;"></i>' 
        : '<i class="fa-solid fa-circle-exclamation" style="color: #EF4444;"></i>';
        
    toast.innerHTML = `
        ${icon}
        <span>${message}</span>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(20px)';
        toast.style.transition = 'all 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// Switch Sidebar sections
function showSection(sectionId) {
    // Manage Sidebar active menu styling
    document.querySelectorAll('.sidebar-menu .menu-item').forEach(item => {
        item.classList.remove('active');
    });
    
    const activeItem = document.getElementById(`menu-${sectionId}`);
    if (activeItem) activeItem.classList.add('active');
    
    // Toggle Section visibility
    const dashboardGrid = document.querySelector('.kpi-grid');
    const chartsSection = document.querySelector('.charts-section');
    const mainDashboardLayout = document.querySelector('.dashboard-layout');
    const violationsSection = document.getElementById('violations-calculator-section');
    
    if (sectionId === 'dashboard') {
        dashboardGrid.style.display = 'grid';
        chartsSection.style.display = 'grid';
        mainDashboardLayout.style.display = 'grid';
        violationsSection.style.display = 'none';
    } else if (sectionId === 'violations') {
        dashboardGrid.style.display = 'none';
        chartsSection.style.display = 'none';
        mainDashboardLayout.style.display = 'none';
        violationsSection.style.display = 'block';
        
        renderViolationsTable();
        populateViolationsClients();
    }
}

// Fetch stats and clients from API
async function refreshDashboard() {
    try {
        await fetchStats();
        await fetchClients();
        showToast("تم تحديث البيانات بنجاح من شيت Excel الفعال!", "success");
    } catch (error) {
        console.error("Refresh error:", error);
        showToast("حدث خطأ أثناء الاتصال بالخادم، يرجى تشغيل الخادم الخلفي.", "error");
    }
}

// Fetch Administration Stats
async function fetchStats() {
    const res = await fetch(`${API_URL}/stats`);
    if (!res.ok) throw new Error("Failed to fetch stats");
    const stats = await res.json();
    
    // Render Stats
    document.getElementById('stat-files').textContent = stats.num_files;
    document.getElementById('stat-qualified').textContent = stats.num_qualified;
    document.getElementById('stat-exceptions').textContent = stats.num_exceptions;
    document.getElementById('stat-rejected').textContent = stats.num_rejected;
    
    document.getElementById('stat-debts').innerHTML = formatCurrency(stats.total_debts) + ' <span style="font-size: 11px; font-weight: normal; color: var(--text-secondary);">ريال</span>';
    document.getElementById('stat-executions').innerHTML = formatCurrency(stats.total_executions) + ' <span style="font-size: 11px; font-weight: normal; color: var(--text-secondary);">ريال</span>';
    document.getElementById('stat-funding').innerHTML = formatCurrency(stats.total_funding) + ' <span style="font-size: 11px; font-weight: normal; color: var(--text-secondary);">ريال</span>';
    document.getElementById('stat-avg-risk').textContent = stats.avg_risk_score;
}

// Fetch Client Lists
async function fetchClients() {
    const res = await fetch(`${API_URL}/clients`);
    if (!res.ok) throw new Error("Failed to fetch clients");
    const data = await res.json();
    clientsData = data;
    
    // Reset search query on dashboard reload
    const searchInput = document.getElementById('client-search-input');
    if (searchInput) searchInput.value = '';

    renderClientsTable(data);
    initCharts(data);
    
    // Auto-select first client if available and none selected
    if (data.length > 0) {
        if (selectedClient) {
            const updated = data.find(c => c.row_idx === selectedClient.row_idx);
            if (updated) selectClient(updated);
            else selectClient(data[0]);
        } else {
            selectClient(data[0]);
        }
    } else {
        renderEmptyInspector();
    }
}

// Render Clients Table with Avatars
function renderClientsTable(clients) {
    const tbody = document.getElementById('clients-tbody');
    tbody.innerHTML = '';
    
    if (clients.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="empty-placeholder">
                    <div class="empty-icon"><i class="fa-solid fa-folder-open"></i></div>
                    <div class="empty-title">لا يوجد عملاء مضافين حالياً</div>
                    <p style="font-size: 11px;">أضف عميلاً جديداً بالضغط على زر "إضافة ملف جديد".</p>
                </td>
            </tr>
        `;
        return;
    }
    
    clients.forEach(c => {
        const tr = document.createElement('tr');
        tr.className = selectedClient && selectedClient.row_idx === c.row_idx ? 'active' : '';
        tr.onclick = () => selectClient(c, true);
        
        // Match Decision Badges
        let decisionBadge = '';
        if (c.decision.includes("مؤهل") && !c.decision.includes("تحفظ")) {
            decisionBadge = `<span class="badge badge-qualified"><i class="fa-solid fa-circle-check"></i> ${c.decision}</span>`;
        } else if (c.decision.includes("تحفظ")) {
            decisionBadge = `<span class="badge badge-reserved"><i class="fa-solid fa-circle-dot"></i> ${c.decision}</span>`;
        } else if (c.decision.includes("استثناء")) {
            decisionBadge = `<span class="badge badge-exception"><i class="fa-solid fa-circle-exclamation"></i> ${c.decision}</span>`;
        } else {
            decisionBadge = `<span class="badge badge-rejected"><i class="fa-solid fa-circle-xmark"></i> ${c.decision}</span>`;
        }
        
        // Risk level class
        let riskClass = 'green';
        if (c.risk_level.includes("مرفوض")) riskClass = 'red';
        else if (c.risk_level.includes("عالية")) riskClass = 'orange';
        else if (c.risk_level.includes("متوسطة")) riskClass = 'yellow';
        
        // Initials avatar
        const initials = c.name.split(" ").slice(0, 2).map(w => w[0]).join("");
        
        tr.innerHTML = `
            <td style="font-family: 'Outfit', sans-serif; font-weight: 700; color: #818CF8;">${c.file_id}</td>
            <td style="font-weight: bold;">
                <div class="avatar-cell">
                    <div class="avatar-circle">${initials}</div>
                    <span>${c.name}</span>
                </div>
            </td>
            <td style="font-family: 'Outfit', sans-serif;" class="currency">${formatCurrency(c.net_sal)}</td>
            <td style="font-family: 'Outfit', sans-serif;">
                <span class="badge-risk ${c.simah >= 650 ? 'green' : c.simah >= 550 ? 'yellow' : 'red'}">${c.simah}</span>
            </td>
            <td style="font-family: 'Outfit', sans-serif; font-weight: 600;">${c.dti_pct}%</td>
            <td style="font-family: 'Outfit', sans-serif;" class="currency">${formatCurrency(c.total_exec_val)}</td>
            <td><span class="badge-risk ${riskClass}">${c.risk_level.split(" ").slice(0,-1).join(" ")}</span></td>
            <td>${decisionBadge}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Select Client & Show tabbed details
function selectClient(client, openModal = false) {
    selectedClient = client;
    
    // Update active row highlighting
    const rows = document.querySelectorAll('#clients-tbody tr');
    clientsData.forEach((c, idx) => {
        if (rows[idx]) {
            if (c.row_idx === client.row_idx) {
                rows[idx].classList.add('active');
            } else {
                rows[idx].classList.remove('active');
            }
        }
    });
    
    // Show Tab headers
    document.getElementById('inspector-tab-headers').style.display = 'flex';
    
    const inspector = document.getElementById('inspector-content');
    const sidebar = document.getElementById('inspector-decision-sidebar');
    
    // Set Decision Class
    let decisionClass = 'rejected';
    if (client.decision.includes("مؤهل") && !client.decision.includes("تحفظ")) decisionClass = 'qualified';
    else if (client.decision.includes("تحفظ")) decisionClass = 'reserved';
    else if (client.decision.includes("استثناء")) decisionClass = 'exception';
    
    // Risk level class
    let riskClass = 'green';
    if (client.risk_level.includes("مرفوض")) riskClass = 'red';
    else if (client.risk_level.includes("عالية")) riskClass = 'orange';
    else if (client.risk_level.includes("متوسطة")) riskClass = 'yellow';
    
    // Populate Decision Sidebar
    if (sidebar) {
        sidebar.innerHTML = `
            <!-- Decision Box -->
            <div class="decision-result-box ${decisionClass}" style="margin-bottom: 0; padding: 20px; width: 100%;">
                <span class="info-label" style="color: inherit; opacity: 0.8; margin-bottom: 6px; font-size: 11px; text-transform: uppercase;">القرار الائتماني النهائي</span>
                <div class="decision-title" style="font-size: 26px; font-weight: 900; margin-bottom: 8px;">${client.decision}</div>
                
                <!-- Score & Percent Row -->
                <div style="display: flex; align-items: center; gap: 16px; margin: 12px 0;">
                    <div class="score-badge-circle" style="width: 70px; height: 70px; border-width: 4px; margin: 0; border-color: ${
                        decisionClass === 'qualified' ? '#10B981' : decisionClass === 'reserved' ? '#F59E0B' : decisionClass === 'exception' ? '#F97316' : '#EF4444'
                    };">
                        <h5 style="font-size: 18px;">${client.total_pts}</h5>
                        <span style="font-size: 8px;">نقطة</span>
                    </div>
                    <div style="text-align: right;">
                        <div style="font-family: 'Outfit'; font-size: 20px; font-weight: 800; color: var(--text-primary);">${client.score_pct}%</div>
                        <div style="font-size: 11px; color: var(--text-secondary);">نسبة الجدارة الائتمانية</div>
                    </div>
                </div>

                <div style="font-size: 11.5px; font-weight: 700; background: rgba(0,0,0,0.2); padding: 4px 10px; border-radius: 20px; margin-bottom: 12px; display: inline-block;">
                    التصنيف الائتماني: [${client.classification}]
                </div>
                
                <div class="decision-reasons" style="font-size: 11px; padding: 10px 12px; border-radius: 8px; text-align: right; width: 100%; border: 1px solid rgba(255,255,255,0.05);">${client.brief_reasons}</div>
            </div>
            
            <!-- Action Buttons -->
            <div style="display: flex; flex-direction: column; gap: 10px; width: 100%;">
                <button class="btn btn-secondary" style="width: 100%; font-size: 12.5px; padding: 10px; display: flex; align-items: center; justify-content: center; gap: 8px;" onclick="openEditClientModal()">
                    <i class="fa-solid fa-user-gear"></i> تعديل ملف العميل
                </button>
                <button class="btn btn-primary" style="width: 100%; font-size: 12.5px; padding: 10px; display: flex; align-items: center; justify-content: center; gap: 8px; background: linear-gradient(135deg, #10B981 0%, #059669 100%); border: none; box-shadow: 0 4px 12px rgba(16,185,129,0.2);" onclick="printExecutionOrder()">
                    <i class="fa-solid fa-print"></i> أمر تنفيذ عملية سداد (PDF)
                </button>
                <button class="btn btn-secondary" style="width: 100%; font-size: 12.5px; padding: 10px; display: flex; align-items: center; justify-content: center; gap: 8px; border-color: rgba(239, 68, 68, 0.3); color: #EF4444;" onclick="deleteClientFile(${client.row_idx})">
                    <i class="fa-solid fa-trash-can"></i> حذف الملف
                </button>
            </div>
        `;
    }
    
    // Populate Tab Content
    inspector.innerHTML = `
        <!-- Tab Content 1: Personal & Debts -->
        <div class="tab-content" id="tab-content-personal">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-user" style="color: #6366F1;"></i> بيانات العميل الأساسية</h4>
                <div class="grid-2col">
                    <div class="info-item">
                        <span class="info-label">رقم الملف (تلقائي)</span>
                        <span class="info-value" style="font-family: 'Outfit'; font-weight: 700; color: #818CF8;">${client.file_id}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">تاريخ الإدخال</span>
                        <span class="info-value" style="font-family: 'Outfit';">${client.date}</span>
                    </div>
                    <div class="info-item" style="grid-column: span 2;">
                        <span class="info-label">اسم العميل</span>
                        <span class="info-value" style="color: var(--text-primary); font-weight: bold; font-size: 15px;">${client.name}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">رقم الهوية</span>
                        <span class="info-value" style="font-family: 'Outfit';">${client.id_num}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">رقم الجوال</span>
                        <span class="info-value" style="font-family: 'Outfit';">${client.mobile}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">العمر</span>
                        <span class="info-value">${client.age} سنة</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">جهة العمل</span>
                        <span class="info-value">${client.employer}</span>
                    </div>
                    <div class="info-item" style="grid-column: span 2;">
                        <span class="info-label">نوع جهة العمل</span>
                        <span class="info-value">${client.emp_type}</span>
                    </div>
                </div>
            </div>
            
            <div class="inspector-section">
                <h4><i class="fa-solid fa-wallet" style="color: #10B981;"></i> الدخل والملاءة المالية</h4>
                <div class="grid-2col">
                    <div class="info-item">
                        <span class="info-label">الراتب الأساسي</span>
                        <span class="info-value currency">${formatCurrency(client.basic_sal)} ريال</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي الراتب</span>
                        <span class="info-value currency">${formatCurrency(client.gross_sal)} ريال</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">صافي الراتب المحول</span>
                        <span class="info-value currency" style="color: #10B981;">${formatCurrency(client.net_sal)} ريال</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">مدة الخدمة</span>
                        <span class="info-value">${client.svc_months} شهر</span>
                    </div>
                </div>
            </div>

            <div class="inspector-section">
                <h4><i class="fa-solid fa-shield-halved" style="color: #F59E0B;"></i> الوضع الائتماني وسمة (SIMAH)</h4>
                <div class="grid-2col">
                    <div class="info-item">
                        <span class="info-label">درجة سمة</span>
                        <span class="info-value" style="font-family: 'Outfit'; font-weight: 700; color: ${client.simah >= 650 ? '#10B981' : client.simah >= 550 ? '#F59E0B' : '#EF4444'};">${client.simah}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">عدد الاستعلامات</span>
                        <span class="info-value" style="font-family: 'Outfit';">${client.inquiries} استعلامات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">تعثر قائم؟</span>
                        <span class="info-value" style="color: ${client.default_status === 'نعم' ? '#EF4444' : '#10B981'}; font-weight: bold;">${client.default_status}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إيقاف خدمات؟</span>
                        <span class="info-value" style="color: ${client.blacklist === 'نعم' ? '#EF4444' : '#10B981'}; font-weight: bold;">${client.blacklist}</span>
                    </div>
                    <div class="info-item" style="grid-column: span 2;">
                        <span class="info-label">حجز على الراتب؟</span>
                        <span class="info-value" style="color: ${client.sal_attach === 'نعم' ? '#EF4444' : '#10B981'}; font-weight: bold;">${client.sal_attach}</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab Content 2: Risk Engine -->
        <div class="tab-content" id="tab-content-risk">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-circle-exclamation" style="color: #F59E0B;"></i> شاشة الالتزامات الائتمانية والمديونيات</h4>
                
                <div class="grid-2col" style="margin-bottom: 16px; background: rgba(0,0,0,0.15); padding: 12px; border-radius: 8px;">
                    <div class="info-item">
                        <span class="info-label">إجمالي الالتزامات (المديونيات)</span>
                        <span class="info-value currency" style="color: #EF4444; font-weight: 800; font-size: 15px;">${formatCurrency(client.total_debts)} ريال</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي الأقساط الشهرية</span>
                        <span class="info-value currency" style="color: var(--primary); font-weight: 800; font-size: 15px;">${formatCurrency(client.monthly_installment)} ريال</span>
                    </div>
                    <div class="info-item" style="grid-column: span 2; border-top: 1px dashed var(--border-color); padding-top: 8px; margin-top: 8px;">
                        <span class="info-label" style="display: block; margin-bottom: 2px;">نسبة الاستقطاع الاحتسابية (DTI)</span>
                        <span class="info-value" style="font-size: 14px; font-family: 'Outfit'; font-weight: bold; color: ${client.dti_pct > 70 ? '#EF4444' : client.dti_pct > 50 ? '#F97316' : '#10B981'};">
                            ${client.dti_pct}% <span style="font-size: 11px; font-weight: normal; color: var(--text-secondary);">(= ${formatCurrency(client.monthly_installment)} ÷ ${formatCurrency(client.net_sal)} × 100)</span>
                        </span>
                    </div>
                </div>

                <div class="grid-2col" style="font-size: 11px; row-gap: 8px;">
                    <div class="info-item">
                        <span class="info-label">عدد القروض الشخصية</span>
                        <span class="info-value">${client.loans_count} قروض</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي القروض الشخصية</span>
                        <span class="info-value currency">${formatCurrency(client.loans_total)} ريال</span>
                    </div>
                    
                    <div class="info-item">
                        <span class="info-label">عدد التمويلات العقارية</span>
                        <span class="info-value">${client.real_estate_count} تمويل</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي التمويل العقاري</span>
                        <span class="info-value currency">${formatCurrency(client.real_estate_total)} ريال</span>
                    </div>
                    
                    <div class="info-item">
                        <span class="info-label">عدد البطاقات الائتمانية</span>
                        <span class="info-value">${client.cards_count} بطاقات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي مبالغ البطاقات</span>
                        <span class="info-value currency">${formatCurrency(client.cards_total)} ريال</span>
                    </div>
                    
                    <div class="info-item">
                        <span class="info-label">عدد قروض شركات التمويل</span>
                        <span class="info-value">${client.finance_cos_count} شركات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي قروض شركات التمويل</span>
                        <span class="info-value currency">${formatCurrency(client.finance_cos_total)} ريال</span>
                    </div>
                </div>
            </div>

            <div class="inspector-section">
                <h4><i class="fa-solid fa-gavel" style="color: #EF4444;"></i> ثالثاً: شاشة التنفيذات القضائية والنزاعات</h4>
                
                <div class="grid-2col" style="margin-bottom: 16px; background: rgba(0,0,0,0.15); padding: 12px; border-radius: 8px;">
                    <div class="info-item">
                        <span class="info-label">إجمالي عدد التنفيذات (الكلي)</span>
                        <span class="info-value" style="font-family: 'Outfit'; font-weight: 800; font-size: 15px; color: ${client.total_exec_count > 0 ? '#EF4444' : 'inherit'};">${client.total_exec_count} طلبات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي مبالغ التنفيذات الكلية</span>
                        <span class="info-value currency" style="font-family: 'Outfit'; font-weight: 800; font-size: 15px; color: ${client.total_exec_val > 0 ? '#EF4444' : '#10B981'};">${formatCurrency(client.total_exec_val)} ريال</span>
                    </div>
                    <div class="info-item" style="grid-column: span 2; border-top: 1px dashed var(--border-color); padding-top: 8px; margin-top: 8px;">
                        <span class="info-label" style="display: block; margin-bottom: 2px;">درجة خطورة التنفيذات (احتساب تلقائي)</span>
                        <span class="info-value" style="font-size: 14px; font-weight: bold;"><span class="badge-risk ${riskClass}">${client.risk_level}</span></span>
                    </div>
                </div>

                <div class="grid-2col" style="font-size: 11px; row-gap: 8px; margin-bottom: 12px;">
                    <div class="info-item">
                        <span class="info-label">عدد طلبات التنفيذ (أمر 46)</span>
                        <span class="info-value">${client.exec_requests_count} طلبات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي مبالغ التنفيذ (أمر 46)</span>
                        <span class="info-value currency">${formatCurrency(client.exec_requests_total)} ريال</span>
                    </div>

                    <div class="info-item">
                        <span class="info-label">عدد تنفيذات الأفراد</span>
                        <span class="info-value">${client.ind_exec_count} طلبات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي تنفيذات الأفراد</span>
                        <span class="info-value currency">${formatCurrency(client.ind_exec_total)} ريال</span>
                    </div>

                    <div class="info-item">
                        <span class="info-label">عدد تنفيذات شركات التمويل</span>
                        <span class="info-value">${client.fin_exec_count} طلبات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي تنفيذات شركات التمويل</span>
                        <span class="info-value currency">${formatCurrency(client.fin_exec_total)} ريال</span>
                    </div>

                    <div class="info-item">
                        <span class="info-label">عدد تنفيذات البنوك</span>
                        <span class="info-value">${client.bank_exec_count} طلبات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي تنفيذات البنوك</span>
                        <span class="info-value currency">${formatCurrency(client.bank_exec_total)} ريال</span>
                    </div>
                </div>
            </div>

            <div class="inspector-section">
                <h4><i class="fa-solid fa-triangle-exclamation" style="color: #F59E0B;"></i> رابعاً: سياسة المخاطر وفحص الرفض المباشر</h4>
                <p style="font-size: 11px; color: var(--text-secondary); margin-bottom: 12px; line-height: 1.5;">
                    مؤشر التدقيق التلقائي لقواعد الرفض المباشر الـ 8 المحددة بالسياسة الائتمانية:
                </p>
                <div style="display: flex; flex-direction: column; gap: 8px; font-size: 11px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.default_status === 'نعم' ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.default_status === 'نعم' ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>1. خلو من التعثر القائم بـ سمة</span>
                        <span style="font-weight: bold; color: ${client.default_status === 'نعم' ? '#EF4444' : '#10B981'};">
                            ${client.default_status === 'نعم' ? 'متعثر' : 'سليم'}
                        </span>
                    </div>
                    
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.blacklist === 'نعم' ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.blacklist === 'نعم' ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>2. خلو من إيقاف الخدمات (Blacklist)</span>
                        <span style="font-weight: bold; color: ${client.blacklist === 'نعم' ? '#EF4444' : '#10B981'};">
                            ${client.blacklist === 'نعم' ? 'موقوف' : 'سليم'}
                        </span>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.sal_attach === 'نعم' ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.sal_attach === 'نعم' ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>3. خلو من حجز الراتب</span>
                        <span style="font-weight: bold; color: ${client.sal_attach === 'نعم' ? '#EF4444' : '#10B981'};">
                            ${client.sal_attach === 'نعم' ? 'محجوز' : 'سليم'}
                        </span>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.exec_requests_count > 8 ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.exec_requests_count > 8 ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>4. عدد طلبات التنفيذ (≤ 8 طلبات)</span>
                        <span style="font-weight: bold; color: ${client.exec_requests_count > 8 ? '#EF4444' : '#10B981'};">
                            ${client.exec_requests_count > 8 ? `متجاوز (${client.exec_requests_count})` : `سليم (${client.exec_requests_count})`}
                        </span>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.total_exec_val > 100000 ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.total_exec_val > 100000 ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>5. إجمالي مبالغ التنفيذ (≤ 100,000 ريال)</span>
                        <span style="font-weight: bold; color: ${client.total_exec_val > 100000 ? '#EF4444' : '#10B981'};">
                            ${client.total_exec_val > 100000 ? `متجاوز (${formatCurrency(client.total_exec_val)})` : `سليم (${formatCurrency(client.total_exec_val)})`}
                        </span>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.fin_exec_count > 3 ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.fin_exec_count > 3 ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>6. تنفيذات شركات التمويل (≤ 3 تنفيذات)</span>
                        <span style="font-weight: bold; color: ${client.fin_exec_count > 3 ? '#EF4444' : '#10B981'};">
                            ${client.fin_exec_count > 3 ? `متجاوز (${client.fin_exec_count})` : `سليم (${client.fin_exec_count})`}
                        </span>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.simah < 550 ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.simah < 550 ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>7. درجة سمة الائتمانية (≥ 550)</span>
                        <span style="font-weight: bold; color: ${client.simah < 550 ? '#EF4444' : '#10B981'};">
                            ${client.simah < 550 ? `منخفضة (${client.simah})` : `سليم (${client.simah})`}
                        </span>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: ${client.dti_pct > 70 ? 'rgba(239,68,68,0.1)' : 'rgba(16,185,129,0.05)'}; border: 1px solid ${client.dti_pct > 70 ? 'rgba(239,68,68,0.2)' : 'rgba(16,185,129,0.1)'}; border-radius: 6px;">
                        <span>8. نسبة الاستقطاع المسموحة (≤ 70%)</span>
                        <span style="font-weight: bold; color: ${client.dti_pct > 70 ? '#EF4444' : '#10B981'};">
                            ${client.dti_pct > 70 ? `متجاوز (${client.dti_pct}%)` : `سليم (${client.dti_pct}%)`}
                        </span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab Content 3: Feasibility & Surpluses -->
        <div class="tab-content" id="tab-content-feasibility">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-chart-pie" style="color: #EC4899;"></i> سادساً: دراسة الجدوى التمويلية (معاملات المدخلات)</h4>
                <div class="grid-2col" style="margin-bottom: 12px;">
                    <div class="info-item">
                        <span class="info-label">مبلغ التمويل المتوقع</span>
                        <span class="info-value currency" style="color: #10B981; font-weight: bold;">${formatCurrency(client.expected_funding)} ريال</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">نسبة أتعاب الشركة</span>
                        <span class="info-value" style="font-family: 'Outfit'; font-weight: bold;">${client.fees_percent * 100}%</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي المديونيات (الالتزامات)</span>
                        <span class="info-value currency" style="color: #EF4444;">${formatCurrency(client.total_debts)} ريال</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي التنفيذات القضائية</span>
                        <span class="info-value currency" style="color: #EF4444;">${formatCurrency(client.total_exec_val)} ريال</span>
                    </div>
                </div>
            </div>

            <div class="inspector-section">
                <h4><i class="fa-solid fa-calculator" style="color: #6366F1;"></i> نتائج الاحتساب التلقائي والفوائض</h4>
                
                <div class="grid-2col" style="row-gap: 12px; margin-bottom: 12px;">
                    <div class="info-item">
                        <span class="info-label">مبلغ السداد المطلوب (الكلي)</span>
                        <span class="info-value currency" style="font-weight: bold; color: #EF4444;">${formatCurrency(client.required_payment)} ريال</span>
                        <span style="font-size: 9px; color: var(--text-secondary); display: block;">(= المديونيات + التنفيذات)</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">قيمة أتعاب الشركة الفعلية</span>
                        <span class="info-value currency" style="font-weight: bold; color: #F59E0B;">${formatCurrency(client.company_fees)} ريال</span>
                        <span style="font-size: 9px; color: var(--text-secondary); display: block;">(= التمويل × نسبة الأتعاب)</span>
                    </div>
                    
                    <div class="info-item" style="grid-column: span 2; background: rgba(0,0,0,0.1); padding: 8px 12px; border-radius: 6px; border: 1px solid var(--border-color);">
                        <span class="info-label">الفائض الإجمالي المتوقع</span>
                        <span class="info-value currency" style="font-weight: bold; color: ${client.gross_surplus > 0 ? '#10B981' : '#EF4444'};">${formatCurrency(client.gross_surplus)} ريال</span>
                        <span style="font-size: 9px; color: var(--text-secondary); display: block;">(= التمويل المتوقع - مبلغ السداد المطلوب)</span>
                    </div>
                </div>

                <div class="grid-2col" style="background: rgba(16,185,129,0.05); padding: 12px; border-radius: 8px; border: 1px solid rgba(16,185,129,0.15); display: flex; flex-direction: row; justify-content: space-between; align-items: center; margin-top: 12px;">
                    <div>
                        <span class="info-label" style="display: block; font-weight: bold; color: var(--text-primary);">صافي الفائض للعميل (المتبقي)</span>
                        <span class="info-value currency" style="color: ${client.net_surplus > 0 ? '#10B981' : '#EF4444'}; font-size: 16px; font-weight: 800;">${formatCurrency(client.net_surplus)} ريال</span>
                        <span style="font-size: 9px; color: var(--text-secondary); display: block;">(= الفائض الإجمالي - أتعاب الشركة)</span>
                    </div>
                    <div style="text-align: left;">
                        <span class="info-label" style="display: block; font-weight: bold;">قرار الجدوى</span>
                        <span class="info-value" style="color: ${client.net_surplus > 0 ? '#10B981' : '#EF4444'}; font-size: 13px; font-weight: bold;">${client.feasibility_decision}</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab Content 4: Workflows -->
        <div class="tab-content" id="tab-content-workflow">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-gears" style="color: #6366F1;"></i> نقاط التقييم التفصيلية الائتمانية</h4>
                <div class="grid-3col" style="font-size: 11px; margin-bottom: 12px;">
                    <div class="info-item">
                        <span class="info-label">درجة سمة (25)</span>
                        <span class="info-value">${client.simah_pts} / 25</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">نسبة الاستقطاع (25)</span>
                        <span class="info-value">${client.dti_pts} / 25</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">جهة العمل (15)</span>
                        <span class="info-value">${client.emp_pts} / 15</span>
                    </div>
                </div>
                <div class="grid-2col" style="font-size: 11px;">
                    <div class="info-item">
                        <span class="info-label">التعثرات وإيقاف الراتب (20)</span>
                        <span class="info-value">${client.default_pts} / 20</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">التنفيذات ومبالغها (20)</span>
                        <span class="info-value">${client.exec_pts} / 20</span>
                    </div>
                </div>
            </div>
            
            <div class="inspector-section">
                <h4><i class="fa-solid fa-signature" style="color: #818CF8;"></i> سابعاً: Workflow ومراحل اعتماد الملف</h4>
                <p style="font-size: 11px; color: var(--text-secondary); margin-bottom: 12px; line-height: 1.5;">
                    مرحلة اعتماد الملف الائتماني والاعتمادات المالية الحالية (تفاعلية):
                </p>
                <div class="form-group" style="margin-top: 10px;">
                    <div style="position: relative; display: flex; align-items: center; width: 100%;">
                        <select id="inspector-workflow-select" class="form-input" style="background: rgba(99, 102, 241, 0.08); border: 1px solid rgba(99, 102, 241, 0.3); color: #818CF8; font-weight: bold; padding: 12px 36px 12px 16px; border-radius: 10px; width: 100%; outline: none; cursor: pointer; transition: all 0.3s;" onchange="changeClientWorkflowStage(this.value)">
                            <option value="جديد" ${client.workflow_stage === 'جديد' ? 'selected' : ''}>1- جديد</option>
                            <option value="تحت التحليل" ${client.workflow_stage === 'تحت التحليل' ? 'selected' : ''}>2- تحت التحليل</option>
                            <option value="معتمد من محلل الائتمان" ${client.workflow_stage === 'معتمد من محلل الائتمان' ? 'selected' : ''}>3- معتمد من محلل الائتمان</option>
                            <option value="معتمد من مدير العمليات" ${client.workflow_stage === 'معتمد من مدير العمليات' ? 'selected' : ''}>4- معتمد من مدير العمليات</option>
                            <option value="تحت المراجعة المالية" ${client.workflow_stage === 'تحت المراجعة المالية' ? 'selected' : ''}>5- تحت المراجعة المالية</option>
                            <option value="معتمد مالياً" ${client.workflow_stage === 'معتمد مالياً' ? 'selected' : ''}>6- معتمد مالياً</option>
                            <option value="مكتمل" ${client.workflow_stage === 'مكتمل' ? 'selected' : ''}>7- مكتمل</option>
                            <option value="مرفوض" ${client.workflow_stage === 'مرفوض' ? 'selected' : ''}>8- مرفوض</option>
                        </select>
                        <i class="fa-solid fa-circle-play" style="position: absolute; right: 14px; color: #818CF8; pointer-events: none;"></i>
                    </div>
                </div>
            </div>

            <div class="inspector-section" style="border-top: 1px dashed var(--border-color); padding-top: 16px;">
                <h4><i class="fa-solid fa-file-signature" style="color: #10B981;"></i> ثامناً: الاعتمادات</h4>
                
                <!-- Analyst approval card -->
                <div style="background: rgba(255,255,255,0.02); border: 1px solid var(--border-color); border-radius: 8px; padding: 10px; margin-bottom: 10px; font-size: 11px;">
                    <div style="display: flex; justify-content: space-between; font-weight: bold; color: var(--primary); margin-bottom: 4px;">
                        <span>اعتماد محلل الائتمان</span>
                        <span style="font-family: 'Outfit'; font-weight: normal; color: var(--text-secondary);">${client.analyst_date || 'غير موقع'}</span>
                    </div>
                    <div><strong>الاسم:</strong> ${client.analyst_name || '---'}</div>
                    <div style="margin-top: 4px; padding: 4px 6px; background: rgba(99,102,241,0.05); border-radius: 4px; color: var(--text-secondary);">
                        <strong>التوصية:</strong> ${client.analyst_recommendation || 'في انتظار كتابة التوصية الائتمانية'}
                    </div>
                </div>

                <!-- Operations approval card -->
                <div style="background: rgba(255,255,255,0.02); border: 1px solid var(--border-color); border-radius: 8px; padding: 10px; margin-bottom: 10px; font-size: 11px;">
                    <div style="display: flex; justify-content: space-between; font-weight: bold; color: #F59E0B; margin-bottom: 4px;">
                        <span>اعتماد مدير العمليات</span>
                        <span style="font-family: 'Outfit'; font-weight: normal; color: var(--text-secondary);">${client.ops_date || 'غير موقع'}</span>
                    </div>
                    <div><strong>الاسم:</strong> ${client.ops_name || '---'}</div>
                    <div style="margin-top: 4px; padding: 4px 6px; background: rgba(245,158,11,0.05); border-radius: 4px; color: var(--text-secondary);">
                        <strong>القرار:</strong> ${client.ops_decision || 'في انتظار قرار مدير العمليات'}
                    </div>
                </div>

                <!-- Finance approval card -->
                <div style="background: rgba(255,255,255,0.02); border: 1px solid var(--border-color); border-radius: 8px; padding: 10px; font-size: 11px;">
                    <div style="display: flex; justify-content: space-between; font-weight: bold; color: #10B981; margin-bottom: 4px;">
                        <span>اعتماد الإدارة المالية</span>
                        <span style="font-family: 'Outfit'; font-weight: normal; color: var(--text-secondary);">${client.finance_date || 'غير موقع'}</span>
                    </div>
                    <div><strong>الاسم:</strong> ${client.finance_name || '---'}</div>
                    <div style="margin-top: 4px; padding: 4px 6px; background: rgba(16,185,129,0.05); border-radius: 4px; color: var(--text-secondary);">
                        <strong>القرار:</strong> ${client.finance_decision || 'في انتظار قرار الإدارة المالية'}
                    </div>
                </div>

                <!-- General notes & completion date -->
                ${client.general_notes || client.completion_date ? `
                    <div style="margin-top: 12px; padding: 10px; background: rgba(0,0,0,0.15); border-radius: 8px; font-size: 11px;">
                        ${client.general_notes ? `<div style="margin-bottom: 4px;"><strong>ملاحظات عامة:</strong> ${client.general_notes}</div>` : ''}
                        ${client.completion_date ? `<div><strong>تاريخ الاكتمال النهائي:</strong> <span style="font-family: 'Outfit';">${client.completion_date}</span></div>` : ''}
                    </div>
                ` : ''}
            </div>
        </div>
    `;
    
    switchInspectorTab(currentInspectorTab);
    
    if (openModal) {
        document.getElementById('detail-modal').classList.add('active');
    }
}

// Switch between inspector detail tabs
function switchInspectorTab(tabId) {
    currentInspectorTab = tabId;
    
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    const activeBtn = document.getElementById(`tab-btn-${tabId}`);
    if (activeBtn) activeBtn.classList.add('active');
    
    const activeContent = document.getElementById(`tab-content-${tabId}`);
    if (activeContent) activeContent.classList.add('active');
}

// Change client workflow stage directly from inspector dropdown
async function changeClientWorkflowStage(newStage) {
    if (!selectedClient) return;
    
    // Copy the selected client and update the workflow stage
    const body = {
        ...selectedClient,
        workflow_stage: newStage
    };
    
    showToast("جاري تحديث مرحلة الملف...", "success");
    
    try {
        const res = await fetch(`${API_URL}/clients`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        
        const data = await res.json();
        if (data.success) {
            showToast(`تم تغيير مرحلة الملف بنجاح إلى: ${newStage}`, "success");
            selectedClient = data.client;
            
            // Refresh dashboard list
            await refreshDashboard();
            
            // Update details view without closing the modal
            selectClient(selectedClient, false);
        } else {
            showToast(data.detail || "فشل تحديث مرحلة الملف.", "error");
        }
    } catch (err) {
        console.error(err);
        showToast("خطأ في الاتصال بالخادم.", "error");
    }
}

// Close Detailed Client Modal
function closeDetailModal() {
    const modal = document.getElementById('detail-modal');
    if (modal) {
        modal.classList.remove('active');
    }
}

// Filter Clients table dynamically (Smart Search)
function filterClientsTable(query) {
    if (!query) {
        renderClientsTable(clientsData);
        return;
    }
    const filtered = clientsData.filter(c => {
        return (c.name && c.name.toLowerCase().includes(query)) ||
               (c.mobile && c.mobile.includes(query)) ||
               (c.file_id && c.file_id.toLowerCase().includes(query)) ||
               (c.id_num && c.id_num.includes(query));
    });
    renderClientsTable(filtered);
}

// Render empty inspector placeholder
function renderEmptyInspector() {
    document.getElementById('inspector-tab-headers').style.display = 'none';
    const sidebar = document.getElementById('inspector-decision-sidebar');
    if (sidebar) sidebar.innerHTML = '';
    const inspector = document.getElementById('inspector-content');
    inspector.innerHTML = `
        <div class="empty-placeholder">
            <div class="empty-icon"><i class="fa-solid fa-user-check"></i></div>
            <div class="empty-title">لم يتم اختيار أي ملف عميل</div>
            <p style="font-size: 11px; color: var(--text-secondary);">اختر ملفاً من الجدول لعرض تفاصيله، محرك المخاطر، الجدوى المالية وقرار الأهلية التفاعلي.</p>
        </div>
    `;
}

// Render Traffic Violations Calculator list
function renderViolationsTable() {
    const tbody = document.getElementById('violations-tbody');
    tbody.innerHTML = '';
    
    const discountType = document.getElementById('violations-discount-select').value;
    
    violationsList.forEach(v => {
        const tr = document.createElement('tr');
        
        // Compute discount & due for this violation row
        const rowGross = v.basic * v.count;
        let rowDiscount = "بدون تخفيض";
        let rowDue = rowGross;
        
        if (v.count > 0) {
            if (discountType === '50') {
                rowDiscount = "خصم 50% مبادرة";
                rowDue = rowGross * 0.5;
            } else if (discountType === '25') {
                rowDiscount = "خصم 25% سداد مبكر";
                rowDue = rowGross * 0.75;
            }
        }
        
        tr.innerHTML = `
            <td style="font-weight: bold; padding: 12px 16px;">${v.name}</td>
            <td style="text-align: center; font-family: 'Outfit';">${formatCurrency(v.basic)} ريال</td>
            <td style="text-align: center; width: 140px;">
                <div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
                    <button class="btn btn-secondary" style="padding: 4px 10px; font-weight: bold; border-radius: 6px;" type="button" onclick="adjustViolCount(${v.id}, -1)">-</button>
                    <span style="font-family: 'Outfit'; font-weight: 700; font-size: 15px; width: 25px; text-align: center;">${v.count}</span>
                    <button class="btn btn-secondary" style="padding: 4px 10px; font-weight: bold; border-radius: 6px;" type="button" onclick="adjustViolCount(${v.id}, 1)">+</button>
                </div>
            </td>
            <td style="text-align: center;">
                <span class="badge ${v.count > 0 && discountType !== 'none' ? 'badge-qualified' : 'badge-risk'}" style="font-size: 11px;">
                    ${rowDiscount}
                </span>
            </td>
            <td style="text-align: center; font-family: 'Outfit'; font-weight: 700; color: #818CF8;">${formatCurrency(rowDue)} ريال</td>
        `;
        tbody.appendChild(tr);
    });
    
    updateViolationsCalc();
}

// Adjust violation count
function adjustViolCount(id, amount) {
    const v = violationsList.find(x => x.id === id);
    if (v) {
        v.count = Math.max(0, v.count + amount);
        renderViolationsTable();
    }
}

// Update violations calculation totals
function updateViolationsCalc() {
    const discountType = document.getElementById('violations-discount-select').value;
    
    let totalGross = 0;
    let totalDue = 0;
    
    violationsList.forEach(v => {
        const rowGross = v.basic * v.count;
        totalGross += rowGross;
        
        if (discountType === '50') {
            totalDue += rowGross * 0.5;
        } else if (discountType === '25') {
            totalDue += rowGross * 0.75;
        } else {
            totalDue += rowGross;
        }
    });
    
    const totalSaved = totalGross - totalDue;
    
    // Render sums
    document.getElementById('viol-total-due').innerHTML = `${formatCurrency(totalDue)} <span style="font-size: 14px; font-weight: normal; color: var(--text-secondary);">ريال</span>`;
    document.getElementById('viol-total-gross').textContent = `${formatCurrency(totalGross)} ريال`;
    document.getElementById('viol-total-saved').textContent = `${formatCurrency(totalSaved)} ريال`;
}

// Populate target client select dropdown in violations
function populateViolationsClients() {
    const select = document.getElementById('viol-client-select');
    select.innerHTML = '<option value="">-- اختر ملف العميل --</option>';
    
    clientsData.forEach(c => {
        const option = document.createElement('option');
        option.value = c.row_idx;
        option.textContent = `${c.file_id} - ${c.name}`;
        
        if (selectedClient && selectedClient.row_idx === c.row_idx) {
            option.selected = true;
        }
        
        select.appendChild(option);
    });
}

// Migrate violation amounts as debt to the client
async function migrateViolationsToClient() {
    const select = document.getElementById('viol-client-select');
    const targetRowIdx = parseInt(select.value);
    
    if (isNaN(targetRowIdx)) {
        showToast("يرجى اختيار ملف العميل لترحيل المبالغ إليه!", "error");
        return;
    }
    
    // Calculate total violations due
    const discountType = document.getElementById('violations-discount-select').value;
    let totalDue = 0;
    
    violationsList.forEach(v => {
        const rowGross = v.basic * v.count;
        if (discountType === '50') totalDue += rowGross * 0.5;
        else if (discountType === '25') totalDue += rowGross * 0.75;
        else totalDue += rowGross;
    });
    
    if (totalDue === 0) {
        showToast("إجمالي مبالغ المخالفات 0 ريال! يرجى إدخال مخالفات أولاً.", "error");
        return;
    }
    
    const client = clientsData.find(c => c.row_idx === targetRowIdx);
    if (!client) {
        showToast("لم يتم العثور على ملف العميل المحدد.", "error");
        return;
    }
    
    if (!confirm(`هل أنت متأكد من رغبتك في ترحيل مبلغ ${formatCurrency(totalDue)} ريال كالتزام إضافي قضائي (تنفيذات) في ملف العميل: ${client.name}؟`)) {
        return;
    }
    
    // We add the traffic violations directly to the existing exec_requests_total (Judicial Executions)
    // and also increase the exec_requests_count (violations represents new executions/claims)
    const newExecTotal = client.exec_requests_total + totalDue;
    const newExecCount = client.exec_requests_count + violationsList.filter(v => v.count > 0).reduce((sum, v) => sum + v.count, 0);
    
    const body = {
        name: client.name,
        id_num: client.id_num,
        mobile: client.mobile,
        age: client.age,
        employer: client.employer,
        emp_type: client.emp_type,
        basic_sal: client.basic_sal,
        gross_sal: client.gross_sal,
        net_sal: client.net_sal,
        svc_months: client.svc_months,
        simah: client.simah,
        inquiries: client.inquiries,
        default_status: client.default_status,
        blacklist: client.blacklist,
        sal_attach: client.sal_attach,
        
        loans_count: client.loans_count,
        loans_total: client.loans_total,
        real_estate_count: client.real_estate_count,
        real_estate_total: client.real_estate_total,
        cards_count: client.cards_count,
        cards_total: client.cards_total,
        finance_cos_count: client.finance_cos_count,
        finance_cos_total: client.finance_cos_total,
        monthly_installment: client.monthly_installment,
        
        exec_requests_count: newExecCount,
        exec_requests_total: newExecTotal,
        ind_exec_count: client.ind_exec_count,
        ind_exec_total: client.ind_exec_total,
        fin_exec_count: client.fin_exec_count,
        fin_exec_total: client.fin_exec_total,
        bank_exec_count: client.bank_exec_count,
        bank_exec_total: client.bank_exec_total,
        fees_percent: client.fees_percent,
        workflow_stage: client.workflow_stage || "جديد",
        analyst_name: client.analyst_name || "",
        analyst_date: client.analyst_date || "",
        analyst_recommendation: client.analyst_recommendation || "",
        ops_name: client.ops_name || "",
        ops_date: client.ops_date || "",
        ops_decision: client.ops_decision || "",
        finance_name: client.finance_name || "",
        finance_date: client.finance_date || "",
        finance_decision: client.finance_decision || "",
        general_notes: client.general_notes || "",
        completion_date: client.completion_date || ""
    };
    
    try {
        const res = await fetch(`${API_URL}/clients`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        
        const data = await res.json();
        if (data.success) {
            showToast(`تم ترحيل مديونية المخالفات بنجاح وإعادة احتساب الملاءة للعميل في Excel!`, "success");
            
            // Reset violations count
            violationsList.forEach(v => v.count = 0);
            renderViolationsTable();
            
            // Go back to main dashboard view & auto-select that client to see updated risk!
            showSection('dashboard');
            selectedClient = data.client;
            
            await refreshDashboard();
        } else {
            showToast(data.detail || "حدث خطأ أثناء ترحيل المخالفات.", "error");
        }
    } catch (err) {
        console.error(err);
        showToast("فشل الاتصال بالخادم لترحيل البيانات.", "error");
    }
}

// Initialize / Render Chart.js Premium Charts in Dark Mode
function initCharts(clients) {
    if (clients.length === 0) return;
    
    let qualifiedCount = 0;
    let reservedCount = 0;
    let exceptionCount = 0;
    let rejectedCount = 0;
    
    clients.forEach(c => {
        if (c.decision.includes("مؤهل") && !c.decision.includes("تحفظ")) qualifiedCount++;
        else if (c.decision.includes("تحفظ")) reservedCount++;
        else if (c.decision.includes("استثناء")) exceptionCount++;
        else if (c.decision.includes("مرفوض")) rejectedCount++;
    });
    
    const pieCtx = document.getElementById('classificationChart').getContext('2d');
    if (classificationChart) classificationChart.destroy();
    
    classificationChart = new Chart(pieCtx, {
        type: 'doughnut',
        data: {
            labels: ['مؤهل', 'مؤهل بتحفظ', 'يحتاج استثناء', 'مرفوض'],
            datasets: [{
                data: [qualifiedCount, reservedCount, exceptionCount, rejectedCount],
                backgroundColor: [
                    '#10B981',
                    '#F59E0B',
                    '#F97316',
                    '#EF4444'
                ],
                borderColor: '#111827',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    rtl: true,
                    labels: {
                        color: '#9CA3AF',
                        font: { family: 'Tajawal', weight: 'bold' }
                    }
                }
            },
            cutout: '65%'
        }
    });
    
    const clientNames = clients.slice(0, 5).map(c => c.name.split(" ")[0]);
    const expectedFundings = clients.slice(0, 5).map(c => c.expected_funding);
    const requiredPayments = clients.slice(0, 5).map(c => c.required_payment);
    const netSurpluses = clients.slice(0, 5).map(c => c.net_surplus);
    
    const barCtx = document.getElementById('financialChart').getContext('2d');
    if (financialChart) financialChart.destroy();
    
    financialChart = new Chart(barCtx, {
        type: 'bar',
        data: {
            labels: clientNames,
            datasets: [
                {
                    label: 'التمويل المتوقع',
                    data: expectedFundings,
                    backgroundColor: 'rgba(16, 185, 129, 0.75)',
                    borderColor: '#10B981',
                    borderWidth: 1
                },
                {
                    label: 'السداد المطلوب',
                    data: requiredPayments,
                    backgroundColor: 'rgba(239, 68, 68, 0.75)',
                    borderColor: '#EF4444',
                    borderWidth: 1
                },
                {
                    label: 'صافي الفائض للعميل',
                    data: netSurpluses,
                    backgroundColor: 'rgba(99, 102, 241, 0.75)',
                    borderColor: '#6366F1',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    rtl: true,
                    labels: {
                        color: '#9CA3AF',
                        font: { family: 'Tajawal', weight: 'bold' }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255,255,255,0.05)' },
                    ticks: { color: '#9CA3AF', font: { family: 'Tajawal' } }
                },
                y: {
                    grid: { color: 'rgba(255,255,255,0.05)' },
                    ticks: { color: '#9CA3AF', font: { family: 'Outfit' } }
                }
            }
        }
    });
}

// Open modal for new client
function openNewClientModal() {
    document.getElementById('modal-title').textContent = "إضافة ملف ائتماني جديد";
    document.getElementById('client-form').reset();
    
    document.getElementById('form-name').value = '';
    document.getElementById('form-id').value = '';
    document.getElementById('form-mobile').value = '';
    document.getElementById('form-age').value = '';
    document.getElementById('form-employer').value = '';
    document.getElementById('form-emp-type').value = 'حكومي';
    document.getElementById('form-basic-sal').value = '';
    document.getElementById('form-gross-sal').value = '';
    document.getElementById('form-net-sal').value = '';
    document.getElementById('form-svc-months').value = '';
    document.getElementById('form-simah').value = '';
    document.getElementById('form-inquiries').value = '';
    document.getElementById('form-default').value = 'لا';
    document.getElementById('form-blacklist').value = 'لا';
    document.getElementById('form-sal-attach').value = 'لا';
    document.getElementById('form-workflow').value = 'جديد';

    
    document.getElementById('form-loans-cnt').value = 0;
    document.getElementById('form-loans-tot').value = 0;
    document.getElementById('form-estate-cnt').value = 0;
    document.getElementById('form-estate-tot').value = 0;
    document.getElementById('form-cards-cnt').value = 0;
    document.getElementById('form-cards-tot').value = 0;
    document.getElementById('form-fin-cnt').value = 0;
    document.getElementById('form-fin-tot').value = 0;
    document.getElementById('form-installment').value = '';
    
    document.getElementById('form-exec-cnt').value = 0;
    document.getElementById('form-exec-tot').value = 0;
    document.getElementById('form-exec-ind-cnt').value = 0;
    document.getElementById('form-exec-ind-tot').value = 0;
    document.getElementById('form-exec-fin-cnt').value = 0;
    document.getElementById('form-exec-fin-tot').value = 0;
    document.getElementById('form-exec-bank-cnt').value = 0;
    document.getElementById('form-exec-bank-tot').value = 0;
    document.getElementById('form-fees-pct').value = '0.10';

    // Approvals
    document.getElementById('form-analyst-name').value = '';
    document.getElementById('form-analyst-date').value = '';
    document.getElementById('form-analyst-rec').value = '';
    
    document.getElementById('form-ops-name').value = '';
    document.getElementById('form-ops-date').value = '';
    document.getElementById('form-ops-dec').value = '';
    
    document.getElementById('form-fin-name').value = '';
    document.getElementById('form-fin-date').value = '';
    document.getElementById('form-fin-dec').value = '';
    
    document.getElementById('form-general-notes').value = '';
    document.getElementById('form-completion-date').value = '';

    document.getElementById('client-modal').classList.add('active');
}

// Open modal for editing selected client
function openEditClientModal() {
    if (!selectedClient) return;
    
    closeDetailModal();
    
    document.getElementById('modal-title').textContent = `تعديل ملف العميل: ${selectedClient.name}`;
    
    document.getElementById('form-name').value = selectedClient.name;
    document.getElementById('form-id').value = selectedClient.id_num;
    document.getElementById('form-mobile').value = selectedClient.mobile;
    document.getElementById('form-age').value = selectedClient.age;
    document.getElementById('form-employer').value = selectedClient.employer;
    document.getElementById('form-emp-type').value = selectedClient.emp_type;
    document.getElementById('form-basic-sal').value = selectedClient.basic_sal;
    document.getElementById('form-gross-sal').value = selectedClient.gross_sal;
    document.getElementById('form-net-sal').value = selectedClient.net_sal;
    document.getElementById('form-svc-months').value = selectedClient.svc_months;
    document.getElementById('form-simah').value = selectedClient.simah;
    document.getElementById('form-inquiries').value = selectedClient.inquiries;
    document.getElementById('form-default').value = selectedClient.default_status;
    document.getElementById('form-blacklist').value = selectedClient.blacklist;
    document.getElementById('form-sal-attach').value = selectedClient.sal_attach;
    
    document.getElementById('form-loans-cnt').value = selectedClient.loans_count;
    document.getElementById('form-loans-tot').value = selectedClient.loans_total;
    document.getElementById('form-estate-cnt').value = selectedClient.real_estate_count;
    document.getElementById('form-estate-tot').value = selectedClient.real_estate_total;
    document.getElementById('form-cards-cnt').value = selectedClient.cards_count;
    document.getElementById('form-cards-tot').value = selectedClient.cards_total;
    document.getElementById('form-fin-cnt').value = selectedClient.finance_cos_count;
    document.getElementById('form-fin-tot').value = selectedClient.finance_cos_total;
    document.getElementById('form-installment').value = selectedClient.monthly_installment;
    
    document.getElementById('form-exec-cnt').value = selectedClient.exec_requests_count;
    document.getElementById('form-exec-tot').value = selectedClient.exec_requests_total;
    document.getElementById('form-exec-ind-cnt').value = selectedClient.ind_exec_count;
    document.getElementById('form-exec-ind-tot').value = selectedClient.ind_exec_total;
    document.getElementById('form-exec-fin-cnt').value = selectedClient.fin_exec_count;
    document.getElementById('form-exec-fin-tot').value = selectedClient.fin_exec_total;
    document.getElementById('form-exec-bank-cnt').value = selectedClient.bank_exec_count;
    document.getElementById('form-exec-bank-tot').value = selectedClient.bank_exec_total;
    document.getElementById('form-fees-pct').value = selectedClient.fees_percent.toFixed(2);
    document.getElementById('form-workflow').value = selectedClient.workflow_stage || 'جديد';

    // Approvals
    document.getElementById('form-analyst-name').value = selectedClient.analyst_name || '';
    document.getElementById('form-analyst-date').value = selectedClient.analyst_date || '';
    document.getElementById('form-analyst-rec').value = selectedClient.analyst_recommendation || '';
    
    document.getElementById('form-ops-name').value = selectedClient.ops_name || '';
    document.getElementById('form-ops-date').value = selectedClient.ops_date || '';
    document.getElementById('form-ops-dec').value = selectedClient.ops_decision || '';
    
    document.getElementById('form-fin-name').value = selectedClient.finance_name || '';
    document.getElementById('form-fin-date').value = selectedClient.finance_date || '';
    document.getElementById('form-fin-dec').value = selectedClient.finance_decision || '';
    
    document.getElementById('form-general-notes').value = selectedClient.general_notes || '';
    document.getElementById('form-completion-date').value = selectedClient.completion_date || '';

    document.getElementById('client-modal').classList.add('active');
}

// Close client modal
function closeClientModal() {
    document.getElementById('client-modal').classList.remove('active');
}

// Handle client form submit (post to API)
async function handleFormSubmit(e) {
    e.preventDefault();
    
    const body = {
        name: document.getElementById('form-name').value,
        id_num: document.getElementById('form-id').value,
        mobile: document.getElementById('form-mobile').value,
        age: parseInt(document.getElementById('form-age').value) || 0,
        employer: document.getElementById('form-employer').value,
        emp_type: document.getElementById('form-emp-type').value,
        basic_sal: parseFloat(document.getElementById('form-basic-sal').value) || 0,
        gross_sal: parseFloat(document.getElementById('form-gross-sal').value) || 0,
        net_sal: parseFloat(document.getElementById('form-net-sal').value) || 0,
        svc_months: parseInt(document.getElementById('form-svc-months').value) || 0,
        simah: parseInt(document.getElementById('form-simah').value) || 0,
        inquiries: parseInt(document.getElementById('form-inquiries').value) || 0,
        default_status: document.getElementById('form-default').value,
        blacklist: document.getElementById('form-blacklist').value,
        sal_attach: document.getElementById('form-sal-attach').value,
        
        loans_count: parseInt(document.getElementById('form-loans-cnt').value || 0),
        loans_total: parseFloat(document.getElementById('form-loans-tot').value || 0),
        real_estate_count: parseInt(document.getElementById('form-estate-cnt').value || 0),
        real_estate_total: parseFloat(document.getElementById('form-estate-tot').value || 0),
        cards_count: parseInt(document.getElementById('form-cards-cnt').value || 0),
        cards_total: parseFloat(document.getElementById('form-cards-tot').value || 0),
        finance_cos_count: parseInt(document.getElementById('form-fin-cnt').value || 0),
        finance_cos_total: parseFloat(document.getElementById('form-fin-tot').value || 0),
        monthly_installment: parseFloat(document.getElementById('form-installment').value) || 0,
        
        exec_requests_count: parseInt(document.getElementById('form-exec-cnt').value || 0),
        exec_requests_total: parseFloat(document.getElementById('form-exec-tot').value || 0),
        ind_exec_count: parseInt(document.getElementById('form-exec-ind-cnt').value || 0),
        ind_exec_total: parseFloat(document.getElementById('form-exec-ind-tot').value || 0),
        fin_exec_count: parseInt(document.getElementById('form-exec-fin-cnt').value || 0),
        fin_exec_total: parseFloat(document.getElementById('form-exec-fin-tot').value || 0),
        bank_exec_count: parseInt(document.getElementById('form-exec-bank-cnt').value || 0),
        bank_exec_total: parseFloat(document.getElementById('form-exec-bank-tot').value || 0),
        fees_percent: parseFloat(document.getElementById('form-fees-pct').value),
        workflow_stage: document.getElementById('form-workflow').value,
        
        analyst_name: document.getElementById('form-analyst-name').value,
        analyst_date: document.getElementById('form-analyst-date').value,
        analyst_recommendation: document.getElementById('form-analyst-rec').value,
        
        ops_name: document.getElementById('form-ops-name').value,
        ops_date: document.getElementById('form-ops-date').value,
        ops_decision: document.getElementById('form-ops-dec').value,
        
        finance_name: document.getElementById('form-fin-name').value,
        finance_date: document.getElementById('form-fin-date').value,
        finance_decision: document.getElementById('form-fin-dec').value,
        
        general_notes: document.getElementById('form-general-notes').value,
        completion_date: document.getElementById('form-completion-date').value
    };
    
    try {
        const res = await fetch(`${API_URL}/clients`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        
        const data = await res.json();
        if (data.success) {
            closeClientModal();
            showToast(data.message, "success");
            
            selectedClient = data.client;
            
            await refreshDashboard();
        } else {
            showToast(data.detail || "حدث خطأ أثناء الحفظ.", "error");
        }
    } catch (err) {
        console.error(err);
        showToast("فشل الاتصال بالخادم لحفظ الملف.", "error");
    }
}

// Delete Client File
async function deleteClientFile(rowIdx) {
    if (rowIdx === -1) return;
    if (!confirm("هل أنت متأكد من رغبتك في حذف هذا الملف الائتماني بشكل نهائي من قاعدة بيانات Excel؟")) return;
    
    try {
        const res = await fetch(`${API_URL}/clients/${rowIdx}`, {
            method: 'DELETE'
        });
        const data = await res.json();
        if (data.success) {
            closeDetailModal();
            showToast(data.message, "success");
            selectedClient = null;
            await refreshDashboard();
        } else {
            showToast(data.detail || "حدث خطأ أثناء الحذف.", "error");
        }
    } catch (err) {
        console.error(err);
        showToast("فشل الاتصال بالخادم لحذف الملف.", "error");
    }
}

// Print Debt Settlement Execution Order
function printExecutionOrder() {
    if (!selectedClient) {
        showToast("يرجى اختيار ملف عميل أولاً لطباعة تقرير أمر التنفيذ!", "error");
        return;
    }
    
    const client = selectedClient;
    
    // Format currency helper
    const fmt = (v) => formatCurrency(v);
    
    // Classification badge color
    let badgeClass = '';
    if (client.classification === 'A') badgeClass = 'badge-green';
    else if (client.classification === 'B') badgeClass = 'badge-yellow';
    else if (client.classification === 'C') badgeClass = 'badge-orange';
    else badgeClass = 'badge-red';
    
    const printWindow = window.open('', '_blank', 'width=900,height=900');
    printWindow.document.write(`
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>أمر تنفيذ عملية سداد - ${client.file_id}</title>
    <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap" rel="stylesheet">
    <style>
        @page {
            size: A4 portrait;
            margin: 10mm 12mm;
        }
        body {
            font-family: 'Tajawal', sans-serif;
            background-color: #ffffff;
            color: #111827;
            margin: 0;
            padding: 10px 15px;
            direction: rtl;
            font-size: 11.5px;
            line-height: 1.4;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #1e3a8a;
            padding-bottom: 8px;
            margin-bottom: 12px;
        }
        .logo-box {
            text-align: right;
        }
        .logo-title {
            font-size: 17px;
            font-weight: 800;
            color: #1e3a8a;
            margin: 0;
        }
        .logo-sub {
            font-size: 10px;
            color: #4b5563;
            margin: 2px 0 0 0;
            font-weight: bold;
        }
        .report-title-container {
            text-align: center;
        }
        .report-title {
            font-size: 15px;
            font-weight: 800;
            color: #ffffff;
            background: #1e3a8a;
            padding: 6px 18px;
            border-radius: 6px;
            display: inline-block;
        }
        .meta-box {
            text-align: left;
            font-size: 11px;
            color: #374151;
        }
        .meta-box div {
            margin-bottom: 2px;
        }
        .section-title {
            font-size: 12.5px;
            font-weight: 800;
            color: #1e3a8a;
            border-bottom: 1.5px solid #1e3a8a;
            padding-bottom: 4px;
            margin: 12px 0 6px 0;
        }
        .table-data {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 10px;
        }
        .table-data th, .table-data td {
            border: 1px solid #cbd5e1;
            padding: 6px 8px;
            text-align: right;
            font-size: 11px;
        }
        .table-data th {
            background-color: #f8fafc;
            color: #1e293b;
            font-weight: bold;
            width: 25%;
        }
        .table-data td {
            width: 25%;
        }
        .grid-2col-layout {
            display: flex;
            gap: 15px;
            margin-bottom: 10px;
        }
        .grid-2col-layout > div {
            flex: 1;
        }
        .grid-3col {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            margin-top: 8px;
            margin-bottom: 10px;
        }
        .approval-card {
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 10px;
            background-color: #f8fafc;
        }
        .approval-card-title {
            font-weight: 800;
            color: #1e3a8a;
            border-bottom: 1px dashed #cbd5e1;
            padding-bottom: 4px;
            margin-bottom: 8px;
            font-size: 11.5px;
            display: flex;
            justify-content: space-between;
        }
        .approval-body {
            font-size: 10.5px;
            color: #374151;
            display: flex;
            flex-direction: column;
            gap: 4px;
            min-height: 70px;
        }
        .signature-line {
            margin-top: auto;
            border-top: 1px dashed #94a3b8;
            padding-top: 4px;
            text-align: center;
            font-weight: bold;
            font-size: 10px;
            color: #64748b;
        }
        .badge-status {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-weight: 800;
            font-size: 11px;
            color: #fff;
        }
        .badge-green { background-color: #10b981; }
        .badge-yellow { background-color: #f59e0b; color: #111827; }
        .badge-orange { background-color: #f97316; }
        .badge-red { background-color: #ef4444; }
        
        .footer {
            margin-top: 15px;
            border-top: 1px solid #cbd5e1;
            padding-top: 8px;
            font-size: 10px;
            color: #64748b;
            display: flex;
            justify-content: space-between;
            font-weight: 500;
        }

        .highlight-red {
            color: #dc2626;
            font-weight: bold;
        }
        .highlight-green {
            color: #16a34a;
            font-weight: bold;
        }

        @media print {
            body {
                padding: 0;
                margin: 0;
                font-size: 11px;
            }
            .no-print {
                display: none;
            }
            .print-btn-container {
                display: none !important;
            }
        }
    </style>
</head>
<body>
    <div class="print-btn-container" style="display: flex; justify-content: center; margin-bottom: 20px; background: #f3f4f6; padding: 10px; border-radius: 8px; border: 1px solid #e5e7eb; gap: 10px;">
        <button onclick="window.print()" style="background: #10b981; color: #fff; border: none; padding: 8px 20px; font-weight: bold; border-radius: 6px; cursor: pointer; font-family: inherit;">
            طباعة التقرير PDF
        </button>
        <button onclick="window.close()" style="background: #ef4444; color: #fff; border: none; padding: 8px 20px; font-weight: bold; border-radius: 6px; cursor: pointer; font-family: inherit;">
            إغلاق النافذة
        </button>
    </div>

    <div class="header">
        <div class="logo-box">
            <h1 class="logo-title">نظام الجدارة الائتمانية</h1>
            <p class="logo-sub">وساطة تمويل وسداد مديونيات</p>
        </div>
        <div class="report-title-container">
            <div class="report-title">أمر تنفيذ عملية سداد</div>
            <div style="font-size: 11px; margin-top: 4px; color: #4b5563; font-weight: bold; font-family: 'Outfit';">Debt Settlement Execution Order</div>
        </div>
        <div class="meta-box">
            <div>رقم الملف: <strong style="font-family: Arial; color: #1e3a8a;">${client.file_id}</strong></div>
            <div>تاريخ الإدخال: <span style="font-family: Arial;">${client.date}</span></div>
            <div>تاريخ الطباعة: <span style="font-family: Arial;">${new Date().toISOString().split('T')[0]}</span></div>
        </div>
    </div>

    <!-- Section 1: Client profile -->
    <div class="section-title">أولاً: بيانات العميل الأساسية</div>
    <table class="table-data">
        <tr>
            <th>اسم العميل</th>
            <td><strong>${client.name}</strong></td>
            <th>رقم الهوية الوطنية</th>
            <td style="font-family: Arial;">${client.id_num}</td>
        </tr>
        <tr>
            <th>رقم الجوال</th>
            <td style="font-family: Arial;">${client.mobile}</td>
            <th>العمر</th>
            <td>${client.age} سنة</td>
        </tr>
        <tr>
            <th>جهة العمل</th>
            <td>${client.employer}</td>
            <th>نوع جهة العمل</th>
            <td>${client.emp_type}</td>
        </tr>
        <tr>
            <th>الراتب الأساسي</th>
            <td>${fmt(client.basic_sal)} ريال</td>
            <th>صافي الراتب المحول</th>
            <td class="highlight-green">${fmt(client.net_sal)} ريال</td>
        </tr>
    </table>

    <div class="grid-2col-layout">
        <!-- Section 2: Evaluation Results -->
        <div>
            <div class="section-title">ثانياً: نتائج التقييم الائتماني</div>
            <table class="table-data" style="margin-bottom: 0;">
                <tr>
                    <th>درجة سمة</th>
                    <td style="font-family: Arial; font-weight: bold;">${client.simah}</td>
                </tr>
                <tr>
                    <th>نسبة الاستقطاع (DTI)</th>
                    <td style="font-family: Arial; font-weight: bold;">${client.dti_pct}%</td>
                </tr>
                <tr>
                    <th>التصنيف الائتماني</th>
                    <td>
                        <span class="badge-status ${badgeClass}">${client.classification}</span>
                    </td>
                </tr>
                <tr>
                    <th>نتيجة التقييم النهائي</th>
                    <td style="font-weight: bold;">${client.decision}</td>
                </tr>
            </table>
        </div>

        <!-- Section 3: Executions -->
        <div>
            <div class="section-title">ثالثاً: ملف التنفيذات القضائية والنزاعات</div>
            <table class="table-data" style="margin-bottom: 0;">
                <tr>
                    <th>إجمالي عدد التنفيذات</th>
                    <td style="font-family: Arial;">${client.total_exec_count} طلبات</td>
                </tr>
                <tr>
                    <th>إجمالي مبالغ التنفيذات</th>
                    <td class="highlight-red">${fmt(client.total_exec_val)} ريال</td>
                </tr>
                <tr>
                    <th>درجة مخاطر التنفيذات</th>
                    <td>${client.risk_level}</td>
                </tr>
                <tr>
                    <th>درجة المخاطرة العامة</th>
                    <td>${client.risk_level}</td>
                </tr>
            </table>
        </div>
    </div>

    <!-- Section 4: Feasibility Study -->
    <div class="section-title">رابعاً: دراسة الجدوى التمويلية وعملية السداد</div>
    <table class="table-data">
        <tr>
            <th>مبلغ التمويل المتوقع</th>
            <td class="highlight-green" style="font-size: 13px;">${fmt(client.expected_funding)} ريال</td>
            <th>إجمالي المديونيات</th>
            <td>${fmt(client.total_debts)} ريال</td>
        </tr>
        <tr>
            <th>إجمالي التنفيذات القضائية</th>
            <td class="highlight-red">${fmt(client.total_exec_val)} ريال</td>
            <th>مبلغ السداد المطلوب (الكلي)</th>
            <td class="highlight-red" style="font-size: 13px;">${fmt(client.required_payment)} ريال</td>
        </tr>
        <tr>
            <th>الفائض المتوقع (الإجمالي)</th>
            <td>${fmt(client.gross_surplus)} ريال</td>
            <th>نسبة أتعاب الشركة</th>
            <td style="font-family: Arial; font-weight: bold;">${client.fees_percent * 100}%</td>
        </tr>
        <tr>
            <th>قيمة أتعاب الشركة الفعلية</th>
            <td>${fmt(client.company_fees)} ريال</td>
            <th>صافي الفائض للعميل (المتبقي)</th>
            <td class="highlight-green" style="font-size: 14px; font-weight: 800;">${fmt(client.net_surplus)} ريال</td>
        </tr>
        <tr>
            <th>قرار الجدوى الاقتصادية</th>
            <td colspan="3" style="font-weight: bold; font-size: 13px;">${client.feasibility_decision}</td>
        </tr>
    </table>

    <!-- Section 5: Approvals -->
    <div class="section-title">خامساً: الاعتمادات وتواقيع دورة العمل</div>
    <div class="grid-3col">
        <!-- Analyst -->
        <div class="approval-card">
            <div class="approval-card-title">
                <span>اعتماد محلل الائتمان</span>
                <span style="font-family: Arial; font-weight: normal; font-size: 10px; color: #4b5563;">${client.analyst_date || 'غير موقع'}</span>
            </div>
            <div class="approval-body">
                <div>الاسم: <strong>${client.analyst_name || '---'}</strong></div>
                <div style="margin-top: 4px; line-height: 1.4;">التوصية: <br><span style="color: #4b5563; font-style: italic;">${client.analyst_recommendation || 'في انتظار كتابة التوصية الائتمانية'}</span></div>
                <div class="signature-line">التوقيع / الختم</div>
            </div>
        </div>
        
        <!-- Operations -->
        <div class="approval-card">
            <div class="approval-card-title">
                <span>اعتماد مدير العمليات</span>
                <span style="font-family: Arial; font-weight: normal; font-size: 10px; color: #4b5563;">${client.ops_date || 'غير موقع'}</span>
            </div>
            <div class="approval-body">
                <div>الاسم: <strong>${client.ops_name || '---'}</strong></div>
                <div style="margin-top: 4px; line-height: 1.4;">القرار: <br><span style="color: #4b5563; font-style: italic;">${client.ops_decision || 'في انتظار قرار مدير العمليات'}</span></div>
                <div class="signature-line">التوقيع / الختم</div>
            </div>
        </div>
        
        <!-- Finance -->
        <div class="approval-card">
            <div class="approval-card-title">
                <span>اعتماد الإدارة المالية</span>
                <span style="font-family: Arial; font-weight: normal; font-size: 10px; color: #4b5563;">${client.finance_date || 'غير موقع'}</span>
            </div>
            <div class="approval-body">
                <div>الاسم: <strong>${client.finance_name || '---'}</strong></div>
                <div style="margin-top: 4px; line-height: 1.4;">القرار: <br><span style="color: #4b5563; font-style: italic;">${client.finance_decision || 'في انتظار قرار الإدارة المالية'}</span></div>
                <div class="signature-line">التوقيع / الختم</div>
            </div>
        </div>
    </div>

    <!-- Footer -->
    <div class="footer">
        <div>نظام تحليل العملاء والجدارة الائتمانية وسداد المديونيات (Excel Professional Workflow System)</div>
        <div>صفحة 1 من 1</div>
    </div>
</body>
</html>
    `);
    printWindow.document.close();
}

// Download Excel File
function downloadExcelFile() {
    window.open(`${API_URL}/download`, '_blank');
    showToast("جاري تحميل ملف Excel الخاص بالنظام...", "success");
}

// Trigger hidden file input click
function triggerExcelUpload() {
    document.getElementById('excel-file-upload-input').click();
}

// Upload Excel File
async function uploadExcelFile(input) {
    if (!input.files || input.files.length === 0) return;
    const file = input.files[0];
    
    if (!confirm(`هل أنت متأكد من استبدال ملف Excel الحالي بالملف المختار: "${file.name}"؟`)) {
        input.value = '';
        return;
    }
    
    showToast("جاري رفع وتحديث ملف Excel...", "success");
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const res = await fetch(`${API_URL}/upload`, {
            method: 'POST',
            body: formData
        });
        
        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || "فشل تحميل الملف");
        }
        
        const data = await res.json();
        showToast(data.message || "تم تحديث ملف Excel بنجاح!", "success");
        refreshDashboard();
    } catch (e) {
        console.error(e);
        showToast(`خطأ في رفع الملف: ${e.message}`, "error");
    } finally {
        input.value = '';
    }
}
