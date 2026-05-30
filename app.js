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
    document.getElementById('stat-surpluses').innerHTML = formatCurrency(stats.total_surpluses) + ' <span style="font-size: 11px; font-weight: normal; color: var(--text-secondary);">ريال</span>';
    document.getElementById('stat-fees').innerHTML = formatCurrency(stats.total_fees) + ' <span style="font-size: 11px; font-weight: normal; color: var(--text-secondary);">ريال</span>';
}

// Fetch Client Lists
async function fetchClients() {
    const res = await fetch(`${API_URL}/clients`);
    if (!res.ok) throw new Error("Failed to fetch clients");
    const data = await res.json();
    clientsData = data;
    
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
        tr.onclick = () => selectClient(c);
        
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
function selectClient(client) {
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
    
    inspector.innerHTML = `
        <!-- Decision Box -->
        <div class="decision-result-box ${decisionClass}">
            <span class="info-label" style="color: inherit; opacity: 0.85; margin-bottom: 2px;">القرار الائتماني النهائي</span>
            <div class="decision-title">${client.decision}</div>
            <div class="decision-reasons">${client.brief_reasons}</div>
            
            <div class="score-badge-circle" style="border-color: ${
                decisionClass === 'qualified' ? '#10B981' : decisionClass === 'reserved' ? '#F59E0B' : decisionClass === 'exception' ? '#F97316' : '#EF4444'
            };">
                <h5>${client.total_pts}</h5>
                <span>نقاط الجدارة /105</span>
            </div>
            
            <div style="font-size: 11px; font-weight: bold; color: var(--text-secondary);">
                نسبة الجدارة: ${client.score_pct}% | تصنيف: [${client.classification}]
            </div>
        </div>
        
        <!-- Action Buttons -->
        <div style="display: flex; gap: 8px; margin-bottom: 16px;">
            <button class="btn btn-secondary" style="flex: 1; font-size: 12px; padding: 8px 12px;" onclick="openEditClientModal()">
                <i class="fa-solid fa-user-gear"></i> تعديل ملف العميل
            </button>
            <button class="btn btn-secondary" style="flex: 1; font-size: 12px; padding: 8px 12px; border-color: rgba(239, 68, 68, 0.3); color: #EF4444;" onclick="deleteClientFile(${client.row_idx})">
                <i class="fa-solid fa-trash-can"></i> حذف الملف
            </button>
        </div>

        <!-- Tab Content 1: Personal & Debts -->
        <div class="tab-content" id="tab-content-personal">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-user" style="color: #6366F1;"></i> بيانات العميل الأساسية</h4>
                <div class="grid-2col">
                    <div class="info-item">
                        <span class="info-label">اسم العميل</span>
                        <span class="info-value" style="color: #818CF8;">${client.name}</span>
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
                    <div class="info-item">
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
                        <span class="info-value currency">${formatCurrency(client.basic_sal)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي الراتب</span>
                        <span class="info-value currency">${formatCurrency(client.gross_sal)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">صافي الراتب المحول</span>
                        <span class="info-value currency" style="color: #10B981;">${formatCurrency(client.net_sal)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">مدة الخدمة</span>
                        <span class="info-value">${client.svc_months} شهر</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab Content 2: Risk Engine -->
        <div class="tab-content" id="tab-content-risk">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-circle-exclamation" style="color: #F59E0B;"></i> الالتزامات ونسبة الاستقطاع</h4>
                <div class="grid-2col" style="margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px dashed var(--border-color);">
                    <div class="info-item">
                        <span class="info-label">إجمالي المديونيات</span>
                        <span class="info-value currency" style="color: #EF4444;">${formatCurrency(client.total_debts)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">الأقساط الشهرية</span>
                        <span class="info-value currency">${formatCurrency(client.monthly_installment)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">نسبة الاستقطاع (DTI)</span>
                        <span class="info-value" style="color: ${client.dti_pct > 70 ? '#EF4444' : client.dti_pct > 50 ? '#F97316' : '#10B981'};">${client.dti_pct}%</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">مستوى المخاطرة الائتمانية</span>
                        <span class="info-value"><span class="badge-risk ${riskClass}">${client.risk_level.split(" ").slice(0,-1).join(" ")}</span></span>
                    </div>
                </div>
                <div class="grid-2col" style="font-size: 11px;">
                    <div class="info-item">
                        <span class="info-label">القروض / البطاقات</span>
                        <span class="info-value" style="font-weight: 500;">
                            ${client.loans_count} شخصي (${formatCurrency(client.loans_total)}) | ${client.cards_count} بطاقات (${formatCurrency(client.cards_total)})
                        </span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">العقاري / شركات التمويل</span>
                        <span class="info-value" style="font-weight: 500;">
                            ${client.real_estate_count} عقاري (${formatCurrency(client.real_estate_total)}) | ${client.finance_cos_count} شركات (${formatCurrency(client.finance_cos_total)})
                        </span>
                    </div>
                </div>
            </div>

            <div class="inspector-section">
                <h4><i class="fa-solid fa-gavel" style="color: #EF4444;"></i> التنفيذات القضائية والنزاعات</h4>
                <div class="grid-2col" style="margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px dashed var(--border-color);">
                    <div class="info-item">
                        <span class="info-label">عدد طلبات التنفيذ (أمر 46)</span>
                        <span class="info-value" style="color: ${client.exec_requests_count > 0 ? '#EF4444' : 'inherit'};">${client.exec_requests_count} طلبات</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي مبالغ التنفيذات</span>
                        <span class="info-value currency" style="color: ${client.exec_requests_total > 0 ? '#EF4444' : 'inherit'};">${formatCurrency(client.exec_requests_total)}</span>
                    </div>
                </div>
                <div class="grid-3col" style="font-size: 11px;">
                    <div class="info-item">
                        <span class="info-label">تنفيذات الأفراد</span>
                        <span class="info-value">${client.ind_exec_count} (${formatCurrency(client.ind_exec_total)})</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">تنفيذات الشركات</span>
                        <span class="info-value" style="color: ${client.fin_exec_count > 3 ? '#EF4444' : 'inherit'};">${client.fin_exec_count} (${formatCurrency(client.fin_exec_total)})</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">تنفيذات البنوك</span>
                        <span class="info-value">${client.bank_exec_count} (${formatCurrency(client.bank_exec_total)})</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab Content 3: Feasibility & Surpluses -->
        <div class="tab-content" id="tab-content-feasibility">
            <div class="inspector-section">
                <h4><i class="fa-solid fa-chart-pie" style="color: #EC4899;"></i> الجدوى التمويلية والفائض الاقتصادي</h4>
                <div class="grid-2col" style="margin-bottom: 16px; padding-bottom: 16px; border-bottom: 1px dashed var(--border-color);">
                    <div class="info-item">
                        <span class="info-label">مبلغ التمويل المتوقع</span>
                        <span class="info-value currency" style="color: #10B981; font-size: 16px;">${formatCurrency(client.expected_funding)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">مبلغ السداد الإجمالي المطلوب</span>
                        <span class="info-value currency" style="color: #EF4444; font-size: 16px;">${formatCurrency(client.required_payment)}</span>
                    </div>
                </div>
                
                <div class="grid-2col">
                    <div class="info-item">
                        <span class="info-label">قيمة أتعاب الشركة المتوقعة (${client.fees_percent * 100}%)</span>
                        <span class="info-value currency" style="color: #F59E0B;">${formatCurrency(client.company_fees)}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">إجمالي الفائض المالي</span>
                        <span class="info-value currency" style="color: ${client.gross_surplus > 0 ? '#10B981' : '#EF4444'};">${formatCurrency(client.gross_surplus)}</span>
                    </div>
                    <div class="info-item" style="grid-column: span 2; background: rgba(0,0,0,0.15); padding: 12px; border-radius: 8px; margin-top: 8px; display: flex; flex-direction: row; justify-content: space-between; align-items: center;">
                        <div>
                            <span class="info-label" style="display: block;">صافي الفائض للعميل (المتبقي بجيبه)</span>
                            <span class="info-value currency" style="color: ${client.net_surplus > 0 ? '#10B981' : '#EF4444'}; font-size: 16px;">${formatCurrency(client.net_surplus)}</span>
                        </div>
                        <div style="text-align: left;">
                            <span class="info-label" style="display: block;">قرار الجدوى</span>
                            <span class="info-value" style="color: var(--text-primary); font-size: 13px; font-weight: bold;">${client.feasibility_decision}</span>
                        </div>
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
                <h4><i class="fa-solid fa-signature" style="color: #818CF8;"></i> Workflow ومراحل الاعتماد والـ Audit</h4>
                <div style="font-size: 12px; color: var(--text-secondary); line-height: 1.8; display: flex; flex-direction: column; gap: 8px;">
                    <div style="display: flex; gap: 8px; align-items: center;">
                        <span class="badge" style="background: rgba(16,185,129,0.1); color: #10B981; padding: 2px 6px;">معتمد</span>
                        <span><strong>محلل الائتمان (أحمد عسيري):</strong> موصى به للتمويل وسداد مديونية العميل</span>
                    </div>
                    <div style="display: flex; gap: 8px; align-items: center;">
                        <span class="badge" style="background: rgba(16,185,129,0.1); color: #10B981; padding: 2px 6px;">معتمد</span>
                        <span><strong>مدير العمليات (خالد الحربي):</strong> معتمد للتنفيذ ومطابقة الجدارة</span>
                    </div>
                    <div style="display: flex; gap: 8px; align-items: center;">
                        <span class="badge" style="background: rgba(16,185,129,0.1); color: #10B981; padding: 2px 6px;">معتمد</span>
                        <span><strong>الإدارة المالية (فهد العتيبي):</strong> مقبول ومعتمد لصرف مبالغ السداد والمديونيات</span>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    switchInspectorTab(currentInspectorTab);
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

// Render empty inspector placeholder
function renderEmptyInspector() {
    document.getElementById('inspector-tab-headers').style.display = 'none';
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
        fees_percent: client.fees_percent
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
        if (c.decision.includes("🟢")) qualifiedCount++;
        else if (c.decision.includes("🟡")) reservedCount++;
        else if (c.decision.includes("🟠")) exceptionCount++;
        else if (c.decision.includes("🔴")) rejectedCount++;
    });
    
    const pieCtx = document.getElementById('classificationChart').getContext('2d');
    if (classificationChart) classificationChart.destroy();
    
    classificationChart = new Chart(pieCtx, {
        type: 'doughnut',
        data: {
            labels: ['🟢 مؤهل', '🟡 مؤهل بتحفظ', '🟠 استثناء', '🔴 مرفوض'],
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

    document.getElementById('client-modal').classList.add('active');
}

// Open modal for editing selected client
function openEditClientModal() {
    if (!selectedClient) return;
    
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
        age: parseInt(document.getElementById('form-age').value),
        employer: document.getElementById('form-employer').value,
        emp_type: document.getElementById('form-emp-type').value,
        basic_sal: parseFloat(document.getElementById('form-basic-sal').value),
        gross_sal: parseFloat(document.getElementById('form-gross-sal').value),
        net_sal: parseFloat(document.getElementById('form-net-sal').value),
        svc_months: parseInt(document.getElementById('form-svc-months').value),
        simah: parseInt(document.getElementById('form-simah').value),
        inquiries: parseInt(document.getElementById('form-inquiries').value),
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
        monthly_installment: parseFloat(document.getElementById('form-installment').value),
        
        exec_requests_count: parseInt(document.getElementById('form-exec-cnt').value || 0),
        exec_requests_total: parseFloat(document.getElementById('form-exec-tot').value || 0),
        ind_exec_count: parseInt(document.getElementById('form-exec-ind-cnt').value || 0),
        ind_exec_total: parseFloat(document.getElementById('form-exec-ind-tot').value || 0),
        fin_exec_count: parseInt(document.getElementById('form-exec-fin-cnt').value || 0),
        fin_exec_total: parseFloat(document.getElementById('form-exec-fin-tot').value || 0),
        bank_exec_count: parseInt(document.getElementById('form-exec-bank-cnt').value || 0),
        bank_exec_total: parseFloat(document.getElementById('form-exec-bank-tot').value || 0),
        fees_percent: parseFloat(document.getElementById('form-fees-pct').value)
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
