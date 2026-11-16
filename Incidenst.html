<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <meta name="theme-color" content="#2c3e50">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <title>Sukiya POS - Quản Lý Sự Cố</title>
  
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  
  <!-- SheetJS CDN (Xuất Excel) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <!-- Chart.js CDN (Biểu đồ thống kê) -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  
  <style>
    ::-webkit-scrollbar { height: 6px; width: 6px; }
    ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
    ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
    .table-scrollbar::-webkit-scrollbar { height: 8px; }
    .table-scrollbar::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 4px; }

    /* Định dạng khi in Biên bản A4 */
    @media print {
      body * { visibility: hidden; }
      #printReportArea, #printReportArea * { visibility: visible; }
      #printReportArea {
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        background: white !important;
        padding: 0 !important;
        margin: 0 !important;
        box-shadow: none !important;
      }
      .no-print { display: none !important; }
    }
  </style>

  <!-- Firebase & Logic Script -->
  <script type="module">
    import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
    import { 
      getFirestore, collection, getDocs, doc, setDoc, deleteDoc, query, where 
    } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
    import { 
      getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged, createUserWithEmailAndPassword 
    } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";

    // 1. CẤU HÌNH FIREBASE
    const firebaseConfig = {
      apiKey: "AIzaSyBsqUZo_6BryZ5nJzxXB4xdSM4shloO9Ps",
      authDomain: "abico-63099.firebaseapp.com",
      projectId: "abico-63099",
      storageBucket: "abico-63099.firebasestorage.app",
      messagingSenderId: "764504395267",
      appId: "1:764504395267:web:2b6090aca2a8d2bdf48836"
    };

    const app = !getApps().length ? initializeApp(firebaseConfig) : getApps()[0];
    const db = getFirestore(app);
    const auth = getAuth(app);

    const adminAppName = "AdminAppCreator_" + Math.random();
    const adminApp = initializeApp(firebaseConfig, adminAppName);
    const adminAuth = getAuth(adminApp);

    // 2. CẤU HÌNH GOOGLE APPS SCRIPT (Gửi Telegram Bot)
    const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTt6B5_cTTEM4cbNUtKuN7079B0Db7984EuP2c34wD7y7o1ZYOVfaqyHmyRwHA1pI/exec";

    // 3. DANH MỤC CHI TIẾT CATEGORY 2
    const category2Options = {
      "Network": ["Wi-Fi", "LAN", "Internet(ISP)", "Router", "Switch", "Access Point"],
      "POS": ["POS Printer", "Scanner", "OES", "POS Network", "ALL POS Error", "POS 01 Cash Error", "POS 02 Cash Error"],
      "Printer": ["KP1", "KP2", "KP3", "KP4", "KP5", "Do not print bill", "Blurry Bill", "All KP"],
      "Server": ["Server Down", "Database", "Backup", "Storage", "Access", "Security"],
      "CCTV": ["Camera", "NVR", "Recording", "Playback", "Camera Offline", "HDD"],
      "Hardware": ["PC", "Monitor", "Keyboard", "Mouse", "Printer", "Scanner", "Tell Phone"],
      "Software": ["Application", "Windows", "Update", "License", "System Error"],
      "Handy": ["Won't power on", "Not charging", "Doesn't connect to Wi-Fi"],
      "Attendance": ["Network Error", "Data not loading", "Won't power on"],
      "KD": ["All KD Error", "KD1", "KD2", "HOT", "OES", "KD Network", "TO & DL", "Do not show the bill"],
      "Khác": ["Khác"]
    };

    let systemKnowledgeBase = [];

    // TẢI DỮ LIỆU KNOWLEDGE BASE TỪ FIREBASE
    async function loadKnowledgeBase() {
        try {
            const qSnap = await getDocs(collection(db, "knowledge_base"));
            systemKnowledgeBase = [];
            qSnap.forEach(doc => {
                systemKnowledgeBase.push({ id: doc.id, ...doc.data() });
            });
        } catch (e) {
            console.log("Lỗi tải KB:", e);
        }
    }

    // 5. DANH SÁCH CHI NHÁNH MẶC ĐỊNH
    const defaultStores = [
      { id: "1001", name: "SK Aeon Mall Bình Tân", code: "1001", address: "Aeon Mall Bình Tân, Q.Bình Tân", ipPrefix: "10.231.1" },
      { id: "1017", name: "Tạ Quang Bửu", code: "1017", address: "Tạ Quang Bửu, Q.8", ipPrefix: "10.231.15" },
      { id: "1018", name: "Lê Thánh Tôn", code: "1018", address: "Lê Thánh Tôn, Q.1", ipPrefix: "10.231.16" },
      { id: "1019", name: "Song Hành", code: "1019", address: "Song Hành, Q.2", ipPrefix: "10.231.17" },
      { id: "1020", name: "Phan Xích Long", code: "1020", address: "Phan Xích Long, Phú Nhuận", ipPrefix: "10.231.18" },
      { id: "1021", name: "Vivo City", code: "1021", address: "Vivo City, Q.7", ipPrefix: "10.231.5" },
      { id: "1027", name: "Lê Văn Sỹ", code: "1027", address: "Lê Văn Sỹ, Tân Bình", ipPrefix: "10.231.24" },
      { id: "1028", name: "Nguyễn Ảnh Thủ", code: "1028", address: "Nguyễn Ảnh Thủ, Hóc Môn", ipPrefix: "10.231.25" },
      { id: "1029", name: "Nguyễn Gia Trí", code: "1029", address: "Nguyễn Gia Trí, Bình Thạnh", ipPrefix: "10.231.26" }
    ];

    let currentStores = [];
    let systemUsers = []; 
    let systemIncidents = [];
    let editingIncidentId = null; 
    let incidentFilterState = { search: "", branch: "Tất cả", status: "Tất cả", dateFrom: "", dateTo: "" };
    let categoryChartInstance = null;
    let branchChartInstance = null;
    let readNotificationIds = JSON.parse(localStorage.getItem('sukiya_read_notifs') || '[]');

    let currentIncidentPage = 1;
    const rowsPerPage = 20;

    window.currentUserRole = 'Nhân viên'; 
    window.currentUserStore = '';
    window.currentUserName = '';
    window.currentUserEmail = '';
    window.currentEmployeeCode = '';

    window.showToast = function(msg, type = 'error') {
      const toast = document.getElementById('sysToast');
      if(!toast) return;
      toast.innerText = msg;
      toast.className = `fixed top-5 right-5 px-5 py-3 rounded-lg shadow-2xl z-[999999] font-bold text-sm text-white transform transition-all ${type === 'success' ? 'bg-green-600' : 'bg-red-600'}`;
      setTimeout(() => { toast.className = 'hidden'; }, 4000);
    };

    window.handleLogin = async function() {
      const email = document.getElementById('loginUsername').value.trim();
      const pass = document.getElementById('loginPassword').value.trim();
      const btnLogin = document.getElementById('btnLogin');
      if(!email || !pass) { showToast("Vui lòng nhập Email và Mật khẩu!"); return; }
      try {
        btnLogin.disabled = true; btnLogin.innerText = 'Đang đăng nhập...';
        await signInWithEmailAndPassword(auth, email, pass);
      } catch (error) {
        showToast("Sai Email đăng nhập hoặc Mật khẩu!");
        btnLogin.disabled = false; btnLogin.innerText = 'Đăng nhập';
      }
    };

    window.handleLogout = async function() { await signOut(auth); };

    onAuthStateChanged(auth, async (user) => {
      const loginScreen = document.getElementById('loginScreen');
      const mainAppContainer = document.getElementById('mainAppContainer');

      if (user) {
        if(loginScreen) loginScreen.style.display = 'none';
        if(mainAppContainer) mainAppContainer.classList.remove('hidden');
        
        window.currentUserEmail = user.email;
        try {
          await loadUsers();
          const q = query(collection(db, "users"), where("email", "==", user.email));
          const querySnapshot = await getDocs(q);
          if (!querySnapshot.empty) {
              const userData = querySnapshot.docs[0].data();
              window.currentUserRole = userData.position || 'Nhân viên';
              window.currentUserStore = userData.storeCode || '';
              window.currentUserName = userData.full_name || user.email;
              window.currentEmployeeCode = userData.employee_code || user.email;
          } else if (user.email === 'gest@gmail.com' || user.email === 'abicovip@gmail.com') {
              window.currentUserRole = 'Quản trị viên';
              window.currentUserName = 'Root Admin';
              window.currentEmployeeCode = 'ADMIN-ROOT';
          } else {
              window.currentUserRole = 'Nhân viên';
              window.currentUserName = user.email;
              window.currentEmployeeCode = user.email;
          }
        } catch (e) {
          window.currentUserName = user.email;
          window.currentEmployeeCode = user.email;
        }

        const roleDisplay = document.getElementById('displayUserRole');
        if(roleDisplay) roleDisplay.innerText = `${window.currentUserName} (${window.currentUserRole})`;
        
        const btnAdmin = document.getElementById('btnAdminUsers');
        const btnAdminKB = document.getElementById('btnAdminKB');
        if (window.currentUserRole === 'Quản trị viên') {
            if(btnAdmin) btnAdmin.classList.remove('hidden');
            if(btnAdminKB) btnAdminKB.classList.remove('hidden');
        } else {
            if(btnAdmin) btnAdmin.classList.add('hidden');
            if(btnAdminKB) btnAdminKB.classList.add('hidden');
        }
        
        await loadKnowledgeBase();
        await loadStores();
      } else {
        if(loginScreen) loginScreen.style.display = 'flex';
        if(mainAppContainer) mainAppContainer.classList.add('hidden');
      }
    });

    function getTodayString() { return new Date().toISOString().split('T')[0]; }
    function getMonthString() { const d = new Date(); return `${d.getFullYear()}${(d.getMonth() + 1).toString().padStart(2, '0')}`; }
    function getCurrentDateTime() {
      const d = new Date();
      return `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')} ${d.getHours().toString().padStart(2, '0')}:${d.getMinutes().toString().padStart(2, '0')}:${d.getSeconds().toString().padStart(2, '0')}`;
    }
    function formatDateTimeView(isoString) {
      if (!isoString) return '--';
      const d = new Date(isoString);
      return d.toLocaleDateString('vi-VN') + ' ' + d.toLocaleTimeString('vi-VN', {hour: '2-digit', minute:'2-digit'});
    }

    function getFullStoreList() {
      const combinedStoresMap = new Map();
      defaultStores.forEach(s => combinedStoresMap.set(s.code, s));
      if (currentStores && currentStores.length > 0) {
          currentStores.forEach(s => combinedStoresMap.set(s.code, s));
      }
      return Array.from(combinedStoresMap.values()).sort((a, b) => a.code.localeCompare(b.code));
    }

    function generateUniqueIncidentCode() {
      let maxNum = 0;
      systemIncidents.forEach(inc => {
        if (inc.id && inc.id.startsWith('SC-')) {
          const numPart = parseInt(inc.id.replace('SC-', ''), 10);
          if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
        }
      });
      return 'SC-' + (maxNum + 1).toString().padStart(6, '0');
    }

    function getAssignedUserName(assigneeKey) {
        if (!assigneeKey) return 'Chưa gán';
        const uObj = systemUsers.find(u => u.employee_code === assigneeKey || u.email === assigneeKey || u.full_name === assigneeKey);
        return uObj ? (uObj.full_name || uObj.email.split('@')[0]) : assigneeKey;
    }

    async function sendTelegramAlert(incData, isNew, customTitle = null) {
      const title = customTitle || (isNew ? "🚨 <b>CẢNH BÁO SỰ CỐ KHẨN CẤP</b> 🚨" : "🔔 <b>CẬP NHẬT SỰ CỐ</b> 🔔");
      let imgHTML = incData.imageUrl ? `<a href="${incData.imageUrl}">&#8205;</a>` : "";
      const message = `
${title}
<b>Mã phiếu:</b> ${incData.id}
<b>Cửa hàng:</b> ${incData.storeCode}
<b>Người báo:</b> ${incData.contactPerson} (${incData.phoneNumber})
<b>Phân loại:</b> ${incData.categories.join(', ')} ${incData.category2 ? `(${incData.category2})` : ''}
<b>Trạng thái:</b> ${incData.status}
<b>Phân công (PIC):</b> ${getAssignedUserName(incData.assignee)}
<b>Mô tả:</b> <i>${incData.issue}</i>
${incData.tempSolution ? `\n<b>Xử lý tạm:</b> ${incData.tempSolution}` : ''}
${incData.result ? `\n<b>Xử lý lâu dài:</b> ${incData.result}` : ''}${imgHTML}`;

      try {
        await fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ action: "sendTelegram", message })
        });
      } catch (error) { console.log("Lỗi Telegram:", error); }
    }

    window.checkSLAReminders = async function() {
        const now = new Date();
        let overdueCount = 0;

        for (const inc of systemIncidents) {
            if (inc.status === 'Mới ghi nhận' && inc.createdAt) {
                const createdTime = new Date(inc.createdAt);
                const diffMinutes = Math.floor((now - createdTime) / (1000 * 60));
                if (diffMinutes >= 30) {
                    overdueCount++;
                    await sendTelegramAlert(inc, false, `⏰ <b>CẢNH BÁO SLA (TRỄ ${diffMinutes} PHÚT CHƯA NHẬN VIỆC)</b> ⏰`);
                }
            }
        }

        if (overdueCount > 0) {
            showToast(`⚠️ Đã gửi cảnh báo Telegram cho ${overdueCount} phiếu quá hạn 30 phút!`, "error");
        } else {
            showToast("✅ Tất cả phiếu đều đang nằm trong thời gian cam kết SLA!", "success");
        }
    };

    function updateInAppNotifications() {
        const bellBadge = document.getElementById('notifBadge');
        const notifList = document.getElementById('notifListContainer');
        if (!bellBadge || !notifList) return;

        const relevantIncidents = systemIncidents.filter(inc => {
            const isMyStore = window.currentUserRole === 'Quản trị viên' || inc.storeCode === window.currentUserStore;
            const isAssignedToMe = inc.assignee === window.currentEmployeeCode || inc.assignee === window.currentUserEmail;
            return (isMyStore && inc.status !== 'Đã xử lý') || (isAssignedToMe && inc.status !== 'Đã xử lý');
        });

        const unreadItems = relevantIncidents.filter(inc => !readNotificationIds.includes(inc.id + '_' + inc.status));

        if (unreadItems.length > 0) {
            bellBadge.innerText = unreadItems.length > 9 ? '9+' : unreadItems.length;
            bellBadge.classList.remove('hidden');
        } else {
            bellBadge.classList.add('hidden');
        }

        if (relevantIncidents.length === 0) {
            notifList.innerHTML = `<div class="p-4 text-center text-xs text-gray-500">Không có thông báo sự cố nào cần xử lý.</div>`;
            return;
        }

        const fullStoreList = getFullStoreList();
        notifList.innerHTML = relevantIncidents.slice(0, 10).map(inc => {
            const notifKey = inc.id + '_' + inc.status;
            const isUnread = !readNotificationIds.includes(notifKey);
            const storeInfo = fullStoreList.find(s => s.code === inc.storeCode);
            const storeName = storeInfo ? storeInfo.name : inc.storeCode;
            const isAssignedToMe = inc.assignee === window.currentEmployeeCode || inc.assignee === window.currentUserEmail;

            return `
            <div onclick="handleNotificationClick('${inc.id}', '${notifKey}')" class="p-3 border-b hover:bg-blue-50 cursor-pointer transition flex items-start gap-2.5 ${isUnread ? 'bg-amber-50/60' : 'bg-white'}">
                <div class="text-lg mt-0.5">${isAssignedToMe ? '🛠️' : '🚨'}</div>
                <div class="flex-1 min-w-0">
                    <div class="flex justify-between items-center">
                        <span class="font-bold text-xs text-red-700">${inc.id} - ${storeName}</span>
                        ${isUnread ? `<span class="w-2 h-2 bg-red-600 rounded-full"></span>` : ''}
                    </div>
                    <p class="text-[11px] text-gray-700 truncate mt-0.5">${inc.issue}</p>
                    <div class="flex justify-between items-center mt-1 text-[10px] text-gray-500">
                        <span>${isAssignedToMe ? 'Được gán cho bạn' : inc.status}</span>
                        <span>${formatDateTimeView(inc.createdAt)}</span>
                    </div>
                </div>
            </div>`;
        }).join('');
    }

    window.toggleNotificationDropdown = function(event) {
        event.stopPropagation();
        const dropdown = document.getElementById('notifDropdown');
        if (dropdown) dropdown.classList.toggle('hidden');
    };

    window.handleNotificationClick = function(incId, notifKey) {
        if (!readNotificationIds.includes(notifKey)) {
            readNotificationIds.push(notifKey);
            localStorage.setItem('sukiya_read_notifs', JSON.stringify(readNotificationIds));
        }
        document.getElementById('notifDropdown').classList.add('hidden');
        updateInAppNotifications();
        window.openIncidentModal(incId);
    };

    window.markAllNotificationsRead = function() {
        systemIncidents.forEach(inc => {
            const key = inc.id + '_' + inc.status;
            if (!readNotificationIds.includes(key)) readNotificationIds.push(key);
        });
        localStorage.setItem('sukiya_read_notifs', JSON.stringify(readNotificationIds));
        updateInAppNotifications();
    };

    document.addEventListener('click', function(e) {
        const dropdown = document.getElementById('notifDropdown');
        const bellBtn = document.getElementById('notifBellBtn');
        if (dropdown && !dropdown.classList.contains('hidden') && !dropdown.contains(e.target) && !bellBtn.contains(e.target)) {
            dropdown.classList.add('hidden');
        }
    });

    window.openKnowledgeBaseModal = function() {
        document.body.classList.add('overflow-hidden');
        document.getElementById('kbModal').classList.remove('hidden');
        renderKnowledgeBaseModalList('');
    };

    window.closeKnowledgeBaseModal = function() {
        document.body.classList.remove('overflow-hidden');
        document.getElementById('kbModal').classList.add('hidden');
    };

    window.renderKnowledgeBaseModalList = function(keyword = '') {
        const container = document.getElementById('kbContentList');
        if (!container) return;
        const kw = keyword.toLowerCase().trim();
        const filtered = systemKnowledgeBase.filter(item => 
            item.category.toLowerCase().includes(kw) || 
            (item.category2 && item.category2.toLowerCase().includes(kw)) ||
            item.title.toLowerCase().includes(kw) || 
            (Array.isArray(item.steps) ? item.steps.join(' ').toLowerCase() : '').includes(kw)
        );

        if (filtered.length === 0) {
            container.innerHTML = `<div class="p-6 text-center text-gray-500 text-sm">Không tìm thấy hướng dẫn phù hợp.</div>`;
            return;
        }

        container.innerHTML = filtered.map(item => {
            const stepsHtml = Array.isArray(item.steps) ? item.steps.map(s => `<li>${s}</li>`).join('') : `<li>${item.steps}</li>`;
            return `
            <div class="bg-white border rounded-xl p-4 shadow-sm hover:border-blue-300 transition">
                <div class="flex items-center gap-2 mb-2">
                    <span class="bg-blue-100 text-blue-800 text-[10px] font-bold px-2.5 py-0.5 rounded-full uppercase">${item.category}</span>
                    ${item.category2 ? `<span class="bg-indigo-100 text-indigo-800 font-bold px-2 py-0.5 rounded text-[10px]">${item.category2}</span>` : ''}
                    <h4 class="font-bold text-slate-800 text-sm ml-1">${item.title}</h4>
                </div>
                <ul class="space-y-1.5 text-xs text-gray-700 bg-slate-50 p-3 rounded-lg border border-slate-100">
                    ${stepsHtml}
                </ul>
            </div>`;
        }).join('');
    };

    window.openPrintReportModal = function(incId) {
        const inc = systemIncidents.find(i => i.id === incId);
        if (!inc) return;

        document.body.classList.add('overflow-hidden');
        document.getElementById('printReportModal').classList.remove('hidden');

        const fullStoreList = getFullStoreList();
        const storeInfo = fullStoreList.find(s => s.code === inc.storeCode);
        const storeName = storeInfo ? storeInfo.name : inc.storeCode;
        const storeAddr = storeInfo ? storeInfo.address : '--';
        const cats = inc.categories && inc.categories.length > 0 ? inc.categories.join(', ') : '--';

        document.getElementById('prTicketId').innerText = inc.id;
        document.getElementById('prCreatedDate').innerText = formatDateTimeView(inc.createdAt);
        document.getElementById('prStoreName').innerText = `${inc.storeCode} - ${storeName}`;
        document.getElementById('prStoreAddr').innerText = storeAddr;
        document.getElementById('prContact').innerText = `${inc.contactPerson || '--'} (${inc.phoneNumber || '--'})`;
        document.getElementById('prCategory').innerText = `${cats} ${inc.category2 ? `(${inc.category2})` : ''}`;
        document.getElementById('prPic').innerText = getAssignedUserName(inc.assignee);
        document.getElementById('prStatus').innerText = inc.status || 'Mới ghi nhận';
        document.getElementById('prIssue').innerText = inc.issue || '--';
        document.getElementById('prTempSol').innerText = inc.tempSolution || 'Chưa có ghi nhận';
        document.getElementById('prLongSol').innerText = inc.result || 'Chưa có ghi nhận';
        document.getElementById('prResolvedDate').innerText = inc.resolvedAt ? formatDateTimeView(inc.resolvedAt) : 'Đang xử lý';

        const imgWrapper = document.getElementById('prImageWrapper');
        const imgEl = document.getElementById('prImage');
        if (inc.imageUrl && inc.imageUrl.trim() !== '') {
            imgEl.src = inc.imageUrl;
            imgWrapper.classList.remove('hidden');
        } else {
            imgWrapper.classList.add('hidden');
        }
    };

    window.closePrintReportModal = function() {
        document.body.classList.remove('overflow-hidden');
        document.getElementById('printReportModal').classList.add('hidden');
    };

    async function loadStores() {
      try {
        const qSnap = await getDocs(collection(db, "stores"));
        if (!qSnap.empty) {
          const loaded = [];
          qSnap.forEach((doc) => loaded.push({ id: doc.id, ...doc.data() }));
          currentStores = loaded;
        } else { currentStores = [...defaultStores]; }
      } catch (e) { currentStores = [...defaultStores]; }
      await loadUsers(); 
      await loadIncidents();
    }

    async function loadUsers() {
      try {
        const qSnap = await getDocs(collection(db, "users"));
        const loaded = [];
        qSnap.forEach((doc) => loaded.push({ id: doc.id, ...doc.data() }));
        systemUsers = loaded;
        if(document.getElementById('userManagerModal') && !document.getElementById('userManagerModal').classList.contains('hidden')){
            renderUserTable();
        }
      } catch (e) { systemUsers = []; }
    }

    window.openUserManager = function() {
        document.body.classList.add('overflow-hidden');
        document.getElementById('userManagerModal').classList.remove('hidden');
        resetUserForm();
        renderUserTable();
        const fullStoreList = getFullStoreList();
        const umStoreSelect = document.getElementById('umStoreCode');
        if(umStoreSelect) umStoreSelect.innerHTML = `<option value="">Tất cả / Văn phòng</option>` + fullStoreList.map(s => `<option value="${s.code}">${s.code} - ${s.name}</option>`).join('');
    };

    window.closeUserManager = function() { 
        document.body.classList.remove('overflow-hidden');
        document.getElementById('userManagerModal').classList.add('hidden'); 
    };

    window.resetUserForm = function() {
        isEditingUser = false;
        document.getElementById('userForm').reset();
        document.getElementById('umEmail').readOnly = false;
        document.getElementById('umEmail').classList.remove('bg-gray-100', 'text-gray-500');
        document.getElementById('umPasswordContainer').classList.remove('hidden');
        document.getElementById('umPassword').required = true;
        document.getElementById('btnSubmitUser').innerText = "Tạo Tài Khoản Mới";
        document.getElementById('umFormTitle').innerText = "Thêm Tài Khoản Mới";
    };

    let isEditingUser = false;
    function renderUserTable() {
        const tbody = document.getElementById('userTableBody');
        if (!tbody) return;
        if (systemUsers.length === 0) { tbody.innerHTML = `<tr><td colspan="5" class="p-6 text-center text-gray-500">Chưa có người dùng.</td></tr>`; return; }
        const fullStoreList = getFullStoreList();
        tbody.innerHTML = systemUsers.map((u) => {
            const storeInfo = fullStoreList.find(s => s.code === u.storeCode);
            const storeName = storeInfo ? storeInfo.name : (u.storeCode || 'Tất cả');
            let roleBadge = u.position === 'Quản trị viên' ? `<span class="bg-red-100 text-red-700 px-2 py-0.5 rounded font-bold text-[10px]">QUẢN TRỊ</span>` : `<span class="bg-gray-100 text-gray-700 px-2 py-0.5 rounded font-bold text-[10px]">NHÂN VIÊN</span>`;
            return `
            <tr class="border-b hover:bg-gray-50">
                <td class="p-2 font-semibold text-indigo-700">${u.email}</td>
                <td class="p-2">${u.full_name || '--'}</td>
                <td class="p-2">${roleBadge}</td>
                <td class="p-2">${storeName}</td>
                <td class="p-2 text-center">
                    <button onclick="editUser('${u.email}')" class="p-1 border border-yellow-400 text-yellow-600 rounded hover:bg-yellow-50" title="Sửa">✏️</button>
                    <button onclick="deleteUserAccess('${u.email}')" class="p-1 border border-red-400 text-red-600 rounded hover:bg-red-50 ml-1" title="Xóa">🗑️</button>
                </td>
            </tr>`;
        }).join('');
    }

    window.editUser = function(email) {
        const user = systemUsers.find(u => u.email === email);
        if(!user) return;
        isEditingUser = true;
        document.getElementById('umFormTitle').innerText = "Chỉnh Sửa Phân Quyền";
        document.getElementById('umEmail').value = user.email;
        document.getElementById('umEmail').readOnly = true;
        document.getElementById('umEmail').classList.add('bg-gray-100', 'text-gray-500'); 
        document.getElementById('umPasswordContainer').classList.add('hidden'); 
        document.getElementById('umPassword').required = false;
        document.getElementById('umFullName').value = user.full_name || '';
        document.getElementById('umRole').value = user.position || 'Nhân viên';
        document.getElementById('umStoreCode').value = user.storeCode || '';
        document.getElementById('btnSubmitUser').innerText = "Lưu Thay Đổi";
    };

    window.saveUserManager = async function() {
        const email = document.getElementById('umEmail').value.trim();
        const password = document.getElementById('umPassword').value.trim();
        const fullName = document.getElementById('umFullName').value.trim();
        const role = document.getElementById('umRole').value;
        const storeCode = document.getElementById('umStoreCode').value;
        const btn = document.getElementById('btnSubmitUser');
        btn.disabled = true; btn.innerText = "Đang lưu...";
        try {
            if (!isEditingUser) {
                try {
                    await createUserWithEmailAndPassword(adminAuth, email, password);
                    await signOut(adminAuth); 
                } catch(authErr) {
                    if(authErr.code !== 'auth/email-already-in-use') throw new Error("Lỗi Auth: " + authErr.message);
                }
            }
            let empCode = systemUsers.find(u => u.email === email)?.employee_code || "EMP-" + Math.floor(Math.random() * 10000);
            await setDoc(doc(db, "users", email), { email, full_name: fullName, position: role, storeCode, employee_code: empCode, updatedAt: new Date().toISOString() }, { merge: true });
            showToast("Lưu tài khoản thành công!", "success");
            resetUserForm();
            await loadUsers(); 
        } catch (e) { showToast("Lỗi: " + e.message); } 
        finally { btn.disabled = false; btn.innerText = isEditingUser ? "Lưu Thay Đổi" : "Tạo Tài Khoản Mới"; }
    };

    window.deleteUserAccess = async function(email) {
        if(!window.confirm(`Xóa quyền của [${email}]?`)) return;
        try {
            await deleteDoc(doc(db, "users", email));
            showToast("Đã xóa!", "success");
            await loadUsers();
        } catch (e) { showToast("Lỗi quyền hạn."); }
    };

    // --- QUẢN LÝ KNOWLEDGE BASE (THÊM / SỬA / XÓA) ---
    window.openKnowledgeManager = function() {
        document.body.classList.add('overflow-hidden');
        document.getElementById('knowledgeManagerModal').classList.remove('hidden');
        resetKBForm();
        renderKBManagerTable('');
    };

    window.closeKnowledgeManager = function() {
        document.body.classList.remove('overflow-hidden');
        document.getElementById('knowledgeManagerModal').classList.add('hidden');
        loadKnowledgeBase();
    };

    window.resetKBForm = function() {
        document.getElementById('kbForm').reset();
        document.getElementById('kbDocId').value = "";
        document.getElementById('kbFormTitle').innerText = "➕ Thêm Giải Pháp Mới";
        document.getElementById('btnSubmitKB').innerText = "Lưu Giải Pháp";
        onKbCategoryChange();
    };

    window.onKbCategoryChange = function() {
        const cat = document.getElementById('kbCategory').value;
        const cat2Select = document.getElementById('kbCategory2');
        if(!cat2Select) return;
        
        cat2Select.innerHTML = `<option value="">-- Chọn Chi tiết lỗi --</option>`;
        if (category2Options[cat]) {
            category2Options[cat].forEach(item => {
                const opt = document.createElement('option');
                opt.value = item;
                opt.innerText = item;
                cat2Select.appendChild(opt);
            });
        }
    };

    window.renderKBManagerTable = function(keyword = '') {
        const tbody = document.getElementById('kbTableBody');
        if (!tbody) return;

        const kw = keyword.toLowerCase().trim();
        const filtered = systemKnowledgeBase.filter(item => 
            item.category.toLowerCase().includes(kw) || 
            (item.category2 && item.category2.toLowerCase().includes(kw)) ||
            item.title.toLowerCase().includes(kw) || 
            (Array.isArray(item.steps) ? item.steps.join(' ').toLowerCase() : '').includes(kw)
        );

        if (filtered.length === 0) {
            tbody.innerHTML = `<tr><td colspan="3" class="p-6 text-center text-gray-500">Chưa có dữ liệu giải pháp nào.</td></tr>`;
            return;
        }

        tbody.innerHTML = filtered.map(item => {
            const stepsHtml = Array.isArray(item.steps) ? item.steps.map(s => `<div>• ${s}</div>`).join('') : item.steps;
            return `
            <tr class="border-b hover:bg-gray-50 align-top">
                <td class="p-2.5">
                    <div class="flex items-center gap-1.5 mb-1">
                        <span class="bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded text-[10px] uppercase">${item.category}</span>
                        ${item.category2 ? `<span class="bg-blue-100 text-blue-800 font-bold px-2 py-0.5 rounded text-[10px]">${item.category2}</span>` : ''}
                    </div>
                    <div class="font-bold text-slate-800">${item.title}</div>
                </td>
                <td class="p-2.5 text-gray-600 max-w-[300px] whitespace-normal leading-relaxed">${stepsHtml}</td>
                <td class="p-2.5 text-center whitespace-nowrap">
                    <button onclick="editKBItem('${item.id}')" class="p-1 border border-yellow-400 text-yellow-600 rounded hover:bg-yellow-50" title="Sửa">✏️</button>
                    <button onclick="deleteKBItem('${item.id}')" class="p-1 border border-red-400 text-red-600 rounded hover:bg-red-50 ml-1" title="Xóa">🗑️</button>
                </td>
            </tr>`;
        }).join('');
    };

    window.saveKnowledgeBaseItem = async function() {
        const docId = document.getElementById('kbDocId').value.trim();
        const category = document.getElementById('kbCategory').value;
        const category2 = document.getElementById('kbCategory2').value;
        const title = document.getElementById('kbTitle').value.trim();
        const rawSteps = document.getElementById('kbSteps').value.trim();
        const steps = rawSteps.split('\n').filter(s => s.trim() !== '');

        if (!category2) {
            showToast("Vui lòng chọn Chi tiết lỗi (Category 2)!", "error");
            return;
        }

        const btn = document.getElementById('btnSubmitKB');
        btn.disabled = true; btn.innerText = "Đang lưu...";

        try {
            const targetId = docId || ("KB_" + Date.now());
            await setDoc(doc(db, "knowledge_base", targetId), {
                category,
                category2,
                title,
                steps,
                updatedAt: new Date().toISOString()
            }, { merge: true });

            showToast("✅ Lưu giải pháp thành công!", "success");
            resetKBForm();
            await loadKnowledgeBase();
            renderKBManagerTable('');
        } catch (e) {
            showToast("Lỗi: " + e.message, "error");
        } finally {
            btn.disabled = false;
            btn.innerText = "Lưu Giải Pháp";
        }
    };

    window.editKBItem = function(id) {
        const item = systemKnowledgeBase.find(i => i.id === id);
        if (!item) return;

        document.getElementById('kbDocId').value = item.id;
        document.getElementById('kbCategory').value = item.category;
        onKbCategoryChange();
        document.getElementById('kbCategory2').value = item.category2 || '';
        document.getElementById('kbTitle').value = item.title;
        document.getElementById('kbSteps').value = Array.isArray(item.steps) ? item.steps.join('\n') : item.steps;
        
        document.getElementById('kbFormTitle').innerText = "✏️ Cập Nhật Giải Pháp";
        document.getElementById('btnSubmitKB').innerText = "Lưu Thay Đổi";
        document.getElementById('kbForm').scrollIntoView({ behavior: 'smooth' });
    };

    window.deleteKBItem = async function(id) {
        if (!window.confirm("Bạn có chắc chắn muốn xóa giải pháp này khỏi hệ thống?")) return;
        try {
            await deleteDoc(doc(db, "knowledge_base", id));
            showToast("Đã xóa giải pháp!", "success");
            await loadKnowledgeBase();
            renderKBManagerTable('');
        } catch (e) {
            showToast("Lỗi khi xóa dữ liệu!", "error");
        }
    };

    window.loadIncidents = async function() {
      const btnRefresh = document.getElementById('btnRefreshDashboard');
      if(btnRefresh) btnRefresh.innerHTML = `<span>⏳</span> Đang tải...`;
      try {
        const qSnap = await getDocs(collection(db, "incidents"));
        const loaded = [];
        qSnap.forEach((doc) => loaded.push({ id: doc.id, ...doc.data() }));
        systemIncidents = loaded.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
      } catch (e) { systemIncidents = []; }
      populateIncidentBranchFilter();
      renderIncidentsTable();
      updateInAppNotifications();
      if(btnRefresh) btnRefresh.innerHTML = `<span>↻</span> Làm mới`;
    };

    function populateIncidentBranchFilter() {
      const branchSelect = document.getElementById('filterIncBranch');
      if (!branchSelect) return;
      const fullStoreList = getFullStoreList();
      branchSelect.innerHTML = `<option value="Tất cả">Chi nhánh: Tất cả</option>` + 
        fullStoreList.map(s => `<option value="${s.code}">${s.code} - ${s.name}</option>`).join('');
    }

    window.applyIncidentFilters = function() {
      currentIncidentPage = 1;
      incidentFilterState.search = document.getElementById('filterIncSearch').value.toLowerCase().trim();
      incidentFilterState.branch = document.getElementById('filterIncBranch').value;
      incidentFilterState.status = document.getElementById('filterIncStatus').value;
      incidentFilterState.dateFrom = document.getElementById('filterIncDateFrom').value;
      incidentFilterState.dateTo = document.getElementById('filterIncDateTo').value;
      renderIncidentsTable();
    };

    window.resetIncidentFilters = function() {
      currentIncidentPage = 1;
      document.getElementById('filterIncSearch').value = "";
      document.getElementById('filterIncBranch').value = "Tất cả";
      document.getElementById('filterIncStatus').value = "Tất cả";
      document.getElementById('filterIncDateFrom').value = "";
      document.getElementById('filterIncDateTo').value = "";
      applyIncidentFilters();
    };

    function getFilteredIncidents() {
      return systemIncidents.filter(inc => {
        if (window.currentUserRole !== 'Quản trị viên' && window.currentUserStore) {
            if (inc.storeCode !== window.currentUserStore) return false;
        }
        if (incidentFilterState.search) {
          const searchStr = `${inc.id} ${inc.storeCode} ${inc.contactPerson} ${inc.phoneNumber} ${inc.issue} ${inc.status} ${inc.result || ''}`.toLowerCase();
          if (!searchStr.includes(incidentFilterState.search)) return false;
        }
        if (incidentFilterState.branch !== "Tất cả" && inc.storeCode !== incidentFilterState.branch) return false;
        if (incidentFilterState.status !== "Tất cả" && inc.status !== incidentFilterState.status) return false;
        if (incidentFilterState.dateFrom || incidentFilterState.dateTo) {
          if (!inc.createdAt) return false;
          const incDateStr = inc.createdAt.split('T')[0];
          if (incidentFilterState.dateFrom && incDateStr < incidentFilterState.dateFrom) return false;
          if (incidentFilterState.dateTo && incDateStr > incidentFilterState.dateTo) return false;
        }
        return true;
      });
    }

    function updateDashboardStats(filteredList) {
        let countTotal = filteredList.length;
        let countNew = 0, countProcessing = 0, countResolved = 0;
        let totalResolutionMinutes = 0;
        let resolvedWithTimeCount = 0;

        filteredList.forEach(inc => {
            if (inc.status === 'Mới ghi nhận') countNew++;
            else if (inc.status === 'Đang xử lý') countProcessing++;
            else if (inc.status === 'Đã xử lý') {
                countResolved++;
                if (inc.createdAt && inc.resolvedAt) {
                    const diffMins = Math.max(1, Math.round((new Date(inc.resolvedAt) - new Date(inc.createdAt)) / (1000 * 60)));
                    totalResolutionMinutes += diffMins;
                    resolvedWithTimeCount++;
                }
            }
        });

        const mttrText = resolvedWithTimeCount > 0 ? `${Math.round(totalResolutionMinutes / resolvedWithTimeCount)} phút` : '--';

        const totalEl = document.getElementById('dashTotal');
        if(totalEl) {
          totalEl.innerText = countTotal;
          document.getElementById('dashNew').innerText = countNew;
          document.getElementById('dashProcessing').innerText = countProcessing;
          document.getElementById('dashResolved').innerText = countResolved;
          const mttrEl = document.getElementById('dashMTTR');
          if(mttrEl) mttrEl.innerText = mttrText;
        }
    }

    function renderCharts(filteredList) {
        const pieCanvas = document.getElementById('categoryChart');
        const barCanvas = document.getElementById('branchChart');
        if (!pieCanvas || !barCanvas || typeof Chart === 'undefined') return;

        const catCounts = {};
        filteredList.forEach(inc => {
            if (inc.categories && Array.isArray(inc.categories)) {
                inc.categories.forEach(cat => {
                    catCounts[cat] = (catCounts[cat] || 0) + 1;
                });
            }
        });

        const catLabels = Object.keys(catCounts);
        const catData = Object.values(catCounts);

        if (categoryChartInstance) categoryChartInstance.destroy();
        categoryChartInstance = new Chart(pieCanvas.getContext('2d'), {
            type: 'doughnut',
            data: {
                labels: catLabels.length > 0 ? catLabels : ['Chưa có dữ liệu'],
                datasets: [{
                    data: catData.length > 0 ? catData : [1],
                    backgroundColor: ['#3b82f6', '#f59e0b', '#10b981', '#ef4444', '#8b5cf6', '#ec4899', '#64748b', '#06b6d4']
                }]
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { position: 'right', labels: { boxWidth: 12, font: { size: 11 } } } } 
            }
        });

        const branchCounts = {};
        const fullStoreList = getFullStoreList();
        filteredList.forEach(inc => {
            const storeInfo = fullStoreList.find(s => s.code === inc.storeCode);
            const storeName = storeInfo ? `${storeInfo.code} - ${storeInfo.name}` : (inc.storeCode || 'Khác');
            branchCounts[storeName] = (branchCounts[storeName] || 0) + 1;
        });

        const branchLabels = Object.keys(branchCounts);
        const branchData = Object.values(branchCounts);

        if (branchChartInstance) branchChartInstance.destroy();
        branchChartInstance = new Chart(barCanvas.getContext('2d'), {
            type: 'bar',
            data: {
                labels: branchLabels.length > 0 ? branchLabels : ['Chưa có dữ liệu'],
                datasets: [{
                    label: 'Số lượng sự cố',
                    data: branchData.length > 0 ? branchData : [0],
                    backgroundColor: '#2c3e50',
                    borderRadius: 6
                }]
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
            }
        });
    }

    function renderPaginationControls(totalPages) {
      const paginationContainer = document.getElementById('paginationControls');
      if (!paginationContainer) return;

      if (totalPages <= 1) {
          paginationContainer.classList.add('hidden');
          paginationContainer.innerHTML = '';
          return;
      }

      paginationContainer.classList.remove('hidden');
      paginationContainer.innerHTML = `
        <div class="text-gray-600 font-semibold">
            Trang <span class="font-bold text-blue-700">${currentIncidentPage}</span> / ${totalPages} (20 dòng/trang)
        </div>
        <div class="flex items-center gap-1.5">
            <button onclick="changeIncidentPage(1)" ${currentIncidentPage === 1 ? 'disabled class="px-3 py-1.5 border rounded bg-gray-100 text-gray-400 cursor-not-allowed"' : 'class="px-3 py-1.5 border rounded bg-white hover:bg-gray-50 font-bold text-gray-700"'}>Đầu</button>
            <button onclick="changeIncidentPage(${currentIncidentPage - 1})" ${currentIncidentPage === 1 ? 'disabled class="px-3 py-1.5 border rounded bg-gray-100 text-gray-400 cursor-not-allowed"' : 'class="px-3 py-1.5 border rounded bg-white hover:bg-gray-50 font-bold text-gray-700"'}>◀ Trước</button>
            <span class="px-3.5 py-1.5 bg-slate-900 text-white font-bold rounded">${currentIncidentPage}</span>
            <button onclick="changeIncidentPage(${currentIncidentPage + 1})" ${currentIncidentPage === totalPages ? 'disabled class="px-3 py-1.5 border rounded bg-gray-100 text-gray-400 cursor-not-allowed"' : 'class="px-3 py-1.5 border rounded bg-white hover:bg-gray-50 font-bold text-gray-700"'}>Sau ▶</button>
            <button onclick="changeIncidentPage(${totalPages})" ${currentIncidentPage === totalPages ? 'disabled class="px-3 py-1.5 border rounded bg-gray-100 text-gray-400 cursor-not-allowed"' : 'class="px-3 py-1.5 border rounded bg-white hover:bg-gray-50 font-bold text-gray-700"'}>Cuối</button>
        </div>
      `;
    }

    window.changeIncidentPage = function(targetPage) {
        currentIncidentPage = targetPage;
        renderIncidentsTable();
        const tableArea = document.getElementById('incidentTableWrapper');
        if (tableArea) tableArea.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    function renderIncidentsTable() {
      const tbody = document.getElementById('incidentTableBody');
      const countText = document.getElementById('incCountText');
      if (!tbody) return;

      const filteredList = getFilteredIncidents();
      updateDashboardStats(filteredList); 
      renderCharts(filteredList);

      const totalPages = Math.ceil(filteredList.length / rowsPerPage) || 1;
      if (currentIncidentPage > totalPages) currentIncidentPage = totalPages;
      if (currentIncidentPage < 1) currentIncidentPage = 1;

      const startIndex = (currentIncidentPage - 1) * rowsPerPage;
      const paginatedList = filteredList.slice(startIndex, startIndex + rowsPerPage);

      if (filteredList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="p-6 text-center text-gray-500">Chưa có sự cố.</td></tr>`;
        if(countText) countText.innerText = `Hiển thị 0 / ${systemIncidents.length} phiếu`;
        renderPaginationControls(0);
        return;
      }

      if(countText) {
        countText.innerText = `Hiển thị dòng ${startIndex + 1} - ${Math.min(startIndex + rowsPerPage, filteredList.length)} trên tổng ${filteredList.length} phiếu`;
      }
      
      const fullStoreList = getFullStoreList();

      tbody.innerHTML = paginatedList.map((inc) => {
        const storeInfo = fullStoreList.find(s => s.code === inc.storeCode);
        const storeName = storeInfo ? storeInfo.name : inc.storeCode;
        
        let cats = inc.categories && inc.categories.length > 0 ? inc.categories.join(', ') : '--';
        if (inc.category2) cats += ` (${inc.category2})`;

        let statusColor = "bg-amber-100 text-amber-800 border-amber-300";
        if(inc.status === 'Đã xử lý') statusColor = "bg-green-100 text-green-800 border-green-300";
        else if (inc.status === 'Đang xử lý') statusColor = "bg-blue-100 text-blue-800 border-blue-300";

        let assigneeName = getAssignedUserName(inc.assignee);

        let quickActionHTML = '';
        if (inc.status === 'Mới ghi nhận') {
            quickActionHTML = `<button onclick="quickAcceptTask('${inc.id}')" class="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded text-[11px] font-bold shadow-sm transition">🚀 Nhận việc</button>`;
        } else if (inc.status === 'Đang xử lý') {
            quickActionHTML = `<button onclick="openIncidentModal('${inc.id}')" class="px-2.5 py-1 bg-green-600 hover:bg-green-700 text-white rounded text-[11px] font-bold shadow-sm transition">✅ Báo cáo Done</button>`;
        }

        let actionHTML = `
          <button onclick="openIncidentModal('${inc.id}')" class="p-1.5 border border-gray-300 rounded hover:bg-blue-50 text-blue-600 font-bold" title="Cập nhật & Xem lịch sử">✏️ Cập nhật</button>
          <button onclick="openPrintReportModal('${inc.id}')" class="p-1.5 border border-gray-300 rounded hover:bg-slate-100 text-slate-700 font-bold ml-1" title="In Biên Bản A4">🖨️</button>
          ${window.currentUserRole === 'Quản trị viên' ? `<button onclick="deleteIncident('${inc.id}')" class="p-1.5 border border-red-200 text-red-600 hover:bg-red-50 rounded ml-1" title="Xóa">🗑️</button>` : ''}
        `;
        
        let resultHTML = '';
        if (inc.status === 'Đã xử lý') {
            resultHTML = `
                <div class="mt-1.5 pt-1.5 border-t border-dashed border-gray-200 text-[11px]">
                    ${inc.tempSolution ? `<div><span class="font-bold text-blue-700">Xử lý tạm:</span> ${inc.tempSolution}</div>` : ''}
                    ${inc.result ? `<div><span class="font-bold text-green-700">Xử lý lâu dài:</span> ${inc.result}</div>` : ''}
                </div>`;
        }

        let imgBadge = (inc.imageUrl && inc.imageUrl.trim() !== "") ? `
          <div class="mt-2">
            <img src="${inc.imageUrl}" onerror="this.style.display='none'" onclick="openImageViewer('${inc.imageUrl}')" alt="Ảnh" class="h-14 w-auto max-w-[120px] object-cover rounded border border-gray-300 shadow-sm cursor-zoom-in">
          </div>` : '';

        return `
        <tr class="border-b border-gray-200 hover:bg-gray-50 text-[12px] text-gray-800">
          <td class="p-3 text-center whitespace-nowrap">
             <div class="font-bold text-red-700 cursor-pointer hover:underline" onclick="openIncidentModal('${inc.id}')">${inc.id}</div>
             <div class="text-[10px] text-gray-500 font-medium mt-0.5">${formatDateTimeView(inc.createdAt)}</div>
          </td>
          <td class="p-3 font-bold whitespace-nowrap">${storeName} <div class="text-[10px] text-gray-500 font-normal">Mã: ${inc.storeCode}</div></td>
          <td class="p-3 whitespace-nowrap">
             <div class="font-semibold text-blue-900">${inc.contactPerson || '--'}</div>
             <div class="text-[10px] font-mono text-gray-600">${inc.phoneNumber || '--'}</div>
          </td>
          <td class="p-3 font-semibold text-amber-700 max-w-[160px] whitespace-normal">${cats}</td>
          <td class="p-3 min-w-[200px] whitespace-normal">
             <div class="line-clamp-2">${inc.issue}</div>
             <div class="text-[10px] text-indigo-700 font-bold mt-1">👤 PIC: ${assigneeName}</div>
             ${imgBadge}
             ${resultHTML}
          </td>
          <td class="p-3 text-center whitespace-nowrap">
            <span class="${statusColor} border px-2 py-1 rounded-full font-bold text-[10px] uppercase block mb-1.5">${inc.status || 'Mới ghi nhận'}</span>
            ${quickActionHTML}
          </td>
          <td class="p-3 text-center whitespace-nowrap"><div class="flex items-center justify-center space-x-1">${actionHTML}</div></td>
        </tr>`;
      }).join('');

      renderPaginationControls(totalPages);
    }

    window.quickAcceptTask = async function(incId) {
        try {
            const inc = systemIncidents.find(i => i.id === incId);
            if(!inc) return;
            inc.status = 'Đang xử lý';
            inc.assignee = window.currentEmployeeCode || window.currentUserEmail;
            inc.updatedAt = new Date().toISOString();

            if (!Array.isArray(inc.auditLog)) inc.auditLog = [];
            inc.auditLog.push({
                time: new Date().toISOString(),
                user: window.currentUserName || window.currentUserEmail,
                action: `Nhận việc xử lý sự cố (Đang xử lý)`
            });

            await setDoc(doc(db, "incidents", incId), inc, { merge: true });
            showToast("🚀 Đã nhận việc!", "success");
            sendTelegramAlert(inc, false).catch(e => console.log(e));
            await loadIncidents();
        } catch(e) { showToast("Lỗi: " + e.message); }
    };

    window.onCategoryChange = function() {
        const checkedCats = [];
        document.querySelectorAll('input[name="incCategory"]:checked').forEach(cb => checkedCats.push(cb.value));
        const cat2Select = document.getElementById('incCategory2');
        if(!cat2Select) return;
        const currentSelectedVal = cat2Select.value;
        cat2Select.innerHTML = `<option value="">-- Chọn chi tiết lỗi (Category 2) --</option>`;
        
        let subList = [];
        checkedCats.forEach(cat => {
            if (category2Options[cat]) subList = subList.concat(category2Options[cat]);
        });
        
        subList = [...new Set(subList)];
        subList.forEach(item => {
            const opt = document.createElement('option');
            opt.value = item;
            opt.innerText = item;
            cat2Select.appendChild(opt);
        });
        cat2Select.value = currentSelectedVal;

        const suggestionBox = document.getElementById('inlineKbSuggestion');
        if (suggestionBox) suggestionBox.classList.add('hidden');
    };

    window.onCategory2Change = function() {
        const selectedCategory2 = document.getElementById('incCategory2').value;
        const suggestionBox = document.getElementById('inlineKbSuggestion');
        const suggestionContent = document.getElementById('inlineKbContent');
        if (!suggestionBox || !suggestionContent) return;

        if (!selectedCategory2) {
            suggestionBox.classList.add('hidden');
            return;
        }

        const matchedKBs = systemKnowledgeBase.filter(kb => 
            kb.category2 === selectedCategory2 || 
            (kb.title && kb.title.toLowerCase().includes(selectedCategory2.toLowerCase()))
        );

        if (matchedKBs.length > 0) {
            suggestionContent.innerHTML = matchedKBs.map(kb => {
                const stepsHtml = Array.isArray(kb.steps) ? kb.steps.map(s => `<li>${s}</li>`).join('') : `<li>${kb.steps}</li>`;
                return `
                <div class="mb-2 last:mb-0">
                    <div class="font-bold text-blue-900">💡 ${kb.title}:</div>
                    <ul class="list-disc pl-4 text-gray-700 mt-0.5 space-y-0.5">
                        ${stepsHtml}
                    </ul>
                </div>`;
            }).join('');
            suggestionBox.classList.remove('hidden');
        } else {
            suggestionBox.classList.add('hidden');
        }
    };

    async function compressImageFile(file, maxWidth = 600, quality = 0.7) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = (event) => {
                const img = new Image();
                img.src = event.target.result;
                img.onload = () => {
                    const canvas = document.createElement('canvas');
                    let width = img.width;
                    let height = img.height;
                    if (width > maxWidth) {
                        height = Math.round((height * maxWidth) / width);
                        width = maxWidth;
                    }
                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, width, height);
                    resolve(canvas.toDataURL('image/jpeg', quality));
                };
                img.onerror = (err) => reject(err);
            };
            reader.onerror = (err) => reject(err);
        });
    }

    window.exportIncidentsToExcel = function() {
        showToast("Đang tạo file Excel...", "success");
        setTimeout(() => {
          const fullStoreList = getFullStoreList();
          const dataToExport = getFilteredIncidents().map((inc, index) => {
              const storeInfo = fullStoreList.find(s => s.code === inc.storeCode) || {name: inc.storeCode};
              let assigneeName = getAssignedUserName(inc.assignee);
              let timeStr = formatDateTimeView(inc.createdAt);
              let thoiDiemNoiDung = `${timeStr}\n${inc.issue || ''}`;
              let dateOnly = inc.createdAt ? inc.createdAt.split('T')[0].replace(/-/g, '/') : '';

              return {
                  "STT": index + 1,
                  "Thời gian": dateOnly,
                  "ID Store": inc.storeCode,
                  "Địa điểm": storeInfo.name,
                  "Staff": inc.contactPerson || '',
                  "Phân loại": inc.categories ? inc.categories.join(', ') : '',
                  "Category 2": inc.category2 || "", 
                  "Ưu tiên": "Normal",
                  "ẢNH": inc.imageUrl || '',
                  "Thời điểm + Nội dung": thoiDiemNoiDung,
                  "Xử lý tạm thời": inc.tempSolution || "",
                  "Xử lý lâu dài": inc.result || "",
                  "PIC": assigneeName,
                  "Status": inc.status === 'Đã xử lý' ? "Done" : (inc.status || 'Mới ghi nhận')
              };
          });

          if(dataToExport.length === 0) { showToast("Không có dữ liệu!"); return; }
          
          const ws = XLSX.utils.json_to_sheet(dataToExport);
          const range = XLSX.utils.decode_range(ws['!ref']);

          for (let R = 1; R <= range.e.r; ++R) {
              const cellRef = XLSX.utils.encode_cell({ c: 8, r: R }); 
              if (!ws[cellRef]) continue;
              const val = ws[cellRef].v;
              if (val && (val.startsWith('data:image') || val.startsWith('http'))) {
                  ws[cellRef] = { 
                      t: 's', 
                      v: "🖼️ Xem Ảnh", 
                      l: { Target: val, Tooltip: "Click để xem ảnh hiện trường" } 
                  };
              } else {
                  ws[cellRef] = { t: 's', v: "" };
              }
          }

          ws['!cols'] = [{wch: 5}, {wch: 12}, {wch: 12}, {wch: 20}, {wch: 15}, {wch: 15}, {wch: 18}, {wch: 10}, {wch: 15}, {wch: 45}, {wch: 25}, {wch: 25}, {wch: 18}, {wch: 15}];
          ws['!rows'] = [{ hpt: 30 }]; 
          for (let R = 1; R <= range.e.r; ++R) ws['!rows'][R] = { hpt: 60 }; 

          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Nhat_Ky_Su_Co");
          XLSX.writeFile(wb, `Bao_Cao_Su_Co_${getTodayString()}.xlsx`);
        }, 500);
    };

    window.previewImage = async function(event) {
        const file = event.target.files[0];
        const container = document.getElementById('imagePreviewContainer');
        const img = document.getElementById('imagePreview');
        if (file) {
            try {
                const compressedBase64 = await compressImageFile(file, 600, 0.7);
                img.src = compressedBase64;
                container.classList.remove('hidden');
            } catch (e) {
                showToast("Lỗi đọc file ảnh");
            }
        } else {
            img.src = "";
            container.classList.add('hidden');
        }
    };

    window.onIncidentStatusChange = function() {
        const status = document.getElementById('incStatus').value;
        const resultContainer = document.getElementById('incResultContainer');
        const tempContainer = document.getElementById('incTempContainer');
        const incResult = document.getElementById('incResult');
        const incTemp = document.getElementById('incTempSolution');

        if (status === 'Đã xử lý') {
            if(resultContainer) resultContainer.classList.remove('hidden');
            if(tempContainer) tempContainer.classList.remove('hidden');
            if(incResult) incResult.required = true;
            if(incTemp) incTemp.required = true;
        } else {
            if(resultContainer) resultContainer.classList.add('hidden');
            if(tempContainer) tempContainer.classList.add('hidden');
            if(incResult) incResult.required = false;
            if(incTemp) incTemp.required = false;
        }
    };

    function renderAuditTimeline(inc) {
        const auditSection = document.getElementById('auditTimelineSection');
        const auditList = document.getElementById('auditTimelineList');
        if (!auditSection || !auditList) return;

        const logs = Array.isArray(inc.auditLog) ? [...inc.auditLog] : [];
        if (logs.length === 0 && inc.createdAt) {
            logs.push({
                time: inc.createdAt,
                user: inc.reportedBy || inc.contactPerson || 'Hệ thống',
                action: 'Khởi tạo phiếu sự cố mới'
            });
        }

        auditSection.classList.remove('hidden');
        auditList.innerHTML = logs.reverse().map(log => `
            <div class="relative pl-4 pb-3 border-l-2 border-indigo-200 last:border-l-0 last:pb-0">
                <div class="absolute -left-[5px] top-1 w-2 h-2 rounded-full bg-indigo-600"></div>
                <div class="text-[11px] text-gray-500">${formatDateTimeView(log.time)} • <span class="font-bold text-indigo-800">${log.user}</span></div>
                <div class="text-xs text-gray-800 font-medium mt-0.5">${log.action}</div>
            </div>
        `).join('');
    }

    window.openIncidentModal = async function(incidentId) {
      document.body.classList.add('overflow-hidden');

      if (typeof incidentId !== 'string') incidentId = null;
      editingIncidentId = incidentId;
      document.getElementById('incidentModal').classList.remove('hidden');
      
      const formErrorDiv = document.getElementById('formErrorMsg');
      if (formErrorDiv) formErrorDiv.classList.add('hidden');

      document.getElementById('incidentForm').reset();
      document.getElementById('imagePreviewContainer').classList.add('hidden');
      document.getElementById('imagePreview').src = "";
      document.getElementById('inlineKbSuggestion').classList.add('hidden');
      document.getElementById('auditTimelineSection').classList.add('hidden');
      document.getElementById('btnModalPrint').classList.add('hidden');
      
      document.getElementById('incMonth').value = getMonthString();
      document.getElementById('incDate').value = getTodayString();
      document.getElementById('incTime').value = getCurrentDateTime();
      document.getElementById('incResultContainer').classList.add('hidden');
      document.getElementById('incTempContainer').classList.add('hidden');
      
      const storeSelect = document.getElementById('incStoreCode');
      const fullStoreList = getFullStoreList();
      storeSelect.innerHTML = `<option value="">-- Chọn Chi nhánh --</option>` + 
          fullStoreList.map(s => `<option value="${s.code}">${s.code} - ${s.name}</option>`).join('');

      const leaderSelect = document.getElementById('incLeader');
      if (systemUsers.length === 0) await loadUsers();
      leaderSelect.innerHTML = `<option value="">-- Chọn Người thực hiện (PIC) --</option>` + 
          systemUsers.map(l => `<option value="${l.employee_code || l.email}">${l.full_name || l.email} (${l.position || 'Nhân viên'})</option>`).join('');

      if (editingIncidentId) {
        if (!systemIncidents || systemIncidents.length === 0) {
            await window.loadIncidents();
        }
        
        const inc = systemIncidents.find(i => i.id === editingIncidentId);
        if (inc) {
          document.getElementById('modalIncidentTitle').innerHTML = `📝 Cập Nhật Phiếu / Báo Cáo Done: ${inc.id}`;
          document.getElementById('incTicket').value = inc.id;
          
          const btnPrint = document.getElementById('btnModalPrint');
          btnPrint.classList.remove('hidden');
          btnPrint.onclick = () => { window.closeIncidentModal(); window.openPrintReportModal(inc.id); };

          const sCode = inc.storeCode || inc.store_code || inc.branch || '';
          if (sCode) {
              storeSelect.value = sCode;
              window.onIncidentStoreChange(); 
          }

          document.getElementById('incContact').value = inc.contactPerson || inc.contact_person || inc.staff || '';
          document.getElementById('incPhone').value = inc.phoneNumber || inc.phone_number || inc.phone || '';
          document.getElementById('incIssue').value = inc.issue || inc.description || inc.noi_dung || '';
          document.getElementById('incLeader').value = inc.assignee || inc.pic || '';
          document.getElementById('incStatus').value = inc.status || inc.trang_thai || 'Mới ghi nhận';
          
          window.onIncidentStatusChange(); 

          document.getElementById('incResult').value = inc.result || inc.xu_ly_lau_dai || '';
          document.getElementById('incTempSolution').value = inc.tempSolution || inc.xu_ly_tam || '';

          const imgUrl = inc.imageUrl || inc.image_url || inc.anh || '';
          if (imgUrl) {
            document.getElementById('imagePreview').src = imgUrl;
            document.getElementById('imagePreviewContainer').classList.remove('hidden');
          }

          const cats = inc.categories || inc.category || inc.phan_loai || [];
          if (Array.isArray(cats)) {
              const checkboxes = document.querySelectorAll('input[name="incCategory"]');
              checkboxes.forEach(cb => { 
                  if(cats.includes(cb.value)) cb.checked = true; 
              });
          }
          
          window.onCategoryChange();
          
          const cat2 = inc.category2 || inc.category_2 || '';
          if(cat2) {
              document.getElementById('incCategory2').value = cat2;
              window.onCategory2Change();
          }

          renderAuditTimeline(inc);
        }
      } else {
        document.getElementById('modalIncidentTitle').innerHTML = `📝 Lập Phiếu Yêu Cầu / New Request`;
        document.getElementById('incTicket').value = generateUniqueIncidentCode();
        window.onCategoryChange();
        if(window.currentUserRole !== 'Quản trị viên' && window.currentUserStore) {
            storeSelect.value = window.currentUserStore;
            window.onIncidentStoreChange();
        }
      }
    };

    window.closeIncidentModal = function() { 
      document.body.classList.remove('overflow-hidden');
      document.getElementById('incidentModal').classList.add('hidden'); 
    };

    window.onIncidentStoreChange = function() {
        const storeCode = document.getElementById('incStoreCode').value;
        const districtEl = document.getElementById('incDistrict');
        const addressEl = document.getElementById('incAddress');

        if(!storeCode) { 
            if(districtEl) districtEl.value = ''; 
            if(addressEl) addressEl.value = ''; 
            return; 
        }
        
        const fullStoreList = getFullStoreList();
        const store = fullStoreList.find(s => s.code === storeCode);
        if(store) {
            if(addressEl) addressEl.value = store.address || '';
            const addrParts = store.address ? store.address.split(',') : [];
            if(districtEl) districtEl.value = addrParts.length > 1 ? addrParts[addrParts.length-1].trim() : '';
        }
    };

    window.saveIncidentToFirebase = async function() {
      const errorBox = document.getElementById('formErrorMsg');
      errorBox.classList.add('hidden'); errorBox.innerText = "";

      const incId = document.getElementById('incTicket').value.trim();
      const storeCode = document.getElementById('incStoreCode').value;
      const contactPerson = document.getElementById('incContact').value.trim();
      const phoneNumber = document.getElementById('incPhone').value.trim();
      const issue = document.getElementById('incIssue').value.trim();
      const category2 = document.getElementById('incCategory2').value;
      const assignee = document.getElementById('incLeader').value;
      const status = document.getElementById('incStatus').value;
      const result = document.getElementById('incResult').value.trim();
      const tempSolution = document.getElementById('incTempSolution').value.trim();

      if (!incId) {
          errorBox.innerText = "Lỗi: Không tìm thấy Mã Phiếu (Ticket ID). Vui lòng tắt bảng và mở lại!";
          errorBox.classList.remove('hidden'); return;
      }
      if (!storeCode || !contactPerson || !phoneNumber || !issue || !category2) {
          errorBox.innerText = "Vui lòng điền đủ thông tin bắt buộc và Chi tiết lỗi (*)!";
          errorBox.classList.remove('hidden'); return;
      }
      if (status === 'Đã xử lý' && (!result || !tempSolution)) {
          errorBox.innerText = "Vui lòng nhập đầy đủ Xử lý tạm thời và Xử lý lâu dài khi hoàn tất Done!";
          errorBox.classList.remove('hidden'); return;
      }

      const categories = [];
      document.querySelectorAll('input[name="incCategory"]:checked').forEach(cb => categories.push(cb.value));

      const submitBtn = document.getElementById('btnSubmitForm');
      if (submitBtn) { submitBtn.disabled = true; submitBtn.innerText = "Đang lưu..."; }

      try {
        let existingInc = systemIncidents.find(i => i.id === incId);
        let imageUrl = existingInc?.imageUrl || "";
        let resolvedAt = existingInc?.resolvedAt || null;
        
        if (status === 'Đã xử lý' && !resolvedAt) resolvedAt = new Date().toISOString(); 
        else if (status !== 'Đã xử lý') resolvedAt = null;

        const previewImgSrc = document.getElementById('imagePreview').src;
        if (previewImgSrc && previewImgSrc.startsWith('data:image')) {
            imageUrl = previewImgSrc;
        }

        let auditLog = Array.isArray(existingInc?.auditLog) ? [...existingInc.auditLog] : [];
        const actorName = window.currentUserName || window.currentUserEmail || "Nhân viên";
        const nowIso = new Date().toISOString();

        const isNewTicket = !existingInc;
        if (isNewTicket) {
            auditLog.push({
                time: nowIso,
                user: actorName,
                action: `Khởi tạo phiếu sự cố mới (${status})`
            });
        } else {
            let changes = [];
            if (existingInc.status !== status) changes.push(`Đổi trạng thái: "${existingInc.status}" ➔ "${status}"`);
            if (existingInc.assignee !== assignee) changes.push(`Phân công PIC: ${getAssignedUserName(assignee)}`);
            if (status === 'Đã xử lý' && existingInc.result !== result) changes.push(`Cập nhật báo cáo xử lý hoàn tất (Done)`);
            if (changes.length === 0) changes.push(`Cập nhật thông tin chi tiết phiếu`);

            auditLog.push({
                time: nowIso,
                user: actorName,
                action: changes.join(' | ')
            });
        }

        const incData = {
          id: incId, 
          storeCode, 
          contactPerson, 
          phoneNumber, 
          categories, 
          category2, 
          issue, 
          assignee, 
          status, 
          imageUrl,
          tempSolution: status === 'Đã xử lý' ? tempSolution : (tempSolution || ''),
          result: status === 'Đã xử lý' ? result : (result || ''),
          resolvedAt, 
          auditLog,
          updatedAt: nowIso
        };

        if (isNewTicket) {
            incData.createdAt = nowIso;
            incData.reportedBy = actorName;
            if (status === 'Đã xử lý') incData.resolvedAt = incData.createdAt; 
        } else {
            incData.createdAt = existingInc?.createdAt || nowIso;
        }

        await setDoc(doc(db, "incidents", incId), incData, { merge: true });
        
        if (isNewTicket) {
            sendTelegramAlert(incData, true).catch(e => console.log(e));
        }

        window.showToast("✅ Lưu phiếu thành công!", "success");
        window.closeIncidentModal();
        await window.loadIncidents();
      } catch (e) { 
        errorBox.innerText = "Lỗi lưu dữ liệu: " + e.message;
        errorBox.classList.remove('hidden');
      } finally {
        if (submitBtn) { submitBtn.disabled = false; submitBtn.innerText = "Lưu Phiếu / Save Ticket"; }
      }
    };

    window.deleteIncident = async function(id) {
      try {
        if(!window.confirm("Xóa phiếu sự cố này?")) return;
        await deleteDoc(doc(db, "incidents", id));
        window.showToast("Đã xóa!", "success");
        await window.loadIncidents();
      } catch(e) { window.showToast("Lỗi xóa."); }
    };

    window.openImageViewer = function(url) {
      document.body.classList.add('overflow-hidden');
      const modal = document.getElementById('imageViewerModal');
      const img = document.getElementById('fullSizeImage');
      img.src = url; modal.classList.remove('hidden');
    };
    window.closeImageViewer = function() { 
      document.body.classList.remove('overflow-hidden');
      document.getElementById('imageViewerModal').classList.add('hidden'); 
    };

    // --- BÁO CÁO TỔNG HỢP THÁNG ---
    window.openMonthlyReportModal = function() {
        document.body.classList.add('overflow-hidden');
        document.getElementById('monthlyReportModal').classList.remove('hidden');
        
        const monthSet = new Set();
        const currentMStr = getMonthString();
        monthSet.add(currentMStr);

        systemIncidents.forEach(inc => {
            if (inc.createdAt) {
                const d = new Date(inc.createdAt);
                const mStr = `${d.getFullYear()}${(d.getMonth() + 1).toString().padStart(2, '0')}`;
                monthSet.add(mStr);
            }
        });

        const sortedMonths = Array.from(monthSet).sort().reverse();
        const selectEl = document.getElementById('reportMonthSelect');
        selectEl.innerHTML = sortedMonths.map(m => {
            const year = m.substring(0, 4);
            const month = m.substring(4, 6);
            return `<option value="${m}" ${m === currentMStr ? 'selected' : ''}>Tháng ${month}/${year}</option>`;
        }).join('');

        generateMonthlyReportData();
    };

    window.closeMonthlyReportModal = function() {
        document.body.classList.remove('overflow-hidden');
        document.getElementById('monthlyReportModal').classList.add('hidden');
    };

    window.generateMonthlyReportData = function() {
        const selectedMonth = document.getElementById('reportMonthSelect').value;
        const year = selectedMonth.substring(0, 4);
        const month = selectedMonth.substring(4, 6);

        const monthIncidents = systemIncidents.filter(inc => {
            if (!inc.createdAt) return false;
            const d = new Date(inc.createdAt);
            const incY = d.getFullYear().toString();
            const incM = (d.getMonth() + 1).toString().padStart(2, '0');
            return incY === year && incM === month;
        });

        let totalInc = monthIncidents.length;
        let resolvedInc = monthIncidents.filter(i => i.status === 'Đã xử lý').length;
        let pendingInc = totalInc - resolvedInc;

        let totalMins = 0;
        let timeCount = 0;
        monthIncidents.forEach(inc => {
            if (inc.status === 'Đã xử lý' && inc.createdAt && inc.resolvedAt) {
                const diff = Math.max(1, Math.round((new Date(inc.resolvedAt) - new Date(inc.createdAt)) / (1000 * 60)));
                totalMins += diff;
                timeCount++;
            }
        });
        const avgMttr = timeCount > 0 ? `${Math.round(totalMins / timeCount)} phút` : '--';

        document.getElementById('repTotalIncidents').innerText = totalInc;
        document.getElementById('repResolvedIncidents').innerText = resolvedInc;
        document.getElementById('repPendingIncidents').innerText = pendingInc;
        document.getElementById('repAvgMttr').innerText = avgMttr;

        const fullStoreList = getFullStoreList();
        const storeStats = {};

        fullStoreList.forEach(s => {
            storeStats[s.code] = { code: s.code, name: s.name, total: 0, resolved: 0, pending: 0, totalMins: 0, timeCount: 0 };
        });

        monthIncidents.forEach(inc => {
            const code = inc.storeCode || 'KHAC';
            if (!storeStats[code]) {
                storeStats[code] = { code: code, name: code, total: 0, resolved: 0, pending: 0, totalMins: 0, timeCount: 0 };
            }
            storeStats[code].total++;
            if (inc.status === 'Đã xử lý') {
                storeStats[code].resolved++;
                if (inc.createdAt && inc.resolvedAt) {
                    const diff = Math.max(1, Math.round((new Date(inc.resolvedAt) - new Date(inc.createdAt)) / (1000 * 60)));
                    storeStats[code].totalMins += diff;
                    storeStats[code].timeCount++;
                }
            } else {
                storeStats[code].pending++;
            }
        });

        const tbody = document.getElementById('monthlyReportTableBody');
        const activeStores = Object.values(storeStats).filter(s => s.total > 0).sort((a, b) => b.total - a.total);

        if (activeStores.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" class="p-6 text-center text-gray-500">Không có dữ liệu sự cố trong tháng này.</td></tr>`;
            return;
        }

        tbody.innerHTML = activeStores.map(s => {
            const rate = s.total > 0 ? Math.round((s.resolved / s.total) * 100) : 0;
            const sMttr = s.timeCount > 0 ? `${Math.round(s.totalMins / s.timeCount)} phút` : '--';
            return `
            <tr class="border-b hover:bg-slate-50">
                <td class="p-2.5 font-bold text-slate-700">${s.code}</td>
                <td class="p-2.5 font-semibold">${s.name}</td>
                <td class="p-2.5 text-center font-bold">${s.total}</td>
                <td class="p-2.5 text-center text-amber-600 font-semibold">${s.pending}</td>
                <td class="p-2.5 text-center text-green-600 font-semibold">${s.resolved}</td>
                <td class="p-2.5 text-center">
                    <span class="px-2 py-0.5 rounded font-bold text-[10px] ${rate >= 80 ? 'bg-green-100 text-green-800' : 'bg-amber-100 text-amber-800'}">${rate}%</span>
                </td>
                <td class="p-2.5 text-center text-indigo-700 font-semibold">${sMttr}</td>
            </tr>`;
        }).join('');
    };

    window.exportMonthlyReportExcel = function() {
        const selectedMonth = document.getElementById('reportMonthSelect').value;
        const year = selectedMonth.substring(0, 4);
        const month = selectedMonth.substring(4, 6);

        showToast("Đang xuất file Excel Báo cáo Tháng...", "success");
        setTimeout(() => {
            const monthIncidents = systemIncidents.filter(inc => {
                if (!inc.createdAt) return false;
                const d = new Date(inc.createdAt);
                return d.getFullYear().toString() === year && (d.getMonth() + 1).toString().padStart(2, '0') === month;
            });

            const fullStoreList = getFullStoreList();
            const exportData = monthIncidents.map((inc, idx) => {
                const storeInfo = fullStoreList.find(s => s.code === inc.storeCode) || {name: inc.storeCode};
                return {
                    "STT": idx + 1,
                    "Mã Phiếu": inc.id,
                    "Thời gian": formatDateTimeView(inc.createdAt),
                    "Mã CH": inc.storeCode,
                    "Tên Cửa Hàng": storeInfo.name,
                    "Người Báo": inc.contactPerson || '',
                    "Điện Thoại": inc.phoneNumber || '',
                    "Phân Loại": inc.categories ? inc.categories.join(', ') : '',
                    "Chi Tiết (Cat 2)": inc.category2 || '',
                    "Mô Tả": inc.issue || '',
                    "Trạng Thái": inc.status || '',
                    "Xử Lý Tạm": inc.tempSolution || '',
                    "Xử Lý Lâu Dài": inc.result || '',
                    "PIC": getAssignedUserName(inc.assignee)
                };
            });

            if (exportData.length === 0) { showToast("Không có dữ liệu trong tháng này để xuất!"); return; }

            const ws = XLSX.utils.json_to_sheet(exportData);
            ws['!cols'] = [{wch: 5}, {wch: 12}, {wch: 18}, {wch: 8}, {wch: 22}, {wch: 15}, {wch: 12}, {wch: 15}, {wch: 15}, {wch: 35}, {wch: 12}, {wch: 20}, {wch: 20}, {wch: 15}];
            
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, `Bao_Cao_Thang_${month}_${year}`);
            XLSX.writeFile(wb, `Bao_Cao_Su_Co_Thang_${month}_${year}.xlsx`);
        }, 500);
    };

  </script>
</head>
<body class="bg-slate-50 text-gray-800 font-sans relative min-h-screen">

  <div id="sysToast" class="hidden"></div>

  <div id="imageViewerModal" class="hidden fixed z-[999999] inset-0 bg-black/90 flex justify-center items-center p-4" onclick="closeImageViewer()">
    <span class="absolute top-5 right-5 text-white text-4xl font-bold cursor-pointer">&times;</span>
    <img id="fullSizeImage" src="" class="max-w-full max-h-full object-contain rounded-lg shadow-2xl">
  </div>

  <!-- ĐĂNG NHẬP -->
  <div id="loginScreen" class="w-full min-h-screen bg-slate-100 absolute top-0 left-0 z-[99999] flex justify-center items-center p-4">
    <div class="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-8 border border-gray-200">
      <div class="flex items-center justify-center gap-3 mb-6">
        <div class="w-12 h-12 bg-amber-600 rounded-lg flex items-center justify-center text-white font-bold text-xl">⚠️</div>
        <div><h1 class="text-xl font-black text-slate-900">QL SỰ CỐ</h1><p class="text-[10px] text-gray-500 uppercase">Hệ Thống Ghi Nhận</p></div>
      </div>
      <form onsubmit="event.preventDefault(); handleLogin();" class="space-y-4">
        <div><label class="block text-sm font-semibold text-gray-700 mb-1">Email đăng nhập</label><input type="email" id="loginUsername" class="w-full p-2.5 border rounded-lg" required></div>
        <div><label class="block text-sm font-semibold text-gray-700 mb-1">Mật khẩu</label><input type="password" id="loginPassword" class="w-full p-2.5 border rounded-lg" required></div>
        <button id="btnLogin" type="submit" class="w-full bg-amber-600 hover:bg-amber-700 text-white font-bold py-2.5 rounded-lg shadow-md">Đăng nhập</button>
      </form>
    </div>
  </div>

  <!-- GIAO DIỆN CHÍNH -->
  <div id="mainAppContainer" class="hidden max-w-7xl mx-auto relative p-2 md:p-6">
    <!-- HEADER BAR -->
    <div class="flex flex-wrap justify-between items-center mb-6 gap-3 bg-white p-3.5 md:p-4 rounded-xl shadow-sm border border-gray-200">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 bg-amber-100 text-amber-700 rounded-lg flex items-center justify-center text-xl font-bold shrink-0">⚠️</div>
        <div>
          <h1 class="text-base md:text-lg font-black text-slate-800">QUẢN LÝ SỰ CỐ KỸ THUẬT</h1>
          <p class="text-[11px] text-gray-500">Theo dõi, xử lý & báo cáo thiết bị</p>
        </div>
      </div>

      <div class="flex items-center gap-2 md:gap-3 flex-wrap">
        <div class="text-xs text-gray-600 font-semibold bg-slate-50 px-3 py-1.5 rounded-lg border hidden lg:block">
          Đang đăng nhập: <span id="displayUserRole" class="text-blue-700 font-bold"></span>
        </div>

        <button onclick="openKnowledgeBaseModal()" class="px-3 py-1.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-md text-xs font-bold hover:bg-blue-100 flex items-center gap-1" title="Tra cứu hướng dẫn xử lý sự cố nhanh">
          <span>💡</span> <span class="hidden sm:inline">Hướng dẫn xử lý</span>
        </button>

        <button onclick="checkSLAReminders()" class="px-3 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-md text-xs font-bold hover:bg-amber-100 flex items-center gap-1" title="Kiểm tra các phiếu quá 30 phút chưa nhận việc">
          <span>⏰</span> <span class="hidden sm:inline">Quét SLA</span>
        </button>

        <!-- Chuông Thông Báo Nội Bộ -->
        <div class="relative">
          <button id="notifBellBtn" onclick="toggleNotificationDropdown(event)" class="p-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 relative transition" title="Thông báo sự cố của bạn">
            <span class="text-sm">🔔</span>
            <span id="notifBadge" class="hidden absolute -top-1.5 -right-1.5 bg-red-600 text-white text-[10px] font-black px-1.5 py-0.2 rounded-full shadow">0</span>
          </button>
          <div id="notifDropdown" class="hidden absolute right-0 mt-2 w-80 bg-white rounded-xl shadow-2xl border border-gray-200 z-[9999] overflow-hidden">
            <div class="bg-slate-800 text-white px-3.5 py-2.5 flex justify-between items-center text-xs">
              <span class="font-bold">🔔 Thông báo sự cố</span>
              <button onclick="markAllNotificationsRead()" class="text-blue-300 hover:underline text-[11px]">Đánh dấu đã đọc</button>
            </div>
            <div id="notifListContainer" class="max-h-72 overflow-y-auto"></div>
          </div>
        </div>

        <button id="btnAdminKB" onclick="openKnowledgeManager()" class="hidden px-3 py-1.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-md text-xs font-bold hover:bg-emerald-100">📖 Quản lý KB</button>
        <button id="btnAdminUsers" onclick="openUserManager()" class="hidden px-3 py-1.5 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded-md text-xs font-bold hover:bg-indigo-100">👥 User</button>
        <button onclick="handleLogout()" class="px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-md text-xs font-bold hover:bg-red-100">🚪 Thoát</button>
      </div>
    </div>

    <!-- DASHBOARD & ANALYTICS -->
    <div class="mb-6 bg-slate-100 p-4 md:p-5 rounded-2xl border shadow-sm">
      <div class="flex justify-between items-center mb-5">
        <h2 class="text-lg md:text-xl font-bold text-gray-800"><span>📈</span> Tổng quan & Biểu đồ thống kê</h2>
        <button id="btnRefreshDashboard" onclick="loadIncidents()" class="bg-white border hover:bg-gray-50 text-gray-700 text-xs md:text-sm px-3.5 py-2 rounded-lg font-semibold shadow-sm"><span>↻</span> Làm mới</button>
      </div>
      
      <div class="grid grid-cols-2 lg:grid-cols-5 gap-3 md:gap-4 mb-6">
        <div class="bg-white p-4 rounded-xl shadow-sm border"><div class="text-2xl font-black text-gray-800" id="dashTotal">0</div><div class="text-xs text-gray-500 mt-1">Tổng sự cố</div></div>
        <div class="bg-white p-4 rounded-xl shadow-sm border"><div class="text-2xl font-black text-gray-800" id="dashNew">0</div><div class="text-xs text-red-600 font-bold mt-1">Mới ghi nhận</div></div>
        <div class="bg-white p-4 rounded-xl shadow-sm border"><div class="text-2xl font-black text-gray-800" id="dashProcessing">0</div><div class="text-xs text-blue-600 font-bold mt-1">Đang xử lý</div></div>
        <div class="bg-white p-4 rounded-xl shadow-sm border"><div class="text-2xl font-black text-gray-800" id="dashResolved">0</div><div class="text-xs text-green-600 font-bold mt-1">Đã xử lý</div></div>
        <div class="bg-white p-4 rounded-xl shadow-sm border col-span-2 lg:col-span-1"><div class="text-2xl font-black text-indigo-700" id="dashMTTR">--</div><div class="text-xs text-indigo-600 font-bold mt-1">TB xử lý (MTTR)</div></div>
      </div>

      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 bg-white p-4 rounded-xl shadow-sm border">
        <div class="relative h-[250px] flex flex-col items-center">
          <h3 class="text-xs font-bold text-gray-600 uppercase mb-2">Tỷ lệ các loại lỗi chính (Category)</h3>
          <div class="w-full h-full relative"><canvas id="categoryChart"></canvas></div>
        </div>
        <div class="relative h-[250px] flex flex-col items-center">
          <h3 class="text-xs font-bold text-gray-600 uppercase mb-2">Số lượng sự cố theo Chi nhánh</h3>
          <div class="w-full h-full relative"><canvas id="branchChart"></canvas></div>
        </div>
      </div>
    </div>

    <!-- BẢNG NHẬT KÝ SỰ CỐ -->
    <div id="incidentTableWrapper" class="w-full">
      <div class="mb-4 flex justify-between items-center flex-col sm:flex-row gap-3">
        <div><h2 class="text-lg md:text-xl font-bold text-gray-800">Nhật ký sự cố & Phân công</h2><p class="text-xs text-gray-500" id="incCountText">Đang tải...</p></div>
        <button onclick="openIncidentModal()" class="w-full sm:w-auto bg-[#2c3e50] hover:bg-[#1a252f] text-white text-sm px-5 py-2.5 rounded-lg font-bold shadow-md"><span>➕</span> Lập phiếu sự cố</button>
      </div>
      
      <div class="bg-white border rounded-xl p-3.5 md:p-4 shadow-sm mb-4">
        <div class="flex flex-wrap items-center gap-2.5 text-xs">
            <input id="filterIncSearch" type="text" oninput="applyIncidentFilters()" placeholder="Tìm mã phiếu, lỗi..." class="flex-grow p-2 border rounded-lg min-w-[180px]">
            <select id="filterIncBranch" onchange="applyIncidentFilters()" class="border rounded-lg px-2 py-2 bg-white font-semibold max-w-[210px]"></select>
            <select id="filterIncStatus" onchange="applyIncidentFilters()" class="border rounded-lg px-2 py-2 bg-white font-semibold max-w-[150px]">
              <option value="Tất cả">Trạng thái: Tất cả</option>
              <option value="Mới ghi nhận">Mới ghi nhận</option>
              <option value="Đang xử lý">Đang xử lý</option>
              <option value="Đã xử lý">Đã xử lý</option>
            </select>
            <input id="filterIncDateFrom" type="date" onchange="applyIncidentFilters()" class="border rounded-lg px-2 py-1.5 bg-white">
            <input id="filterIncDateTo" type="date" onchange="applyIncidentFilters()" class="border rounded-lg px-2 py-1.5 bg-white">
            <button onclick="resetIncidentFilters()" class="text-red-600 font-bold ml-auto">↺ Xóa lọc</button>
            <button onclick="openMonthlyReportModal()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-2 rounded-lg font-bold shadow">📊 Báo cáo Tháng</button>
            <button onclick="exportIncidentsToExcel()" class="bg-green-600 hover:bg-green-700 text-white px-3.5 py-2 rounded-lg font-bold shadow">📈 Xuất Excel</button>
        </div>
      </div>

      <div class="overflow-x-auto rounded-lg border bg-white shadow-sm">
        <table class="w-full text-left border-collapse min-w-[820px]">
          <thead>
            <tr class="bg-slate-900 text-white text-[11px] uppercase tracking-wider">
              <th class="p-3">MÃ PHIẾU & THỜI GIAN</th>
              <th class="p-3">CHI NHÁNH</th>
              <th class="p-3">LIÊN LẠC</th>
              <th class="p-3">LOẠI LỖI</th>
              <th class="p-3 w-1/3">MÔ TẢ & PIC</th>
              <th class="p-3 text-center">TRẠNG THÁI</th>
              <th class="p-3 text-center">THAO TÁC</th>
            </tr>
          </thead>
          <tbody id="incidentTableBody"></tbody>
        </table>
      </div>

      <!-- THANH ĐIỀU HƯỚNG PHÂN TRANG -->
      <div id="paginationControls" class="hidden flex flex-wrap justify-between items-center gap-2 mt-4 px-4 py-3 bg-white border rounded-xl shadow-sm text-xs"></div>
    </div>

    <!-- MODAL 1: THƯ VIỆN GIẢI PHÁP MẪU (FAQ) -->
    <div id="kbModal" class="hidden fixed z-[9999] inset-0 flex items-center justify-center p-2 sm:p-4 overscroll-contain">
      <div class="fixed inset-0 bg-slate-900/60" onclick="closeKnowledgeBaseModal()"></div>
      <div class="relative bg-slate-50 rounded-xl shadow-2xl max-w-3xl w-full border overflow-hidden flex flex-col z-10 max-h-[88vh]">
        <div class="bg-blue-800 text-white px-5 py-3.5 flex justify-between items-center shrink-0">
          <h3 class="font-bold text-sm">💡 Thư Viện Hướng Dẫn Xử Lý Sự Cố Nhanh (FAQ)</h3>
          <button onclick="closeKnowledgeBaseModal()" class="text-lg hover:text-gray-300">✕</button>
        </div>
        <div class="p-4 bg-white border-b shrink-0">
          <input type="text" oninput="renderKnowledgeBaseModalList(this.value)" placeholder="Nhập từ khóa tìm hướng dẫn (VD: Printer, POS, Wi-Fi, Handy...)" class="w-full p-2.5 border rounded-lg text-sm">
        </div>
        <div id="kbContentList" class="p-4 space-y-4 overflow-y-auto overscroll-contain"></div>
      </div>
    </div>

    <!-- MODAL 2: IN BIÊN BẢN NGHIỆM THU SỰ CỐ (A4) -->
    <div id="printReportModal" class="hidden fixed z-[99999] inset-0 flex items-center justify-center p-2 sm:p-4 overscroll-contain">
      <div class="fixed inset-0 bg-slate-900/70 no-print" onclick="closePrintReportModal()"></div>
      <div class="relative bg-white rounded-xl shadow-2xl max-w-3xl w-full border overflow-hidden flex flex-col z-10 max-h-[92vh]">
        <div class="bg-slate-800 text-white px-5 py-3 flex justify-between items-center shrink-0 no-print">
          <h3 class="font-bold text-sm">🖨️ Xem Trước Biên Bản Nghiệm Thu Sự Cố</h3>
          <div class="flex items-center gap-2">
            <button onclick="window.print()" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-1.5 rounded text-xs font-bold shadow">🖨️ In Ngay (A4)</button>
            <button onclick="closePrintReportModal()" class="text-lg px-2 hover:text-gray-300">✕</button>
          </div>
        </div>

        <div id="printReportArea" class="p-8 overflow-y-auto overscroll-contain text-gray-900 text-sm leading-relaxed">
          <div class="flex justify-between items-start border-b-2 border-gray-800 pb-4 mb-6">
            <div>
              <h2 class="text-xl font-black uppercase tracking-wide">BIÊN BẢN GHI NHẬN & XỬ LÝ SỰ CỐ</h2>
              <p class="text-xs text-gray-600 mt-0.5">Hệ thống Quản lý Vận hành Kỹ thuật - Sukiya POS</p>
            </div>
            <div class="text-right">
              <div class="font-bold text-base text-red-700" id="prTicketId">SC-000000</div>
              <div class="text-xs text-gray-600">Ngày lập: <span id="prCreatedDate"></span></div>
            </div>
          </div>

          <div class="grid grid-cols-2 gap-4 mb-6 bg-gray-50 p-4 rounded border border-gray-300 text-xs">
            <div><span class="font-bold">Cửa hàng / Chi nhánh:</span> <span id="prStoreName"></span></div>
            <div><span class="font-bold">Kỹ thuật phụ trách (PIC):</span> <span id="prPic"></span></div>
            <div><span class="font-bold">Địa chỉ:</span> <span id="prStoreAddr"></span></div>
            <div><span class="font-bold">Trạng thái hiện tại:</span> <span id="prStatus" class="font-bold uppercase"></span></div>
            <div><span class="font-bold">Người báo sự cố:</span> <span id="prContact"></span></div>
            <div><span class="font-bold">Thời gian hoàn tất:</span> <span id="prResolvedDate"></span></div>
            <div class="col-span-2"><span class="font-bold">Phân loại thiết bị:</span> <span id="prCategory"></span></div>
          </div>

          <div class="space-y-4 mb-6">
            <div>
              <h4 class="font-bold uppercase text-xs text-gray-700 border-b pb-1 mb-1.5">1. Mô tả chi tiết sự cố:</h4>
              <p id="prIssue" class="p-3 border rounded bg-white min-h-[50px] whitespace-pre-line"></p>
            </div>
            <div>
              <h4 class="font-bold uppercase text-xs text-gray-700 border-b pb-1 mb-1.5">2. Biện pháp xử lý tạm thời:</h4>
              <p id="prTempSol" class="p-3 border rounded bg-white min-h-[40px] whitespace-pre-line"></p>
            </div>
            <div>
              <h4 class="font-bold uppercase text-xs text-gray-700 border-b pb-1 mb-1.5">3. Giải pháp khắc phục triệt để (Lâu dài):</h4>
              <p id="prLongSol" class="p-3 border rounded bg-white min-h-[50px] whitespace-pre-line"></p>
            </div>
            <div id="prImageWrapper" class="hidden">
              <h4 class="font-bold uppercase text-xs text-gray-700 border-b pb-1 mb-2">4. Hình ảnh hiện trường đính kèm:</h4>
              <img id="prImage" src="" class="max-h-56 border rounded p-1">
            </div>
          </div>

          <div class="grid grid-cols-2 gap-8 pt-6 mt-8 border-t border-gray-300 text-center text-xs">
            <div>
              <div class="font-bold uppercase">ĐẠI DIỆN CỬA HÀNG</div>
              <div class="text-gray-500 italic mt-0.5">(Ký và ghi rõ họ tên)</div>
              <div class="h-20"></div>
            </div>
            <div>
              <div class="font-bold uppercase">KỸ THUẬT VIÊN XỬ LÝ (IT)</div>
              <div class="text-gray-500 italic mt-0.5">(Ký và ghi rõ họ tên)</div>
              <div class="h-20"></div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- MODAL 3: BÁO CÁO TỔNG HỢP THÁNG -->
    <div id="monthlyReportModal" class="hidden fixed z-[99999] inset-0 flex items-center justify-center p-2 sm:p-4 overscroll-contain">
      <div class="fixed inset-0 bg-slate-900/70" onclick="closeMonthlyReportModal()"></div>
      <div class="relative bg-white rounded-xl shadow-2xl max-w-4xl w-full border overflow-hidden flex flex-col z-10 max-h-[90vh]">
        <div class="bg-indigo-900 text-white px-5 py-3.5 flex justify-between items-center shrink-0">
          <div class="flex items-center gap-2">
            <span class="text-lg">📊</span>
            <h3 class="font-bold text-sm">Báo Cáo Tổng Hợp Vận Hành & Sự Cố Theo Tháng</h3>
          </div>
          <div class="flex items-center gap-2">
            <select id="reportMonthSelect" onchange="generateMonthlyReportData()" class="bg-indigo-800 text-white text-xs px-2.5 py-1 rounded border border-indigo-700 font-semibold"></select>
            <button onclick="exportMonthlyReportExcel()" class="bg-green-600 hover:bg-green-700 text-white px-3 py-1 rounded text-xs font-bold shadow">📈 Xuất Excel Báo Cáo</button>
            <button onclick="closeMonthlyReportModal()" class="text-lg px-2 hover:text-gray-300">✕</button>
          </div>
        </div>

        <div class="p-5 overflow-y-auto overscroll-contain text-xs">
          <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
            <div class="bg-slate-50 p-3 rounded-lg border text-center">
              <div class="text-xl font-black text-slate-800" id="repTotalIncidents">0</div>
              <div class="text-[11px] text-gray-500 mt-0.5">Tổng số sự cố tháng</div>
            </div>
            <div class="bg-blue-50 p-3 rounded-lg border border-blue-200 text-center">
              <div class="text-xl font-black text-blue-700" id="repResolvedIncidents">0</div>
              <div class="text-[11px] text-blue-600 mt-0.5">Đã xử lý xong (Done)</div>
            </div>
            <div class="bg-amber-50 p-3 rounded-lg border border-amber-200 text-center">
              <div class="text-xl font-black text-amber-700" id="repPendingIncidents">0</div>
              <div class="text-[11px] text-amber-600 mt-0.5">Đang xử lý / Mới</div>
            </div>
            <div class="bg-indigo-50 p-3 rounded-lg border border-indigo-200 text-center">
              <div class="text-xl font-black text-indigo-700" id="repAvgMttr">--</div>
              <div class="text-[11px] text-indigo-600 mt-0.5">Thời gian MTTR TB</div>
            </div>
          </div>

          <div class="font-bold text-slate-800 mb-2">Chi tiết tổng hợp theo từng Chi nhánh / Cửa hàng:</div>
          <div class="overflow-x-auto rounded-lg border">
            <table class="w-full text-left border-collapse">
              <thead>
                <tr class="bg-slate-800 text-white text-[11px] uppercase">
                  <th class="p-2.5">Mã CN</th>
                  <th class="p-2.5">Tên Chi Nhánh</th>
                  <th class="p-2.5 text-center">Tổng Phiếu</th>
                  <th class="p-2.5 text-center">Mới / Đang xử lý</th>
                  <th class="p-2.5 text-center">Đã xong</th>
                  <th class="p-2.5 text-center">Tỷ lệ hoàn thành</th>
                  <th class="p-2.5 text-center">MTTR Trung bình</th>
                </tr>
              </thead>
              <tbody id="monthlyReportTableBody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- MODAL 4: QUẢN LÝ KNOWLEDGE BASE (THÊM / SỬA / XÓA) -->
    <div id="knowledgeManagerModal" class="hidden fixed z-[9999] inset-0 flex items-center justify-center p-2 sm:p-4 overscroll-contain">
      <div class="fixed inset-0 bg-slate-900/60" onclick="closeKnowledgeManager()"></div>
      <div class="relative bg-slate-50 rounded-lg shadow-2xl max-w-5xl w-full flex flex-col z-10 border max-h-[90vh] overflow-hidden">
        <div class="bg-emerald-800 text-white px-4 py-3 flex justify-between items-center shrink-0">
          <h3 class="font-bold text-sm">📖 Quản Lý Kho Dữ Liệu Giải Pháp (Knowledge Base)</h3>
          <button onclick="closeKnowledgeManager()" class="text-xl font-bold">&times;</button>
        </div>
        <div class="p-5 grid grid-cols-1 lg:grid-cols-3 gap-6 overflow-y-auto overscroll-contain">
          <form id="kbForm" onsubmit="event.preventDefault(); saveKnowledgeBaseItem();" class="space-y-3 bg-white p-4 rounded border shadow-sm">
              <h4 id="kbFormTitle" class="font-bold text-emerald-800 text-sm border-b pb-1.5">➕ Thêm Giải Pháp Mới</h4>
              <input type="hidden" id="kbDocId">
              
              <div>
                <label class="text-xs font-semibold text-gray-700">Phân loại chính (Category) *</label>
                <select id="kbCategory" onchange="onKbCategoryChange()" class="w-full p-2 border rounded text-xs font-semibold bg-white" required>
                  <option value="Network">Network</option>
                  <option value="POS">POS</option>
                  <option value="Printer">Printer</option>
                  <option value="Server">Server</option>
                  <option value="CCTV">CCTV</option>
                  <option value="Hardware">Hardware</option>
                  <option value="Software">Software</option>
                  <option value="Handy">Handy</option>
                  <option value="Attendance">Attendance</option>
                  <option value="KD">KD</option>
                  <option value="Khác">Khác</option>
                </select>
              </div>

              <div>
                <label class="text-xs font-semibold text-gray-700">Chi tiết lỗi (Category 2) *</label>
                <select id="kbCategory2" class="w-full p-2 border rounded text-xs font-semibold text-blue-900 bg-white" required>
                  <option value="">-- Chọn Chi tiết lỗi --</option>
                </select>
              </div>

              <div>
                <label class="text-xs font-semibold text-gray-700">Tiêu đề lỗi / Triệu chứng *</label>
                <input id="kbTitle" type="text" placeholder="VD: Máy in không ra bill..." class="w-full p-2 border rounded text-xs" required>
              </div>

              <div>
                <label class="text-xs font-semibold text-gray-700">Các bước xử lý (Mỗi dòng 1 bước) *</label>
                <textarea id="kbSteps" rows="5" placeholder="1. Kiểm tra nguồn...&#10;2. Kiểm tra dây mạng..." class="w-full p-2 border rounded text-xs leading-relaxed" required></textarea>
              </div>

              <div class="flex gap-2 pt-1">
                  <button type="submit" id="btnSubmitKB" class="flex-1 bg-emerald-700 hover:bg-emerald-800 text-white py-2 rounded font-bold text-xs shadow">Lưu Giải Pháp</button>
                  <button type="button" onclick="resetKBForm()" class="px-3 bg-gray-200 hover:bg-gray-300 text-gray-700 py-2 rounded font-bold text-xs">Làm mới</button>
              </div>
          </form>

          <div class="lg:col-span-2 bg-white p-4 rounded border shadow-sm flex flex-col">
              <div class="flex justify-between items-center mb-3">
                  <h4 class="font-bold text-xs text-gray-700 uppercase">Danh sách giải pháp trong hệ thống</h4>
                  <input type="text" oninput="renderKBManagerTable(this.value)" placeholder="Tìm kiếm nhanh..." class="p-1.5 border rounded text-xs max-w-[200px]">
              </div>
              <div class="overflow-auto max-h-[420px] border rounded">
                  <table class="w-full text-left text-xs">
                      <thead class="bg-slate-100 text-gray-700 sticky top-0">
                          <tr class="border-b">
                              <th class="p-2.5">Danh mục / Tiêu đề</th>
                              <th class="p-2.5">Các bước xử lý</th>
                              <th class="p-2.5 text-center">Thao tác</th>
                          </tr>
                      </thead>
                      <tbody id="kbTableBody"></tbody>
                  </table>
              </div>
          </div>
        </div>
      </div>
    </div>

    <!-- MODAL 5: USER MANAGER -->
    <div id="userManagerModal" class="hidden fixed z-[9999] inset-0 flex items-center justify-center p-2 sm:p-4 overscroll-contain">
      <div class="fixed inset-0 bg-slate-900/60" onclick="closeUserManager()"></div>
      <div class="relative bg-slate-50 rounded-lg shadow-2xl max-w-5xl w-full flex flex-col z-10 border max-h-[90vh] overflow-hidden">
        <div class="bg-[#4a3fcb] text-white px-4 py-3 flex justify-between items-center shrink-0"><h3 class="font-bold">Quản Lý Tài Khoản</h3><button onclick="closeUserManager()" class="text-xl">&times;</button></div>
        <div class="p-5 grid grid-cols-1 lg:grid-cols-3 gap-6 overflow-y-auto overscroll-contain">
          <form id="userForm" onsubmit="event.preventDefault(); saveUserManager();" class="space-y-3 bg-white p-4 rounded border">
              <h4 id="umFormTitle" class="font-bold text-[#4a3fcb]">Thêm User</h4>
              <div><label class="text-xs font-semibold">Email *</label><input id="umEmail" type="email" class="w-full p-2 border rounded text-sm" required></div>
              <div id="umPasswordContainer"><label class="text-xs font-semibold">Mật khẩu *</label><input id="umPassword" type="text" class="w-full p-2 border rounded text-sm" required></div>
              <div><label class="text-xs font-semibold">Họ và Tên NV</label><input id="umFullName" type="text" class="w-full p-2 border rounded text-sm"></div>
              <div><label class="text-xs font-semibold">Chức vụ *</label><select id="umRole" class="w-full p-2 border rounded text-sm"><option value="Nhân viên">Nhân viên</option><option value="Quản trị viên">Quản trị viên</option></select></div>
              <div><label class="text-xs font-semibold">Chi nhánh</label><select id="umStoreCode" class="w-full p-2 border rounded text-sm"></select></div>
              <button type="submit" id="btnSubmitUser" class="w-full bg-[#4a3fcb] text-white py-2 rounded font-bold text-sm">Lưu Tài Khoản</button>
          </form>
          <div class="lg:col-span-2 bg-white p-4 rounded border overflow-auto max-h-[450px]">
              <table class="w-full text-left text-sm">
                  <thead><tr class="border-b text-gray-500 text-xs"><th>Email</th><th>Họ Tên</th><th>Chức vụ</th><th class="text-center">Thao tác</th></tr></thead>
                  <tbody id="userTableBody"></tbody>
              </table>
          </div>
        </div>
      </div>
    </div>

    <!-- MODAL 6: PHIẾU SỰ CỐ -->
    <div id="incidentModal" class="hidden fixed z-[9999] inset-0 flex items-center justify-center p-2 sm:p-4 overscroll-contain">
      <div class="fixed inset-0 bg-slate-900/60" onclick="closeIncidentModal()"></div>
      <div class="relative bg-white rounded-xl shadow-2xl max-w-3xl w-full border overflow-hidden flex flex-col z-10 max-h-[90vh]">
        <div class="bg-[#2c3e50] text-white px-5 py-3 flex justify-between items-center shrink-0">
          <h3 id="modalIncidentTitle" class="font-bold text-sm">📝 Lập Phiếu / Cập Nhật</h3>
          <div class="flex items-center gap-2">
            <button type="button" id="btnModalPrint" class="hidden bg-slate-700 hover:bg-slate-600 text-white px-3 py-1 rounded text-xs font-bold border border-slate-500">🖨️ In Biên Bản</button>
            <button type="button" onclick="closeIncidentModal()" class="text-lg hover:text-gray-300 ml-1">✕</button>
          </div>
        </div>
        
        <form id="incidentForm" onsubmit="event.preventDefault(); saveIncidentToFirebase();" class="p-4 md:p-5 space-y-4 text-[13px] overflow-y-auto overscroll-contain">
          <div class="grid grid-cols-2 gap-3 md:gap-4">
            <div><label class="text-gray-500">Mã Phiếu</label><input id="incTicket" type="text" class="w-full p-2 bg-gray-50 font-bold border rounded" readonly></div>
            <div><label class="text-gray-500">Thời gian</label><input id="incTime" type="text" class="w-full p-2 bg-gray-50 border rounded" readonly><input type="hidden" id="incMonth"><input type="hidden" id="incDate"></div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 md:gap-4">
            <div><label class="text-gray-500">*Mã Cửa Hàng</label><select id="incStoreCode" onchange="onIncidentStoreChange()" class="w-full p-2 border rounded font-semibold" required></select></div>
            <div><label class="text-gray-500">Khu vực / Địa chỉ</label><input id="incDistrict" type="text" class="w-full p-2 bg-gray-50 border rounded" readonly><input type="hidden" id="incAddress"></div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 md:gap-4">
            <div><label class="text-gray-500">*Người liên lạc</label><input id="incContact" type="text" class="w-full p-2 border rounded" required></div>
            <div><label class="text-gray-500">*Số điện thoại</label><input id="incPhone" type="tel" class="w-full p-2 border rounded" required></div>
          </div>

          <div class="border-t pt-3">
            <label class="block font-bold text-gray-700 mb-2">*Phân loại chính / Category</label>
            <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2 mb-3">
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Network" onchange="onCategoryChange()"> Network</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="POS" onchange="onCategoryChange()"> POS</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Printer" onchange="onCategoryChange()"> Printer</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Server" onchange="onCategoryChange()"> Server</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="CCTV" onchange="onCategoryChange()"> CCTV</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Hardware" onchange="onCategoryChange()"> Hardware</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Software" onchange="onCategoryChange()"> Software</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Handy" onchange="onCategoryChange()"> Handy</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Attendance" onchange="onCategoryChange()"> Attendance</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="KD" onchange="onCategoryChange()"> KD</label>
              <label class="flex items-center gap-1.5"><input type="checkbox" name="incCategory" value="Khác" onchange="onCategoryChange()"> Khác</label>
            </div>

            <div class="mt-2">
              <label class="block font-bold text-gray-700 mb-1">*Chi tiết lỗi (Category 2)</label>
              <select id="incCategory2" onchange="onCategory2Change()" class="w-full p-2 border rounded font-semibold text-blue-900 bg-white" required>
                <option value="">-- Chọn Phân loại chính ở trên trước --</option>
              </select>
            </div>

            <!-- Hộp gợi ý giải pháp mẫu chỉ hiện khi chọn Category 2 -->
            <div id="inlineKbSuggestion" class="hidden mt-3 bg-blue-50/90 border border-blue-200 rounded-lg p-3 text-xs">
              <div id="inlineKbContent"></div>
            </div>
          </div>

          <div>
            <label class="block font-bold text-gray-700 mb-1">*Mô tả chi tiết / Description</label>
            <textarea id="incIssue" rows="3" class="w-full p-2 border rounded" required></textarea>
            <div class="mt-2">
              <label class="block font-semibold">📸 Chụp / Đính kèm ảnh hiện trường (Hỗ trợ Camera di động)</label>
              <input type="file" id="incImage" accept="image/*" capture="environment" onchange="previewImage(event)" class="w-full p-1.5 border rounded text-xs bg-slate-50">
              <div id="imagePreviewContainer" class="hidden mt-2"><img id="imagePreview" class="h-24 object-contain rounded border shadow-sm"></div>
            </div>
          </div>

          <div class="border-t pt-3">
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 md:gap-4 mb-3">
              <div><label class="font-bold">Phân công kỹ thuật (PIC)</label><select id="incLeader" class="w-full p-2 border rounded font-semibold text-indigo-700"></select></div>
              <div><label class="font-bold">Trạng thái phiếu</label>
                <select id="incStatus" onchange="onIncidentStatusChange()" class="w-full p-2 border rounded font-semibold">
                  <option value="Mới ghi nhận">1. Mới ghi nhận</option>
                  <option value="Đang xử lý">2. Đang xử lý</option>
                  <option value="Đã xử lý">3. Đã xử lý (Done)</option>
                </select>
              </div>
            </div>

            <div id="incTempContainer" class="hidden mt-2">
              <label class="block font-bold text-indigo-800 mb-1">*Xử lý tạm thời:</label>
              <input id="incTempSolution" type="text" placeholder="Nhập xử lý tạm thời..." class="w-full p-2 border rounded bg-indigo-50/50">
            </div>

            <div id="incResultContainer" class="hidden mt-2 bg-green-50 p-3 rounded border border-green-200">
              <label class="block font-bold text-green-800 mb-1">*Xử lý lâu dài:</label>
              <textarea id="incResult" rows="2" placeholder="Nhập xử lý lâu dài..." class="w-full p-2 border rounded"></textarea>
            </div>
          </div>

          <div id="auditTimelineSection" class="hidden border-t pt-3">
            <h4 class="font-bold text-xs text-gray-600 uppercase mb-2.5">🕒 Lịch sử thao tác & Cập nhật phiếu</h4>
            <div id="auditTimelineList" class="bg-slate-50 p-3.5 rounded-lg border border-slate-200 max-h-40 overflow-y-auto"></div>
          </div>

          <div class="text-center pt-2">
            <div id="formErrorMsg" class="hidden bg-red-100 text-red-700 font-bold p-2 rounded mb-3 text-xs"></div>
            <button type="submit" id="btnSubmitForm" class="w-full sm:w-auto px-8 py-2.5 bg-[#2c3e50] hover:bg-[#1a252f] text-white rounded-lg font-bold shadow">Lưu Phiếu / Save Ticket</button>
          </div>
        </form>
      </div>
    </div>
  </div>
</body>
</html>
