// 29 Institutes Database Generation + Admin
const INSTITUTES = [
  {
    id: 0,
    code: "admin",
    name: "AP Gurukulam (State Admin)",
    username: "admin",
    password: "admin@123"
  }
];
const INSTITUTE_DATA = [
  { name: "EMRS(Co-Edu),Bhamini", code: "28111401524" },
  { name: "EMRS(Co-Edu),Meliaputi", code: "28112600218" },
  { name: "EMRS(Co-Edu),Anasabhadra", code: "28120701907" },
  { name: "EMRS(Co-Edu),Kotikapenta", code: "28121204105" },
  { name: "EMRS(Co-Edu),Kurupam", code: "28120305828" },
  { name: "EMRS(Co-Edu),GLPuram", code: "28120207003" },
  { name: "EMRS(Co-Edu),K.Veedhi", code: "28131306710" },
  { name: "EMRS(Co-Edu),Ananthagiri", code: "28137100093" },
  { name: "EMRS(Co-Edu),Arakuvalley", code: "28137100094" },
  { name: "EMRS(Co-Edu),Paderu", code: "28137100095" },
  { name: "EMRS(Co-Edu),Hukumpeta", code: "28137100099" },
  { name: "EMRS(Co-Edu),Peddabayalu", code: "28137100092" },
  { name: "EMRS(Co-Edu),G.Madugula", code: "28137100100" },
  { name: "EMRS(Co-Edu),Koyyuru", code: "28137100101" },
  { name: "EMRS(Co-Edu),Dumbriguda", code: "28130403207" },
  { name: "EMRS(Co-Edu),Munchigput", code: "28130115711" },
  { name: "EMRS(Co-Edu),Chintapalli", code: "28131210110" },
  { name: "EMRS(Co-Edu),Y.Ramavaram", code: "28140208902" },
  { name: "EMRS(Co-Edu),Maredumilli", code: "28140104901" },
  { name: "EMRS(Co-Edu),R.C. Varam", code: "28147100147" },
  { name: "EMRS(Co-Edu),Addatheegala", code: "28147100145" },
  { name: "EMRS(Co-Edu),Chintoor", code: "28146205301" },
  { name: "EMRS(Co-Edu),Rajavommangi", code: "28140400202" },
  { name: "EMRS(Co-Edu),Buttayagudem", code: "28150202542" },
  { name: "EMRS(Co-Edu),Dornala", code: "28180700734" },
  { name: "EMRS(Co-Edu),Kodavaluru", code: "28191701107" },
  { name: "EMRS(Co-Edu),Ojili", code: "28193703508" },
  { name: "EMRS(Co-Edu),B.N.Khandriga", code: "28231601515" }
];

INSTITUTE_DATA.forEach((inst, index) => {
  const id = index + 1;
  const uniquePassword = "Gest@123";

  INSTITUTES.push({
    id: id,
    code: inst.code,
    name: inst.name,
    username: inst.code,
    password: uniquePassword
  });
});

// Student Data Store
let students = [];
let currentInstitute = null;

// Column Mappings for Fuzzy Matching Excel Headers
const COLUMN_MAPPINGS = {
  sNo: ['s.no.', 'sno', 'sl.no.', 'slno', 'serial number', 'serialno'],
  candidateId: ['candidate id', 'id', 'candidateid', 'roll no', 'roll number', 'rollno'],
  hallTicketNo: ['hall ticket no', 'hall ticket number', 'hallticket no', 'hallticketnumber', 'hall ticket', 'hallticket', 'ticket no'],
  candidateName: ['candidate name', 'name', 'student name', 'candidatename', 'studentname'],
  aadharNumber: ['aadhar number', 'aadhar', 'aadhar card', 'aadharno', 'aadhar no', 'uidai'],
  dob: ['dob', 'date of birth', 'birth date', 'birthdate', 'd.o.b'],
  gender: ['gender', 'sex', 'male/female'],
  parentName: ['parent/guardian name', 'parent guardian name', 'parent name', 'guardian name', 'father name', 'fathername', 'parent/guardian', 'father/guardian name'],
  designation: ['designation', 'post', 'role', 'job title', 'jobtitle', 'designation details'],
  caste: ['caste', 'category', 'social status'],
  physicallyChallenged: ['are you physically challenged ?', 'physically challenged', 'physically challenged ?', 'handicapped', 'pwd', 'ph'],
  mobileNumber: ['mobile number', 'mobile', 'mobile no', 'phone', 'phone number', 'contact no', 'contact number'],
  appliedClass: ['applied class', 'class', 'appliedclass', 'admission class'],
  address: ['address', 'candidate address', 'communication address', 'permanent address'],
  principalDetails: ['principal contact details', 'principal details', 'principal contact', 'principal phone', 'principal contact detail is'],
  examDateTime: ['exam date & time', 'exam date and time', 'exam date & tim e', 'exam datetime', 'exam date', 'exam time', 'exam date & time e'],
  examCenterAddress: ['exam center address', 'exam centre address', 'exam center', 'exam centre', 'exam center code & address', 'exam center details'],
  instituteCode: ['institute code', 'institutecode', 'center code', 'centercode', 'exam center code'],
  photo: ['photo', 'photopath', 'photo path', 'candidate photo', 'student photo'],
  signature: ['signature', 'signaturepath', 'signature path', 'candidate signature', 'student signature']
};

// Default Mock Data Distributed Across Multiple Institutes
const MOCK_STUDENTS = [
  {
    candidateId: "CAD1001",
    hallTicketNo: "HT20261001",
    candidateName: "Ramanjaneyulu K",
    aadharNumber: "********1234",
    dob: "15-08-1995",
    gender: "MALE",
    parentName: "K. Satyanarayana",
    designation: "POST GRADUATE TEACHER (PGT) - ENGLISH",
    caste: "ST",
    physicallyChallenged: "NO",
    mobileNumber: "9876543210",
    appliedClass: "N/A",
    address: "H.No. 4-123, Near Hanuman Temple, Bhamini, Srikakulam District, Andhra Pradesh - 532456",
    principalDetails: "9440000000",
    examDateTime: "14-07-2026 10:00 AM to 12:00 PM",
    examCenterAddress: "EMRS (Co-Edu), Kurupam, Vizianagaram District, AP",
    instituteCode: "28111401524"
  }
];

// Initialize application on load
window.addEventListener('DOMContentLoaded', () => {
  const storedStudents = localStorage.getItem('all_students');
  if (storedStudents) {
    try {
      students = JSON.parse(storedStudents);
      if (!Array.isArray(students) || students.length === 0) {
        students = JSON.parse(JSON.stringify(MOCK_STUDENTS));
        saveToStorage();
      } else {
        // Clean up cached student Aadhar numbers to ensure they have exactly 8 stars
        students.forEach(student => {
          if (student.aadharNumber) {
            const match = student.aadharNumber.toString().trim().match(/^\*+(\d{4})$/);
            if (match) {
              student.aadharNumber = '********' + match[1];
            }
          }
        });
      }
    } catch (e) {
      students = JSON.parse(JSON.stringify(MOCK_STUDENTS));
      saveToStorage();
    }
  } else {
    students = JSON.parse(JSON.stringify(MOCK_STUDENTS));
    saveToStorage();
  }

  setupEventListeners();
  checkSessionState();
  loadExcelData(true);
});

// Setup event listeners
function setupEventListeners() {
  const loginForm = document.getElementById('loginForm');
  loginForm.addEventListener('submit', handleLogin);

  document.getElementById('logoutBtn').addEventListener('click', handleLogout);

  const searchInput = document.getElementById('searchQuery');
  searchInput.addEventListener('input', renderStudentList);

  document.getElementById('printSelectedBtn').addEventListener('click', () => printHallTickets(true));
  document.getElementById('printAllBtn').addEventListener('click', () => printHallTickets(false));

  const selectAllCheckbox = document.getElementById('selectAll');
  selectAllCheckbox.addEventListener('change', (e) => {
    const checkboxes = document.querySelectorAll('.student-checkbox');
    checkboxes.forEach(cb => cb.checked = e.target.checked);
    updateSelectedCount();
  });

  const adminFilterSelect = document.getElementById('adminInstFilter');
  adminFilterSelect.addEventListener('change', () => {
    renderStudentList();
  });

  // Clear error message when user starts typing
  document.getElementById('username').addEventListener('input', () => {
    document.getElementById('loginError').classList.add('hidden');
  });
  document.getElementById('password').addEventListener('input', () => {
    document.getElementById('loginError').classList.add('hidden');
  });

  // Handle single candidate print/download button clicks
  document.getElementById('studentGrid').addEventListener('click', (e) => {
    const btn = e.target.closest('.print-single-btn');
    if (btn) {
      e.stopPropagation();
      const sNo = btn.dataset.sno;
      const student = students.find(s => String(s.sNo) === String(sNo));
      if (student) {
        printHallTickets(false, [student]);
      }
    }
  });
}

// Fetch and load the Excel file automatically
function loadExcelData(isSilent = false) {
  fetch('students.xlsx?t=' + Date.now())
    .then(res => {
      if (!res.ok) throw new Error('File not found on server.');
      return res.arrayBuffer();
    })
    .then(ab => {
      const data = new Uint8Array(ab);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      if (jsonData.length === 0) {
        return;
      }

      const parsedStudents = jsonData.map(mapRowToStudent);
      const validStudents = parsedStudents.filter(s => s.candidateId || s.candidateName);

      if (validStudents.length > 0) {
        students = validStudents;
        saveToStorage();
        renderStudentList();
        if (!isSilent) showToast('Refreshed student data from students.xlsx!');
      }
    })
    .catch(err => {
      console.warn('Automatic students.xlsx load bypassed. Active data loaded from storage/mock.', err);
      renderStudentList();
    });
}



// Session State Management
function checkSessionState() {
  const savedCode = localStorage.getItem('loggedInInstituteCode');
  if (savedCode) {
    const matched = INSTITUTES.find(inst => inst.code === savedCode);
    if (matched) {
      logInSession(matched);
      return;
    }
  }

  document.getElementById('loginOverlay').classList.remove('hidden');
  document.getElementById('appDashboard').classList.remove('active');
}

function handleLogin(e) {
  e.preventDefault();
  const user = document.getElementById('username').value.trim();
  const pass = document.getElementById('password').value.trim();

  const matched = INSTITUTES.find(inst => inst.username === user && inst.password === pass);
  const errorDiv = document.getElementById('loginError');

  if (matched) {
    errorDiv.classList.add('hidden');
    logInSession(matched);
    showToast(`Welcome back, ${matched.name}!`);
  } else {
    errorDiv.classList.remove('hidden');
    // Force Lucide icons to render in the dynamically shown error alert
    lucide.createIcons();
  }
}

function logInSession(institute) {
  currentInstitute = institute;
  localStorage.setItem('loggedInInstituteCode', institute.code);

  document.getElementById('sessionInstName').innerText = institute.name;
  document.getElementById('sessionInstCode').innerText = `Code: ${institute.code}`;

  // Handle Admin filter wrapper visibility and populate options dynamically
  const filterWrapper = document.getElementById('adminFilterWrapper');
  const filterSelect = document.getElementById('adminInstFilter');

  if (institute.code === 'admin') {
    filterWrapper.classList.remove('hidden');
    if (filterSelect.options.length <= 1) {
      // Sort and populate all 29 unique institute options
      const sortedInsts = INSTITUTES.filter(i => i.code !== 'admin').sort((a, b) => a.name.localeCompare(b.name));
      sortedInsts.forEach(inst => {
        const opt = document.createElement('option');
        opt.value = inst.code;
        opt.textContent = `${inst.name} (${inst.code})`;
        filterSelect.appendChild(opt);
      });
    }
    filterSelect.value = 'all';
  } else {
    filterWrapper.classList.add('hidden');
  }

  document.getElementById('loginOverlay').classList.add('hidden');
  document.getElementById('appDashboard').classList.add('active');

  renderStudentList();
}

function handleLogout() {
  if (confirm('Are you sure you want to log out?')) {
    currentInstitute = null;
    localStorage.removeItem('loggedInInstituteCode');

    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
    document.getElementById('loginError').classList.add('hidden');

    document.getElementById('adminFilterWrapper').classList.add('hidden');
    document.getElementById('loginOverlay').classList.remove('hidden');
    document.getElementById('appDashboard').classList.remove('active');

    showToast('Logged out successfully.');
  }
}

// Local Storage Sync helper
function saveToStorage() {
  localStorage.setItem('all_students', JSON.stringify(students));
}

// Clean string for fuzzy lookup
function cleanString(str) {
  if (!str) return '';
  return str.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
}

// Fuzzy find keys in Excel row
function mapRowToStudent(row) {
  const student = {};
  const rowKeys = Object.keys(row);

  for (const [targetKey, synonyms] of Object.entries(COLUMN_MAPPINGS)) {
    let matchedValue = '';

    for (const key of rowKeys) {
      const cleanKey = cleanString(key);
      const isMatch = synonyms.some(synonym => cleanString(synonym) === cleanKey || cleanKey.includes(cleanString(synonym)));

      if (isMatch) {
        matchedValue = row[key];
        break;
      }
    }

    student[targetKey] = matchedValue !== undefined && matchedValue !== null ? matchedValue.toString().trim() : '';
  }

  // Format Aadhar if needed
  if (student.aadharNumber) {
    const match = student.aadharNumber.toString().trim().match(/^\*+(\d{4})$/);
    if (match) {
      student.aadharNumber = '********' + match[1];
    } else {
      const clean = student.aadharNumber.toString().replace(/[^0-9]/g, '');
      if (clean.length === 12) {
        student.aadharNumber = '********' + clean.slice(-4);
      } else if (clean.length === 4) {
        student.aadharNumber = '********' + clean;
      } else if (student.aadharNumber.length === 4) {
        student.aadharNumber = '********' + student.aadharNumber;
      }
    }
  }

  // Clean up numerical strings (remove trailing .0)
  for (const key of ['sNo', 'candidateId', 'aadharNumber', 'mobileNumber', 'instituteCode']) {
    if (student[key]) {
      student[key] = student[key].toString().trim().replace(/\.0$/, '');
    }
  }

  return student;
}

// Get scoped list of students for logged-in institute
function getScopedStudents() {
  if (!currentInstitute) return [];
  if (currentInstitute.code === 'admin') {
    const filterVal = document.getElementById('adminInstFilter').value;
    if (filterVal === 'all') {
      return students;
    }
    return students.filter(s => s.instituteCode === filterVal);
  }
  return students.filter(s => s.instituteCode === currentInstitute.code);
}

// Render the grid / list of students on the Dashboard
function renderStudentList() {
  const container = document.getElementById('studentGrid');
  const searchQuery = document.getElementById('searchQuery').value.toLowerCase();

  container.innerHTML = '';

  const scopedList = getScopedStudents();

  const filteredStudents = scopedList.filter(student => {
    return (
      (student.candidateName || '').toLowerCase().includes(searchQuery) ||
      (student.candidateId || '').toLowerCase().includes(searchQuery) ||
      (student.hallTicketNo || '').toLowerCase().includes(searchQuery) ||
      (student.examCenterAddress || '').toLowerCase().includes(searchQuery)
    );
  });

  if (filteredStudents.length === 0) {
    const isAdmin = currentInstitute.code === 'admin';
    container.innerHTML = `
      <div class="empty-state">
        <i data-lucide="users-round" class="empty-icon"></i>
        <h3>No students found</h3>
        <p>${isAdmin ? 'Ensure your "students.xlsx" file contains student records.' : `Ensure your "students.xlsx" file contains records matching institute code <strong>${currentInstitute.code}</strong>.`}</p>
      </div>
    `;
    lucide.createIcons();
    return;
  }

  filteredStudents.forEach(student => {
    const card = document.createElement('div');
    card.className = 'student-card';

    // Find the institute name
    const inst = INSTITUTES.find(i => i.code === student.instituteCode);
    const instName = inst ? inst.name : `Code: ${student.instituteCode || 'Unknown'}`;

    const photoUrl = student.photo || (student.sNo ? `photos/p${student.sNo}.jpg` : `photos/${(student.candidateId || '').toUpperCase()}.jpg`);

    card.innerHTML = `
      <div class="card-header">
        <label class="checkbox-container">
          <input type="checkbox" class="student-checkbox" value="${student.sNo || student.candidateId}" checked>
          <span class="checkmark"></span>
        </label>
        <span class="hallticket-badge">${(student.hallTicketNo || '').replace(/\s+/g, '') || 'No Hall Ticket #'}</span>
      </div>
      <div class="card-body student-card-flex">
        <div class="student-card-photo-wrapper">
          <img src="${photoUrl}" class="student-card-photo" alt="Photo" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';">
          <div class="student-card-photo-placeholder" style="display:none;">
            <i data-lucide="user" class="card-avatar-icon"></i>
          </div>
        </div>
        <div class="student-info">
          <h4 class="student-name">${student.candidateName || 'Unnamed Student'}</h4>
          <p class="student-id">ID: <span>${student.candidateId || 'N/A'}</span></p>
          <p class="student-detail">DOB: <span>${student.dob || 'N/A'}</span> | Gender: <span>${student.gender || 'N/A'}</span></p>
          <p class="student-detail">Aadhar: <span>${student.aadharNumber || 'N/A'}</span></p>
          <p class="student-detail">Institute: <span>${instName}</span></p>
        </div>
      </div>
      <div class="card-footer">
        <button class="btn btn-sm btn-outline print-single-btn" data-sno="${student.sNo}">
          <i data-lucide="download" class="btn-icon-small"></i> Download Hall Ticket
        </button>
      </div>
    `;

    container.appendChild(card);
  });

  lucide.createIcons();
  updateSelectedCount();
}

// Update Selected Count
function updateSelectedCount() {
  const checked = document.querySelectorAll('.student-checkbox:checked').length;
  document.getElementById('selectedCountBadge').innerText = `${checked} Selected`;
}

// Dynamic toast message
function showToast(message) {
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerText = message;
  document.body.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('visible');
  }, 100);

  setTimeout(() => {
    toast.classList.remove('visible');
    setTimeout(() => {
      toast.remove();
    }, 300);
  }, 3500);
}

// Build and trigger the print page
function printHallTickets(onlySelected = false, specificStudents = null) {
  if (!currentInstitute) return;

  let studentsToPrint = [];
  if (specificStudents) {
    studentsToPrint = specificStudents;
  } else {
    const scopedList = getScopedStudents();
    if (onlySelected) {
      const selectedIds = Array.from(document.querySelectorAll('.student-checkbox:checked')).map(cb => cb.value);
      studentsToPrint = scopedList.filter(s => selectedIds.includes(String(s.sNo)) || selectedIds.includes(s.candidateId));
    } else {
      studentsToPrint = scopedList;
    }
  }

  if (studentsToPrint.length === 0) {
    alert('No students selected for printing.');
    return;
  }

  const printContainer = document.getElementById('printContainer');
  printContainer.innerHTML = '';

  studentsToPrint.forEach(student => {
    const photoUrl = student.photo || (student.sNo ? `photos/p${student.sNo}.jpg` : `photos/${(student.candidateId || '').toUpperCase()}.jpg`);
    const sigUrl = student.signature || (student.sNo ? `signatures/s${student.sNo}.jpg` : `signatures/${(student.candidateId || '').toUpperCase()}_sig.jpg`);

    const ticketDiv = document.createElement('div');
    ticketDiv.className = 'printable-hallticket';

    ticketDiv.innerHTML = `
      <div class="ticket-wrapper">
        <div class="ticket-header">
          <div class="header-logo-left">
            <img src="Emblem_of_Andhra_Pradesh.svg.webp" class="header-logo" alt="AP Govt Logo">
          </div>
          <div class="header-text-container">
            <h1>AP Tribal Welfare Residential Educational Institutions Society (Gurukulam)</h1>
            <h2>National Education Society for Tribal Students (NESTS) - Under the Ministry of Tribal Affairs , New Delhi</h2>
          </div>
          <div class="header-logo-right">
            <img src="aptwreis.jpg" class="header-logo" alt="Gurukulam Logo">
            <img src="EMRS-logo-2025.png" class="header-logo" alt="EMRS Logo">
          </div>
        </div>
        
        <div class="exam-title-banner">
          HALL TICKET FOR Language Proficiency for Regular Staff EMRS - 2026
        </div>
        
        <table class="details-table">
          <tbody>
            <!-- Row 1 -->
            <tr>
              <td class="col-width-22">
                <span class="label">Candidate ID:</span>
                <span class="value semibold">${student.candidateId || '&nbsp;'}</span>
              </td>
              <td class="col-width-22">
                <span class="label">Aadhar Number:</span>
                <span class="value">${student.aadharNumber || '&nbsp;'}</span>
              </td>
              <td class="col-width-31">
                <span class="label">Candidate Name:</span>
                <span class="value bold uppercase">${student.candidateName || '&nbsp;'}</span>
              </td>
              <td class="col-width-25 center-content bg-light">
                <span class="label block">Hall Ticket No :</span>
                <span class="value ticket-no-text">${(student.hallTicketNo || '').replace(/\s+/g, '') || '&nbsp;'}</span>
              </td>
            </tr>
            
            <!-- Row 2 -->
            <tr>
              <td>
                <span class="label">Designation:</span>
                <span class="value uppercase">${student.designation || '&nbsp;'}</span>
              </td>
              <td>
                <span class="label">Gender:</span>
                <span class="value">${student.gender || '&nbsp;'}</span>
              </td>
              <td>
                <span class="label">Mobile Number:</span>
                <span class="value">${student.mobileNumber || '&nbsp;'}</span>
              </td>
              <!-- Photo & Signature Cell (Row 2 Col 4, spans Row 2 & 3) -->
              <td class="photo-cell" rowspan="2">
                <div class="photo-box">
                  <div class="photo-wrapper">
                    <img src="${photoUrl}" class="student-photo" alt="Photo" onerror="handleImageError(this, 'photo')">
                    <div class="photo-placeholder" style="display:none;">Affix Passport<br>Size Photo</div>
                  </div>
                  <div class="signature-box" style="margin: 0 auto;">
                    <img src="${sigUrl}" class="student-sig" alt="Signature" onerror="handleImageError(this, 'signature')">
                    <div class="sig-placeholder" style="display:none;">Candidate Signature</div>
                  </div>
                </div>
              </td>
            </tr>
            
            <!-- Row 3 -->
            <tr>
              <td colspan="3">
                <span class="label">Candidate Address:</span>
                <span class="value address-text">${student.address || '&nbsp;'}</span>
              </td>
            </tr>
            
            <!-- Row 4 -->
            <tr class="exam-info-row">
              <td class="center-content">
                <span class="label block">Exam Date & Time</span>
                <span class="value bold exam-datetime-text">15-07-2026<br>10:00 AM to 12:00 PM</span>
              </td>
              <td colspan="2">
                <span class="label block text-left">Exam Center Address</span>
                <span class="value bold exam-center-text">${(student.examCenterAddress || '').replace(/\n/g, '<br>')}</span>
              </td>
              <td class="center-content">
                <span class="label block">Exam/Test</span>
                <span class="value bold exam-center-text" style="font-size: 10pt; line-height: 1.4;">Language Proficiency Test - Telugu</span>
              </td>
            </tr>
          </tbody>
        </table>
        
        <div class="instructions-section">
          <h3>Instructions:</h3>
          <ul>
            <li>Bring a Blue or Black ballpoint pen to the exam.</li>
            <li>Arrive at the exam center at least 30 minutes before the exam starts.</li>
            <li>Electronic gadgets are not allowed in the examination hall.</li>
          </ul>
        </div>
        
        <div class="signature-row" style="display: flex; justify-content: space-between; margin-top: 30px; padding: 10px 30px 20px 30px;">
          <div style="text-align: center; width: 180px;">
            <div style="height: 35px;"></div>
            <div style="border-top: 1px solid #000000; padding-top: 5px; font-size: 9.5pt; font-weight: bold;">Signature of the Candidate</div>
          </div>
          <div style="text-align: center; width: 180px;">
            <div style="height: 35px;"></div>
            <div style="border-top: 1px solid #000000; padding-top: 5px; font-size: 9.5pt; font-weight: bold;">Attestation of Principal(with seal)</div>
          </div>
        </div>
      </div>
    `;

    printContainer.appendChild(ticketDiv);
  });

  // Wait for all images inside printContainer to load before triggering print dialog
  const images = printContainer.querySelectorAll('img');
  const promises = Array.from(images).map(img => {
    return new Promise((resolve) => {
      if (img.complete) {
        resolve();
      } else {
        img.addEventListener('load', resolve);
        img.addEventListener('error', resolve); // Resolve on error too to prevent getting stuck
      }
    });
  });

  Promise.all(promises).then(() => {
    // Add a tiny delay to ensure the browser has rendered the loaded images in the layout
    setTimeout(() => {
      window.print();
    }, 150);
  });
}

// Global handler for missing photos and signatures to show a text placeholder
window.handleImageError = function(img, type) {
  img.onerror = null;
  img.style.display = 'none';
  const placeholder = img.nextElementSibling;
  if (placeholder) {
    placeholder.style.display = type === 'photo' ? 'flex' : 'block';
  }
};
