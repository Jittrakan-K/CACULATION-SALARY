/**
 * APP.JS - ระบบคำนวณเงินเดือน & ปฏิทินวันทำงาน
 * CACULATION SALARY
 */

// ==========================================================================
// 1. ฐานข้อมูลปฏิทินเริ่มต้นปี 2569 (2026)
// ==========================================================================
const DEFAULT_CALENDAR_2026 = {
  yearCE: 2026,
  yearBE: 2569,
  title: '2569 / 2026',
  nationalHolidays: [
    '2026-01-01', '2026-01-02', '2026-03-03', '2026-04-06', '2026-04-13',
    '2026-04-14', '2026-04-15', '2026-05-01', '2026-05-04', '2026-06-01',
    '2026-06-03', '2026-07-28', '2026-07-29', '2026-08-12', '2026-10-13',
    '2026-10-23', '2026-12-05', '2026-12-07', '2026-12-10', '2026-12-31'
  ],
  memorialHolidays: [
    '2026-09-19' // วันหยุดประเพณีจ่ายเงิน (Memorial day with pay)
  ],
  bridgeHolidays: [
    '2026-01-03' // วันหยุดพิเศษ
  ],
  orbrayWorkDays: [
    '2026-10-17' // วันทำงานพิเศษ
  ],
  workingSaturdays: [
    '2026-01-17', '2026-02-21', '2026-03-14', '2026-03-28',
    '2026-04-04', '2026-04-18', '2026-05-09', '2026-06-20',
    '2026-07-04', '2026-07-25', '2026-08-01', '2026-08-15', '2026-08-29',
    '2026-10-17', '2026-10-31', '2026-12-12'
  ],
  customDayOverrides: {}
};

// ==========================================================================
// 2. ข้อมูลตัวอย่างจากไฟล์ Excel (21/08/2026 - 20/09/2026)
// ==========================================================================
const EXCEL_SAMPLE_ATTENDANCE = [
  { date: '2026-08-21', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-08-22', inTime: '08:00', outTime: '19:20', ot: 0, hol: 8, otHol: 2, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-08-23', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 },
  { date: '2026-08-24', inTime: '08:00', outTime: '18:20', ot: 1, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-08-25', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-08-26', inTime: '08:00', outTime: '18:20', ot: 1, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-08-27', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-08-28', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-08-29', inTime: '08:00', outTime: '17:50', ot: 0.5, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-08-30', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 },
  { date: '2026-08-31', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-01', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-09-02', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-03', inTime: '08:00', outTime: '18:20', ot: 1, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-04', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-09-05', inTime: '08:00', outTime: '19:20', ot: 0, hol: 8, otHol: 2, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-06', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 },
  { date: '2026-09-07', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-08', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-09-09', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-10', inTime: '08:00', outTime: '18:20', ot: 1, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-11', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-09-12', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 },
  { date: '2026-09-13', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 },
  { date: '2026-09-14', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-15', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-09-16', inTime: '08:00', outTime: '19:20', ot: 2, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-17', inTime: '08:00', outTime: '18:20', ot: 1, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 1 },
  { date: '2026-09-18', inTime: '08:00', outTime: '17:00', ot: 0, hol: 0, otHol: 0, workDay: 1, come: 1, otDay: 0 },
  { date: '2026-09-19', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 },
  { date: '2026-09-20', inTime: '', outTime: '', ot: 0, hol: 0, otHol: 0, workDay: 0, come: 0, otDay: 0 }
];

// ==========================================================================
// 3. State Management (สถานะหลักของระบบ)
// ==========================================================================
let allCalendars = {};      // บันทึกปฏิทินทุกปี { '2026': {...}, '2027': {...} }
let activeYearCE = 2026;    // ปี ค.ศ. ปัจจุบันที่กำลังดู
let currentPeriod = null;   // งวดการจ่ายปัจจุบัน
let currentAttendance = []; // รายการลงเวลา 31 วันของงวดปัจจุบัน

// ข้อมูลพนักงานและข้อมูลคงที่
const salaryProfile = {
  companyName: 'CACULATION SALARY',
  empCode: '20523',
  empName: 'JITTRAKAN K.',
  department: 'PRODUCTION TECHNOLOGY',
  empType: 'รายเดือน',
  periodMonth: 'กันยายน',
  periodRound: 'งวดจ่ายเงินเดือน',
  payDate: '30/09/2569',
  userName: 'JITTRAKAN K.',
  printDate: '13/09/2569'
};

// โครงสร้างอัตราค่าเงินเริ่มต้นตามรูปที่ 2 (สามารถแก้ไขได้ตลอดเวลา)
const DEFAULT_SALARY_CONFIG = {
  baseSalary: 20000,
  diligenceFullAmount: 650,
  transportationAllowance: 1230,
  foodPerDay: 20,
  otMealPerDay: 15,
  otDivisorHours: 240,
  ot15Multiplier: 1.5,
  ot1Multiplier: 1.0,
  ot3Multiplier: 3.0,
  ssoDeduction: 875,
  ssoMaxBase: 17500
};

let currentSalaryConfig = { ...DEFAULT_SALARY_CONFIG };

// ==========================================================================
// 4. ระบบจัดเก็บข้อมูล LocalStorage (Storage Manager)
// ==========================================================================
// 4. ระบบจัดเก็บข้อมูล LocalStorage (Storage Manager Scoped by User)
// ==========================================================================
const STORAGE_KEY_CALENDARS = 'orbray_salary_calendars_v2';
const STORAGE_KEY_YEAR = 'orbray_salary_active_year_v2';
const STORAGE_KEY_ATTENDANCE = 'orbray_salary_attendance_store_v2';
const STORAGE_KEY_SALARY_CONFIG = 'salary_rates_config_v2';

function getScopedUserKey(baseKey) {
  const uid = (typeof getActiveEditingUserId === 'function') ? getActiveEditingUserId() : 'user_admin';
  return `${baseKey}_${uid}`;
}

function syncSalaryProfileWithActiveUser() {
  if (typeof getActiveEditingUser === 'function') {
    const u = getActiveEditingUser();
    if (u) {
      salaryProfile.companyName = (u.companyName || 'CACULATION SALARY').toUpperCase();
      salaryProfile.empCode = (u.empCode || (u.id === 'user_admin' ? '20523' : ('EMP-' + (u.pin || '001')))).toUpperCase();
      salaryProfile.empName = u.name;
      salaryProfile.department = u.department || 'ฝ่ายปฏิบัติการ';
      salaryProfile.userName = u.role === 'admin' ? 'ADMIN' : u.name;
    }
  }
}

function loadStorageData() {
  try {
    const rawCals = localStorage.getItem(STORAGE_KEY_CALENDARS);
    if (rawCals) {
      allCalendars = JSON.parse(rawCals);
    } else {
      allCalendars = { '2026': JSON.parse(JSON.stringify(DEFAULT_CALENDAR_2026)) };
    }

    const savedYear = localStorage.getItem(STORAGE_KEY_YEAR);
    if (savedYear && allCalendars[savedYear]) {
      activeYearCE = parseInt(savedYear, 10);
    } else {
      activeYearCE = 2026;
      if (!allCalendars['2026']) {
        allCalendars['2026'] = JSON.parse(JSON.stringify(DEFAULT_CALENDAR_2026));
      }
    }

    // โหลดการตั้งค่าโครงสร้างค่าเงินเฉพาะของ User ที่กำลัง Active
    const uid = (typeof getActiveEditingUserId === 'function') ? getActiveEditingUserId() : 'user_admin';
    const userCfgKey = getScopedUserKey(STORAGE_KEY_SALARY_CONFIG);
    const rawSalaryCfg = localStorage.getItem(userCfgKey);
    if (rawSalaryCfg) {
      currentSalaryConfig = Object.assign({}, DEFAULT_SALARY_CONFIG, JSON.parse(rawSalaryCfg));
    } else if (uid === 'user_admin') {
      const legacyCfg = localStorage.getItem(STORAGE_KEY_SALARY_CONFIG);
      currentSalaryConfig = legacyCfg ? Object.assign({}, DEFAULT_SALARY_CONFIG, JSON.parse(legacyCfg)) : { ...DEFAULT_SALARY_CONFIG };
    } else {
      currentSalaryConfig = { ...DEFAULT_SALARY_CONFIG };
    }

    // อัปเดตข้อมูลพนักงานในสลิปตาม Active User
    syncSalaryProfileWithActiveUser();
  } catch (e) {
    console.error('Storage load failed, using default', e);
    allCalendars = { '2026': JSON.parse(JSON.stringify(DEFAULT_CALENDAR_2026)) };
    activeYearCE = 2026;
    currentSalaryConfig = { ...DEFAULT_SALARY_CONFIG };
  }
}

function saveCalendarsToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY_CALENDARS, JSON.stringify(allCalendars));
    localStorage.setItem(STORAGE_KEY_YEAR, activeYearCE.toString());
  } catch (e) {
    console.warn('Storage save failed', e);
  }
}

function saveAttendanceToStorage(periodId, rows) {
  try {
    const storeKey = getScopedUserKey(STORAGE_KEY_ATTENDANCE);
    let store = {};
    const raw = localStorage.getItem(storeKey);
    if (raw) store = JSON.parse(raw);
    store[periodId] = rows;
    localStorage.setItem(storeKey, JSON.stringify(store));

    // ซิงค์ข้อมูลลง Cloud
    if (typeof syncDataToCloud === 'function') {
      syncDataToCloud('attendance_' + periodId, rows);
    }
  } catch (e) {
    console.warn('Attendance save failed', e);
  }
}

function loadAttendanceFromStorage(periodId) {
  try {
    const uid = (typeof getActiveEditingUserId === 'function') ? getActiveEditingUserId() : 'user_admin';
    const storeKey = getScopedUserKey(STORAGE_KEY_ATTENDANCE);
    let raw = localStorage.getItem(storeKey);
    // เฉพาะ user_admin เท่านั้นที่อนุญาตให้อ่าน legacy fallback ถ้ายังไม่มี scoped key
    if (!raw && uid === 'user_admin') {
      raw = localStorage.getItem(STORAGE_KEY_ATTENDANCE);
    }
    if (raw) {
      const store = JSON.parse(raw);
      if (store[periodId] && Array.isArray(store[periodId])) {
        return store[periodId];
      }
    }
  } catch (e) {
    console.warn('Attendance load failed', e);
  }
  return null;
}

// ==========================================================================
// 5. ระบบคำนวณวันและรอบการจ่ายเงินเดือน (Calendar & Periods Generator)
// ==========================================================================
const THAI_MONTHS = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

const ENGLISH_MONTHS = [
  'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
  'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'
];

function getLastDayOfMonth(year, monthIndex) {
  const d = new Date(year, monthIndex + 1, 0);
  return d.getDate();
}

function formatDay2Digit(n) {
  return n < 10 ? '0' + n : '' + n;
}

function generatePayrollPeriodsForYear(yearCE) {
  const yearBE = yearCE + 543;
  const periods = [];

  for (let m = 0; m < 12; m++) {
    const monthNum = m + 1;
    const monthName = `${THAI_MONTHS[m]} ${yearBE}`;
    const periodId = `${yearCE}-${formatDay2Digit(monthNum)}`;

    let startYear = yearCE;
    let startMonth = m - 1;
    let startYearBE = yearBE;
    if (startMonth < 0) {
      startMonth = 11;
      startYear = yearCE - 1;
      startYearBE = yearBE - 1;
    }
    const startMonthName = THAI_MONTHS[startMonth];
    const endMonthName = THAI_MONTHS[m];
    const startMonthNameEN = ENGLISH_MONTHS[startMonth];
    const endMonthNameEN = ENGLISH_MONTHS[m];
    const startDate = `${startYear}-${formatDay2Digit(startMonth + 1)}-21`;
    const endDate = `${yearCE}-${formatDay2Digit(monthNum)}-20`;

    const lastDay = getLastDayOfMonth(yearCE, m);
    const payDate = `${formatDay2Digit(lastDay)}/${formatDay2Digit(monthNum)}/${yearBE}`;
    const displayName = `${endMonthNameEN} ${yearCE} (เงินจ่ายสิ้นเดือน)`;

    periods.push({
      id: periodId,
      yearCE,
      yearBE,
      monthIndex: m,
      monthName,
      startMonthName,
      endMonthName,
      startMonthNameEN,
      endMonthNameEN,
      displayName,
      startYear,
      startYearBE,
      startDate,
      endDate,
      payDate
    });
  }
  return periods;
}

function getOrbrayDayInfo(dateStr) {
  const yearCE = parseInt(dateStr.split('-')[0], 10);
  const cal = allCalendars[yearCE] || allCalendars[activeYearCE] || DEFAULT_CALENDAR_2026;

  if (cal.customDayOverrides && cal.customDayOverrides[dateStr]) {
    const ov = cal.customDayOverrides[dateStr];
    return getDayTypeConfig(ov.type, ov.name);
  }

  if (cal.nationalHolidays && cal.nationalHolidays.includes(dateStr)) {
    return getDayTypeConfig('NATIONAL_HOLIDAY');
  }
  if (cal.memorialHolidays && cal.memorialHolidays.includes(dateStr)) {
    return getDayTypeConfig('MEMORIAL_HOLIDAY');
  }
  if (cal.bridgeHolidays && cal.bridgeHolidays.includes(dateStr)) {
    return getDayTypeConfig('BRIDGE_HOLIDAY');
  }
  if (cal.orbrayWorkDays && cal.orbrayWorkDays.includes(dateStr)) {
    return getDayTypeConfig('ORBRAY_WORKDAY');
  }

  const d = new Date(dateStr);
  const dayOfWeek = d.getDay();

  if (dayOfWeek === 0) {
    return getDayTypeConfig('SUNDAY');
  }

  if (dayOfWeek === 6) {
    if (cal.workingSaturdays && cal.workingSaturdays.includes(dateStr)) {
      return getDayTypeConfig('WORKING_SATURDAY');
    } else {
      return getDayTypeConfig('SATURDAY_HOLIDAY');
    }
  }

  return getDayTypeConfig('NORMAL_WORKDAY');
}

function getDayTypeConfig(type, customName = null) {
  switch (type) {
    case 'NATIONAL_HOLIDAY':
      return { type, name: customName || 'วันหยุดนักขัตฤกษ์', isWorkDay: false, badgeClass: 'badge-national', calClass: 'day-national-holiday' };
    case 'MEMORIAL_HOLIDAY':
      return { type, name: customName || 'วันหยุดประเพณีจ่ายเงิน', isWorkDay: false, badgeClass: 'badge-memorial', calClass: 'day-memorial-holiday' };
    case 'ORBRAY_WORKDAY':
      return { type, name: customName || 'วันทำงานพิเศษ', isWorkDay: true, badgeClass: 'badge-orbray', calClass: 'day-orbray-work' };
    case 'WORKING_SATURDAY':
      return { type, name: customName || 'เสาร์ทำงาน', isWorkDay: true, badgeClass: 'badge-work-sat', calClass: 'day-working-sat' };
    case 'SATURDAY_HOLIDAY':
      return { type, name: customName || 'เสาร์หยุด', isWorkDay: false, badgeClass: 'badge-sat-holiday', calClass: 'day-sat-holiday' };
    case 'SUNDAY':
      return { type, name: customName || 'วันอาทิตย์หยุด', isWorkDay: false, badgeClass: 'badge-sunday', calClass: 'day-sunday' };
    case 'BRIDGE_HOLIDAY':
      return { type, name: customName || 'วันหยุดพิเศษ', isWorkDay: false, badgeClass: 'badge-bridge', calClass: 'day-bridge' };
    default:
      return { type: 'NORMAL_WORKDAY', name: customName || 'วันทำงานปกติ', isWorkDay: true, badgeClass: 'badge-workday', calClass: '' };
  }
}

// ==========================================================================
// 6. การคำนวณ OT จากเวลาออกงาน
// ==========================================================================
function getOTFromTimeOut(timeOutStr) {
  if (!timeOutStr) return 0;
  const t = timeOutStr.trim();
  switch (t) {
    case '17:00': return 0;
    case '17:30': return 0.5;
    case '17:50': return 0.5;
    case '18:00': return 1.0;
    case '18:20': return 1.0;
    case '18:30': return 1.0;
    case '18:50': return 1.5;
    case '19:00': return 1.5;
    case '19:20': return 2.0;
    case '19:50': return 2.5;
    default:
      const parts = t.split(':');
      if (parts.length === 2) {
        const h = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        const totalMinutes = (h * 60 + m) - (17 * 60);
        if (totalMinutes <= 0) return 0;
        return Math.floor((totalMinutes / 60) * 2) / 2;
      }
      return 0;
  }
}

function updateRowOT(row) {
  if (!row.workDay || !row.outTime) {
    row.ot = 0;
    row.hol = 0;
    row.otHol = 0;
    row.otDay = 0;
    return;
  }
  const dayInfo = getOrbrayDayInfo(row.date);

  if (dayInfo.isWorkDay) {
    row.ot = getOTFromTimeOut(row.outTime);
    row.hol = 0;
    row.otHol = 0;
  } else {
    row.ot = 0;
    const parts = row.outTime.split(':');
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const totalMinutes = h * 60 + m;

    const normalEndMinutes = 17 * 60;
    const startMinutes = 8 * 60;
    const lunchStart = 12 * 60;
    const lunchEnd = 13 * 60;

    if (totalMinutes <= normalEndMinutes) {
      let workMins = 0;
      if (totalMinutes <= lunchStart) {
        workMins = totalMinutes - startMinutes;
      } else if (totalMinutes <= lunchEnd) {
        workMins = lunchStart - startMinutes;
      } else {
        workMins = (totalMinutes - startMinutes) - 60;
      }
      if (workMins < 0) workMins = 0;
      row.hol = Math.floor((workMins / 60) * 2) / 2;
      row.otHol = 0;
    } else {
      row.hol = 8.0;
      row.otHol = getOTFromTimeOut(row.outTime);
    }
  }

  row.otDay = (row.ot >= 1 || row.hol >= 1 || row.otHol >= 1) ? 1 : 0;
}

// ==========================================================================
// 7. การสร้างและจัดการตารางเข้า-ออกงาน (Timesheet Table)
// ==========================================================================
function buildAttendanceTable() {
  const tbody = document.getElementById('attendanceTbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const thaiDays = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];

  const sundayOutOptions = [
    '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', '12:00',
    '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00',
    '16:30', '17:00', '17:30', '17:50', '18:00', '18:20', '18:30', '18:50',
    '19:00', '19:20'
  ];

  const normalOutOptions = [
    '17:00', '17:30', '17:50', '18:20', '18:50', '19:20', '19:50'
  ];

  currentAttendance.forEach((row, idx) => {
    const d = new Date(row.date);
    const dayName = thaiDays[d.getDay()];
    const dayInfo = getOrbrayDayInfo(row.date);

    const tr = document.createElement('tr');
    if (!dayInfo.isWorkDay) {
      tr.classList.add('row-holiday');
    }

    const isSundayOrHoliday = !dayInfo.isWorkDay;
    const availableOptions = isSundayOrHoliday ? sundayOutOptions : normalOutOptions;

    let optionsHtml = `<option value="" ${!row.outTime ? 'selected' : ''}>-</option>`;
    availableOptions.forEach(opt => {
      const isSelected = (row.outTime === opt || (opt === '18:20' && row.outTime === '18:00'));
      optionsHtml += `<option value="${opt}" ${isSelected ? 'selected' : ''}>${opt}</option>`;
    });

    tr.innerHTML = `
      <td class="text-center font-monospace">${row.date.split('-').slice(1).join('/')}</td>
      <td class="text-center font-weight-bold">${dayName}</td>
      <td class="text-center"><span class="day-badge ${dayInfo.badgeClass}">${dayInfo.name}</span></td>
      <td class="text-center font-monospace ${row.workDay ? 'text-primary font-weight-bold' : 'text-muted'}">
        <span class="time-in-val">${row.workDay ? '08:00' : '-'}</span>
      </td>
      <td>
        <select class="form-control form-control-sm form-select text-center" 
          onchange="onAttendanceTimeChange(${idx}, 'out', this.value)" 
          ${!row.workDay ? 'disabled' : ''}>
          ${optionsHtml}
        </select>
      </td>
      <td class="text-right font-monospace ${row.workDay && row.ot > 0 ? 'text-primary font-weight-bold' : 'text-muted'}">
        ${row.workDay && row.ot ? row.ot.toFixed(1) : '0'}
      </td>
      <td class="text-right font-monospace ${row.workDay && row.hol > 0 ? 'text-primary font-weight-bold' : 'text-muted'}">
        ${row.workDay && row.hol ? row.hol.toFixed(1) : '0'}
      </td>
      <td class="text-right font-monospace ${row.workDay && row.otHol > 0 ? 'text-primary font-weight-bold' : 'text-muted'}">
        ${row.workDay && row.otHol ? row.otHol.toFixed(1) : '0'}
      </td>
      <td class="text-center">
        <input type="checkbox" ${row.workDay ? 'checked' : ''} onchange="onAttendanceToggleWorkDay(${idx}, this.checked)">
      </td>
      <td class="text-center">
        <span class="${row.come ? 'text-success font-weight-bold' : 'text-muted'}">${row.come ? '1' : '0'}</span>
      </td>
      <td class="text-center">
        <span class="${row.otDay ? 'text-primary font-weight-bold' : 'text-muted'}">${row.otDay ? '1' : '0'}</span>
      </td>
    `;
    tbody.appendChild(tr);
  });

  recalculateAttendanceTotals();
}

function onAttendanceTimeChange(index, field, value) {
  const row = currentAttendance[index];
  if (field === 'out') {
    row.outTime = value ? value.trim() : '';
  }
  updateRowOT(row);
  buildAttendanceTable();
  recalculateSalary();
  saveCurrentAttendance();
}

function onAttendanceToggleWorkDay(index, isChecked) {
  const row = currentAttendance[index];
  row.workDay = isChecked ? 1 : 0;
  row.come = isChecked ? 1 : 0;

  if (isChecked) {
    row.inTime = '08:00';
    if (!row.outTime) {
      row.outTime = '17:00';
    }
    updateRowOT(row);
  } else {
    row.inTime = '';
    row.outTime = '';
    row.ot = 0;
    row.hol = 0;
    row.otHol = 0;
    row.otDay = 0;
  }

  buildAttendanceTable();
  recalculateSalary();
  saveCurrentAttendance();
}

function recalculateAttendanceTotals() {
  let totalOT15 = 0;
  let totalHol = 0;
  let totalOTHol = 0;
  let totalTargetDays = 0;
  let totalCameDays = 0;
  let totalOTDays = 0;

  currentAttendance.forEach(row => {
    totalOT15 += parseFloat(row.ot) || 0;
    totalHol += parseFloat(row.hol) || 0;
    totalOTHol += parseFloat(row.otHol) || 0;
    totalTargetDays += parseInt(row.workDay, 10) || 0;
    totalCameDays += parseInt(row.come, 10) || 0;
    totalOTDays += parseInt(row.otDay, 10) || 0;
  });

  const elOT15 = document.getElementById('attTotalOT15');
  const elHol = document.getElementById('attTotalHol');
  const elOTHol = document.getElementById('attTotalOTHol');
  const elTarget = document.getElementById('attTotalTargetDays');
  const elCame = document.getElementById('attTotalCameDays');
  const elOTDays = document.getElementById('attTotalOTDays');

  if (elOT15) elOT15.innerText = totalOT15.toFixed(1);
  if (elHol) elHol.innerText = totalHol.toFixed(1);
  if (elOTHol) elOTHol.innerText = totalOTHol.toFixed(1);
  if (elTarget) elTarget.innerText = totalTargetDays;
  if (elCame) elCame.innerText = totalCameDays;
  if (elOTDays) elOTDays.innerText = totalOTDays;

  return { totalOT15, totalHol, totalOTHol, totalTargetDays, totalCameDays, totalOTDays };
}

function autoFillNormalWorkdays() {
  currentAttendance.forEach(row => {
    const dayInfo = getOrbrayDayInfo(row.date);
    if (dayInfo.isWorkDay) {
      row.workDay = 1;
      row.come = 1;
      row.inTime = '08:00';
      row.outTime = '17:00';
      row.ot = 0;
      row.hol = 0;
      row.otHol = 0;
      row.otDay = 0;
    } else {
      row.workDay = 0;
      row.come = 0;
      row.inTime = '';
      row.outTime = '';
      row.ot = 0;
      row.hol = 0;
      row.otHol = 0;
      row.otDay = 0;
    }
  });
  buildAttendanceTable();
  recalculateSalary();
  saveCurrentAttendance();
}

function clearAttendanceTable() {
  if (!confirm('คุณต้องการล้างข้อมูลเวลาทั้งหมดในตารางใช่หรือไม่?')) return;
  currentAttendance.forEach(row => {
    row.workDay = 0;
    row.come = 0;
    row.inTime = '';
    row.outTime = '';
    row.ot = 0;
    row.hol = 0;
    row.otHol = 0;
    row.otDay = 0;
  });
  buildAttendanceTable();
  recalculateSalary();
  saveCurrentAttendance();
}

function generateAttendanceForPeriod(startStr, endStr) {
  const rows = [];
  let cur = new Date(startStr);
  const end = new Date(endStr);

  while (cur <= end) {
    const dStr = cur.toISOString().split('T')[0];
    const dayInfo = getOrbrayDayInfo(dStr);
    rows.push({
      date: dStr,
      inTime: dayInfo.isWorkDay ? '08:00' : '',
      outTime: dayInfo.isWorkDay ? '17:00' : '',
      ot: 0,
      hol: 0,
      otHol: 0,
      workDay: dayInfo.isWorkDay ? 1 : 0,
      come: dayInfo.isWorkDay ? 1 : 0,
      otDay: 0
    });
    cur.setDate(cur.getDate() + 1);
  }
  currentAttendance = rows;
}

function saveCurrentAttendance() {
  if (currentPeriod && currentAttendance) {
    saveAttendanceToStorage(currentPeriod.id, currentAttendance);
  }
}

function loadAttendanceFromExcel() {
  currentAttendance = JSON.parse(JSON.stringify(EXCEL_SAMPLE_ATTENDANCE));
  salaryProfile.periodMonth = 'กันยายน';
  salaryProfile.payDate = '30/09/2569';
  buildAttendanceTable();
  recalculateSalary();
  saveCurrentAttendance();
  alert('โหลดข้อมูลลงเวลา 31 วันจากไฟล์ Excel DATA SALARY 210826-200926.xlsx สำเร็จแล้ว!');
}

// ==========================================================================
// 8. ระบบคำนวณเงินเดือนตามสูตร Excel (Salary Calculation Engine)
// ==========================================================================
function recalculateSalary() {
  const totals = recalculateAttendanceTotals();

  const B2_targetDays = totals.totalTargetDays;
  const D2_cameDays = totals.totalCameDays;
  const F2_otDays = totals.totalOTDays;
  const G2_ot15Hours = totals.totalOT15;
  const H2_ot1Hours = totals.totalHol;
  const I2_ot3Hours = totals.totalOTHol;

  const baseSalary = typeof currentSalaryConfig.baseSalary !== 'undefined' ? Number(currentSalaryConfig.baseSalary) : DEFAULT_SALARY_CONFIG.baseSalary;
  const diligenceFullAmount = Number(currentSalaryConfig.diligenceFullAmount) || 0;
  const transportationAllowance = Number(currentSalaryConfig.transportationAllowance) || 0;
  const foodPerDay = Number(currentSalaryConfig.foodPerDay) || 0;
  const otMealPerDay = Number(currentSalaryConfig.otMealPerDay) || 0;
  const otDivisor = Number(currentSalaryConfig.otDivisorHours) || 240;
  const ot15Multiplier = Number(currentSalaryConfig.ot15Multiplier) || 1.5;
  const ot1Multiplier = Number(currentSalaryConfig.ot1Multiplier) || 1.0;
  const ot3Multiplier = Number(currentSalaryConfig.ot3Multiplier) || 3.0;
  const ssoDeduction = typeof currentSalaryConfig.ssoDeduction !== 'undefined' ? Number(currentSalaryConfig.ssoDeduction) : 0;

  const C2_dailyRate = B2_targetDays > 0 ? (baseSalary / B2_targetDays) : 0;
  const E2_actualSalary = (B2_targetDays > 0 && D2_cameDays >= B2_targetDays)
    ? baseSalary
    : Math.round(C2_dailyRate * D2_cameDays);

  const J2_foodAllowance = foodPerDay * D2_cameDays;
  const K2_otMealAllowance = otMealPerDay * F2_otDays;
  const L2_travelAllowance = transportationAllowance;
  const M2_diligenceAllowance = (B2_targetDays > 0 && D2_cameDays >= B2_targetDays) ? diligenceFullAmount : 0;

  const hourlyOTRate = otDivisor > 0 ? (baseSalary / otDivisor) : 0;
  const N2_ot15Amount = Math.round(hourlyOTRate * ot15Multiplier * G2_ot15Hours);
  const O2_ot1Amount = Math.round(hourlyOTRate * ot1Multiplier * H2_ot1Hours);
  const P2_ot3Amount = Math.round(hourlyOTRate * ot3Multiplier * I2_ot3Hours);

  // ยอดหักประกันสังคม: ถ้าไม่มีรายได้เลยและไม่มีฐานเงินเดือน (เช่น ยังไม่ได้เริ่มงาน) ให้ยอดหักเป็น 0
  const totalEarnings = E2_actualSalary + J2_foodAllowance + K2_otMealAllowance + L2_travelAllowance +
                        M2_diligenceAllowance + N2_ot15Amount + O2_ot1Amount + P2_ot3Amount;
  const Q2_sso = (totalEarnings === 0 && baseSalary === 0) ? 0 : ssoDeduction;
  const totalDeductions = Q2_sso;
  const netPay = Math.max(0, totalEarnings - totalDeductions);

  // อัปเดตการ์ด Dashboard ด้านบน
  const elTotalIncome = document.getElementById('dashTotalIncome');
  const elTotalDeduct = document.getElementById('dashTotalDeduct');
  const elNetPay = document.getElementById('dashNetPay');
  if (elTotalIncome) elTotalIncome.innerText = formatCurrency(totalEarnings);
  if (elTotalDeduct) elTotalDeduct.innerText = formatCurrency(totalDeductions);
  if (elNetPay) elNetPay.innerText = formatCurrency(netPay);

  // อัปเดต Hero Header สถิติสำคัญ
  setElText('heroTargetDays', B2_targetDays);
  setElText('heroNetPay', formatCurrency(netPay));
  const totalOTHours = (G2_ot15Hours + H2_ot1Hours + I2_ot3Hours).toFixed(1);
  setElText('heroTotalOTHours', totalOTHours);

  // อัปเดตการ์ดภาพรวมบันทึกเวลาทำงานบนหน้าคำนวณเงินเดือน (Quick Overview)
  setElText('quickCameDays', D2_cameDays);
  setElText('quickTargetDays', B2_targetDays);
  setElText('quickTotalOTHours', totalOTHours);
  setElText('quickTotalOTDays', F2_otDays);

  // อัปเดตแถบ KPI บนหน้าบันทึกเวลาออกงาน (Attendance Page)
  setElText('attCardTargetDays', `${B2_targetDays} วัน`);
  setElText('attCardCameDays', `${D2_cameDays} วัน`);
  setElText('attCardOTHours', `${totalOTHours} ชม.`);
  setElText('attCardNetPay', `${formatCurrency(netPay)} บาท`);

  // อัปเดตรายการคำนวณเงินเดือนในการ์ด (รูปที่ 2)
  setElText('calcBaseSalary', formatCurrency(baseSalary));
  setElText('calcTargetDays', B2_targetDays);
  setElText('calcDailyRate', formatCurrency(C2_dailyRate));
  setElText('calcCameDays', D2_cameDays);
  setElText('calcActualSalary', formatCurrency(E2_actualSalary));

  setElText('calcFoodDays', D2_cameDays);
  setElText('calcFoodTotal', formatCurrency(J2_foodAllowance));

  setElText('calcOTMealDays', F2_otDays);
  setElText('calcOTMealTotal', formatCurrency(K2_otMealAllowance));

  setElText('calcTravelTotal', formatCurrency(L2_travelAllowance));

  const diligenceStatusEl = document.getElementById('calcDiligenceStatus');
  if (diligenceStatusEl) {
    if (B2_targetDays > 0 && D2_cameDays >= B2_targetDays) {
      diligenceStatusEl.innerText = 'ครบวันทำงาน';
      diligenceStatusEl.className = 'badge-status-ok';
    } else {
      diligenceStatusEl.innerText = 'ขาด/ไม่ครบวันทำงาน';
      diligenceStatusEl.className = 'day-badge badge-sunday';
    }
  }
  setElText('calcDiligenceTotal', formatCurrency(M2_diligenceAllowance));

  setElText('calcHourlyOTRate', formatCurrency(hourlyOTRate));
  setElText('calcOT15Hours', G2_ot15Hours.toFixed(1));
  setElText('calcOT15Total', formatCurrency(N2_ot15Amount));

  setElText('calcOT1Hours', H2_ot1Hours.toFixed(1));
  setElText('calcOT1Total', formatCurrency(O2_ot1Amount));

  setElText('calcOT3Hours', I2_ot3Hours.toFixed(1));
  setElText('calcOT3Total', formatCurrency(P2_ot3Amount));

  setElText('calcSSOTotal', Q2_sso > 0 ? `-${formatCurrency(Q2_sso)}` : '0.00');
  setElText('calcSumIncome', formatCurrency(totalEarnings));
  setElText('calcSumDeduct', totalDeductions > 0 ? `-${formatCurrency(totalDeductions)}` : '0.00');
  setElText('calcGrandNetPay', formatCurrency(netPay));
  setElText('cardPayDateDisplay', salaryProfile.payDate || '-');

  // ปรับปรุงข้อความระบุอัตราบนหน้าจอให้ตรงกับการตั้งค่าปัจจุบัน
  updateRateLabelsOnCards();

  updatePayslip({
    actualSalary: E2_actualSalary,
    diligence: M2_diligenceAllowance,
    travel: L2_travelAllowance,
    food: J2_foodAllowance,
    otMeal: K2_otMealAllowance,
    ot15: N2_ot15Amount,
    ot1: O2_ot1Amount,
    ot3: P2_ot3Amount,
    totalEarnings,
    totalDeductions,
    netPay
  });
}

function updateRateLabelsOnCards() {
  setElText('lblRateDiligence', formatCurrency(currentSalaryConfig.diligenceFullAmount));
  setElText('lblFoodTitle', `ค่าอาหาร (${currentSalaryConfig.foodPerDay} บ./วัน)`);
  setElText('lblFoodPerDay', currentSalaryConfig.foodPerDay);
  setElText('lblOTMealTitle', `อาหารโอที (${currentSalaryConfig.otMealPerDay} บ./วัน)`);
  setElText('lblOTMealPerDay', currentSalaryConfig.otMealPerDay);
  setElText('calcTravelTotal', formatCurrency(currentSalaryConfig.transportationAllowance));
  
  const baseFormatted = formatCurrency(currentSalaryConfig.baseSalary);
  const div = currentSalaryConfig.otDivisorHours;
  setElText('lblOT15Title', `ค่าล่วงเวลา OT ${currentSalaryConfig.ot15Multiplier} เท่า (วันปกติ)`);
  setElText('lblOT15Mult', currentSalaryConfig.ot15Multiplier);
  setElText('lblOT15Base', baseFormatted);
  setElText('lblOT15Divisor', div);

  setElText('lblOT1Title', `ค่าล่วงเวลา OT ${currentSalaryConfig.ot1Multiplier} เท่า (วันหยุด 8 ชม.)`);
  setElText('lblOT1Mult', currentSalaryConfig.ot1Multiplier);
  setElText('lblOT1Base', baseFormatted);
  setElText('lblOT1Divisor', div);

  setElText('lblOT3Title', `ค่าล่วงเวลา OT ${currentSalaryConfig.ot3Multiplier} เท่า (วันหยุดหลัง 17:00)`);
  setElText('lblOT3Mult', currentSalaryConfig.ot3Multiplier);
  setElText('lblOT3Base', baseFormatted);
  setElText('lblOT3Divisor', div);

  setElText('lblSSOMaxBase', formatCurrency(currentSalaryConfig.ssoMaxBase));
}

function updatePayslip(calc) {
  setElText('slipPeriodMonth', salaryProfile.periodMonth);
  setElText('slipPayDate', salaryProfile.payDate);
  setElText('slipCompanyName', salaryProfile.companyName || 'CACULATION SALARY');
  setElText('slipEmpCode', salaryProfile.empCode || '-');
  setElText('slipEmpName', salaryProfile.empName || '-');
  setElText('slipDepartment', salaryProfile.department || '-');

  const tbodyIncome = document.getElementById('slipIncomeRows');
  if (tbodyIncome) {
    const items = [
      { name: 'เงินเดือน', amount: calc.actualSalary },
      { name: 'เบี้ยขยัน', amount: calc.diligence },
      { name: 'เงินช่วยเหลือค่าเดินทาง', amount: calc.travel },
      { name: 'ค่าอาหาร', amount: calc.food },
      { name: 'อาหารโอที', amount: calc.otMeal },
      { name: 'OT*1.5', amount: calc.ot15 },
      { name: 'OT*1', amount: calc.ot1 },
      { name: 'OT*3', amount: calc.ot3 }
    ];

    tbodyIncome.innerHTML = '';
    items.forEach(item => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${item.name}</td><td class="text-right">${formatCurrency(item.amount)}</td>`;
      tbodyIncome.appendChild(tr);
    });
  }

  const tbodyDeduct = document.getElementById('slipDeductRows');
  if (tbodyDeduct) {
    tbodyDeduct.innerHTML = `<tr><td>ประกันสังคม</td><td class="text-right">${formatCurrency(calc.totalDeductions)}</td></tr>`;
  }

  setElText('slipTotalIncome', formatCurrency(calc.totalEarnings));
  setElText('slipTotalDeduct', formatCurrency(calc.totalDeductions));
  setElText('slipNetPay', formatCurrency(calc.netPay));
}

function setElText(id, text) {
  const el = document.getElementById(id);
  if (el) el.innerText = text;
}

function formatCurrency(val) {
  const num = parseFloat(val);
  if (isNaN(num)) return '0.00';
  return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ==========================================================================
// 9. ระบบจัดการปฏิทินแบบเพิ่มปี (Dynamic Multi-Year Calendar UI)
// ==========================================================================
function renderOrbrayCalendarGrid() {
  const grid = document.getElementById('orbrayCalendarGrid');
  if (!grid) return;
  grid.innerHTML = '';

  const cal = allCalendars[activeYearCE] || allCalendars['2026'];
  const yearCE = cal.yearCE;
  const yearBE = cal.yearBE;

  setElText('calHeaderYear', `${yearBE} / ${yearCE}`);

  const deleteBtn = document.getElementById('btnDeleteCurrentYear');
  if (deleteBtn) {
    deleteBtn.style.display = (yearCE === 2026 && Object.keys(allCalendars).length <= 1) ? 'none' : 'inline-flex';
  }

  const thaiDayHeaders = ['S', 'M', 'T', 'W', 'TH', 'F', 'S'];

  for (let m = 0; m < 12; m++) {
    const monthCard = document.createElement('div');
    monthCard.className = 'cal-month-card';

    const lastDay = getLastDayOfMonth(yearCE, m);
    const firstDayDate = new Date(yearCE, m, 1);
    const firstDayOfWeek = firstDayDate.getDay();

    const monthNum = m + 1;
    const periodId = `${yearCE}-${formatDay2Digit(monthNum)}`;

    let html = `
      <div class="cal-month-header">
        <div>
          <h3>${m + 1}. ${THAI_MONTHS[m]}</h3>
          <span class="th-name">${yearBE}</span>
        </div>
        <button type="button" class="btn btn-sm btn-outline-primary" onclick="selectPeriodMonth('${periodId}')" title="เลือกงวดนี้">
          คำนวณงวดนี้
        </button>
      </div>
      <table class="cal-mini-table">
        <thead>
          <tr>
            ${thaiDayHeaders.map((dh, i) => `<th class="${i === 0 ? 'col-sun' : ''}">${dh}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
    `;

    let currentDay = 1;
    for (let r = 0; r < 6; r++) {
      if (currentDay > lastDay) break;
      html += '<tr>';
      for (let c = 0; c < 7; c++) {
        if (r === 0 && c < firstDayOfWeek) {
          html += '<td class="cal-day-cell"></td>';
        } else if (currentDay > lastDay) {
          html += '<td class="cal-day-cell"></td>';
        } else {
          const dateStr = `${yearCE}-${formatDay2Digit(m + 1)}-${formatDay2Digit(currentDay)}`;
          const dayInfo = getOrbrayDayInfo(dateStr);
          const customClass = dayInfo.calClass || '';

          html += `
            <td class="cal-day-cell">
              <span class="cal-day-num ${customClass}" 
                title="${dayInfo.name} (${dateStr})"
                onclick="openEditDayModal('${dateStr}')">
                ${currentDay}
              </span>
            </td>
          `;
          currentDay++;
        }
      }
      html += '</tr>';
    }

    html += '</tbody></table>';
    monthCard.innerHTML = html;
    grid.appendChild(monthCard);
  }
}

function updateYearSelectDropdown() {
  const select = document.getElementById('appYearSelect');
  if (!select) return;
  select.innerHTML = '';

  const years = Object.keys(allCalendars).map(Number).sort((a, b) => a - b);
  years.forEach(y => {
    const opt = document.createElement('option');
    opt.value = y;
    const cal = allCalendars[y];
    opt.innerText = `พ.ศ. ${cal.yearBE} (${cal.yearCE})`;
    if (y === activeYearCE) opt.selected = true;
    select.appendChild(opt);
  });
}

function onAppYearChange() {
  const select = document.getElementById('appYearSelect');
  if (!select) return;
  activeYearCE = parseInt(select.value, 10);
  saveCalendarsToStorage();

  refreshPeriodSelector();
  renderOrbrayCalendarGrid();
}

function refreshPeriodSelector() {
  const select1 = document.getElementById('payrollPeriodSelect');
  const select2 = document.getElementById('payrollPeriodSelectAtt');
  const selects = [select1, select2].filter(Boolean);
  if (selects.length === 0) return;

  const periods = generatePayrollPeriodsForYear(activeYearCE);
  selects.forEach(select => {
    select.innerHTML = '';
    periods.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.id;
      opt.innerText = p.displayName;
      select.appendChild(opt);
    });
  });

  const currentMonthNum = new Date().getMonth() + 1;
  const matchPeriod = periods.find(p => p.monthIndex + 1 === currentMonthNum) || periods[0];
  selects.forEach(select => {
    select.value = matchPeriod.id;
  });

  onPayrollPeriodSelectChange();
}

function onPayrollPeriodSelectAttChange() {
  const selectAtt = document.getElementById('payrollPeriodSelectAtt');
  const selectMain = document.getElementById('payrollPeriodSelect');
  if (selectAtt && selectMain) {
    selectMain.value = selectAtt.value;
  }
  onPayrollPeriodSelectChange();
}

function onPayrollPeriodSelectChange() {
  const select = document.getElementById('payrollPeriodSelect');
  const selectAtt = document.getElementById('payrollPeriodSelectAtt');
  if (!select && !selectAtt) return;

  const activeSelect = select || selectAtt;
  const periodVal = activeSelect.value;
  if (select && select.value !== periodVal) select.value = periodVal;
  if (selectAtt && selectAtt.value !== periodVal) selectAtt.value = periodVal;

  const periods = generatePayrollPeriodsForYear(activeYearCE);
  const found = periods.find(p => p.id === periodVal);
  if (!found) return;

  currentPeriod = found;
  salaryProfile.periodMonth = `${found.endMonthNameEN} ${found.yearCE}`;
  salaryProfile.payDate = found.payDate;

  const uid = (typeof getActiveEditingUserId === 'function') ? getActiveEditingUserId() : 'user_admin';
  const savedAtt = loadAttendanceFromStorage(found.id);
  if (savedAtt) {
    currentAttendance = savedAtt;
  } else if (found.id === '2026-09' && uid === 'user_admin') {
    currentAttendance = JSON.parse(JSON.stringify(EXCEL_SAMPLE_ATTENDANCE));
  } else {
    generateAttendanceForPeriod(found.startDate, found.endDate);
    autoFillNormalWorkdays();
  }

  const periodSubtitle = document.getElementById('timesheetPeriodSubtitle');
  if (periodSubtitle) {
    periodSubtitle.innerText = `${found.displayName} (ตัดรอบ ${found.startDate} ถึง ${found.endDate} | จ่ายสิ้นเดือน ${found.payDate})`;
  }

  const btnExcel = document.getElementById('btnLoadExcelSample');
  if (btnExcel) {
    btnExcel.style.display = (found.id === '2026-09' && uid === 'user_admin') ? 'inline-flex' : 'none';
  }

  buildAttendanceTable();
  recalculateSalary();
}

function navPrevPeriod() {
  const select = document.getElementById('payrollPeriodSelect') || document.getElementById('payrollPeriodSelectAtt');
  if (!select || select.selectedIndex <= 0) return;
  const newIndex = select.selectedIndex - 1;
  const selectMain = document.getElementById('payrollPeriodSelect');
  const selectAtt = document.getElementById('payrollPeriodSelectAtt');
  if (selectMain) selectMain.selectedIndex = newIndex;
  if (selectAtt) selectAtt.selectedIndex = newIndex;
  onPayrollPeriodSelectChange();
}

function navNextPeriod() {
  const select = document.getElementById('payrollPeriodSelect') || document.getElementById('payrollPeriodSelectAtt');
  if (!select || select.selectedIndex >= select.options.length - 1) return;
  const newIndex = select.selectedIndex + 1;
  const selectMain = document.getElementById('payrollPeriodSelect');
  const selectAtt = document.getElementById('payrollPeriodSelectAtt');
  if (selectMain) selectMain.selectedIndex = newIndex;
  if (selectAtt) selectAtt.selectedIndex = newIndex;
  onPayrollPeriodSelectChange();
}

function selectPeriodMonth(periodId) {
  const select = document.getElementById('payrollPeriodSelect');
  const selectAtt = document.getElementById('payrollPeriodSelectAtt');
  if (select) select.value = periodId;
  if (selectAtt) selectAtt.value = periodId;
  onPayrollPeriodSelectChange();
  const calSection = document.getElementById('view-calendar');
  const btnToggle = document.getElementById('btnToggleCalendar');
  if (calSection) calSection.style.display = 'none';
  if (btnToggle) btnToggle.innerText = '📅 ปฏิทินวันทำงาน';
  switchView('main');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ==========================================================================
// 10. ระบบเพิ่มปีปฏิทินใหม่ (Add Year Modal)
// ==========================================================================
function openAddYearModal() {
  const nextBE = activeYearCE + 543 + 1;
  const inpBE = document.getElementById('inpNewYearBE');
  if (inpBE) inpBE.value = nextBE;
  document.getElementById('addYearModal').style.display = 'flex';
}

function closeAddYearModal() {
  document.getElementById('addYearModal').style.display = 'none';
}

function confirmAddNewYear() {
  const inpBE = document.getElementById('inpNewYearBE');
  const yearBE = parseInt(inpBE.value, 10);
  if (!yearBE || yearBE < 2400 || yearBE > 3000) {
    alert('กรุณาระบุปี พ.ศ. ให้ถูกต้อง');
    return;
  }
  const yearCE = yearBE - 543;

  if (allCalendars[yearCE]) {
    alert(`ปี พ.ศ. ${yearBE} (${yearCE}) มีอยู่ในระบบแล้ว`);
    activeYearCE = yearCE;
    updateYearSelectDropdown();
    onAppYearChange();
    closeAddYearModal();
    return;
  }

  const satPreset = document.getElementById('selSatPreset').value;

  const newCal = {
    yearCE,
    yearBE,
    title: `${yearBE} / ${yearCE}`,
    nationalHolidays: [],
    memorialHolidays: [],
    bridgeHolidays: [],
    orbrayWorkDays: [],
    workingSaturdays: [],
    customDayOverrides: {}
  };

  let cur = new Date(yearCE, 0, 1);
  const end = new Date(yearCE, 11, 31);
  let satIndex = 0;

  while (cur <= end) {
    if (cur.getDay() === 6) {
      const dStr = cur.toISOString().split('T')[0];
      if (satPreset === 'ALL_WORK') {
        newCal.workingSaturdays.push(dStr);
      } else if (satPreset === 'ALT_WORK') {
        if (satIndex % 2 === 1) {
          newCal.workingSaturdays.push(dStr);
        }
      }
      satIndex++;
    }
    cur.setDate(cur.getDate() + 1);
  }

  allCalendars[yearCE] = newCal;
  activeYearCE = yearCE;
  saveCalendarsToStorage();

  updateYearSelectDropdown();
  onAppYearChange();
  closeAddYearModal();

  alert(`เพิ่มปีปฏิทิน พ.ศ. ${yearBE} (${yearCE}) เรียบร้อยแล้ว! สามารถคลิกกำหนดวันหยุดได้ในปฏิทิน`);
}

function deleteCurrentYear() {
  const years = Object.keys(allCalendars);
  if (years.length <= 1) {
    alert('ไม่สามารถลบปีปฏิทินได้ เนื่องจากต้องมีอย่างน้อย 1 ปีในระบบ');
    return;
  }

  const cal = allCalendars[activeYearCE];
  if (!confirm(`คุณต้องการลบปฏิทินปี พ.ศ. ${cal.yearBE} (${cal.yearCE}) ใช่หรือไม่?`)) return;

  delete allCalendars[activeYearCE];
  activeYearCE = Object.keys(allCalendars).map(Number)[0];
  saveCalendarsToStorage();

  updateYearSelectDropdown();
  onAppYearChange();
}

// ==========================================================================
// 11. กล่องแก้ไขสถานะวันในปฏิทิน (Edit Day Modal)
// ==========================================================================
function openEditDayModal(dateStr) {
  const modal = document.getElementById('editDayModal');
  if (!modal) return;

  const dayInfo = getOrbrayDayInfo(dateStr);
  const d = new Date(dateStr);
  const thaiDays = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

  document.getElementById('modalTargetDateStr').value = dateStr;
  document.getElementById('modalDayDateDisplay').innerText = `${dateStr} (${thaiDays[d.getDay()]})`;

  const select = document.getElementById('modalDayTypeSelect');
  select.value = dayInfo.type;

  const nameInput = document.getElementById('modalHolidayNameInput');
  const cal = allCalendars[activeYearCE] || allCalendars['2026'];
  const override = cal.customDayOverrides ? cal.customDayOverrides[dateStr] : null;

  nameInput.value = override ? (override.name || '') : (dayInfo.name !== 'วันทำงานปกติ' ? dayInfo.name : '');
  onModalDayTypeSelectChange();

  modal.style.display = 'flex';
}

function closeEditDayModal() {
  document.getElementById('editDayModal').style.display = 'none';
}

function onModalDayTypeSelectChange() {
  const select = document.getElementById('modalDayTypeSelect');
  const group = document.getElementById('modalHolidayNameGroup');
  if (select.value === 'NORMAL_WORKDAY' || select.value === 'SUNDAY') {
    group.style.display = 'none';
  } else {
    group.style.display = 'block';
  }
}

function saveDayStatusFromModal() {
  const dateStr = document.getElementById('modalTargetDateStr').value;
  const newType = document.getElementById('modalDayTypeSelect').value;
  const holidayName = document.getElementById('modalHolidayNameInput').value.trim();

  const yearCE = parseInt(dateStr.split('-')[0], 10);
  const cal = allCalendars[yearCE];
  if (!cal) return;

  if (!cal.customDayOverrides) cal.customDayOverrides = {};

  cal.customDayOverrides[dateStr] = {
    type: newType,
    name: holidayName || null
  };

  saveCalendarsToStorage();
  renderOrbrayCalendarGrid();

  if (currentPeriod && currentAttendance) {
    const foundRow = currentAttendance.find(r => r.date === dateStr);
    if (foundRow) {
      updateRowOT(foundRow);
      buildAttendanceTable();
      recalculateSalary();
      saveCurrentAttendance();
    }
  }

  closeEditDayModal();
}

// ==========================================================================
// 12. การสลับแท็บและพิมพ์สลิป (View Navigation & Printing)
// ==========================================================================
function toggleCalendarView() {
  const calSection = document.getElementById('view-calendar');
  const btnToggle = document.getElementById('btnToggleCalendar');
  if (!calSection) return;

  if (calSection.style.display === 'none' || !calSection.style.display) {
    calSection.style.display = 'block';
    if (btnToggle) btnToggle.innerText = '✖ ปิดปฏิทินวันทำงาน';
    calSection.scrollIntoView({ behavior: 'smooth' });
  } else {
    calSection.style.display = 'none';
    if (btnToggle) btnToggle.innerText = '📅 ปฏิทินวันทำงาน';
  }
}

let currentActiveView = 'main';

function closeSlipView() {
  switchView(currentActiveView === 'attendance' ? 'attendance' : 'main');
}

function switchView(viewName) {
  const mainView = document.getElementById('view-main');
  const attView = document.getElementById('view-attendance');
  const calView = document.getElementById('view-calendar');
  const slipView = document.getElementById('view-slip');

  const navBtnMain = document.getElementById('navBtnMain');
  const navBtnAtt = document.getElementById('navBtnAttendance');
  const navBtnCal = document.getElementById('navBtnCalendar');
  const navBtnSlip = document.getElementById('navBtnSlip');

  // Track the primary view
  if (viewName === 'attendance' || viewName === 'timesheet') {
    currentActiveView = 'attendance';
  } else if (viewName === 'main' || viewName === 'salary') {
    currentActiveView = 'main';
  }

  // Hide all views first
  if (mainView) mainView.style.display = 'none';
  if (attView) attView.style.display = 'none';
  if (calView) calView.style.display = 'none';
  if (slipView) slipView.style.display = 'none';

  // Remove active class from nav buttons
  [navBtnMain, navBtnAtt, navBtnCal, navBtnSlip].forEach(btn => {
    if (btn) btn.classList.remove('active');
  });

  if (viewName === 'main' || viewName === 'salary') {
    if (mainView) mainView.style.display = 'block';
    if (navBtnMain) navBtnMain.classList.add('active');
  } else if (viewName === 'attendance' || viewName === 'timesheet') {
    if (attView) attView.style.display = 'block';
    if (navBtnAtt) navBtnAtt.classList.add('active');
  } else if (viewName === 'calendar') {
    if (currentActiveView === 'attendance') {
      if (attView) attView.style.display = 'block';
    } else {
      if (mainView) mainView.style.display = 'block';
    }
    if (calView) calView.style.display = 'block';
    if (navBtnCal) navBtnCal.classList.add('active');
  } else if (viewName === 'slip') {
    if (slipView) slipView.style.display = 'block';
    if (navBtnSlip) navBtnSlip.classList.add('active');
  }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function prepareAndPrint() {
  switchView('slip');
  setTimeout(() => {
    window.print();
  }, 350);
}

// ==========================================================================
// 13. ระบบจัดการหน้าต่างตั้งค่าค่าเงิน (Salary & Rates Settings Modal - Image 2)
// ==========================================================================
function openSettingsSalaryModal() {
  const modal = document.getElementById('settingsSalaryModal');
  if (!modal) return;

  // นำค่าปัจจุบันใส่ลงในฟอร์ม
  setInputValue('cfgBaseSalary', typeof currentSalaryConfig.baseSalary !== 'undefined' ? currentSalaryConfig.baseSalary : DEFAULT_SALARY_CONFIG.baseSalary);
  setInputValue('cfgTransportationAllowance', currentSalaryConfig.transportationAllowance ?? 0);
  setInputValue('cfgDiligenceFullAmount', currentSalaryConfig.diligenceFullAmount ?? 0);
  setInputValue('cfgFoodPerDay', currentSalaryConfig.foodPerDay ?? 0);
  setInputValue('cfgOtMealPerDay', currentSalaryConfig.otMealPerDay ?? 0);
  setInputValue('cfgOtDivisorHours', currentSalaryConfig.otDivisorHours ?? 240);
  setInputValue('cfgOt15Multiplier', currentSalaryConfig.ot15Multiplier ?? 1.5);
  setInputValue('cfgOt1Multiplier', currentSalaryConfig.ot1Multiplier ?? 1.0);
  setInputValue('cfgOt3Multiplier', currentSalaryConfig.ot3Multiplier ?? 3.0);
  setInputValue('cfgSsoDeduction', typeof currentSalaryConfig.ssoDeduction !== 'undefined' ? currentSalaryConfig.ssoDeduction : DEFAULT_SALARY_CONFIG.ssoDeduction);
  setInputValue('cfgSsoMaxBase', currentSalaryConfig.ssoMaxBase ?? 17500);

  updateLiveOTPreview();
  modal.style.display = 'flex';
}

function closeSettingsSalaryModal() {
  const modal = document.getElementById('settingsSalaryModal');
  if (modal) modal.style.display = 'none';
}

function updateLiveOTPreview() {
  const base = parseFloat(document.getElementById('cfgBaseSalary')?.value) || 0;
  const divisor = parseFloat(document.getElementById('cfgOtDivisorHours')?.value) || 240;
  const rate = divisor > 0 ? (base / divisor) : 0;
  const previewEl = document.getElementById('cfgLiveOTRate');
  if (previewEl) previewEl.innerText = formatCurrency(rate);
}

function parseFormNumber(val, defaultVal = 0) {
  if (val === '' || val === null || val === undefined) return defaultVal;
  const n = parseFloat(val);
  return isNaN(n) ? defaultVal : n;
}

function saveSalaryConfigFromModal() {
  const newConfig = {
    baseSalary: parseFormNumber(document.getElementById('cfgBaseSalary')?.value, 0),
    transportationAllowance: parseFormNumber(document.getElementById('cfgTransportationAllowance')?.value, 0),
    diligenceFullAmount: parseFormNumber(document.getElementById('cfgDiligenceFullAmount')?.value, 0),
    foodPerDay: parseFormNumber(document.getElementById('cfgFoodPerDay')?.value, 0),
    otMealPerDay: parseFormNumber(document.getElementById('cfgOtMealPerDay')?.value, 0),
    otDivisorHours: parseFormNumber(document.getElementById('cfgOtDivisorHours')?.value, 240),
    ot15Multiplier: parseFormNumber(document.getElementById('cfgOt15Multiplier')?.value, 1.5),
    ot1Multiplier: parseFormNumber(document.getElementById('cfgOt1Multiplier')?.value, 1.0),
    ot3Multiplier: parseFormNumber(document.getElementById('cfgOt3Multiplier')?.value, 3.0),
    ssoDeduction: parseFormNumber(document.getElementById('cfgSsoDeduction')?.value, 0),
    ssoMaxBase: parseFormNumber(document.getElementById('cfgSsoMaxBase')?.value, 17500)
  };

  currentSalaryConfig = newConfig;
  try {
    const userCfgKey = getScopedUserKey(STORAGE_KEY_SALARY_CONFIG);
    localStorage.setItem(userCfgKey, JSON.stringify(currentSalaryConfig));
    const activeUid = (typeof getActiveEditingUserId === 'function') ? getActiveEditingUserId() : 'user_admin';
    if (activeUid === 'user_admin') {
      localStorage.setItem(STORAGE_KEY_SALARY_CONFIG, JSON.stringify(currentSalaryConfig));
    }
  } catch (e) {
    console.warn('LocalStorage save failed', e);
  }

  // ซิงค์ขึ้น Cloud ถ้าล็อกอินอยู่
  if (typeof syncDataToCloud === 'function') {
    syncDataToCloud('salaryConfig', currentSalaryConfig);
  }

  closeSettingsSalaryModal();
  updateRateLabelsOnCards();
  recalculateSalary();
  showToastNotification('✅ บันทึกโครงสร้างค่าเงินและเบี้ยเลี้ยงเรียบร้อยแล้ว!');
}

function clearSalaryConfigToZero() {
  if (confirm('คุณต้องการปรับยอดเงินเดือน เบี้ยเลี้ยง และเงินหักทั้งหมดให้เป็น 0 ใช่หรือไม่?\n\n(สำหรับกรณีเดือนที่ยังไม่ได้เข้าทำงาน หรือยังไม่มีรายได้)')) {
    setInputValue('cfgBaseSalary', 0);
    setInputValue('cfgTransportationAllowance', 0);
    setInputValue('cfgDiligenceFullAmount', 0);
    setInputValue('cfgFoodPerDay', 0);
    setInputValue('cfgOtMealPerDay', 0);
    setInputValue('cfgSsoDeduction', 0);
    setInputValue('cfgSsoMaxBase', 0);
    updateLiveOTPreview();
    showToastNotification('🧹 เคลียร์ค่าเงินและเบี้ยเลี้ยงเป็น 0 ทั้งหมดแล้ว (กรุณากด "💾 บันทึกการตั้งค่า" เพื่อยืนยัน)');
  }
}

function resetSalaryConfigToDefault() {
  if (confirm('คุณต้องการคืนค่าอัตราเงินและเบี้ยเลี้ยงเป็นค่ามาตรฐานตามรูปที่ 2 หรือไม่?')) {
    currentSalaryConfig = { ...DEFAULT_SALARY_CONFIG };
    setInputValue('cfgBaseSalary', DEFAULT_SALARY_CONFIG.baseSalary);
    setInputValue('cfgTransportationAllowance', DEFAULT_SALARY_CONFIG.transportationAllowance);
    setInputValue('cfgDiligenceFullAmount', DEFAULT_SALARY_CONFIG.diligenceFullAmount);
    setInputValue('cfgFoodPerDay', DEFAULT_SALARY_CONFIG.foodPerDay);
    setInputValue('cfgOtMealPerDay', DEFAULT_SALARY_CONFIG.otMealPerDay);
    setInputValue('cfgOtDivisorHours', DEFAULT_SALARY_CONFIG.otDivisorHours);
    setInputValue('cfgOt15Multiplier', DEFAULT_SALARY_CONFIG.ot15Multiplier);
    setInputValue('cfgOt1Multiplier', DEFAULT_SALARY_CONFIG.ot1Multiplier);
    setInputValue('cfgOt3Multiplier', DEFAULT_SALARY_CONFIG.ot3Multiplier);
    setInputValue('cfgSsoDeduction', DEFAULT_SALARY_CONFIG.ssoDeduction);
    setInputValue('cfgSsoMaxBase', DEFAULT_SALARY_CONFIG.ssoMaxBase);
    updateLiveOTPreview();
  }
}

function setInputValue(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}

// ==========================================================================
// 14. ระบบจัดการ PIN Authentication & ผู้ใช้งาน (ADMIN PIN: 20523)
// ==========================================================================
function openAuthModal() {
  const modal = document.getElementById('authModal');
  if (!modal) return;
  clearAuthAlerts();
  clearPinInput();
  modal.style.display = 'flex';
  const pinInput = document.getElementById('pinDisplayInput');
  if (pinInput) {
    setTimeout(() => pinInput.focus(), 150);
  }
}

function closeAuthModal() {
  const modal = document.getElementById('authModal');
  if (modal) modal.style.display = 'none';
  clearAuthAlerts();
  clearPinInput();
}

function clearAuthAlerts() {
  const err = document.getElementById('authErrorAlert');
  const succ = document.getElementById('authSuccessAlert');
  if (err) { err.style.display = 'none'; err.innerText = ''; }
  if (succ) { succ.style.display = 'none'; succ.innerText = ''; }
}

function showAuthError(msg) {
  const err = document.getElementById('authErrorAlert');
  if (err) { err.innerText = msg; err.style.display = 'block'; }
}

function showAuthSuccess(msg) {
  const succ = document.getElementById('authSuccessAlert');
  if (succ) { succ.innerText = msg; succ.style.display = 'block'; }
}

function pressPinDigit(digit) {
  const input = document.getElementById('pinDisplayInput');
  if (!input) return;
  clearAuthAlerts();
  if (input.value.length < 8) {
    input.value += digit;
  }
}

function backspacePinDigit() {
  const input = document.getElementById('pinDisplayInput');
  if (!input) return;
  clearAuthAlerts();
  input.value = input.value.slice(0, -1);
}

function clearPinInput() {
  const input = document.getElementById('pinDisplayInput');
  if (input) input.value = '';
}

function togglePinVisibility() {
  const input = document.getElementById('pinDisplayInput');
  const icon = document.getElementById('pinVisibilityIcon');
  if (!input) return;
  if (input.type === 'password') {
    input.type = 'text';
    if (icon) icon.innerText = '🙈';
  } else {
    input.type = 'password';
    if (icon) icon.innerText = '👁️';
  }
}

function handlePinKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    submitPinLogin();
  }
}

async function submitPinLogin() {
  clearAuthAlerts();
  const input = document.getElementById('pinDisplayInput');
  const pin = input ? input.value.trim() : '';

  if (!pin) {
    showAuthError('กรุณาระบุรหัส PIN');
    return;
  }

  const btn = document.getElementById('btnAuthSubmit');
  const prevText = btn ? btn.innerText : '';
  if (btn) { btn.disabled = true; btn.innerText = 'กำลังตรวจสอบ...'; }

  try {
    const res = await loginWithPin(pin);
    if (res.success) {
      const roleText = res.user.role === 'admin' ? 'ผู้ดูแลระบบ (ADMIN)' : 'พนักงาน';
      showAuthSuccess(`เข้าสู่ระบบสำเร็จ! สวัสดีคุณ ${res.user.name} (${roleText})`);
      setTimeout(() => {
        closeAuthModal();
        reloadUserDataAfterAuthChange();
        showToastNotification(`เข้าสู่ระบบสำเร็จ: ${res.user.name}`);
      }, 500);
    } else {
      showAuthError(res.error || 'รหัส PIN ไม่ถูกต้อง');
    }
  } catch (err) {
    showAuthError(err.message || 'เกิดข้อผิดพลาดในการตรวจสอบรหัส PIN');
  } finally {
    if (btn) { btn.disabled = false; btn.innerText = prevText; }
  }
}

function quickLoginAdmin() {
  const input = document.getElementById('pinDisplayInput');
  if (input) input.value = '20523';
  submitPinLogin();
}

async function handleLogout() {
  if (confirm('คุณต้องการออกจากระบบหรือไม่?')) {
    await logoutCurrentUser();
    reloadUserDataAfterAuthChange();
    showToastNotification('ออกจากระบบเรียบร้อยแล้ว');
  }
}

/**
 * โหลดข้อมูลและรีเฟรชหน้าจอทั้งหมดหลังจากผู้ใช้เปลี่ยน
 */
function reloadUserDataAfterAuthChange() {
  loadStorageData();
  refreshPeriodSelector();
  renderOrbrayCalendarGrid();
  updateRateLabelsOnCards();
  recalculateSalary();
  updateAdminScopeBanner();
  updateNavbarAuthUI(currentUser, isFirebaseOnline);
}

/**
 * สลับเปลี่ยนพนักงานที่กำลังดู/แก้ไขข้อมูล (เฉพาะ ADMIN)
 */
function switchEditingUser(targetUserId) {
  if (!targetUserId) return;

  // ตรวจสอบสิทธิ์: เฉพาะ ADMIN (JITTRAKAN K.) เท่านั้นที่สามารถดูหรือสลับข้อมูลของพนักงานคนอื่นได้
  if (!currentUser || currentUser.role !== 'admin' || currentUser.id !== 'user_admin') {
    showToastNotification('⚠️ คุณไม่มีสิทธิ์ดูข้อมูลของพนักงานคนอื่น เฉพาะ ADMIN เท่านั้น');
    return;
  }

  setActiveEditingUserId(targetUserId);
  const targetUser = getActiveEditingUser();

  // อัปเดตข้อมูลและคำนวณใหม่
  loadStorageData();
  refreshPeriodSelector();
  renderOrbrayCalendarGrid();
  updateRateLabelsOnCards();
  recalculateSalary();
  updateAdminScopeBanner();

  // อัปเดต Dropdown ใน Navbar ให้ตรงกัน
  const navSelect = document.getElementById('adminUserNavSelect');
  if (navSelect) navSelect.value = targetUserId;

  showToastNotification(`สลับไปยังข้อมูลของ: ${targetUser.name}`);
}

/**
 * อัปเดตแบนเนอร์แจ้งเตือนสถานะการแก้ไขข้อมูลของ Admin
 */
function updateAdminScopeBanner() {
  const banners = [
    document.getElementById('adminScopeBanner'),
    document.getElementById('adminScopeBannerAtt')
  ].filter(Boolean);
  if (banners.length === 0) return;

  const currentActiveId = getActiveEditingUserId();
  const isAdmin = currentUser && currentUser.role === 'admin';

  if (isAdmin && currentActiveId !== 'user_admin') {
    const targetUser = getActiveEditingUser();
    const bannerHtml = `
      <div class="admin-scope-banner-info">
        <span style="font-size:1.3rem;">👑</span>
        <div>
          กำลังดูและแก้ไขข้อมูลของ: <strong>${escapeHtml(targetUser.name)}</strong> (${escapeHtml(targetUser.department || 'พนักงาน')})
          <div style="font-size:0.78rem; color:#0369a1; margin-top:2px;">
            ⚡ สิทธิ์ ADMIN: คุณสามารถแก้ไขเวลาทำงาน ปรับโครงสร้างค่าจ้าง และพิมพ์สลิปของพนักงานคนนี้ได้
          </div>
        </div>
      </div>
      <div>
        <button type="button" class="btn btn-xs btn-outline-primary" onclick="switchEditingUser('user_admin')">
          ↩️ กลับไปที่ข้อมูล ADMIN
        </button>
      </div>
    `;
    banners.forEach(b => {
      b.innerHTML = bannerHtml;
      b.style.display = 'flex';
    });
  } else {
    banners.forEach(b => {
      b.style.display = 'none';
    });
  }
}

/**
 * อัปเดต Navbar ส่วนแสดงสถานะการล็อกอินและ User Switcher
 */
function updateNavbarAuthUI(user, isOnline) {
  const container = document.getElementById('navAuthArea');
  if (!container) return;

  if (user) {
    const activeUserId = getActiveEditingUserId();

    if (user.role === 'admin') {
      // ผู้ใช้เป็น ADMIN: แสดง Badge + Dropdown สลับผู้ใช้ + ปุ่มบัญชีของฉัน + ปุ่มจัดการผู้ใช้ + ออกจากระบบ
      const allUsers = getAllSystemUsers();
      const optionsHtml = allUsers.map(u => {
        const selected = (u.id === activeUserId) ? 'selected' : '';
        const roleLabel = u.role === 'admin' ? '👑 ' : '👤 ';
        return `<option value="${u.id}" ${selected}>${roleLabel}${escapeHtml(u.name)} (รหัส: ${escapeHtml(u.empCode || '-')}, PIN: ${u.pin})</option>`;
      }).join('');

      container.innerHTML = `
        <div class="user-chip-menu">
          <span class="badge-admin-gold">👑 ADMIN</span>
          <div class="nav-user-switcher-box" title="เลือกพนักงานเพื่อดูหรือแก้ไขข้อมูล">
            <span class="nav-user-switcher-label">พนักงาน:</span>
            <select id="adminUserNavSelect" class="nav-user-switcher-select" onchange="switchEditingUser(this.value)">
              ${optionsHtml}
            </select>
          </div>
          <button type="button" class="btn btn-xs btn-outline-primary" onclick="openAccountModal()" title="ดูและจัดการข้อมูลบัญชีโปรไฟล์ของฉัน">
            👤 บัญชีของฉัน
          </button>
          <button type="button" class="btn btn-xs btn-outline-primary" onclick="openManageUsersModal()" title="จัดการรายชื่อและรหัส PIN พนักงาน">
            👥 จัดการผู้ใช้
          </button>
          <button type="button" class="btn btn-xs btn-outline-danger" onclick="handleLogout()" title="ออกจากระบบ">
            ออก
          </button>
        </div>
      `;
    } else {
      // พนักงานทั่วไป: แสดงชื่อ + ปุ่มบัญชีของฉัน + ปุ่มออกจากระบบ
      container.innerHTML = `
        <div class="user-chip-menu">
          <div class="user-chip" onclick="openAccountModal()" style="cursor: pointer;" title="คลิกเพื่อดูและจัดการข้อมูลบัญชีของฉัน">
            <div class="user-avatar-initials">${user.name.charAt(0).toUpperCase()}</div>
            <span class="user-name-text">${escapeHtml(user.name)}</span>
            <span class="status-dot dot-online" title="เข้าสู่ระบบแล้ว"></span>
          </div>
          <button type="button" class="btn btn-xs btn-outline-primary" onclick="openAccountModal()" title="ดูและจัดการข้อมูลบัญชีของฉัน">
            👤 บัญชีของฉัน
          </button>
          <button type="button" class="btn btn-xs btn-outline-danger" onclick="handleLogout()" title="ออกจากระบบ">
            ออก
          </button>
        </div>
      `;
    }
  } else {
    // ยังไม่ได้ล็อกอิน: แสดงปุ่มเข้าสู่ระบบ PIN
    container.innerHTML = `
      <button type="button" class="btn btn-sm btn-primary btn-cta" onclick="openAuthModal()">
        🔐 เข้าสู่ระบบ PIN
      </button>
    `;
  }

  updateAdminScopeBanner();
}

/**
 * ==========================================================================
 * ระบบจัดการบัญชีผู้ใช้งานส่วนตัว (My Account Profile Modal)
 * ==========================================================================
 */
function openAccountModal() {
  const modal = document.getElementById('accountModal');
  if (!modal) return;

  const user = currentUser || (typeof getActiveEditingUser === 'function' ? getActiveEditingUser() : null);
  if (!user) {
    showToastNotification('⚠️ กรุณาเข้าสู่ระบบก่อนจัดการบัญชี');
    openAuthModal();
    return;
  }

  const alertBox = document.getElementById('accountAlert');
  if (alertBox) {
    alertBox.style.display = 'none';
    alertBox.textContent = '';
  }

  // เติมข้อมูลลงในการ์ดสรุปโปรไฟล์
  const avatarLarge = document.getElementById('accAvatarLarge');
  if (avatarLarge) avatarLarge.textContent = user.role === 'admin' ? '👑' : (user.avatar || '👤');

  const summaryName = document.getElementById('accSummaryName');
  if (summaryName) summaryName.textContent = user.name;

  const summaryRole = document.getElementById('accSummaryRole');
  if (summaryRole) {
    summaryRole.className = user.role === 'admin' ? 'badge-account-role' : 'badge-account-role-user';
    summaryRole.textContent = user.role === 'admin' ? '👑 ผู้ดูแลระบบ (ADMIN)' : '👤 พนักงานทั่วไป';
  }

  const summaryCode = document.getElementById('accSummaryCode');
  if (summaryCode) summaryCode.textContent = 'ID: ' + (user.empCode || (user.id === 'user_admin' ? '20523' : ('EMP-' + user.pin)));

  const summaryDept = document.getElementById('accSummaryDept');
  if (summaryDept) summaryDept.textContent = user.department || 'ฝ่ายปฏิบัติการ';

  // เติมฟอร์ม
  const elName = document.getElementById('accName');
  if (elName) elName.value = user.name || '';

  const elCompany = document.getElementById('accCompany');
  if (elCompany) elCompany.value = user.companyName || 'CACULATION SALARY';

  const elEmpCode = document.getElementById('accEmpCode');
  if (elEmpCode) elEmpCode.value = user.empCode || (user.id === 'user_admin' ? '20523' : ('EMP-' + user.pin));

  const elDept = document.getElementById('accDept');
  if (elDept) elDept.value = user.department || '';

  const elPin = document.getElementById('accPin');
  if (elPin) {
    elPin.value = user.pin || '';
    elPin.type = 'password';
  }
  const pinIcon = document.getElementById('accPinIcon');
  if (pinIcon) pinIcon.textContent = '👁️';

  modal.style.display = 'flex';
}

function closeAccountModal() {
  const modal = document.getElementById('accountModal');
  if (modal) modal.style.display = 'none';
}

function toggleAccPinVisibility() {
  const input = document.getElementById('accPin');
  const icon = document.getElementById('accPinIcon');
  if (!input || !icon) return;

  if (input.type === 'password') {
    input.type = 'text';
    icon.textContent = '🙈';
  } else {
    input.type = 'password';
    icon.textContent = '👁️';
  }
}

function handleSaveAccountProfile(e) {
  e.preventDefault();
  const alertBox = document.getElementById('accountAlert');
  if (alertBox) {
    alertBox.style.display = 'none';
    alertBox.textContent = '';
  }

  const name = document.getElementById('accName')?.value.trim();
  const companyName = document.getElementById('accCompany')?.value.trim();
  const empCode = document.getElementById('accEmpCode')?.value.trim();
  const department = document.getElementById('accDept')?.value.trim();
  const pin = document.getElementById('accPin')?.value.trim();

  if (!name) {
    if (alertBox) {
      alertBox.className = 'alert alert-danger';
      alertBox.textContent = 'กรุณาระบุชื่อ-นามสกุล';
      alertBox.style.display = 'block';
    }
    return;
  }

  if (!pin || !/^\d{4,8}$/.test(pin)) {
    if (alertBox) {
      alertBox.className = 'alert alert-danger';
      alertBox.textContent = 'รหัส PIN ต้องเป็นตัวเลขความยาว 4–8 หลัก';
      alertBox.style.display = 'block';
    }
    return;
  }

  const res = updateCurrentAccountProfile({
    name,
    companyName: (companyName || 'CACULATION SALARY').toUpperCase(),
    empCode: (empCode || '-').toUpperCase(),
    department: department || '-',
    pin
  });

  if (res.success) {
    // ซิงก์ข้อมูลพนักงานไปยัง salaryProfile และการคำนวณ
    syncSalaryProfileWithActiveUser();
    calculatePayroll();
    updateNavbarAuthUI(currentUser, isFirebaseOnline);

    closeAccountModal();
    showToastNotification(`✅ บันทึกข้อมูลบัญชี (${name}) เรียบร้อยแล้ว!`);
  } else {
    if (alertBox) {
      alertBox.className = 'alert alert-danger';
      alertBox.textContent = res.error || 'เกิดข้อผิดพลาดในการบันทึกข้อมูล';
      alertBox.style.display = 'block';
    } else {
      alert(res.error || 'เกิดข้อผิดพลาดในการบันทึกข้อมูล');
    }
  }
}

/**
 * ระบบจัดการรายชื่อผู้ใช้และ PIN (Manage Users Modal - Admin Only: JITTRAKAN K.)
 */
function openManageUsersModal() {
  if (!currentUser || currentUser.role !== 'admin' || currentUser.id !== 'user_admin') {
    showToastNotification('⚠️ คุณต้องเข้าสู่ระบบเป็น ADMIN (JITTRAKAN K.) เพื่อจัดการผู้ใช้');
    openAuthModal();
    return;
  }
  const modal = document.getElementById('manageUsersModal');
  if (!modal) return;
  renderManageUsersTable();
  modal.style.display = 'flex';
}

function closeManageUsersModal() {
  const modal = document.getElementById('manageUsersModal');
  if (modal) modal.style.display = 'none';
}

function renderManageUsersTable() {
  const tbody = document.getElementById('manageUsersTableBody');
  if (!tbody) return;

  const users = getAllSystemUsers();
  tbody.innerHTML = users.map(u => {
    const isMasterAdmin = (u.id === 'user_admin');
    const badge = isMasterAdmin 
      ? '<span class="user-role-badge-admin">👑 ADMIN</span>' 
      : '<span class="user-role-badge-user">👤 พนักงาน</span>';

    const actionHtml = isMasterAdmin 
      ? '<span class="badge-pill-locked" title="บัญชีผู้ดูแลระบบหลัก (ไม่สามารถลบหรือแก้ไขสิทธิ์ได้)">🔒 บัญชีหลัก</span>' 
      : `
        <div class="user-action-group">
          <button type="button" class="btn btn-outline-primary btn-action-pill" onclick="openEditUserModal('${u.id}')" title="แก้ไขข้อมูลพนักงาน">
            ✏️ แก้ไข
          </button>
          <button type="button" class="btn btn-outline-danger btn-action-pill" onclick="confirmDeleteUser('${u.id}')" title="ลบพนักงานคนนี้">
            🗑️ ลบ
          </button>
        </div>
      `;

    return `
      <tr>
        <td style="text-align: center;">${badge}</td>
        <td><strong style="color: #0c4a6e;">${escapeHtml(u.name)}</strong></td>
        <td style="text-align: center;"><span class="badge-account-code">${escapeHtml(u.empCode || '-')}</span></td>
        <td><span style="color: #475569;">${escapeHtml(u.department || '-')}</span></td>
        <td style="text-align: center;"><span class="pin-code-badge">${escapeHtml(u.pin)}</span></td>
        <td style="text-align: center;">${actionHtml}</td>
      </tr>
    `;
  }).join('');
}

function handleAddNewUserSubmit(e) {
  e.preventDefault();
  const name = document.getElementById('newUserName')?.value.trim();
  const empCode = document.getElementById('newUserEmpCode')?.value.trim();
  const pin = document.getElementById('newUserPin')?.value.trim();
  const dept = document.getElementById('newUserDept')?.value.trim();

  const res = registerNewUser(name, pin, 'user', dept, 'CACULATION SALARY', empCode);
  if (res.success) {
    document.getElementById('newUserName').value = '';
    if (document.getElementById('newUserEmpCode')) document.getElementById('newUserEmpCode').value = '';
    document.getElementById('newUserPin').value = '';
    document.getElementById('newUserDept').value = '';
    renderManageUsersTable();
    updateNavbarAuthUI(currentUser, isFirebaseOnline);
    showToastNotification(`✅ เพิ่มพนักงาน ${name} (รหัส: ${res.user.empCode}, PIN: ${pin}) เรียบร้อยแล้ว!`);
  } else {
    alert(res.error || 'ไม่สามารถเพิ่มผู้ใช้ได้');
  }
}

/**
 * หน้าต่าง Modal แก้ไขข้อมูลพนักงาน (Admin Only)
 */
function openEditUserModal(userId) {
  if (!currentUser || currentUser.role !== 'admin' || currentUser.id !== 'user_admin') {
    showToastNotification('⚠️ คุณไม่มีสิทธิ์แก้ไขข้อมูลพนักงาน เฉพาะ ADMIN เท่านั้น');
    return;
  }
  const users = getAllSystemUsers(true);
  const u = users.find(x => x.id === userId);
  if (!u) return;

  const modal = document.getElementById('editUserModal');
  if (!modal) return;

  const alertBox = document.getElementById('editUserAlert');
  if (alertBox) {
    alertBox.style.display = 'none';
    alertBox.textContent = '';
  }

  const idInput = document.getElementById('editUserId');
  const nameInput = document.getElementById('editUserName');
  const empCodeInput = document.getElementById('editUserEmpCode');
  const deptInput = document.getElementById('editUserDept');
  const pinInput = document.getElementById('editUserPin');

  if (idInput) idInput.value = u.id;
  if (nameInput) nameInput.value = u.name || '';
  if (empCodeInput) empCodeInput.value = u.empCode || '';
  if (deptInput) deptInput.value = u.department || '';
  if (pinInput) pinInput.value = u.pin || '';

  modal.style.display = 'flex';
}

function closeEditUserModal() {
  const modal = document.getElementById('editUserModal');
  if (modal) modal.style.display = 'none';
}

function handleSaveEditUserSubmit(e) {
  e.preventDefault();
  const userId = document.getElementById('editUserId')?.value;
  const name = document.getElementById('editUserName')?.value.trim();
  const empCode = document.getElementById('editUserEmpCode')?.value.trim();
  const dept = document.getElementById('editUserDept')?.value.trim();
  const pin = document.getElementById('editUserPin')?.value.trim();
  const alertBox = document.getElementById('editUserAlert');

  if (!userId || !name) {
    if (alertBox) {
      alertBox.textContent = 'กรุณาระบุชื่อ-นามสกุลพนักงาน';
      alertBox.style.display = 'block';
    }
    return;
  }

  if (!pin || !/^\d{4,8}$/.test(pin)) {
    if (alertBox) {
      alertBox.textContent = 'รหัส PIN ต้องเป็นตัวเลขความยาว 4–8 หลักเท่านั้น';
      alertBox.style.display = 'block';
    }
    return;
  }

  const res = updateSystemUser(userId, {
    name: name.toUpperCase(),
    empCode: (empCode || '-').toUpperCase(),
    department: dept || '-',
    pin: pin
  });

  if (res.success) {
    renderManageUsersTable();
    updateNavbarAuthUI(currentUser, isFirebaseOnline);
    if (getActiveEditingUserId() === userId) {
      syncSalaryProfileWithActiveUser();
      calculatePayroll();
    }
    closeEditUserModal();
    showToastNotification(`✅ อัปเดตข้อมูลพนักงาน "${name}" สำเร็จ!`);
  } else {
    if (alertBox) {
      alertBox.textContent = res.error || 'เกิดข้อผิดพลาดในการบันทึกข้อมูล';
      alertBox.style.display = 'block';
    } else {
      alert(res.error || 'เกิดข้อผิดพลาดในการบันทึกข้อมูล');
    }
  }
}

function confirmDeleteUser(userId) {
  if (userId === 'user_admin') {
    alert('ไม่สามารถลบบัญชีผู้ดูแลระบบหลัก (JITTRAKAN K.) ได้');
    return;
  }
  const users = getAllSystemUsers();
  const u = users.find(x => x.id === userId);
  if (!u) return;

  if (confirm(`คุณต้องการลบพนักงาน "${u.name}" (รหัส: ${u.empCode || '-'}, PIN: ${u.pin}) ออกจากระบบหรือไม่?`)) {
    const res = deleteSystemUser(userId);
    if (res.success) {
      renderManageUsersTable();
      updateNavbarAuthUI(currentUser, isFirebaseOnline);
      reloadUserDataAfterAuthChange();
      showToastNotification(`🗑️ ลบพนักงาน "${u.name}" เรียบร้อยแล้ว`);
    } else {
      alert(res.error || 'ไม่สามารถลบได้');
    }
  }
}

function confirmClearAllStaffUsers() {
  const users = getAllSystemUsers();
  const staffUsers = users.filter(u => u.id !== 'user_admin');
  if (staffUsers.length === 0) {
    showToastNotification('ℹ️ ปัจจุบันไม่มีรายชื่อพนักงานอื่นในระบบ (มีเฉพาะ ADMIN)');
    return;
  }

  if (confirm(`คุณต้องการลบรายชื่อพนักงานทั้งหมดจำนวน ${staffUsers.length} คน ออกจากระบบหรือไม่?\n\n* บัญชีผู้ดูแลระบบหลัก (JITTRAKAN K.) จะยังคงอยู่ตามปกติ`)) {
    clearAllStaffUsers();
    renderManageUsersTable();
    updateNavbarAuthUI(currentUser, isFirebaseOnline);
    reloadUserDataAfterAuthChange();
    showToastNotification('🗑️ ลบรายชื่อพนักงานทั้งหมดเรียบร้อยแล้ว');
  }
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ==========================================================================
// 15. ระบบตั้งค่า Firebase Project Credentials Modal
// ==========================================================================
function openFirebaseConfigModal() {
  closeAuthModal();
  const modal = document.getElementById('firebaseConfigModal');
  if (!modal) return;
  const cfg = typeof getActiveFirebaseConfig === 'function' ? getActiveFirebaseConfig() : DEFAULT_FIREBASE_CONFIG;
  setInputValue('fbCfgApiKey', cfg.apiKey || '');
  setInputValue('fbCfgAuthDomain', cfg.authDomain || '');
  setInputValue('fbCfgProjectId', cfg.projectId || '');
  setInputValue('fbCfgStorageBucket', cfg.storageBucket || '');
  setInputValue('fbCfgMessagingSenderId', cfg.messagingSenderId || '');
  setInputValue('fbCfgAppId', cfg.appId || '');
  modal.style.display = 'flex';
}

function closeFirebaseConfigModal() {
  const modal = document.getElementById('firebaseConfigModal');
  if (modal) modal.style.display = 'none';
}

function saveFirebaseConfigFromModal() {
  const newCfg = {
    apiKey: document.getElementById('fbCfgApiKey')?.value.trim(),
    authDomain: document.getElementById('fbCfgAuthDomain')?.value.trim(),
    projectId: document.getElementById('fbCfgProjectId')?.value.trim(),
    storageBucket: document.getElementById('fbCfgStorageBucket')?.value.trim(),
    messagingSenderId: document.getElementById('fbCfgMessagingSenderId')?.value.trim(),
    appId: document.getElementById('fbCfgAppId')?.value.trim()
  };

  if (typeof saveFirebaseConfig === 'function') {
    saveFirebaseConfig(newCfg);
  }
  closeFirebaseConfigModal();
  alert('บันทึกคอนฟิก Firebase เรียบร้อยแล้ว! ระบบจะรีโหลดเพื่อเริ่มต้นการเชื่อมต่อใหม่');
  location.reload();
}

function resetFirebaseConfigDefault() {
  if (confirm('คุณต้องการรีเซ็ตค่า Firebase Config เป็นค่าเริ่มต้นหรือไม่?')) {
    if (typeof saveFirebaseConfig === 'function') {
      saveFirebaseConfig(DEFAULT_FIREBASE_CONFIG);
    }
    closeFirebaseConfigModal();
    location.reload();
  }
}

// ระบบ Toast Notification แจ้งเตือนสวยงาม
function showToastNotification(message) {
  let toast = document.getElementById('appGlobalToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'appGlobalToast';
    toast.className = 'app-toast';
    document.body.appendChild(toast);
  }
  toast.innerText = message;
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}

// ==========================================================================
// 16. เริ่มต้นระบบ (DOM Ready Initialization)
// ==========================================================================
window.addEventListener('DOMContentLoaded', () => {
  loadStorageData();
  updateYearSelectDropdown();
  refreshPeriodSelector();
  renderOrbrayCalendarGrid();
  updateRateLabelsOnCards();

  // ลงทะเบียนติดตามสถานะผู้ใช้จาก Firebase
  if (typeof addAuthStateListener === 'function') {
    addAuthStateListener(updateNavbarAuthUI);
  }

  updateAdminScopeBanner();
});
