/**
 * ระบบคำนวณเงินเดือน & ออกใบแจ้งรายได้ ออบเรย์ (ประเทศไทย) จำกัด
 * อ้างอิงสูตรคำนวณจากไฟล์ Excel: DATA SALARY 210826-200926.xlsx
 * และปฏิทินวันทำงานปี 2569 / 2026 จาก Image (1).jpg และ Image (2).jpg
 */

// ==========================================================================
// 1. ฐานข้อมูลปฏิทินวันทำงาน Orbray 2569 / 2026 (ตามภาพ Image 1 & 2)
// ==========================================================================

const ORBRAY_CALENDAR_2026 = {
  // วันหยุดนักขัตฤกษ์ (วงรีสีเหลือง)
  nationalHolidays: [
    '2026-01-01', // วันขึ้นปีใหม่
    '2026-03-03', // วันมาฆบูชา
    '2026-04-13', '2026-04-14', '2026-04-15', '2026-04-16', // วันสงกรานต์
    '2026-05-01', // วันแรงงานแห่งชาติ
    '2026-06-01', // วันวิสาขบูชา (ชดเชย)
    '2026-07-28', '2026-07-29', // วันเฉลิมฯ ร.10 / วันอาสาฬหบูชา
    '2026-08-12', // วันแม่แห่งชาติ
    '2026-10-13', // วันนวมินทรมหาราช ร.9
    '2026-10-23', // วันปิยมหาราช
    '2026-12-05', // วันพ่อแห่งชาติ ร.9
    '2026-12-28', // วันสมเด็จพระเจ้าตากสินฯ / วันหยุดพิเศษ
    '2026-12-31'  // วันสิ้นปี
  ],
  // วันหยุดประเพณีจ่ายเงิน (วงรีสีม่วง: Memorial day with pay)
  memorialHolidays: [
    '2026-09-19',
    '2026-12-30'
  ],
  // วันหยุดพิเศษ/เชื่อมวันหยุด (วงรีสีดำ วันธรรมดา)
  bridgeHolidays: [
    '2026-01-02',
    '2026-12-29'
  ],
  // วันทำงานพิเศษ Orbray (วงรีสีแดง: Orbray day)
  orbrayWorkDays: [
    '2026-10-17'
  ],
  // วันเสาร์ที่เป็นวันทำงานปกติ (ตัวเลขอักษรสีดำ ไม่มีวงกลม)
  workingSaturdays: [
    '2026-01-17',
    '2026-02-21',
    '2026-03-14', '2026-03-28',
    '2026-04-04', '2026-04-18',
    '2026-05-09',
    '2026-06-20',
    '2026-07-04', '2026-07-25',
    '2026-08-01', '2026-08-15', '2026-08-29',
    '2026-10-17', '2026-10-31',
    '2026-12-12'
  ]
};

// ข้อมูล 12 งวดการจ่ายประจำปี 2569 (ตัดรอบทุกวันที่ 20 จ่ายทุกสิ้นเดือน)
const PAYROLL_PERIODS_2026 = [
  { id: '2026-01', monthName: 'มกราคม 2569', startDate: '2025-12-21', endDate: '2026-01-20', payDate: '31/01/2569' },
  { id: '2026-02', monthName: 'กุมภาพันธ์ 2569', startDate: '2026-01-21', endDate: '2026-02-20', payDate: '28/02/2569' },
  { id: '2026-03', monthName: 'มีนาคม 2569', startDate: '2026-02-21', endDate: '2026-03-20', payDate: '31/03/2569' },
  { id: '2026-04', monthName: 'เมษายน 2569', startDate: '2026-03-21', endDate: '2026-04-20', payDate: '30/04/2569' },
  { id: '2026-05', monthName: 'พฤษภาคม 2569', startDate: '2026-04-21', endDate: '2026-05-20', payDate: '31/05/2569' },
  { id: '2026-06', monthName: 'มิถุนายน 2569', startDate: '2026-05-21', endDate: '2026-06-20', payDate: '30/06/2569' },
  { id: '2026-07', monthName: 'กรกฎาคม 2569', startDate: '2026-06-21', endDate: '2026-07-20', payDate: '31/07/2569' },
  { id: '2026-08', monthName: 'สิงหาคม 2569', startDate: '2026-07-21', endDate: '2026-08-20', payDate: '31/08/2569' },
  { id: '2026-09', monthName: 'กันยายน 2569', startDate: '2026-08-21', endDate: '2026-09-20', payDate: '30/09/2569' },
  { id: '2026-10', monthName: 'ตุลาคม 2569', startDate: '2026-09-21', endDate: '2026-10-20', payDate: '31/10/2569' },
  { id: '2026-11', monthName: 'พฤศจิกายน 2569', startDate: '2026-10-21', endDate: '2026-11-20', payDate: '30/11/2569' },
  { id: '2026-12', monthName: 'ธันวาคม 2569', startDate: '2026-11-21', endDate: '2026-12-20', payDate: '31/12/2569' }
];

// ==========================================================================
// 2. ข้อมูลตัวอย่างจากไฟล์ Excel: DATA SALARY 210826-200926.xlsx
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
// 3. ฟังก์ชันระบุประเภทวันตามปฏิทิน Orbray
// ==========================================================================
function getOrbrayDayInfo(dateStr) {
  const d = new Date(dateStr);
  const dayOfWeek = d.getDay(); // 0 = Sun, 6 = Sat

  if (ORBRAY_CALENDAR_2026.nationalHolidays.includes(dateStr)) {
    return { type: 'NATIONAL_HOLIDAY', name: 'วันหยุดนักขัตฤกษ์', isWorkDay: false, badgeClass: 'badge-national' };
  }
  if (ORBRAY_CALENDAR_2026.memorialHolidays.includes(dateStr)) {
    return { type: 'MEMORIAL_HOLIDAY', name: 'วันหยุดประเพณีจ่ายเงิน', isWorkDay: false, badgeClass: 'badge-memorial' };
  }
  if (ORBRAY_CALENDAR_2026.bridgeHolidays.includes(dateStr)) {
    return { type: 'BRIDGE_HOLIDAY', name: 'วันหยุดพิเศษ', isWorkDay: false, badgeClass: 'badge-bridge' };
  }
  if (dayOfWeek === 0) {
    return { type: 'SUNDAY', name: 'วันอาทิตย์หยุด', isWorkDay: false, badgeClass: 'badge-sunday' };
  }
  if (dayOfWeek === 6) {
    if (ORBRAY_CALENDAR_2026.workingSaturdays.includes(dateStr)) {
      return { type: 'WORKING_SATURDAY', name: 'เสาร์ทำงาน', isWorkDay: true, badgeClass: 'badge-work-sat' };
    } else {
      return { type: 'SATURDAY_HOLIDAY', name: 'เสาร์หยุด', isWorkDay: false, badgeClass: 'badge-sat-holiday' };
    }
  }
  return { type: 'NORMAL_WORKDAY', name: 'วันทำงานปกติ', isWorkDay: true, badgeClass: 'badge-workday' };
}

// ==========================================================================
// 4. คำนวณชั่วโมง OT จากเวลาออกงานตามสูตร SWITCH ใน Excel
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
      // คำนวณตามจริงถ้าเป็นเวลาอื่น
      const parts = t.split(':');
      if (parts.length === 2) {
        const h = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        const totalMinutes = (h * 60 + m) - (17 * 60);
        if (totalMinutes <= 0) return 0;
        return Math.floor((totalMinutes / 60) * 2) / 2; // ปัดลงเป็นขั้น 0.5 ชม.
      }
      return 0;
  }
}

// ==========================================================================
// 5. โครงสร้างสถานะและข้อมูลหลักของแอปพลิเคชัน
// ==========================================================================
let currentPeriod = PAYROLL_PERIODS_2026.find(p => p.id === '2026-09'); // เริ่มต้นที่งวด ก.ย. (21 ส.ค. - 20 ก.ย.)
let currentAttendance = JSON.parse(JSON.stringify(EXCEL_SAMPLE_ATTENDANCE));

let salaryProfile = {
  companyName: 'ออบเรย์ (ประเทศไทย) จำกัด',
  empName: '20523:นายจิตรกาล เกตุประสาท',
  department: '9003.2:PRODUCTION TECHNOLOGY',
  empType: 'รายเดือน',
  bankAccount: '4202160927',
  periodMonth: 'กันยายน',
  periodRound: 'งวดจ่ายเงินเดือน',
  payDate: '30/09/2569',
  userName: 'JITTRAKAN',
  printDate: '13/09/2569'
};

let baseSalary = 20000;
let transportationAllowance = 1230;
let diligenceFullAmount = 650;
let ssoDeduction = 875;

let taxAllowances = {
  personal: 60000,
  expenses: 100000,
  sso: 8908
};

let accumulatedData = {
  income: 174389.00,
  taxEmp: 0.00,
  taxComp: 0.00,
  sso: 5408.00,
  fundEmp: 0.00,
  fundComp: 0.00
};

// ==========================================================================
// 6. การสลับหน้าจอ (Tab Navigation)
// ==========================================================================
function switchView(viewName) {
  document.querySelectorAll('.view-panel').forEach(el => el.classList.remove('active-panel'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

  const targetPanel = document.getElementById('view-' + viewName);
  const targetBtn = document.getElementById('tab-' + viewName + '-btn');

  if (targetPanel) targetPanel.classList.add('active-panel');
  if (targetBtn) targetBtn.classList.add('active');

  if (viewName === 'slip') {
    updatePayslip();
  }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function prepareAndPrint() {
  switchView('slip');
  setTimeout(() => {
    window.print();
  }, 250);
}

// ==========================================================================
// 7. ตารางลงเวลา IN-OUT (Time Tracking & OT Sheet)
// ==========================================================================
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
    // วันทำงานปกติ: OT นับเป็น OT 1.5
    row.ot = getOTFromTimeOut(row.outTime);
    row.hol = 0;
    row.otHol = 0;
  } else {
    // วันหยุด / วันอาทิตย์ / เสาร์หยุด / นักขัตฤกษ์:
    row.ot = 0;

    const parts = row.outTime.split(':');
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const totalMinutes = h * 60 + m;

    const normalEndMinutes = 17 * 60; // 17:00
    const startMinutes = 8 * 60;      // 08:00
    const lunchStart = 12 * 60;       // 12:00
    const lunchEnd = 13 * 60;         // 13:00

    if (totalMinutes <= normalEndMinutes) {
      // ทำงานช่วง 08:00 - 17:00 ในวันหยุด นับเป็น Holiday (OT 1.0)
      let workMins = 0;
      if (totalMinutes <= lunchStart) {
        workMins = totalMinutes - startMinutes;
      } else if (totalMinutes <= lunchEnd) {
        workMins = lunchStart - startMinutes; // 4 ชม.
      } else {
        workMins = (totalMinutes - startMinutes) - 60; // หักพักเที่ยง 1 ชม.
      }
      if (workMins < 0) workMins = 0;
      row.hol = Math.floor((workMins / 60) * 2) / 2; // ปัดขั้น 0.5 ชม.
      row.otHol = 0;
    } else {
      // ทำงานเกิน 17:00 ในวันหยุด (ครบ 8 ชม. Holiday OT 1.0 + OT หลัง 17:00 เป็น OT Hol 3.0)
      row.hol = 8.0;
      row.otHol = getOTFromTimeOut(row.outTime);
    }
  }

  // วันทำโอที: OT >= 1 ชม. (รวมทั้งวันปกติและวันหยุด)
  row.otDay = (row.ot >= 1 || row.hol >= 1 || row.otHol >= 1) ? 1 : 0;
}

function buildAttendanceTable() {
  const tbody = document.getElementById('attendanceTbody');
  tbody.innerHTML = '';

  const thaiDays = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];

  // ตัวเลือกเวลาออกสำหรับวันอาทิตย์ / วันหยุด (ตั้งแต่ 08:30 ทุก 30 นาที จนถึง 19:20)
  const sundayOutOptions = [
    '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', '12:00',
    '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00',
    '16:30', '17:00', '17:30', '17:50', '18:00', '18:20', '18:30', '18:50',
    '19:00', '19:20'
  ];

  // ตัวเลือกเวลาออกสำหรับวันทำงานปกติ
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
}

function onAttendanceManualOT(index, field, val) {
  currentAttendance[index][field] = parseFloat(val) || 0;
  const row = currentAttendance[index];
  row.otDay = (row.ot >= 1 || row.hol >= 1 || row.otHol >= 1) ? 1 : 0;
  recalculateAttendanceTotals();
  recalculateSalary();
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

  // แสดงผลในแถวสรุปท้ายตาราง IN-OUT
  document.getElementById('attTotalOT15').innerText = totalOT15.toFixed(1);
  document.getElementById('attTotalHol').innerText = totalHol.toFixed(1);
  document.getElementById('attTotalOTHol').innerText = totalOTHol.toFixed(1);
  document.getElementById('attTotalTargetDays').innerText = totalTargetDays;
  document.getElementById('attTotalCameDays').innerText = totalCameDays;
  document.getElementById('attTotalOTDays').innerText = totalOTDays;

  return { totalOT15, totalHol, totalOTHol, totalTargetDays, totalCameDays, totalOTDays };
}

// โหลดข้อมูลตัวอย่างจากไฟล์ Excel
function loadAttendanceFromExcel() {
  currentAttendance = JSON.parse(JSON.stringify(EXCEL_SAMPLE_ATTENDANCE));
  salaryProfile.periodMonth = 'กันยายน';
  salaryProfile.payDate = '30/09/2569';
  buildAttendanceTable();
  recalculateSalary();
  alert('โหลดข้อมูลลงเวลา 31 วันจากไฟล์ Excel DATA SALARY 210826-200926.xlsx สำเร็จแล้ว!');
}

// โหลดข้อมูลตามสลิปสิงหาคม 2569
function loadAttendanceFromAugSlip() {
  // สร้างตาราง 21 ก.ค. - 20 ส.ค. 2569
  generateAttendanceForPeriod('2026-07-21', '2026-08-20');
  // ปรับยอด OT ให้ตรงกับสลิป (OT1.5=26.5, OT1=8, OT3=2)
  let assigned15 = 26.5;
  let assigned1 = 8.0;
  let assigned3 = 2.0;

  currentAttendance.forEach(row => {
    const dayInfo = getOrbrayDayInfo(row.date);
    if (dayInfo.isWorkDay) {
      row.workDay = 1;
      row.come = 1;
      row.inTime = '08:00';
      row.outTime = '17:00';
    } else {
      row.workDay = 0;
      row.come = 0;
      row.inTime = '';
      row.outTime = '';
    }
  });

  // เสาร์ 8 ส.ค. ทำงานวันหยุด 8 ชม. + OT 2 ชม.
  let satHol = currentAttendance.find(r => r.date === '2026-08-08');
  if (satHol) {
    satHol.inTime = '08:00';
    satHol.outTime = '19:20';
    satHol.come = 1;
    satHol.hol = 8;
    satHol.otHol = 2;
    satHol.workDay = 1;
    satHol.otDay = 1;
  }

  // กระจาย OT 1.5 ให้ได้ 26.5 ชม.
  let needed = 26.5;
  for (let r of currentAttendance) {
    if (needed <= 0) break;
    const dayInfo = getOrbrayDayInfo(r.date);
    if (dayInfo.isWorkDay) {
      if (needed >= 2) {
        r.outTime = '19:20';
        r.ot = 2.0;
        r.otDay = 1;
        needed -= 2;
      } else if (needed >= 1.5) {
        r.outTime = '18:50';
        r.ot = 1.5;
        r.otDay = 1;
        needed -= 1.5;
      } else if (needed >= 1) {
        r.outTime = '18:20';
        r.ot = 1.0;
        r.otDay = 1;
        needed -= 1;
      } else if (needed >= 0.5) {
        r.outTime = '17:50';
        r.ot = 0.5;
        needed -= 0.5;
      }
    }
  }

  salaryProfile.periodMonth = 'สิงหาคม';
  salaryProfile.payDate = '31/08/2569';

  buildAttendanceTable();
  recalculateSalary();
  alert('โหลดข้อมูลสลิปเงินเดือนประจำเดือนสิงหาคม 2569 เรียบร้อยแล้ว!');
}

// เติมวันทำงานปกติครบทุกวัน
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
}

// ล้างเวลาทั้งหมด
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
}

// สร้างตารางช่วงวันที่สำหรับรอบที่เลือก
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

// ==========================================================================
// 8. ระบบคำนวณเงินเดือนตามสูตรชีต CAL SALARY ใน Excel
// ==========================================================================
function recalculateSalary() {
  const totals = recalculateAttendanceTotals();

  const B2_targetDays = totals.totalTargetDays; // จำนวนวัน (ในเดือนนี้)
  const D2_cameDays = totals.totalCameDays;     // จำนวนวัน (ที่มาทำงาน)
  const F2_otDays = totals.totalOTDays;         // จำนวนวัน (ทำโอที >= 1 ชม.)
  const G2_ot15Hours = totals.totalOT15;        // จำนวนชั่วโมง (OT*1.5)
  const H2_ot1Hours = totals.totalHol;          // จำนวนชั่วโมง (OT*1 / Holiday)
  const I2_ot3Hours = totals.totalOTHol;        // จำนวนชั่วโมง (OT*3 / OT Holiday)

  // C2: เงินที่ได้ต่อวัน = A2 / B2
  const C2_dailyRate = B2_targetDays > 0 ? (baseSalary / B2_targetDays) : 0;

  // E2: เงินที่ได้ทั้งหมด (ไม่รวมโอที) = C2 * D2
  const E2_actualSalary = (B2_targetDays > 0 && D2_cameDays >= B2_targetDays)
    ? baseSalary
    : Math.round(C2_dailyRate * D2_cameDays);

  // J2: ค่าอาหาร = 20 * D2
  const J2_foodAllowance = 20 * D2_cameDays;

  // K2: อาหารโอที = 15 * F2
  const K2_otMealAllowance = 15 * F2_otDays;

  // L2: เงินช่วยเหลือค่าเดินทาง = 1,230 บาท
  const L2_travelAllowance = transportationAllowance;

  // M2: เบี้ยขยัน = IF(B2=D2, 650, 0)
  const M2_diligenceAllowance = (B2_targetDays > 0 && D2_cameDays >= B2_targetDays) ? diligenceFullAmount : 0;

  // อัตราต่อชั่วโมงสำหรับการคิด OT: (A2 / 30 / 8) = A2 / 240
  const hourlyOTRate = baseSalary / 240;

  // N2: OT*1.5 = (A2 / 240) * (1.5 * G2) ปัดเศษตามความเหมาะสม
  const N2_ot15Amount = Math.round(hourlyOTRate * 1.5 * G2_ot15Hours);

  // O2: OT*1 = (A2 / 240) * (1 * H2) ปัดเศษตามความเหมาะสม (เช่น 666.67 -> 667)
  const O2_ot1Amount = Math.round(hourlyOTRate * 1.0 * H2_ot1Hours);

  // P2: OT*3 = (A2 / 240) * (3 * I2) ปัดเศษตามความเหมาะสม
  const P2_ot3Amount = Math.round(hourlyOTRate * 3.0 * I2_ot3Hours);

  // Q2: ประกันสังคม = 875 บาท
  const Q2_sso = ssoDeduction;

  // รวมเงินได้ (Earnings)
  const totalEarnings = E2_actualSalary + J2_foodAllowance + K2_otMealAllowance + L2_travelAllowance + 
                        M2_diligenceAllowance + N2_ot15Amount + O2_ot1Amount + P2_ot3Amount;

  // รวมเงินหัก (Deductions)
  const totalDeductions = Q2_sso;

  // เงินได้สุทธิ (Net Pay)
  const netPay = totalEarnings - totalDeductions;

  // อัปเดตค่าแสดงผลใน UI หน้าคำนวณ
  document.getElementById('calcBaseSalary').innerText = formatCurrency(baseSalary);
  document.getElementById('calcTargetDays').innerText = B2_targetDays;
  document.getElementById('calcDailyRate').innerText = formatCurrency(C2_dailyRate);
  document.getElementById('calcCameDays').innerText = D2_cameDays;
  document.getElementById('calcActualSalary').innerText = formatCurrency(E2_actualSalary);

  document.getElementById('calcFoodDays').innerText = D2_cameDays;
  document.getElementById('calcFoodTotal').innerText = formatCurrency(J2_foodAllowance);

  document.getElementById('calcOTMealDays').innerText = F2_otDays;
  document.getElementById('calcOTMealTotal').innerText = formatCurrency(K2_otMealAllowance);

  document.getElementById('calcTravelTotal').innerText = formatCurrency(L2_travelAllowance);

  document.getElementById('calcDiligenceStatus').innerText = (B2_targetDays > 0 && D2_cameDays >= B2_targetDays) ? 'ครบวันทำงาน' : 'ไม่ครบวันทำงาน';
  document.getElementById('calcDiligenceTotal').innerText = formatCurrency(M2_diligenceAllowance);

  document.getElementById('calcHourlyOTRate').innerText = formatCurrency(hourlyOTRate);
  document.getElementById('calcOT15Hours').innerText = G2_ot15Hours.toFixed(1);
  document.getElementById('calcOT15Total').innerText = formatCurrency(N2_ot15Amount);

  document.getElementById('calcOT1Hours').innerText = H2_ot1Hours.toFixed(1);
  document.getElementById('calcOT1Total').innerText = formatCurrency(O2_ot1Amount);

  document.getElementById('calcOT3Hours').innerText = I2_ot3Hours.toFixed(1);
  document.getElementById('calcOT3Total').innerText = formatCurrency(P2_ot3Amount);

  document.getElementById('calcSSOTotal').innerText = formatCurrency(Q2_sso);

  // สรุปบน Dashboard Cards
  document.getElementById('dashTotalIncome').innerText = formatCurrency(totalEarnings);
  document.getElementById('dashTotalDeduct').innerText = formatCurrency(totalDeductions);
  document.getElementById('dashNetPay').innerText = formatCurrency(netPay);

  // ซิงค์ไปยังใบแจ้งรายได้
  updatePayslip({
    salary: E2_actualSalary,
    food: J2_foodAllowance,
    otMeal: K2_otMealAllowance,
    travel: L2_travelAllowance,
    diligence: M2_diligenceAllowance,
    ot15: N2_ot15Amount,
    ot1: O2_ot1Amount,
    ot3: P2_ot3Amount,
    sso: Q2_sso,
    totalEarnings,
    totalDeductions,
    netPay
  });
}

// ==========================================================================
// 9. อัปเดตข้อมูลใบแจ้งรายได้ (Payslip Preview & Print)
// ==========================================================================
function updatePayslip(calcData) {
  // ซิงค์ข้อมูลส่วนหัว
  document.getElementById('slipCompanyName').innerText = salaryProfile.companyName;
  document.getElementById('slipEmpName').innerText = salaryProfile.empName;
  document.getElementById('slipDepartment').innerText = salaryProfile.department;
  document.getElementById('slipEmpType').innerText = salaryProfile.empType;
  document.getElementById('slipBankAccount').innerText = salaryProfile.bankAccount;
  document.getElementById('slipPeriodMonth').innerText = salaryProfile.periodMonth;
  document.getElementById('slipPeriodRound').innerText = salaryProfile.periodRound;
  document.getElementById('slipPayDate').innerText = salaryProfile.payDate;
  document.getElementById('slipUserName').innerText = salaryProfile.userName;
  document.getElementById('slipPrintDate').innerText = salaryProfile.printDate;

  if (!calcData) return;

  // คอลัมน์ที่ 1: รายการเงินได้
  const tbodyIncome = document.getElementById('slipIncomeRows');
  tbodyIncome.innerHTML = '';

  const incomeItems = [
    { name: 'อาหาร', amount: calcData.food },
    { name: 'อาหารโอที', amount: calcData.otMeal },
    { name: 'เงินช่วยเหลือค่าเดินทาง', amount: calcData.travel },
    { name: 'เบี้ยขยัน', amount: calcData.diligence },
    { name: 'OT*1.5', amount: calcData.ot15 },
    { name: 'OT*1', amount: calcData.ot1 },
    { name: 'OT*3', amount: calcData.ot3 },
    { name: 'เงินเดือน', amount: calcData.salary }
  ];

  incomeItems.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${item.name}</td>
      <td class="text-right">${formatCurrency(item.amount)}</td>
    `;
    tbodyIncome.appendChild(tr);
  });

  // คอลัมน์ที่ 2: รายการเงินหัก
  const tbodyDeduct = document.getElementById('slipDeductRows');
  tbodyDeduct.innerHTML = '';
  const trDeduct = document.createElement('tr');
  trDeduct.innerHTML = `
    <td>ประกันสังคมหักจากพนักงาน</td>
    <td class="text-right text-deduct-highlight">${formatCurrency(calcData.sso)}</td>
  `;
  tbodyDeduct.appendChild(trDeduct);

  // สรุปแถบด้านล่าง
  document.getElementById('slipTotalIncome').innerText = formatCurrency(calcData.totalEarnings);
  document.getElementById('slipTotalDeduct').innerText = formatCurrency(calcData.totalDeductions);
  document.getElementById('slipNetPay').innerText = formatCurrency(calcData.netPay);

  // คอลัมน์ที่ 3: ค่าลดหย่อน
  document.getElementById('slipTaxDedPerson').innerText = formatCurrency(taxAllowances.personal);
  document.getElementById('slipTaxDedExpense').innerText = formatCurrency(taxAllowances.expenses);
  document.getElementById('slipTaxDedSSO').innerText = formatCurrency(taxAllowances.sso);
  document.getElementById('slipTotalTaxDed').innerText = formatCurrency(taxAllowances.personal + taxAllowances.expenses + taxAllowances.sso);

  // คอลัมน์ที่ 3: รายการเงินสะสม
  document.getElementById('slipAccIncome').innerText = formatCurrency(accumulatedData.income);
  document.getElementById('slipAccTaxEmp').innerText = formatCurrency(accumulatedData.taxEmp);
  document.getElementById('slipAccTaxComp').innerText = formatCurrency(accumulatedData.taxComp);
  document.getElementById('slipAccSSO').innerText = formatCurrency(accumulatedData.sso);
  document.getElementById('slipAccFundEmp').innerText = formatCurrency(accumulatedData.fundEmp);
  document.getElementById('slipAccFundComp').innerText = formatCurrency(accumulatedData.fundComp);
}

// ==========================================================================
// 10. สร้างปฏิทินบริษัท Orbray 2569 ครบทั้ง 12 เดือน (Orbray Calendar Grid)
// ==========================================================================
function renderOrbrayCalendarGrid() {
  const container = document.getElementById('orbrayCalendarGrid');
  if (!container) return;
  container.innerHTML = '';

  const monthNamesEn = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  const monthNamesTh = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

  for (let m = 0; m < 12; m++) {
    const card = document.createElement('div');
    card.className = 'cal-month-card';

    const monthHeader = document.createElement('div');
    monthHeader.className = 'cal-month-header';
    monthHeader.innerHTML = `
      <div>
        <h3>${monthNamesEn[m]}</h3>
        <span class="th-name">${monthNamesTh[m]} 2569</span>
      </div>
      <button type="button" class="btn btn-sm btn-outline-primary" onclick="selectPeriodMonth(${m + 1})">
        คำนวณงวดนี้
      </button>
    `;
    card.appendChild(monthHeader);

    // ตารางปฏิทิน
    const table = document.createElement('table');
    table.className = 'cal-mini-table';
    table.innerHTML = `
      <thead>
        <tr>
          <th class="col-sun">S</th>
          <th>M</th>
          <th>T</th>
          <th>W</th>
          <th>TH</th>
          <th>F</th>
          <th>S</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

    const tbody = table.querySelector('tbody');
    const firstDay = new Date(2026, m, 1);
    const lastDay = new Date(2026, m + 1, 0);
    const startDayOfWeek = firstDay.getDay(); // 0 = Sun
    const totalDays = lastDay.getDate();

    let tr = document.createElement('tr');
    // วันว่างก่อนวันที่ 1
    for (let i = 0; i < startDayOfWeek; i++) {
      tr.appendChild(document.createElement('td'));
    }

    for (let day = 1; day <= totalDays; day++) {
      const dStr = `2026-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const dayInfo = getOrbrayDayInfo(dStr);

      const td = document.createElement('td');
      td.className = 'cal-day-cell';

      const span = document.createElement('span');
      span.className = 'cal-day-num';
      span.innerText = day;

      // แต่งสไตล์ตามประเภทวันให้ตรงกับภาพปฏิทิน
      if (dayInfo.type === 'NATIONAL_HOLIDAY') {
        span.classList.add('day-national-holiday');
        td.title = dayInfo.name;
      } else if (dayInfo.type === 'MEMORIAL_HOLIDAY') {
        span.classList.add('day-memorial-holiday');
        td.title = dayInfo.name;
      } else if (dayInfo.type === 'ORBRAY_WORKDAY') {
        span.classList.add('day-orbray-work');
        td.title = dayInfo.name;
      } else if (dayInfo.type === 'SATURDAY_HOLIDAY' || dayInfo.type === 'BRIDGE_HOLIDAY') {
        span.classList.add('day-sat-holiday');
        td.title = dayInfo.name;
      } else if (dayInfo.type === 'SUNDAY') {
        span.classList.add('day-sunday');
      } else if (dayInfo.type === 'WORKING_SATURDAY') {
        span.classList.add('day-working-sat');
        td.title = 'วันเสาร์ทำงานปกติ';
      }

      td.appendChild(span);
      tr.appendChild(td);

      if ((startDayOfWeek + day) % 7 === 0 || day === totalDays) {
        tbody.appendChild(tr);
        tr = document.createElement('tr');
      }
    }

    card.appendChild(table);
    container.appendChild(card);
  }
}

// เลือกงวดเดือนจากปฏิทิน
function selectPeriodMonth(monthNumber) {
  const periodId = `2026-${String(monthNumber).padStart(2, '0')}`;
  const found = PAYROLL_PERIODS_2026.find(p => p.id === periodId);
  if (!found) return;

  currentPeriod = found;
  document.getElementById('payrollPeriodSelect').value = found.id;
  onPayrollPeriodSelectChange();
  switchView('timesheet');
}

function onPayrollPeriodSelectChange() {
  const select = document.getElementById('payrollPeriodSelect');
  const found = PAYROLL_PERIODS_2026.find(p => p.id === select.value);
  if (!found) return;

  currentPeriod = found;
  salaryProfile.periodMonth = found.monthName.split(' ')[0];
  salaryProfile.payDate = found.payDate;

  // ถ้าเลือกงวดกันยายน ให้โหลดข้อมูลตัวอย่างจากไฟล์ Excel
  if (found.id === '2026-09') {
    currentAttendance = JSON.parse(JSON.stringify(EXCEL_SAMPLE_ATTENDANCE));
  } else {
    // สร้างตารางวันสำหรับรอบวิกนั้น
    generateAttendanceForPeriod(found.startDate, found.endDate);
    autoFillNormalWorkdays();
  }

  buildAttendanceTable();
  recalculateSalary();
}

// ==========================================================================
// 11. ฟังก์ชันช่วยเหลือและเริ่มต้นระบบ (Helpers & Initialization)
// ==========================================================================
function formatCurrency(val) {
  const num = parseFloat(val);
  if (isNaN(num)) return '0.00';
  return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function updateBaseSalaryInput(val) {
  baseSalary = parseFloat(val) || 20000;
  recalculateSalary();
}

window.addEventListener('DOMContentLoaded', () => {
  // สร้างตัวเลือกงวดการจ่าย
  const select = document.getElementById('payrollPeriodSelect');
  if (select) {
    select.innerHTML = '';
    PAYROLL_PERIODS_2026.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.id;
      opt.innerText = `งวด ${p.monthName} (ตัดรอบ ${p.startDate.split('-').slice(1).join('/')} - ${p.endDate.split('-').slice(1).join('/')} | จ่ายสิ้นเดือน ${p.payDate})`;
      if (p.id === currentPeriod.id) opt.selected = true;
      select.appendChild(opt);
    });
  }

  // สร้างตาราง IN-OUT
  buildAttendanceTable();

  // สร้างปฏิทิน Orbray 2569
  renderOrbrayCalendarGrid();

  // คำนวณเงินเดือนเริ่มต้น
  recalculateSalary();
});
