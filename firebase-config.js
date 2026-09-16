/**
 * FIREBASE-CONFIG.JS
 * ระบบจัดการผู้ใช้งานและสิทธิ์การเข้าถึงด้วยรหัส PIN (4-8 หลัก)
 * - ADMIN PIN: 20523 (ผู้ดูแลระบบ สามารถแก้ไขข้อมูลของตนเองและผู้ใช้อื่นได้ทั้งหมด)
 * - รองรับ Multi-User, PIN Login, On-Screen Keypad, User Switcher, และ Firestore Cloud Sync
 */

const STORAGE_KEY_FIREBASE_CFG = 'custom_firebase_config_v1';
const STORAGE_KEY_AUTH_USER = 'salary_auth_current_user_v1';
const STORAGE_KEY_USERS_LIST = 'salary_app_users_v1';
const STORAGE_KEY_ACTIVE_EDITING_USER = 'salary_active_editing_user_id_v1';

// ค่าเริ่มต้นคอนฟิก Firebase
const DEFAULT_FIREBASE_CONFIG = {
  apiKey: "AIzaSyDemoPlaceholderOnlyForTesting123456",
  authDomain: "caculation-salary-app.firebaseapp.com",
  projectId: "caculation-salary-app",
  storageBucket: "caculation-salary-app.appspot.com",
  messagingSenderId: "1234567890",
  appId: "1:1234567890:web:abcdef123456"
};

// ข้อมูลผู้ใช้เริ่มต้นในระบบ
const DEFAULT_SYSTEM_USERS = [
  {
    id: 'user_admin',
    name: 'ADMIN',
    companyName: 'CACULATION SALARY',
    empCode: '20523',
    department: 'ผู้ดูแลระบบสูงสุด',
    pin: '20523',
    role: 'admin',
    avatar: '👑',
    createdAt: '2026-01-01T00:00:00.000Z'
  },
  {
    id: 'user_staff1',
    name: 'พนักงาน 01',
    companyName: 'CACULATION SALARY',
    empCode: 'EMP-001',
    department: 'ฝ่ายผลิต/ปฏิบัติการ',
    pin: '1111',
    role: 'user',
    avatar: '👤',
    createdAt: '2026-01-01T00:00:00.000Z'
  },
  {
    id: 'user_staff2',
    name: 'พนักงาน 02',
    companyName: 'CACULATION SALARY',
    empCode: 'EMP-002',
    department: 'ฝ่ายเทคนิค/ซ่อมบำรุง',
    pin: '2222',
    role: 'user',
    avatar: '👤',
    createdAt: '2026-01-01T00:00:00.000Z'
  }
];

let firebaseApp = null;
let firebaseAuth = null;
let firebaseDb = null;
let currentUser = null;
let isFirebaseOnline = false;
let activeEditingUserId = 'user_admin';

// Callback แจ้งเตือนเมื่อสถานะผู้ใช้เปลี่ยน
let onAuthStateListeners = [];

function addAuthStateListener(callback) {
  if (typeof callback === 'function') {
    onAuthStateListeners.push(callback);
    callback(currentUser, isFirebaseOnline);
  }
}

function notifyAuthState(user) {
  currentUser = user;
  if (user) {
    localStorage.setItem(STORAGE_KEY_AUTH_USER, JSON.stringify(user));
    // ถ้าไม่มี active editing user หรือ user ไม่ใช่ admin ให้ตั้ง active editing user เป็น user ตัวเอง
    if (user.role !== 'admin' || !localStorage.getItem(STORAGE_KEY_ACTIVE_EDITING_USER)) {
      setActiveEditingUserId(user.id);
    }
  } else {
    localStorage.removeItem(STORAGE_KEY_AUTH_USER);
    setActiveEditingUserId('user_admin');
  }

  onAuthStateListeners.forEach(cb => {
    try { cb(currentUser, isFirebaseOnline); } catch (err) { console.error('Auth state cb error:', err); }
  });
}

/**
 * ดึงรายชื่อผู้ใช้ทั้งหมดในระบบ
 */
function getAllSystemUsers() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_USERS_LIST);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // ตรวจสอบว่ามี ADMIN PIN 20523 อยู่เสมอ
        const hasAdmin = parsed.some(u => u.pin === '20523' && u.role === 'admin');
        if (!hasAdmin) {
          parsed.unshift(DEFAULT_SYSTEM_USERS[0]);
        }
        // ตรวจสอบและใส่ค่าเริ่มต้นสำหรับข้อมูลที่เพิ่มเข้ามาใหม่ (Company, EmpCode)
        let modified = !hasAdmin;
        parsed.forEach(u => {
          if (!u.companyName) {
            u.companyName = 'CACULATION SALARY';
            modified = true;
          } else {
            u.companyName = u.companyName.toUpperCase();
          }
          if (!u.empCode) {
            u.empCode = u.id === 'user_admin' ? '20523' : ('EMP-' + (u.pin || '001'));
            modified = true;
          } else {
            u.empCode = u.empCode.toUpperCase();
          }
          if (u.bankAccount) {
            delete u.bankAccount;
            modified = true;
          }
        });
        if (modified) {
          saveSystemUsers(parsed);
        }
        return parsed;
      }
    }
  } catch (e) {
    console.warn('Cannot parse system users, seeding defaults', e);
  }

  // ถ้ายังไม่มี ให้บันทึกชุดเริ่มต้น
  saveSystemUsers(DEFAULT_SYSTEM_USERS);
  return DEFAULT_SYSTEM_USERS;
}

/**
 * บันทึกรายชื่อผู้ใช้ทั้งหมด
 */
function saveSystemUsers(users) {
  try {
    localStorage.setItem(STORAGE_KEY_USERS_LIST, JSON.stringify(users));
    if (isFirebaseOnline && firebaseDb) {
      firebaseDb.collection('app_system').doc('users_directory').set({
        users: users,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true }).catch(err => console.warn('Firestore users sync warning:', err));
    }
    return true;
  } catch (e) {
    console.error('Failed to save system users', e);
    return false;
  }
}

/**
 * ตรวจสอบและดึง Active Editing User ID
 */
function getActiveEditingUserId() {
  const saved = localStorage.getItem(STORAGE_KEY_ACTIVE_EDITING_USER);
  if (saved) return saved;
  if (currentUser) return currentUser.id;
  return 'user_admin';
}

/**
 * เปลี่ยน User ID ที่กำลังแก้ไขข้อมูลอยู่ (สำหรับ Admin สลับดู/แก้ไขข้อมูลพนักงานคนอื่น)
 */
function setActiveEditingUserId(userId) {
  activeEditingUserId = userId;
  localStorage.setItem(STORAGE_KEY_ACTIVE_EDITING_USER, userId);
  return activeEditingUserId;
}

/**
 * ดึงข้อมูลผู้ใช้ที่กำลังถูกแก้ไขข้อมูลอยู่
 */
function getActiveEditingUser() {
  const targetId = getActiveEditingUserId();
  const users = getAllSystemUsers();
  return users.find(u => u.id === targetId) || users[0] || DEFAULT_SYSTEM_USERS[0];
}

/**
 * เข้าสู่ระบบด้วยรหัส PIN (4-8 หลัก)
 */
async function loginWithPin(pin) {
  if (!pin) {
    return { success: false, error: 'กรุณาระบุรหัส PIN' };
  }

  const cleanPin = String(pin).trim();
  if (cleanPin.length < 4 || cleanPin.length > 8) {
    return { success: false, error: 'รหัส PIN ต้องมีความยาว 4–8 หลัก' };
  }

  const users = getAllSystemUsers();
  const foundUser = users.find(u => u.pin === cleanPin);

  if (!foundUser) {
    return { 
      success: false, 
      error: 'รหัส PIN ไม่ถูกต้อง กรุณาตรวจสอบและลองใหม่อีกครั้ง' 
    };
  }

  // ล็อกอินสำเร็จ
  notifyAuthState(foundUser);
  
  // กำหนด active editing user
  setActiveEditingUserId(foundUser.id);

  return { success: true, user: foundUser };
}

/**
 * เพิ่มผู้ใช้งานใหม่ (เฉพาะ Admin)
 */
function registerNewUser(name, pin, role = 'user', department = '', companyName = 'CACULATION SALARY', empCode = '') {
  if (!name || !name.trim()) {
    return { success: false, error: 'กรุณาระบุชื่อพนักงาน' };
  }

  const cleanPin = String(pin).trim();
  if (!/^\d{4,8}$/.test(cleanPin)) {
    return { success: false, error: 'รหัส PIN ต้องเป็นตัวเลข 4–8 หลักเท่านั้น' };
  }

  const users = getAllSystemUsers();
  if (users.some(u => u.pin === cleanPin)) {
    return { success: false, error: 'รหัส PIN นี้ถูกใช้งานแล้ว กรุณาเลือกรหัสอื่น' };
  }

  const newUser = {
    id: 'user_' + Date.now().toString(36) + '_' + Math.random().toString(36).substring(2, 6),
    name: name.trim(),
    pin: cleanPin,
    role: role,
    avatar: role === 'admin' ? '👑' : '👤',
    companyName: (companyName || 'CACULATION SALARY').trim().toUpperCase(),
    empCode: (empCode || ('EMP-' + cleanPin)).trim().toUpperCase(),
    department: department.trim() || 'แผนกทั่วไป',
    createdAt: new Date().toISOString()
  };

  users.push(newUser);
  saveSystemUsers(users);

  return { success: true, user: newUser };
}

/**
 * อัปเดตข้อมูลผู้ใช้งาน (แก้ไขชื่อ, PIN, แผนก, บริษัท, รหัสพนักงาน)
 */
function updateSystemUser(userId, updateData) {
  const users = getAllSystemUsers();
  const idx = users.findIndex(u => u.id === userId);
  if (idx === -1) {
    return { success: false, error: 'ไม่พบผู้ใช้นี้ในระบบ' };
  }

  // ตรวจสอบ PIN ซ้ำ (ถ้ามีการเปลี่ยน PIN)
  if (updateData.pin !== undefined && updateData.pin !== null && updateData.pin !== '') {
    const cleanPin = String(updateData.pin).trim();
    if (!/^\d{4,8}$/.test(cleanPin)) {
      return { success: false, error: 'รหัส PIN ต้องเป็นตัวเลข 4–8 หลัก' };
    }
    const pinOwner = users.find(u => u.pin === cleanPin && u.id !== userId);
    if (pinOwner) {
      return { success: false, error: 'รหัส PIN นี้ถูกใช้โดยพนักงานคนอื่นแล้ว' };
    }
    users[idx].pin = cleanPin;
  }

  if (updateData.name !== undefined) users[idx].name = updateData.name.trim();
  if (updateData.department !== undefined) users[idx].department = updateData.department.trim();
  if (updateData.companyName !== undefined) users[idx].companyName = updateData.companyName.trim().toUpperCase();
  if (updateData.empCode !== undefined) users[idx].empCode = updateData.empCode.trim().toUpperCase();
  if (updateData.role && users[idx].id !== 'user_admin') {
    users[idx].role = updateData.role;
    users[idx].avatar = updateData.role === 'admin' ? '👑' : '👤';
  }

  saveSystemUsers(users);

  // ถ้าอัปเดต user ที่กำลังล็อกอินอยู่ ให้ sync ทันที
  if (currentUser && currentUser.id === userId) {
    notifyAuthState(users[idx]);
  }

  return { success: true, user: users[idx] };
}

/**
 * อัปเดตข้อมูลโปรไฟล์บัญชีของผู้ใช้ปัจจุบัน
 */
function updateCurrentAccountProfile(profileData) {
  const targetId = (currentUser) ? currentUser.id : getActiveEditingUserId();
  return updateSystemUser(targetId, profileData);
}

/**
 * ลบผู้ใช้งาน (ยกเว้น ADMIN ตัวหลัก)
 */
function deleteSystemUser(userId) {
  if (userId === 'user_admin') {
    return { success: false, error: 'ไม่สามารถลบบัญชี ADMIN สูงสุดได้' };
  }

  let users = getAllSystemUsers();
  users = users.filter(u => u.id !== userId);
  saveSystemUsers(users);

  // ถ้าลบ user ที่กำลังถูก active อยู่ ให้รีเซ็ตกลับมาเป็น ADMIN
  if (getActiveEditingUserId() === userId) {
    setActiveEditingUserId('user_admin');
  }

  return { success: true };
}

/**
 * ออกจากระบบ
 */
async function logoutCurrentUser() {
  notifyAuthState(null);
  setActiveEditingUserId('user_admin');
  return { success: true };
}

/**
 * ดึงการตั้งค่า Firebase
 */
function getActiveFirebaseConfig() {
  try {
    const custom = localStorage.getItem(STORAGE_KEY_FIREBASE_CFG);
    if (custom) return JSON.parse(custom);
  } catch (e) {
    console.warn('Cannot parse stored firebase config', e);
  }
  return DEFAULT_FIREBASE_CONFIG;
}

/**
 * บันทึกการตั้งค่า Firebase
 */
function saveFirebaseConfig(cfg) {
  try {
    localStorage.setItem(STORAGE_KEY_FIREBASE_CFG, JSON.stringify(cfg));
    return true;
  } catch (e) {
    console.error('Failed to save firebase config', e);
    return false;
  }
}

/**
 * เริ่มต้นการทำงานของ Firebase
 */
function initFirebaseApp() {
  const cfg = getActiveFirebaseConfig();

  if (typeof firebase !== 'undefined' && firebase.initializeApp) {
    try {
      if (!firebase.apps.length) {
        firebaseApp = firebase.initializeApp(cfg);
      } else {
        firebaseApp = firebase.app();
      }
      if (firebase.auth) firebaseAuth = firebase.auth();
      if (firebase.firestore) firebaseDb = firebase.firestore();
      isFirebaseOnline = true;
      console.log('Firebase initialized successfully!');
    } catch (err) {
      console.warn('Firebase init note (Local PIN Mode active):', err.message);
      isFirebaseOnline = false;
    }
  } else {
    isFirebaseOnline = false;
  }

  // ตรวจสอบผู้ใช้ที่เคยล็อกอินค้างไว้
  checkStoredCurrentUser();
}

function checkStoredCurrentUser() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_AUTH_USER);
    if (raw) {
      const u = JSON.parse(raw);
      const all = getAllSystemUsers();
      const match = all.find(x => x.id === u.id);
      if (match) {
        notifyAuthState(match);
        return;
      }
    }
  } catch (e) {}

  notifyAuthState(null);
}

/**
 * Cloud Sync สโคปตาม User ID
 */
async function syncDataToCloud(docKey, data, userId = null) {
  if (!isFirebaseOnline || !firebaseDb) return false;
  const targetId = userId || getActiveEditingUserId();
  try {
    await firebaseDb.collection('user_salaries').doc(targetId).set({
      [docKey]: data,
      lastUpdated: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    return true;
  } catch (err) {
    console.warn('Cloud sync error:', err.message);
    return false;
  }
}

/**
 * Cloud Load สโคปตาม User ID
 */
async function loadDataFromCloud(docKey, userId = null) {
  if (!isFirebaseOnline || !firebaseDb) return null;
  const targetId = userId || getActiveEditingUserId();
  try {
    const doc = await firebaseDb.collection('user_salaries').doc(targetId).get();
    if (doc.exists && doc.data()[docKey]) {
      return doc.data()[docKey];
    }
  } catch (err) {
    console.warn('Cloud load error:', err.message);
  }
  return null;
}

// เริ่มต้นระบบเมื่อโหลดเสร็จ
window.addEventListener('DOMContentLoaded', () => {
  initFirebaseApp();
});
