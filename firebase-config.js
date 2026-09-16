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

// ข้อมูลผู้ใช้เริ่มต้นในระบบ (มีเพียง ADMIN หลักคนเดียว)
const DEFAULT_SYSTEM_USERS = [
  {
    id: 'user_admin',
    name: 'JITTRAKAN K.',
    companyName: 'CACULATION SALARY',
    empCode: '20523',
    department: 'PRODUCTION TECHNOLOGY',
    pin: '20523',
    role: 'admin',
    avatar: '👑',
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
    // กฎเหล็ก: ถ้าผู้ใช้ไม่ใช่ ADMIN (JITTRAKAN K.) จะต้องบังคับดูเฉพาะข้อมูลตัวเองเท่านั้น
    if (user.role === 'admin' && user.id === 'user_admin') {
      if (!localStorage.getItem(STORAGE_KEY_ACTIVE_EDITING_USER)) {
        setActiveEditingUserId('user_admin');
      }
    } else {
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
 * ดึงรายชื่อผู้ใช้ทั้งหมดในระบบ (พร้อมระบบล้างบัญชีตกค้าง/ซ้ำซ้อนอัตโนมัติ)
 * กฎความเป็นส่วนตัว: ถ้าผู้ใช้เป็นพนักงานทั่วไป จะสามารถเห็นได้เฉพาะข้อมูลของตนเองเท่านั้น
 */
function getAllSystemUsers(internalBypass = false) {
  if (!internalBypass && currentUser && (currentUser.role !== 'admin' || currentUser.id !== 'user_admin')) {
    return [ currentUser ];
  }

  try {
    const raw = localStorage.getItem(STORAGE_KEY_USERS_LIST);
    if (raw) {
      let parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        let modified = false;

        // 1. ลบบัญชีตกค้าง/ซ้ำซ้อน และบัญชีในรูปที่ผู้ใช้แจ้งให้ลบ:
        //    - PIN 09718455 ทั้งหมด
        //    - ชื่อที่มีเครื่องหมายคำพูด ' หรือ " (เช่น 'JITTRAKAN K.)
        //    - บัญชีที่ชื่อ JITTRAKAN K. หรือ PIN 20523 หรือ รหัส 20523 แต่ไม่ใช่ user_admin
        const originalCount = parsed.length;
        parsed = parsed.filter(u => {
          if (!u) return false;
          const uPin = String(u.pin || '').trim();
          const uName = String(u.name || '').trim();
          const uCode = String(u.empCode || '').trim();

          // ลบทุกบัญชีที่มี PIN 09718455
          if (uPin === '09718455') return false;

          // ลบทุกบัญชีที่มีเครื่องหมายคำพูดในชื่อ เช่น 'JITTRAKAN K.
          if (uName.startsWith("'") || uName.startsWith('"') || uName.includes("'JITTRAKAN")) return false;

          // ลบทุกบัญชีซ้ำซ้อนของ JITTRAKAN K. หรือ PIN 20523 หรือรหัส 20523 ที่ไม่ใช่ user_admin
          if (u.id !== 'user_admin' && (uName.toUpperCase() === 'JITTRAKAN K.' || uPin === '20523' || uCode === '20523')) {
            return false;
          }

          return true;
        });

        if (parsed.length !== originalCount) {
          modified = true;
        }

        // 2. ตรวจสอบและสร้าง/อัปเดต Master Admin (JITTRAKAN K., PIN 20523) เพียงหนึ่งเดียว
        let adminUser = parsed.find(u => u.id === 'user_admin');
        if (!adminUser) {
          adminUser = { ...DEFAULT_SYSTEM_USERS[0] };
          parsed.unshift(adminUser);
          modified = true;
        }

        // บังคับค่าที่ถูกต้องแน่นอนของ ADMIN หลัก
        adminUser.id = 'user_admin';
        adminUser.name = 'JITTRAKAN K.';
        adminUser.companyName = 'CACULATION SALARY';
        adminUser.empCode = '20523';
        adminUser.department = 'PRODUCTION TECHNOLOGY';
        adminUser.pin = '20523';
        adminUser.role = 'admin';
        adminUser.avatar = '👑';

        // 3. กรองบัญชีที่มี ID หรือ PIN ซ้ำกันออก
        const seenIds = new Set();
        const seenPins = new Set();
        seenPins.add('20523'); // จอง PIN ของ Admin ไว้
        const cleanList = [adminUser];
        seenIds.add('user_admin');

        for (const u of parsed) {
          if (u.id === 'user_admin') continue;
          const pinKey = String(u.pin || '').trim();
          if (!seenIds.has(u.id) && !seenPins.has(pinKey)) {
            seenIds.add(u.id);
            seenPins.add(pinKey);
            cleanList.push(u);
          } else {
            modified = true;
          }
        }
        parsed = cleanList;

        // 4. บัญชีที่เหลือทั้งหมด (ถ้ามี) บังคับเป็น role = 'user'
        parsed.forEach(u => {
          if (u.id !== 'user_admin') {
            if (u.role !== 'user') {
              u.role = 'user';
              u.avatar = '👤';
              modified = true;
            }
          }
          if (!u.companyName) {
            u.companyName = 'CACULATION SALARY';
            modified = true;
          } else {
            u.companyName = u.companyName.toUpperCase();
          }
          if (!u.empCode) {
            u.empCode = (u.id === 'user_admin' ? '20523' : ('EMP-' + (u.pin || '001'))).toUpperCase();
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

        // ตรวจสอบและแก้ไข auth user ใน LocalStorage ถ้าค้างข้อมูลผิดอยู่
        const currentAuthRaw = localStorage.getItem(STORAGE_KEY_AUTH_USER);
        if (currentAuthRaw) {
          try {
            const authObj = JSON.parse(currentAuthRaw);
            if (authObj && (String(authObj.pin).trim() === '09718455' || authObj.name?.includes("'") || authObj.id !== 'user_admin')) {
              localStorage.setItem(STORAGE_KEY_AUTH_USER, JSON.stringify(adminUser));
              currentUser = adminUser;
            }
          } catch(e) {}
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
 * ล้างรายชื่อพนักงานทั้งหมด เหลือไว้เฉพาะ ADMIN หลัก
 */
function clearAllStaffUsers() {
  const adminUser = {
    id: 'user_admin',
    name: 'JITTRAKAN K.',
    companyName: 'CACULATION SALARY',
    empCode: '20523',
    department: 'PRODUCTION TECHNOLOGY',
    pin: '20523',
    role: 'admin',
    avatar: '👑',
    createdAt: '2026-01-01T00:00:00.000Z'
  };
  const list = [adminUser];
  saveSystemUsers(list);
  setActiveEditingUserId('user_admin');
  return { success: true, count: 1, users: list };
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
 * กฎเหล็ก: ถ้าผู้ใช้ไม่ใช่ ADMIN (JITTRAKAN K.) บังคับคืนค่า ID ของตัวเองเท่านั้น ห้ามดูข้อมูลคนอื่น
 */
function getActiveEditingUserId() {
  if (currentUser) {
    if (currentUser.role !== 'admin' || currentUser.id !== 'user_admin') {
      return currentUser.id;
    }
    const saved = localStorage.getItem(STORAGE_KEY_ACTIVE_EDITING_USER);
    return saved || 'user_admin';
  }
  return 'user_admin';
}

/**
 * เปลี่ยน User ID ที่กำลังดู/แก้ไขข้อมูล (เฉพาะ ADMIN เท่านั้นที่สามารถสลับดูคนอื่นได้)
 */
function setActiveEditingUserId(userId) {
  if (currentUser && (currentUser.role !== 'admin' || currentUser.id !== 'user_admin')) {
    activeEditingUserId = currentUser.id;
    localStorage.setItem(STORAGE_KEY_ACTIVE_EDITING_USER, currentUser.id);
    return currentUser.id;
  }
  activeEditingUserId = userId || 'user_admin';
  localStorage.setItem(STORAGE_KEY_ACTIVE_EDITING_USER, activeEditingUserId);
  return activeEditingUserId;
}

/**
 * ดึงข้อมูลผู้ใช้ที่กำลังถูกแก้ไขข้อมูลอยู่
 * กฎความเป็นส่วนตัว: พนักงานทั่วไปจะได้รับเฉพาะข้อมูลของตนเองเสมอ
 */
function getActiveEditingUser() {
  if (currentUser && (currentUser.role !== 'admin' || currentUser.id !== 'user_admin')) {
    return currentUser;
  }
  const targetId = getActiveEditingUserId();
  const users = getAllSystemUsers(true);
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

  // ดึงรายชื่อทั้งหมด (internal bypass) เพื่อตรวจสอบรหัส PIN
  const users = getAllSystemUsers(true);
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
  if (currentUser && currentUser.role !== 'admin') {
    return { success: false, error: 'เฉพาะผู้ดูแลระบบ (ADMIN: JITTRAKAN K.) เท่านั้นที่สามารถเพิ่มผู้ใช้ได้' };
  }
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

  // สิทธิ์ของพนักงานใหม่ต้องเป็น 'user' เสมอ (Admin มีแค่ JITTRAKAN K. เท่านั้น)
  const newUser = {
    id: 'user_' + Date.now().toString(36) + '_' + Math.random().toString(36).substring(2, 6),
    name: name.trim(),
    pin: cleanPin,
    role: 'user',
    avatar: '👤',
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
  // สิทธิ์: ถ้าไม่ใช่ Admin และกำลังจะแก้ไข id คนอื่น ให้ปฏิเสธทันที
  if (currentUser && currentUser.role !== 'admin' && currentUser.id !== userId) {
    return { success: false, error: 'คุณไม่มีสิทธิ์แก้ไขข้อมูลบัญชีของผู้อื่น' };
  }

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

  // บังคับบทบาท: เฉพาะ user_admin เท่านั้นที่เป็น admin ที่เหลือทุกคนเป็น user
  if (users[idx].id === 'user_admin') {
    users[idx].role = 'admin';
    users[idx].avatar = '👑';
  } else {
    users[idx].role = 'user';
    users[idx].avatar = '👤';
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
 * ลบผู้ใช้งาน (ยกเว้น ADMIN JITTRAKAN K.)
 */
function deleteSystemUser(userId) {
  if (currentUser && currentUser.role !== 'admin') {
    return { success: false, error: 'เฉพาะผู้ดูแลระบบ (ADMIN: JITTRAKAN K.) เท่านั้นที่สามารถลบผู้ใช้ได้' };
  }
  if (userId === 'user_admin') {
    return { success: false, error: 'ไม่สามารถลบบัญชี ADMIN หลัก (JITTRAKAN K.) ได้' };
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
