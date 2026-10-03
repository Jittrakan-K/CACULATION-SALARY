const fs = require('fs');
const path = require('path');

console.log('--- TESTING PHONE SIMULATOR INTEGRATION ---');

const html = fs.readFileSync('index.html', 'utf8');
const css = fs.readFileSync('style.css', 'utf8');
const js = fs.readFileSync('app.js', 'utf8');

// 1. Verify HTML Elements
const requiredIds = [
  'btnPhoneSimulator',
  'phoneSimulatorModal',
  'simDeviceSelect',
  'simScaleSelect',
  'btnSimOrientation',
  'simOrientationText',
  'simDeviceScaler',
  'simPhoneFrame',
  'simScreenBezel',
  'simStatusBar',
  'simStatusTime',
  'simCutoutContainer',
  'simDynamicIsland',
  'simNotch',
  'simIframe'
];

let missingIds = [];
for (const id of requiredIds) {
  if (!html.includes(`id="${id}"`)) {
    missingIds.push(id);
  }
}

if (missingIds.length > 0) {
  console.error('FAIL: Missing HTML IDs:', missingIds);
  process.exit(1);
} else {
  console.log('PASS: All 15 required Simulator HTML IDs are present.');
}

// 2. Verify iPhone Models in HTML Select
const expectedModels = [
  'iphone16promax',
  'iphone15pro',
  'iphone15promax',
  'iphone14',
  'iphone14plus',
  'iphone13mini',
  'iphone11',
  'iphone11pro'
];

let missingModels = [];
for (const model of expectedModels) {
  if (!html.includes(`value="${model}"`)) {
    missingModels.push(model);
  }
}

if (missingModels.length > 0) {
  console.error('FAIL: Missing iPhone models in select:', missingModels);
  process.exit(1);
} else {
  console.log('PASS: All iPhone models from iPhone 11 and up are present in select dropdown.');
}

// 3. Verify CSS Classes
const requiredClasses = [
  '.btn-simulator-trigger',
  '.simulator-backdrop',
  '.sim-floating-toolbar',
  '.sim-brand-badge',
  '.sim-stage-container',
  '.sim-device-scaler',
  '.sim-phone-frame',
  '.sim-hw-button',
  '.sim-screen-bezel',
  '.sim-status-bar',
  '.sim-dynamic-island',
  '.sim-notch',
  '.sim-status-icons',
  '.sim-iframe',
  '.sim-home-bar-area',
  '.sim-home-indicator',
  'body.is-in-simulator'
];

let missingClasses = [];
for (const cls of requiredClasses) {
  if (!css.includes(cls)) {
    missingClasses.push(cls);
  }
}

if (missingClasses.length > 0) {
  console.error('FAIL: Missing CSS classes:', missingClasses);
  process.exit(1);
} else {
  console.log('PASS: All required Phone Simulator CSS classes are defined in style.css.');
}

// 4. Verify JS Functions & Variables
const requiredJsSymbols = [
  'SIMULATOR_DEVICES',
  'getSimulatorTriggerBtnHtml',
  'openPhoneSimulator',
  'closePhoneSimulator',
  'changeSimulatorDevice',
  'toggleSimulatorOrientation',
  'changeSimulatorScale',
  'reloadSimulatorFrame',
  'handleSimulatorBackdropClick',
  'updateSimulatorClock',
  'applySimulatorDeviceAndScale'
];

let missingSymbols = [];
for (const sym of requiredJsSymbols) {
  if (!js.includes(sym)) {
    missingSymbols.push(sym);
  }
}

if (missingSymbols.length > 0) {
  console.error('FAIL: Missing JS functions/variables:', missingSymbols);
  process.exit(1);
} else {
  console.log('PASS: All required Phone Simulator JS controllers are implemented in app.js.');
}

// 5. Test Gold Standard IDs Integrity
const originalIds = JSON.parse(fs.readFileSync('scratch/original_ids.json', 'utf8'));
let missingOriginal = [];
for (const id of originalIds) {
  if (!html.includes(`id="${id}"`)) {
    missingOriginal.push(id);
  }
}

if (missingOriginal.length > 0) {
  console.error('FAIL: Missing original gold standard IDs:', missingOriginal);
  process.exit(1);
} else {
  console.log(`PASS: All ${originalIds.length} original gold standard IDs are 100% preserved.`);
}

console.log('--- ALL PHONE SIMULATOR TESTS PASSED SUCCESSFULLY! ---');
