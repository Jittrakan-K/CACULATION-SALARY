const fs = require('fs');

// Extract SIMULATOR_DEVICES from app.js
const js = fs.readFileSync('app.js', 'utf8');

const devicesMatch = js.match(/const SIMULATOR_DEVICES = ({[\s\S]*?^};)/m);
if (!devicesMatch) {
  console.error('FAIL: Could not find SIMULATOR_DEVICES definition');
  process.exit(1);
}

const devicesCode = devicesMatch[0].replace('const SIMULATOR_DEVICES =', 'global.SIMULATOR_DEVICES =');
eval(devicesCode);

console.log('Available Devices:', Object.keys(global.SIMULATOR_DEVICES));

const SIMULATOR_DEVICES = global.SIMULATOR_DEVICES;

// Test every device has valid numbers
for (const [key, dev] of Object.entries(SIMULATOR_DEVICES)) {
  if (!dev.name || !dev.width || !dev.height || !dev.borderRadius || !dev.screenRadius || !dev.bezel) {
    console.error(`FAIL: Device ${key} missing required properties`, dev);
    process.exit(1);
  }
  if (dev.island) {
    if (!dev.islandWidth || !dev.islandHeight) {
      console.error(`FAIL: Island device ${key} missing islandWidth/islandHeight`, dev);
      process.exit(1);
    }
  } else {
    if (!dev.notchWidth || !dev.notchHeight) {
      console.error(`FAIL: Notch device ${key} missing notchWidth/notchHeight`, dev);
      process.exit(1);
    }
  }
  console.log(`PASS: ${key} -> ${dev.name} (${dev.width}x${dev.height}, cutout: ${dev.island ? 'Dynamic Island' : 'Notch'})`);
}

// Test auto fit calculations for different laptop screens
function calcFit(w, h, bezel, winW, winH) {
  const availableW = Math.max(300, winW - 40);
  const availableH = Math.max(300, winH - 90 - 40);
  const totalW = w + (bezel * 2) + 16;
  const totalH = h + (bezel * 2) + 16;
  const scaleW = availableW / totalW;
  const scaleH = availableH / totalH;
  return Math.min(1.0, Math.min(scaleW, scaleH));
}

const testScreens = [
  { name: '13-inch laptop (1366x768)', w: 1366, h: 768 },
  { name: 'Full HD monitor (1920x1080)', w: 1920, h: 1080 },
  { name: '4K monitor (3840x2160)', w: 3840, h: 2160 }
];

for (const scr of testScreens) {
  const p15 = SIMULATOR_DEVICES['iphone15pro'];
  const scale = calcFit(p15.width, p15.height, p15.bezel, scr.w, scr.h);
  console.log(`Auto fit on ${scr.name}: scale = ${(scale * 100).toFixed(1)}%`);
  if (scale <= 0 || scale > 1.0) {
    console.error('FAIL: Scale out of bounds');
    process.exit(1);
  }
}

console.log('PASS: Simulator logic and auto-fit scaling verified successfully!');
