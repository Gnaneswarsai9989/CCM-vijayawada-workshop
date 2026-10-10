const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const port = 9334;
const userDataDir = path.join(__dirname, 'scratch_chrome_user_data');

async function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function run() {
  console.log('Launching headless Chrome on port', port);
  const chromeProc = spawn(chromePath, [
    '--headless=new',
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${userDataDir}`,
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-gpu',
    '--window-size=1280,900'
  ]);

  chromeProc.stderr.on('data', d => {});
  chromeProc.stdout.on('data', d => {});

  let versionData = null;
  for (let i = 0; i < 30; i++) {
    await wait(300);
    try {
      const res = await fetch(`http://127.0.0.1:${port}/json/version`);
      if (res.ok) {
        versionData = await res.json();
        break;
      }
    } catch (e) {}
  }

  if (!versionData) {
    console.error('Failed to connect to Chrome CDP');
    chromeProc.kill();
    return;
  }

  console.log('Chrome CDP ready! Connecting WS...');
  const newTabRes = await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, { method: 'PUT' });
  const tab = await newTabRes.json();
  const ws = new WebSocket(tab.webSocketDebuggerUrl);

  let id = 1;
  const pending = new Map();

  ws.onmessage = (event) => {
    const msg = JSON.parse(event.data);
    if (msg.id && pending.has(msg.id)) {
      const { resolve, reject } = pending.get(msg.id);
      pending.delete(msg.id);
      if (msg.error) reject(msg.error);
      else resolve(msg.result);
    }
  };

  await new Promise(r => ws.onopen = r);

  function send(method, params = {}) {
    const msgId = id++;
    return new Promise((resolve, reject) => {
      pending.set(msgId, { resolve, reject });
      ws.send(JSON.stringify({ id: msgId, method, params }));
    });
  }

  await send('Page.enable');
  await send('Runtime.enable');

  async function testPage(filePath, pageName) {
    console.log(`\n=== Testing ${pageName} ===`);
    const fileUrl = 'file:///' + path.resolve(filePath).replace(/\\/g, '/');
    await send('Page.navigate', { url: fileUrl });
    await wait(1500);

    // If there is a popup, dismiss it immediately for clean testing
    await send('Runtime.evaluate', {
      expression: `(() => {
        const popup = document.getElementById('workshopPopup');
        const overlay = document.getElementById('popupOverlay');
        if (popup) popup.remove();
        if (overlay) overlay.remove();
        // disable smooth scroll temporarily for instant testing
        document.documentElement.style.scrollBehavior = 'auto';
      })()`
    });

    // Check dimensions
    let dimRes = await send('Runtime.evaluate', {
      expression: `({
        scrollHeight: document.documentElement.scrollHeight,
        clientHeight: document.documentElement.clientHeight,
        innerWidth: window.innerWidth,
        innerHeight: window.innerHeight
      })`,
      returnByValue: true
    });
    console.log('Page dimensions:', dimRes.result.value);

    // Initial check
    let evalRes = await send('Runtime.evaluate', {
      expression: `(() => {
        const nav = document.querySelector('.nav');
        if (!nav) return { error: 'No .nav found' };
        const rect = nav.getBoundingClientRect();
        const style = window.getComputedStyle(nav);
        return {
          scrollY: window.scrollY,
          position: style.position,
          top: style.top,
          zIndex: style.zIndex,
          hasScrolledClass: nav.classList.contains('scrolled'),
          rect: { top: rect.top, left: rect.left, width: rect.width, height: rect.height }
        };
      })()`,
      returnByValue: true
    });
    console.log('At scroll top (0):', JSON.stringify(evalRes.result.value, null, 2));

    // Scroll down 800px
    await send('Runtime.evaluate', { expression: 'window.scrollTo(0, 800);' });
    await wait(300);

    evalRes = await send('Runtime.evaluate', {
      expression: `(() => {
        const nav = document.querySelector('.nav');
        if (!nav) return { error: 'No .nav found' };
        const rect = nav.getBoundingClientRect();
        const style = window.getComputedStyle(nav);
        return {
          scrollY: window.scrollY,
          position: style.position,
          top: style.top,
          zIndex: style.zIndex,
          hasScrolledClass: nav.classList.contains('scrolled'),
          rect: { top: rect.top, left: rect.left, width: rect.width, height: rect.height }
        };
      })()`,
      returnByValue: true
    });
    console.log('After scrolling to 600px:', JSON.stringify(evalRes.result.value, null, 2));

    // Capture screenshot
    const shot = await send('Page.captureScreenshot', { format: 'png' });
    const shotPath = path.join(__dirname, `screenshot_${pageName}_scrolled.png`);
    fs.writeFileSync(shotPath, Buffer.from(shot.data, 'base64'));
    console.log(`Screenshot saved to ${shotPath}`);

    // Scroll down 1500px
    await send('Runtime.evaluate', { expression: 'window.scrollTo(0, 1500);' });
    await wait(500);

    evalRes = await send('Runtime.evaluate', {
      expression: `(() => {
        const nav = document.querySelector('.nav');
        if (!nav) return { error: 'No .nav found' };
        const rect = nav.getBoundingClientRect();
        return {
          scrollY: window.scrollY,
          hasScrolledClass: nav.classList.contains('scrolled'),
          rectTopInViewport: rect.top
        };
      })()`,
      returnByValue: true
    });
    console.log('After scrolling to 1500px:', JSON.stringify(evalRes.result.value, null, 2));

    // Test mobile viewport
    console.log(`Testing mobile viewport (375x667) on ${pageName}...`);
    await send('Emulation.setDeviceMetricsOverride', {
      width: 375,
      height: 667,
      deviceScaleFactor: 2,
      mobile: true
    });
    await wait(500);

    // Find overflowing elements
    const overflowReport = await send('Runtime.evaluate', {
      expression: `(() => {
        const docWidth = document.documentElement.offsetWidth;
        const scrollW = document.documentElement.scrollWidth;
        const bodyScrollW = document.body.scrollWidth;
        const wideElements = [];
        document.querySelectorAll('*').forEach(el => {
          const r = el.getBoundingClientRect();
          if (r.right > docWidth + 2) {
            wideElements.push({
              tag: el.tagName,
              id: el.id,
              className: el.className,
              right: Math.round(r.right),
              width: Math.round(r.width)
            });
          }
        });
        return { docWidth, scrollW, bodyScrollW, wideCount: wideElements.length, allElements: wideElements };
      })()`,
      returnByValue: true
    });
    console.log('Mobile overflow report for ' + pageName + ':', JSON.stringify(overflowReport.result.value, null, 2));

    await send('Runtime.evaluate', { expression: 'window.scrollTo(0, 400);' });
    await wait(500);

    evalRes = await send('Runtime.evaluate', {
      expression: `(() => {
        const nav = document.querySelector('.nav');
        const rect = nav ? nav.getBoundingClientRect() : null;
        return {
          scrollY: window.scrollY,
          hasScrolledClass: nav ? nav.classList.contains('scrolled') : false,
          rect: rect ? { top: rect.top, width: rect.width, height: rect.height } : null
        };
      })()`,
      returnByValue: true
    });
    console.log('Mobile scrolled state:', JSON.stringify(evalRes.result.value, null, 2));

    const mobileShot = await send('Page.captureScreenshot', { format: 'png' });
    const mobileShotPath = path.join(__dirname, `screenshot_${pageName}_mobile.png`);
    fs.writeFileSync(mobileShotPath, Buffer.from(mobileShot.data, 'base64'));
    console.log(`Mobile screenshot saved to ${mobileShotPath}`);

    // Reset emulation
    await send('Emulation.clearDeviceMetricsOverride');
  }

  await testPage('index.html', 'index');
  await testPage('register.html', 'register');

  ws.close();
  chromeProc.kill();
  try {
    fs.rmSync(userDataDir, { recursive: true, force: true });
  } catch(e) {}
  console.log('\nAll done!');
}

run().catch(err => {
  console.error('Error running test:', err);
  process.exit(1);
});
