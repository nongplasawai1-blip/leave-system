// ลิงก์เว็บแอป Apps Script (ลิงก์ที่ลงท้ายด้วย /exec) วางระหว่างเครื่องหมายคำพูด
window.LEAVE_API_URL = "https://script.google.com/macros/s/AKfycbwZAWWR95UGyTxLXP4THb5b7eeysym3ln6Ic01x9zgHvX2gL0xe406YlBC3By6SecFy/exec"; // อบต.หนองปลาสะวาย

// ---- ไม่ต้องแก้ด้านล่าง: ทำให้หน้าเว็บบน GitHub Pages เรียกคำสั่งใน Apps Script ผ่านลิงก์ /exec ----
(function () {
  if (window.google && window.google.script && window.google.script.run) return; // เปิดผ่าน /exec ใช้ของเดิม
  function call(fn, args) {
    var url = String(window.LEAVE_API_URL || '').trim();
    if (!url) return Promise.reject(new Error('ยังไม่ได้ตั้งค่าลิงก์ระบบ: ใส่ลิงก์ /exec ในไฟล์ docs/config.js'));
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // text/plain ไม่ต้องผ่าน CORS preflight
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow'
    }).then(function (r) {
      if (!r.ok) throw new Error('เชื่อมต่อระบบไม่สำเร็จ (HTTP ' + r.status + ')');
      return r.json();
    }, function () {
      throw new Error('เชื่อมต่อระบบไม่ได้ ตรวจสอบอินเทอร์เน็ต หรือลิงก์ใน docs/config.js');
    }).then(function (j) {
      if (!j || !j.ok) throw new Error((j && j.error) || 'ระบบตอบกลับไม่ถูกต้อง');
      return j.result;
    });
  }
  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        if (typeof name !== 'string') return undefined;
        return function () {
          call(name, Array.prototype.slice.call(arguments)).then(
            function (r) { if (ok) ok(r); },
            function (e) { if (fail) fail(e); });
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = runner(null, null);
})();
