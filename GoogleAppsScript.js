// ซอร์สโค้ดสำหรับนำไปใช้ใน Google Apps Script (Extensions > Apps Script)
// ของ Google Sheets ที่ใช้เก็บข้อมูลการลงทะเบียน

// กำหนดโควต้าจำกัดจำนวนของแต่ละสถานที่จัดงาน
var VENUE_QUOTAS = {
  "สำนักงานสาธารณสุขจังหวัดสตูล และโรงพยาบาลสตูล": 120,
  "โรงพยาบาลละงู": 80,
  "โรงพยาบาลควนกาหลง": 60,
  "โรงพยาบาลควนโดน": 60,
  "โรงพยาบาลมะนัง": 40,
  "โรงพยาบาลทุ่งหว้า": 40,
  "โรงพยาบาลท่าแพ": 50
};

// ฟังก์ชันช่วยดึงและสร้างชีตแรกพร้อมหัวข้อตาราง (กรณีเป็นชีตว่างเปล่า)
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "ประทับเวลา", 
      "ชื่อ", 
      "นามสกุล", 
      "ระดับชั้น", 
      "ประเภทโรงเรียน", 
      "จังหวัด", 
      "อำเภอ", 
      "ตำบล", 
      "ชื่อโรงเรียน", 
      "เบอร์โทรศัพท์", 
      "สถานที่ดูงาน"
    ]);
  }
  return sheet;
}

// ฟังก์ชันช่วยส่งออกข้อมูลในรูปแบบ JSON พร้อมเปิดการเข้าถึงจากโดเมนอื่น (CORS)
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

// จัดการคำขอแบบ GET (ดึงสถิติ และการค้นหาผู้ลงทะเบียน)
function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "getStats") {
    return handleGetStats();
  } else if (action === "search") {
    var phone = e.parameter.phone;
    return handleSearch(phone);
  }
  
  return createJsonResponse({ status: "error", message: "ไม่พบ Action ที่ระบุ" });
}

// จัดการคำขอแบบ POST (ลงทะเบียนใหม่ และแก้ไขข้อมูล)
function doPost(e) {
  var result;
  try {
    // แยกวิเคราะห์ข้อมูล JSON จาก body ที่ส่งมาเป็น text/plain (เลี่ยงปัญหา CORS preflight)
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    
    if (action === "register") {
      result = handleRegister(postData);
    } else if (action === "update") {
      result = handleUpdate(postData);
    } else {
      result = { status: "error", message: "ไม่พบ Action ที่ระบุสำหรับ POST" };
    }
  } catch (err) {
    result = { status: "error", message: "ข้อผิดพลาดระบบ: " + err.toString() };
  }
  
  return createJsonResponse(result);
}

// การดึงสถิติจำนวนคนลงทะเบียนและโควต้าคงเหลือของแต่ละสถานที่
function handleGetStats() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  
  // สร้างตัวนับค่าเริ่มต้นสำหรับทุกสถานที่
  var counts = {};
  for (var key in VENUE_QUOTAS) {
    counts[key] = 0;
  }
  
  // นับจำนวนการลงทะเบียนจริงจากชีต (ข้ามแถวแรกซึ่งเป็น Header)
  for (var i = 1; i < data.length; i++) {
    var venue = data[i][10]; // คอลัมน์ที่ 11: สถานที่ดูงาน
    if (venue && counts.hasOwnProperty(venue)) {
      counts[venue]++;
    }
  }
  
  var venuesList = [];
  for (var key in VENUE_QUOTAS) {
    var quota = VENUE_QUOTAS[key];
    var registered = counts[key];
    var remaining = Math.max(0, quota - registered);
    var status = "ว่าง";
    
    if (remaining === 0) {
      status = "เต็ม";
    } else if (remaining <= 10) {
      status = "ใกล้เต็ม";
    }
    
    venuesList.push({
      name: key,
      quota: quota,
      registered: registered,
      remaining: remaining,
      status: status
    });
  }
  
  return createJsonResponse({
    status: "success",
    totalRegistered: data.length - 1,
    venues: venuesList
  });
}

// ค้นหาข้อมูลผู้ลงทะเบียนด้วยเบอร์โทรศัพท์
function handleSearch(phone) {
  if (!phone) {
    return createJsonResponse({ status: "error", message: "กรุณาระบุเบอร์โทรศัพท์ในการค้นหา" });
  }
  
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var results = [];
  var searchPhone = String(phone).trim();
  
  // ค้นหาแถวที่ตรงกับเบอร์โทรศัพท์ (ข้ามแถว Header)
  for (var i = 1; i < data.length; i++) {
    var rowPhone = String(data[i][9]).trim();
    if (rowPhone === searchPhone) {
      results.push({
        rowNumber: i + 1, // หมายเลขแถวในชีต (1-indexed)
        timestamp: data[i][0],
        firstName: data[i][1],
        lastName: data[i][2],
        grade: data[i][3],
        schoolType: data[i][4],
        province: data[i][5],
        district: data[i][6],
        subDistrict: data[i][7],
        schoolName: data[i][8],
        phone: data[i][9],
        site: data[i][10]
      });
    }
  }
  
  return createJsonResponse({
    status: "success",
    count: results.length,
    data: results
  });
}

// บันทึกการลงทะเบียนใหม่
function handleRegister(data) {
  // เปิดใช้ตัวล็อกสคริปต์เพื่อความถูกต้องเมื่อมีคนส่งข้อมูลมาพร้อมกัน
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // รอคิวก่อนระบบปฏิเสธงาน (สูงสุด 15 วินาที)
  } catch (e) {
    return { status: "error", message: "ระบบมีการใช้งานหนาแน่น กรุณารอสักครู่แล้วลองส่งใหม่อีกครั้ง" };
  }
  
  try {
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    
    var firstName = String(data.firstName || "").trim();
    var lastName = String(data.lastName || "").trim();
    var phone = String(data.phone || "").trim();
    var site = String(data.site || "").trim();
    
    if (!firstName || !lastName || !phone || !site) {
      return { status: "error", message: "กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วน" };
    }
    
    // ตรวจสอบชื่อ - นามสกุลซ้ำในระบบ
    for (var i = 1; i < values.length; i++) {
      var rFirstName = String(values[i][1]).trim();
      var rLastName = String(values[i][2]).trim();
      if (rFirstName.toLowerCase() === firstName.toLowerCase() && rLastName.toLowerCase() === lastName.toLowerCase()) {
        return { status: "error", message: "คุณเคยลงทะเบียนในชื่อและนามสกุลนี้แล้วในระบบ" };
      }
    }
    
    // ตรวจสอบความถูกต้องและโควต้าของสถานที่จัดงาน
    if (!VENUE_QUOTAS.hasOwnProperty(site)) {
      return { status: "error", message: "ไม่มีสถานที่จัดงานที่คุณเลือกในระบบ" };
    }
    
    var quota = VENUE_QUOTAS[site];
    var currentCount = 0;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][10]).trim() === site) {
        currentCount++;
      }
    }
    
    if (currentCount >= quota) {
      return { status: "error", message: "ขออภัย สถานที่นี้ลงทะเบียนเต็มโควต้า (" + quota + " คน) แล้ว" };
    }
    
    // เพิ่มข้อมูลลงในแถวใหม่
    sheet.appendRow([
      new Date(),
      firstName,
      lastName,
      data.grade || "",
      data.schoolType || "",
      data.province || "",
      data.district || "",
      data.subDistrict || "",
      data.schoolName || "",
      phone,
      site
    ]);
    
    return { status: "success", message: "ลงทะเบียนเรียบร้อยแล้ว!" };
  } catch (err) {
    return { status: "error", message: "ไม่สามารถบันทึกข้อมูลได้: " + err.toString() };
  } finally {
    // ปลดล็อกเพื่อให้ผู้ใช้อื่นสามารถเขียนข้อมูลต่อได้
    lock.releaseLock();
  }
}

// อัปเดตข้อมูลผู้ลงทะเบียน (แก้ไขข้อมูล)
function handleUpdate(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { status: "error", message: "ระบบมีการใช้งานหนาแน่น กรุณารอสักครู่แล้วลองส่งใหม่อีกครั้ง" };
  }
  
  try {
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var rowNumber = parseInt(data.rowNumber);
    
    if (!rowNumber || rowNumber <= 1 || rowNumber > values.length) {
      return { status: "error", message: "ไม่พบแถวข้อมูลที่ต้องการแก้ไขในฐานข้อมูล" };
    }
    
    var firstName = String(data.firstName || "").trim();
    var lastName = String(data.lastName || "").trim();
    var phone = String(data.phone || "").trim();
    var site = String(data.site || "").trim();
    
    if (!firstName || !lastName || !phone || !site) {
      return { status: "error", message: "กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วน" };
    }
    
    // ตรวจสอบชื่อ - นามสกุลซ้ำกับคนอื่นๆ ในระบบ (ยกเว้นแถวของตัวเอง)
    for (var i = 1; i < values.length; i++) {
      if (i + 1 === rowNumber) continue;
      
      var rFirstName = String(values[i][1]).trim();
      var rLastName = String(values[i][2]).trim();
      if (rFirstName.toLowerCase() === firstName.toLowerCase() && rLastName.toLowerCase() === lastName.toLowerCase()) {
        return { status: "error", message: "ชื่อและนามสกุลนี้ถูกลงทะเบียนโดยผู้ใช้อื่นแล้วในระบบ" };
      }
    }
    
    // ตรวจสอบโควต้าของสถานที่ กรณีเปลี่ยนสถานที่ดูงาน
    var oldSite = String(values[rowNumber - 1][10]).trim();
    if (oldSite !== site) {
      if (!VENUE_QUOTAS.hasOwnProperty(site)) {
        return { status: "error", message: "ไม่มีสถานที่จัดงานที่คุณเลือกในระบบ" };
      }
      
      var quota = VENUE_QUOTAS[site];
      var currentCount = 0;
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][10]).trim() === site) {
          currentCount++;
        }
      }
      
      if (currentCount >= quota) {
        return { status: "error", message: "ขออภัย สถานที่ดูงานใหม่ที่คุณต้องการเปลี่ยน มีโควต้าเต็มแล้ว" };
      }
    }
    
    // อัปเดตข้อมูลในแถวเดิม (อ้างอิงช่วงเซลล์ดัชนี rowNumber ตั้งแต่คอลัมน์ที่ 2 ถึง 11)
    var range = sheet.getRange(rowNumber, 2, 1, 10);
    range.setValues([[
      firstName,
      lastName,
      data.grade || "",
      data.schoolType || "",
      data.province || "",
      data.district || "",
      data.subDistrict || "",
      data.schoolName || "",
      phone,
      site
    ]]);
    
    return { status: "success", message: "แก้ไขข้อมูลการลงทะเบียนสำเร็จ!" };
  } catch (err) {
    return { status: "error", message: "ไม่สามารถแก้ไขข้อมูลได้: " + err.toString() };
  } finally {
    lock.releaseLock();
  }
}
