var ss = SpreadsheetApp.getActiveSpreadsheet();
var FOLDER_PDF_ID = "ID Folder PDF"; //ไอดี_โฟลเดอร์_PDF
var FOLDER_IMG_ID = "ID Folder Image";  //ไอดี_โฟลเดอร์_IMAGE
var FOLDER_VDO_ID = "ID Folder VDO";  //ไอดี_โฟลเดอร์_VDO
var CERT_BG_ID = "ID Temp GG Slide";  //ไอดี_เทมเพลตเกียรติบัตร
var SIGN_ID = "ID ลายเซ็นต์";  //ไอดี_ลายเซ็น

function doGet() {  
  recordVisit(); // นับสถิติผู้เข้าชม
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('e-Learning | การสร้างบทเรียนออนไลน์')
      .setFaviconUrl("https://img2.pic.in.th/pic/-3DGlow.png")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getLessons() {
  // ลองดึงจาก Cache ก่อน
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get("all_lessons_data");
    if (cached != null) {
      return JSON.parse(cached);
    }
  } catch (e) {
    // ถ้าอ่าน Cache ไม่ได้ หรือมี Error ให้ข้ามไปอ่านจาก Sheet แทน
  }

  var lessonSheet = ss.getSheetByName('lessons');
  var examSheet = ss.getSheetByName('exams');
  var settingSheet = ss.getSheetByName('settings');
  
  var unitOrder = [];
  if (settingSheet) {
    var settingData = settingSheet.getDataRange().getValues();
    for(var i=1; i<settingData.length; i++){
      if(settingData[i][2]) unitOrder.push(String(settingData[i][2]).trim());
    }
  }

  var lessons = [];
  // ดึงเนื้อหาบทเรียน
  if (lessonSheet.getLastRow() > 1) {
    var lessonData = lessonSheet.getRange(2, 1, lessonSheet.getLastRow() - 1, 11).getValues();
    lessons = lessonData.map(r => ({
      id: r[0], 
      unit: String(r[1]).trim(), 
      topic: r[2], 
      type: r[3], 
      content: r[4], 
      link: r[5],
      mediaType: String(r[10] || '').toLowerCase().trim() 
    }));
  }

  // ดึงหัวข้อสอบ
  if (examSheet.getLastRow() > 1) {
    var examData = examSheet.getRange(2, 1, examSheet.getLastRow() - 1, 3).getValues();
    var uniqueExams = {};
    examData.forEach(r => {
      if(!String(r[2]).includes('Final')) {
         var key = r[1] + '_' + r[2];
         if (!uniqueExams[key]) {
           uniqueExams[key] = {
             id: key, unit: String(r[1]).trim(), topic: r[2], type: 'test', content: '', link: ''
           };
         }
      }
    });
    for (var key in uniqueExams) lessons.push(uniqueExams[key]);
  }
  
  // เพิ่ม Final Test
  var hasFinal = false;
  var finalUnitName = "Final Examination";
  if (examSheet.getLastRow() > 1) {
     var eData = examSheet.getDataRange().getValues();
     for(var i=1; i<eData.length; i++){
        if(String(eData[i][2]).includes('Final')){
           hasFinal = true;
           finalUnitName = eData[i][1]; 
           break;
        }
     }
  }
  
  if(hasFinal){
    lessons.push({
      id: 'FINAL_TEST_ID', unit: 'FINAL_TEST_ZONE', displayUnit: finalUnitName,
      topic: 'Final Test', type: 'test', content: '', link: ''
    });
  }

  var result = { lessons: lessons, unitOrder: unitOrder };
  
  // *** แก้ไขสำคัญ: ใส่ try-catch ป้องกัน Error "Argument too large" ***
  try {
    cache.put("all_lessons_data", JSON.stringify(result), 1200);
  } catch (e) {
    Logger.log("Cache Error (Data too big): " + e.toString());
    // ไม่ต้องทำอะไร ปล่อยให้ระบบทำงานต่อโดยไม่ต้อง Cache
  }
  
  return result;
}

function clearAllCaches() {
  var cache = CacheService.getScriptCache();
  cache.remove("all_lessons_data");
}

// --- Exam System ---
function getExamQuestions(unit, testType) {
  var cacheKey = "exam_" + unit + "_" + testType;
  var cache = CacheService.getScriptCache();
  
  try {
    var cached = cache.get(cacheKey);
    if (cached != null) {
      return JSON.parse(cached);
    }
  } catch (e) {
    // Ignore cache error
  }

  var sheet = ss.getSheetByName('exams');
  var data = sheet.getDataRange().getValues();
  var questions = [];
  
  var typeCheck = String(testType).toLowerCase();
  var isFinalTest = typeCheck.includes('final') || typeCheck.includes('achievement');

  for (var i = 1; i < data.length; i++) {
    var rowUnit = String(data[i][1]).trim();
    var rowType = String(data[i][2]).trim();
    
    var isMatch = false;
    if (rowUnit == unit) {
       if (rowType == testType) isMatch = true;
       else if (isFinalTest && (String(rowType).toLowerCase().includes('final') || String(rowType).toLowerCase().includes('achievement'))) {
         isMatch = true;
       }
       
       if (isMatch) {
         questions.push({
           id: data[i][0],
           qType: data[i][3],
           question: data[i][4],
           choices: data[i][5],
           mediaLink: data[i][7],
           mediaType: String(data[i][8] || '').toLowerCase().trim()
         });
       }
    }
  }

  var examDuration = 0;
  if (isFinalTest) {
      try {
        var setSheet = ss.getSheetByName('settings');
        var val = setSheet.getRange(2, 5).getValue(); 
        examDuration = parseInt(val);
        if (isNaN(examDuration) || examDuration <= 0) examDuration = 30; 
      } catch (e) { examDuration = 30; }
  }

  var result = { questions: questions, duration: examDuration };
  try {
    cache.put(cacheKey, JSON.stringify(result), 900);
  } catch (e) {
    Logger.log("Cache Exam Error: " + e.toString());
  }
  
  return result;
}

function processAndSaveExam(username, examId, userAnswers, unit, testType) {
  var sheet = ss.getSheetByName('exams');
  var data = sheet.getDataRange().getValues();
  
  var questionsMap = {};
  var typeCheck = String(testType).toLowerCase();
  var isFinalTest = typeCheck.includes('final') || typeCheck.includes('achievement');

  for (var i = 1; i < data.length; i++) {
    var rowUnit = String(data[i][1]).trim();
    var rowType = String(data[i][2]).trim();
    var isMatch = false;
    
    if (rowUnit == unit) {
       if (rowType == testType) isMatch = true;
       else if (isFinalTest && (String(rowType).toLowerCase().includes('final') || String(rowType).toLowerCase().includes('achievement'))) {
         isMatch = true;
       }
    }
    
    if(isMatch) {
      questionsMap[data[i][0]] = {
        type: data[i][3],
        answer: data[i][6],
        choices: data[i][5]
      };
    }
  }

  var score = 0;
  var totalQuestions = 0;
  
  for (var qId in questionsMap) {
    totalQuestions++;
    var qData = questionsMap[qId];
    var userAns = userAnswers[qId]; // รับค่ามาเป็น String (อาจมี | คั่น)
    var isCorrect = false;

    if (!userAns) {
    } else if (qData.type === 'matching') {
        // Logic Matching เดิม
        var correctPairs = String(qData.choices || '').split('|');
        var userPairs = String(userAns).split('|'); 
        var allPairsCorrect = true;
        var userMap = {};
        userPairs.forEach(p => { var s=p.split(':'); if(s.length>1) userMap[s[0]] = s[1]; });

        for(var k=0; k<correctPairs.length; k++){
            var pair = correctPairs[k].split(':');
            if(pair.length < 2) continue;
            var key = pair[0];
            var val = pair[1];
            if (String(userMap[key]).trim() !== String(val).trim()) {
                allPairsCorrect = false; 
                break;
            }
        }
        if(correctPairs.length > 0 && allPairsCorrect) isCorrect = true;

    } else {
        // *** LOGIC: รองรับหลายคำตอบ (Multi-part answers) ***
        var correctStr = String(qData.answer).trim();
        var userStr = String(userAns).trim();
        
        // ถ้าเฉลยมีเครื่องหมาย | แสดงว่าเป็นข้อสอบหลายช่องว่าง
        if (correctStr.includes('|')) {
            var correctArr = correctStr.split('|');
            var userArr = userStr.split('|');
            
            // ต้องตอบครบทุกช่อง และถูกทุกช่อง ถึงจะได้คะแนน
            if (correctArr.length === userArr.length) {
                var allPartsCorrect = true;
                for (var p = 0; p < correctArr.length; p++) {
                    if (correctArr[p].trim().toLowerCase() !== userArr[p].trim().toLowerCase()) {
                        allPartsCorrect = false;
                        break;
                    }
                }
                if (allPartsCorrect) isCorrect = true;
            }
        } else {
            // กรณีคำตอบเดียว (เหมือนเดิม)
            if (userStr.toLowerCase() === correctStr.toLowerCase()) {
                isCorrect = true;
            }
        }
    }
    
    if (isCorrect) score++;
  }

  var saveResult = saveExamScore(username, examId, score, totalQuestions);
  
  return {
    status: true,
    score: score,
    total: totalQuestions,
    percent: (totalQuestions > 0) ? Math.round((score / totalQuestions) * 100) : 0,
    certNo: saveResult.certNo
  };
}

// --- Progress & Cert ---
function getStudentProgressData(username) {
  var scoreSheet = ss.getSheetByName('scores');
  var data = scoreSheet.getDataRange().getValues();
  
  var userHistory = data.filter(function(r) { 
    return r[0] == username; 
  }).map(function(r) {
    return {
      id: r[1], // Col B
      score: r[2], // Col C
      status: r[3], // Col D
      percent: r[5], // Col F
      certNo: r[6]   // Col G
    };
  });
  
  return userHistory;
}

// --- Save Functions (Auto MediaType) ---
function saveContent(formData) {
  try {
    const sheet = ss.getSheetByName('lessons');
    const id = new Date().getTime().toString(); 
    
    var content = formData.content_desc || '';
    // Process Images
    content = processBase64Images(content); 

    var link = formData.content_link || '';
    
    if (!link) {
        link = extractLinkFromHtmlOnly(content);
    }
    // แปลงเป็น Direct Link ถ้าจำเป็น
    if (link) link = convertGoogleDriveToDirectUrl(link);

    // Audio Url Fix (Old Logic support)
    var audioUrl = extractAudioUrlFromContent(content);
    if (audioUrl && !link) {
       link = convertGoogleDriveToDirectUrl(audioUrl);
       content = convertAudioLinksInContent(content);
    }

    var quizData = extractQuizDataFromHtml(content);
    var autoMediaType = detectMediaType(content, link);
    
    sheet.appendRow([
      id, formData.content_unit, formData.content_topic, 'content',
      content, link, quizData.qType, quizData.question, quizData.choices, quizData.answer,
      autoMediaType 
    ]);
    
    clearAllCaches(); 
    
    return { status: "ok", message: "บันทึกเนื้อหาเรียบร้อยแล้ว" };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function saveExamQuestion(form) {
  var sheet = ss.getSheetByName('exams');
  var id = new Date().getTime().toString();
  var choices = "";
  var answer = form.exam_answer || "";
  
  if(form.exam_type == 'mcq' || form.exam_type == 'complex') choices = [form.choice_1, form.choice_2, form.choice_3, form.choice_4].join('|');
  else if(form.exam_type == 'tf') choices = "True|False";
  else choices = form.choices || "";
  
  var questionHtml = form.exam_question || "";
  questionHtml = processBase64Images(questionHtml); // Process Images

  // *** NEW LOGIC: ดึง Media Link จากโจทย์โดยอัตโนมัติ ***
  var mediaLink = extractMediaUrlFromHtml(questionHtml); 
  if (!mediaLink) mediaLink = extractLinkFromHtmlOnly(questionHtml);
  if (mediaLink && mediaLink.includes('drive.google.com')) {
    mediaLink = convertGoogleDriveToDirectUrl(mediaLink);
  }

  var autoMediaType = detectMediaType(questionHtml, mediaLink);

  sheet.appendRow([
    id, form.exam_unit, form.exam_cat, form.exam_type,
    questionHtml, choices, answer, mediaLink,
    autoMediaType 
  ]);
  
  // ลบ Cache เฉพาะส่วนที่เกี่ยวข้อง
  try {
     var cache = CacheService.getScriptCache();
     cache.remove("exam_" + form.exam_unit + "_" + form.exam_cat);
  } catch(e){}

  return {status: true, msg: "บันทึกข้อสอบเรียบร้อย"};
}

// --- Migration & Helpers ---
function runAutoMigration() {
  migrateLessonsMediaType();
  migrateExamsMediaType();
  return "อัปเดตข้อมูล MediaType เรียบร้อยแล้ว!";
}

function migrateLessonsMediaType() {
  var sheet = ss.getSheetByName('lessons');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, 6); 
  var data = range.getValues();
  var updates = [];
  for (var i = 0; i < data.length; i++) {
    var type = detectMediaType(data[i][4], data[i][5]);
    updates.push([type]);
  }
  sheet.getRange(2, 11, updates.length, 1).setValues(updates);
}

function migrateExamsMediaType() {
  var sheet = ss.getSheetByName('exams');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, 8);
  var data = range.getValues();
  var updates = [];
  for (var i = 0; i < data.length; i++) {
    var type = detectMediaType(data[i][4], data[i][7]);
    updates.push([type]);
  }
  sheet.getRange(2, 9, updates.length, 1).setValues(updates);
}

function detectMediaType(html, link) {
  var str = (html || "") + (link || "");
  str = str.toLowerCase();
  
  // 1. เช็ค Video ก่อน (Youtube, MP4)
  if (str.includes('youtube') || str.includes('youtu.be') || str.includes('.mp4') || 
     (link && link.includes('drive.google.com') && link.includes('preview'))) {
      return 'video';
  }
  
  // 2. เช็ค Audio (mp3, audio tag)
  // ระวัง: ไฟล์รูปบางทีก็มี export=download ดังนั้นต้องเช็ค audio tag หรือนามสกุล .mp3 เป็นหลัก
  if (str.includes('<audio') || str.includes('.mp3') || str.includes('.wav') || str.includes('.ogg')) {
      return 'audio';
  }
  
  // 3. ถ้าเป็นลิงก์ Google Drive และไม่เข้าข่าย Video/Audio ข้างบน -> ให้ถือว่าเป็น Image
  if (link && link.includes('drive.google.com')) {
      return 'image';
  }

  // 4. เช็ค Image ทั่วไป (นามสกุลไฟล์)
  if (str.includes('<img') || (link && link.match(/\.(jpg|jpeg|png|gif|bmp|webp)$/i))) {
      return 'image';
  }

  return '';
}

// --- Audio/Media Helpers ---
function extractMediaUrlFromHtml(html) {
  if (!html) return '';
  var yt = html.match(/(https?:\/\/(?:www\.)?(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/)[a-zA-Z0-9_-]+)/);
  if (yt) return yt[1];
  var drive = html.match(/(https?:\/\/drive\.google\.com\/file\/d\/[a-zA-Z0-9_-]+)/);
  if (drive) return drive[1];
  return '';
}

function convertGoogleDriveToDirectUrl(url) {
  if (!url) return "";
  
  // ดึง File ID
  var fileId = null;
  var match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) fileId = match1[1];
  else {
      var match2 = url.match(/id=([a-zA-Z0-9_-]+)/);
      if (match2) fileId = match2[1];
  }

  if (fileId) {
    return 'https://drive.google.com/uc?export=download&id=' + fileId;
  }
  
  return url;
}

function extractAudioUrlFromContent(html) {
  if (!html) return '';
  var match = html.match(/<audio[^>]*src=["']([^"']+)["']/i);
  if (match) return match[1];
  var match2 = html.match(/<source[^>]*src=["']([^"']+)["']/i);
  if (match2) return match2[1];
  return '';
}

function convertAudioLinksInContent(html) {
  if (!html) return html;
  return html.replace(/(<audio[^>]*src=["'])([^"']+)(["'][^>]*>)/gi, function(m, p1, url, p2) {
      return p1 + convertGoogleDriveToDirectUrl(url) + p2;
  });
}

// ฟังก์ชันช่วยแกะข้อมูล Quiz จาก HTML (Embedded Quiz)
function extractQuizDataFromHtml(html) {
  if (!html) return { qType: '', question: '', choices: '', answer: '' };
  
  // ค้นหา div ที่มี class embedded-quiz
  // ใช้ Regex แบบกว้างๆ เพื่อกัน Summernote เปลี่ยน format
  var typeMatch = html.match(/data-type\\?=\\?["']([^"']+)["']/);
  var qMatch = html.match(/data-q\\?=\\?["']([^"']+)["']/);
  var choicesMatch = html.match(/data-choices\\?=\\?["']([^"']+)["']/);
  var ansMatch = html.match(/data-ans\\?=\\?["']([^"']+)["']/);
  
  // Fallback: ถ้าหา data-q ไม่เจอ ให้ลองหาจาก Text ใน Tag
  var questionText = qMatch ? qMatch[1] : '';
  if (!questionText) {
     // พยายามหาข้อความหลังจาก icon หรือใน h6
     var textMatch = html.match(/<h6[^>]*>.*?<\/i>(.*?)<\/h6>/);
     if(textMatch) questionText = textMatch[1].replace(/<[^>]+>/g, '').trim();
  }

  return {
    qType: typeMatch ? typeMatch[1] : '',
    question: questionText,
    choices: choicesMatch ? choicesMatch[1] : '',
    answer: ansMatch ? ansMatch[1] : ''
  };
}


// --- ฟังก์ชันดึงภาพพื้นหลังเกียรติบัตรและลายเซ็น) ---
function getCertBackgroundData() {
  try {
    // 1. ดึงภาพพื้นหลัง
    var bgFile = DriveApp.getFileById(CERT_BG_ID); 
    var bgBlob = bgFile.getBlob();
    var bgBase64 = Utilities.base64Encode(bgBlob.getBytes());
    var bgMime = bgBlob.getContentType();
    
    // 2. ดึงภาพลายเซ็น (เพิ่มใหม่)
    var signBase64 = "";
    var signMime = "image/png"; // ค่า Default
    
    // ตรวจสอบว่ามีการใส่ SIGN_ID ไว้หรือไม่
    if (typeof SIGN_ID !== 'undefined' && SIGN_ID) {
       var signFile = DriveApp.getFileById(SIGN_ID);
       var signBlob = signFile.getBlob();
       signBase64 = Utilities.base64Encode(signBlob.getBytes());
       signMime = signBlob.getContentType();
    }
    
    return { 
      status: 'ok', 
      base64: bgBase64, 
      mime: bgMime,
      signBase64: signBase64, // ส่งข้อมูลลายเซ็นไปด้วย
      signMime: signMime
    };

  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// --- User System ---
function registerUser(form) {
  var sheet = ss.getSheetByName('user');
  var lastRow = sheet.getLastRow();
  var users = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  
  if(users.includes(form.reg_user)) {
    return {status: false, msg: "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว"};
  }
  
  var newId = lastRow; // Simple running ID
  var fullName = form.reg_prefix + form.reg_fname + " " + form.reg_lname;
  
  sheet.appendRow([newId, fullName, form.reg_user, form.reg_pass, 'student', form.reg_status]);
  return {status: true, msg: "สมัครสมาชิกสำเร็จ"};
}

function loginUser(user, pass) {
  var sheet = ss.getSheetByName('user');
  var data = sheet.getDataRange().getValues();
  
  for(var i = 1; i < data.length; i++) {
    if(data[i][2] == user && data[i][3] == pass) {
      return {
        status: true, 
        name: data[i][1], 
        role: data[i][4],
        username: data[i][2]
      };
    }
  }
  return {status: false, msg: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง"};
}

// --- ส่วนที่เพิ่มใหม่สำหรับ Admin Dashboard ---

// 1. ดึงข้อมูลทั้งหมดสำหรับ Admin (User, Lessons, Exams Grouped)
function getAdminAllData() {
  var userSheet = ss.getSheetByName('user');
  var lessonSheet = ss.getSheetByName('lessons');
  var examSheet = ss.getSheetByName('exams');

  // A. ดึงข้อมูล Users
  var users = [];
  if (userSheet.getLastRow() > 1) {
    var uData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 5).getValues();
    // No. (ใช้ index), Name, Username, Password, Role, ID(col 1)
    users = uData.map((r, i) => ({
      no: i + 1,
      id: r[0],
      name: r[1],
      user: r[2],
      pass: r[3],
      role: r[4]
    }));
  }

  // B. ดึงข้อมูล Lessons
  var lessons = [];
  if (lessonSheet.getLastRow() > 1) {
    var lData = lessonSheet.getRange(2, 1, lessonSheet.getLastRow() - 1, 3).getValues();
    // ID, Unit, Topic
    lessons = lData.map((r, i) => ({
      no: i + 1,
      id: r[0],
      unit: r[1],
      topic: r[2]
    }));
  }

  // C. ดึงและจัดกลุ่ม Exams
  var exams = [];
  if (examSheet.getLastRow() > 1) {
    var eData = examSheet.getRange(2, 1, examSheet.getLastRow() - 1, 4).getValues(); 
    // ID, Unit, TestType, QType
    
    // ใช้ Object เพื่อจัดกลุ่ม (Group By Unit + TestType + QType)
    var groups = {};
    
    eData.forEach(r => {
      var key = r[1] + '|' + r[2] + '|' + r[3]; // Key สำหรับ Group
      if (!groups[key]) {
        groups[key] = {
          unit: r[1],
          testType: r[2],
          qType: r[3],
          count: 0,
          ids: [] // เก็บ ID ของข้อสอบในกลุ่มนี้ไว้เผื่อลบ
        };
      }
      groups[key].count++;
      groups[key].ids.push(r[0]);
    });

    // แปลง Object กลับเป็น Array
    var index = 1;
    for (var k in groups) {
      exams.push({
        no: index++,
        unit: groups[k].unit,
        testType: groups[k].testType,
        qType: groups[k].qType,
        count: groups[k].count,
        ids: groups[k].ids.join(',') // ส่ง ID ทั้งหมดไปเป็น string ขั้นด้วย comma
      });
    }
  }

  return { users: users, lessons: lessons, exams: exams };
}

// 2. ฟังก์ชันลบข้อมูล (Delete)
function deleteAdminItem(type, id) {
  var sheetName = '';
  if (type === 'user') sheetName = 'user';
  else if (type === 'lesson') sheetName = 'lessons';
  else if (type === 'exam') sheetName = 'exams';
  
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  // กรณี Exam ลบเป็นกลุ่ม (id ที่ส่งมาคือ list ของ id เช่น "123,124,125")
  if (type === 'exam') {
    var idsToDelete = id.split(',');
    // ต้องลบจากล่างขึ้นบนเพื่อไม่ให้ index เพี้ยน
    for (var i = data.length - 1; i >= 1; i--) {
      // เช็คว่า ID ของแถวนี้ อยู่ในรายการที่จะลบไหม
      if (idsToDelete.includes(String(data[i][0]))) {
        sheet.deleteRow(i + 1);
      }
    }
  } else {
    // กรณี User และ Lesson ลบแถวเดียว
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) == String(id)) { // เทียบ ID (Column A)
        sheet.deleteRow(i + 1);
        break; // เจอแล้วลบเลย
      }
    }
  }
  
  return { status: true };
}


// 5. บันทึกเนื้อหาพร้อมไฟล์ (ดึงแบบฝึกหัดไปลง Col G-J)
function saveContentWithFile(formData, base64File, fileName, mimeType) {
  try {
    const sheet = ss.getSheetByName('lessons');
    const id = new Date().getTime().toString(); 
    let mediaLink = formData.content_link || "";
    
    if (base64File) {
        const decodedBlob = Utilities.newBlob(Utilities.base64Decode(base64File), mimeType, fileName);
        const folderId = FOLDER_IMG_ID;
        if (folderId) {
            const folder = DriveApp.getFolderById(folderId);
            const uploadedFile = folder.createFile(decodedBlob);
            uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            mediaLink = uploadedFile.getUrl();
        } else {
            return { status: "error", message: "ไม่พบ Folder ID" };
        }
    }
    
    // Process Base64 Images ในเนื้อหาด้วย
    var content = formData.content_desc || '';
    content = processBase64Images(content);

    var quizData = extractQuizDataFromHtml(content);
    var autoMediaType = detectMediaType(content, link || mediaLink);

    sheet.appendRow([
      id, formData.content_unit, formData.content_topic, 'content', 
      content, mediaLink, quizData.qType, quizData.question, quizData.choices, quizData.answer,
      autoMediaType
    ]);
    
    clearAllCaches(); // ล้าง Cache
    
    return { status: "ok", message: "บันทึกเนื้อหาพร้อมไฟล์เรียบร้อยแล้ว" };
  } catch (e) {
    Logger.log("Error in saveContentWithFile: " + e.toString());
    return { status: "error", message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function processBase64Images(htmlContent) {
  if (!htmlContent) return "";
  
  // Regex หา src="data:image/..."
  var regex = /<img[^>]+src="data:image\/([a-zA-Z]*);base64,([^"]*)"[^>]*>/g;
  
  return htmlContent.replace(regex, function(match, imageType, base64Data) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/" + imageType, "image_" + new Date().getTime() + "." + imageType);
      
      // บันทึกลง Folder รูปภาพ (ใช้ FOLDER_IMG_ID ที่ประกาศไว้ต้นไฟล์)
      var folder = DriveApp.getFolderById(FOLDER_IMG_ID);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // คืนค่าเป็น URL ของรูปภาพใน Drive (Direct Link)
      // ใช้ thumbnail link หรือ download link แทน lh3/lh5 เพื่อความง่ายใน GAS
      return match.replace(/src="[^"]*"/, 'src="https://drive.google.com/uc?export=view&id=' + file.getId() + '"');
      
    } catch(e) {
      return match; // ถ้า Error ให้ใช้ Base64 เดิม
    }
  });
}

// --- Admin: ตั้งค่าเกียรติบัตร ---
function setCertStartNumber(num) {
  var sheet = ss.getSheetByName('settings');
  sheet.getRange("B2").setValue(num);
  return {status: true, msg: "บันทึกเลขที่เริ่มต้นเรียบร้อย"};
}


// ฟังก์ชันดึง/ออกเลขที่เกียรติบัตร (ตรวจสอบว่ามีแล้วหรือยัง ถ้าไม่มีก็รันใหม่)
function getMyCertNumber(username) {
  var scoreSheet = ss.getSheetByName('scores');
  var settingSheet = ss.getSheetByName('settings');
  var userSheet = ss.getSheetByName('user'); // เอาไว้ดึงชื่อจริง
  
  // 1. หาชื่อจริงของ User
  var userData = userSheet.getDataRange().getValues();
  var fullName = username; // ค่าเริ่มต้น
  for(var j=1; j<userData.length; j++){
    if(userData[j][2] == username){
       fullName = userData[j][1]; // ดึงชื่อจริง (Name)
       break;
    }
  }

  var data = scoreSheet.getDataRange().getValues();
  
  // 2. เช็คว่า User นี้เคยได้เลขเกียรติบัตรหรือยัง (ดูคอลัมน์ G / Index 6)
  for(var i=1; i<data.length; i++){
    if(data[i][0] == username && String(data[i][1]).includes('Final') && data[i][6]){
       return { status: 'ok', certNo: data[i][6], fullName: fullName };
    }
  }
  
  // 3. ถ้ายังไม่มี ให้รันเลขใหม่จาก Settings
  var currentRun = settingSheet.getRange("B2").getValue(); 
  var nextRun = parseInt(currentRun) + 1;
  
  // อัปเดตเลขใหม่ลง Settings
  settingSheet.getRange("B2").setValue(nextRun);
  
  // Format เลข เช่น 00009/2568
  var year = new Date().getFullYear() + 543;
  var certNo = String(nextRun).padStart(5, '0') + "/" + year;
  
  // 4. บันทึกเลขลงใน Score Sheet (บันทึกย้อนหลังใส่แถว Final ล่าสุดของคนนั้น)
  for(var i=data.length-1; i>=1; i--){ 
     if(data[i][0] == username && String(data[i][1]).includes('Final')){
        scoreSheet.getRange(i+1, 7).setValue(certNo); // Col G
        break;
     }
  }
  
  return { status: 'ok', certNo: certNo, fullName: fullName };
}

// --- 3. ดึงข้อมูลประวัติผู้ได้รับเกียรติบัตร (Admin Datatable) ---
function getCertHistoryList() {
  try {
    var sheet = ss.getSheetByName('scores');
    var userSheet = ss.getSheetByName('user');
    
    if (!sheet || !userSheet) return [];

    var data = sheet.getDataRange().getValues();
    var users = userSheet.getDataRange().getValues();
    
    // สร้าง Map ชื่อจริง
    var userMap = {};
    users.forEach(function(r) {
       if(r.length > 2) userMap[r[2]] = r[1]; 
    });
    
    var certs = [];
    
    // วนลูปข้อมูล (เริ่มจากล่าสุดล่างขึ้นบน จะได้เห็นคนล่าสุดก่อน)
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      
      // ป้องกันข้อมูลไม่ครบ
      if(row.length < 7) continue;

      var lessonId = String(row[1]).toUpperCase(); // *** แปลงเป็นตัวใหญ่หมดก่อนเช็ค ***
      var certNo = String(row[6]);
      
      // เช็คว่า ID มีคำว่า FINAL และมีเลขเกียรติบัตร
      if(lessonId.includes('FINAL') && certNo && certNo.trim() !== '' && certNo !== 'undefined'){ 
        
        // จัดการวันที่ (ใช้ค่าดิบถ้าแปลงไม่ได้)
        var dateStr = String(row[4]);
        try {
           if (row[4] instanceof Date) {
              dateStr = Utilities.formatDate(row[4], "GMT+7", "dd/MM/yyyy HH:mm");
           }
        } catch(e) {}

        certs.push({
          no: certs.length + 1,
          name: userMap[row[0]] || row[0], // ถ้าไม่มีชื่อจริงให้ใช้ User แทน (รองรับ a1)
          certNo: certNo,
          date: dateStr,
          score: row[5]
        });
      }
    }
    
    return certs;
    
  } catch (e) {
    Logger.log("Error: " + e.toString());
    return []; 
  }
}

// --- 4. แก้ไขปัญหารูป/เสียง (ใช้ Direct Link แทน Base64) ---
function getDirectUrl(fileId) {
   return "https://drive.google.com/uc?export=download&id=" + fileId;
}

function getUserProgress(username) {
  var sheet = ss.getSheetByName('scores');
  var data = sheet.getDataRange().getValues();
  var progress = {};
  // กรองเอาเฉพาะของ user นี้
  data.forEach(r => {
    if(r[0] == username) {
      progress[r[1]] = {score: r[2], status: r[3]};
    }
  });
  return progress;
}

function saveProgress(username, lessonId, score, status) {
  var sheet = ss.getSheetByName('scores');
  var time = new Date();
  sheet.appendRow([username, lessonId, score, status, time]);
  return {status: true};
}

// --- Stats & Cert ---
function recordVisit() {
  var sheet = ss.getSheetByName('visitor_logs');
  // ถ้ายังไม่มีชีท ให้สร้างใหม่ (กันพลาด)
  if (!sheet) {
    sheet = ss.insertSheet('visitor_logs');
  }
  // บันทึกวันที่และเวลาปัจจุบันลงไป
  sheet.appendRow([new Date()]);
}

function getCertNumber() {
  var sheet = ss.getSheetByName('settings');
  var current = sheet.getRange("B2").getValue();
  var next = parseInt(current) + 1;
  // Format เช่น 00001/2568
  var year = new Date().getFullYear() + 543;
  var numStr = String(current).padStart(5, '0');
 
  return numStr + "/" + year;
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// --- ฟังก์ชันคำนวณความก้าวหน้า ---

function markLessonComplete(username, lessonId) {
  var scoreSheet = ss.getSheetByName('scores');
  var lessonSheet = ss.getSheetByName('lessons');
  
  // 1. ตรวจสอบว่าเคยบันทึกหรือยัง
  var existing = false;
  var data = scoreSheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++){
    if(data[i][0] == username && data[i][1] == lessonId){
       existing = true;
       break;
    }
  }
  
  // 2. ถ้ายังไม่เคย ให้บันทึก
  if(!existing){
    scoreSheet.appendRow([username, lessonId, 100, 'completed', new Date()]);
  }
  
  // 3. คำนวณเปอร์เซ็นต์ใหม่ทันที
  return calculatePercent(username);
}

function calculatePercent(username){
  var lessonSheet = ss.getSheetByName('lessons');
  var scoreSheet = ss.getSheetByName('scores');
  
  // นับจำนวนบทเรียนทั้งหมด (เฉพาะที่เป็น Content ไม่รวมข้อสอบ)
  var totalLessons = lessonSheet.getLastRow() - 1; 
  if(totalLessons < 1) totalLessons = 1;

  // นับจำนวนที่เรียนจบแล้วของผู้ใช้นี้
  var completedCount = 0;
  var scoreData = scoreSheet.getDataRange().getValues();
  
  // กรองดูว่า user นี้เรียน lesson id ไหนไปแล้วบ้าง (นับเฉพาะที่เป็น content)
  var completedLessons = [];
  scoreData.forEach(r => {
      if(r[0] == username && r[3] == 'completed'){
         if(completedLessons.indexOf(r[1]) === -1) {
             completedLessons.push(r[1]);
         }
      }
  });
  
  completedCount = completedLessons.length;
  
  var percent = (completedCount / totalLessons) * 100;
  if(percent > 100) percent = 100;
  
  return Math.round(percent);
}

// --- Dashboard Logic ---

function getUserDashboardStats(username) {
    if (!username) return {}; // ป้องกันการเข้าถึงหากไม่มี username

    var lessonSheet = ss.getSheetByName('lessons');
    var scoreSheet = ss.getSheetByName('scores');
    var examSheet = ss.getSheetByName('exams'); // ดึง sheet exams มาเพื่อนับคะแนนเต็ม

    // 1. สร้าง Map คะแนนเต็ม (Full Score Map) จาก Sheet Exams
    var examFullScoreMap = {};
    if (examSheet && examSheet.getLastRow() > 1) {
        // ดึงคอลัมน์ Unit(B) และ Cat(C)
        var examData = examSheet.getRange(2, 2, examSheet.getLastRow() - 1, 2).getValues(); 
        examData.forEach(r => {
            var unit = r[0] ? String(r[0]).trim() : ''; // Unit Name
            var cat = r[1] ? String(r[1]).trim() : '';   // Test Type (Pre-test, Post-test, Final-test)
            if (unit && cat) {
                var examId = unit + '_' + cat; // เช่น "Unit 1 : คำศัพท์_Pre-test"
                examFullScoreMap[examId] = (examFullScoreMap[examId] || 0) + 1;
            }
        });
    }

    // 2. ดึงโครงสร้างบทเรียนและกำหนด Unit Structure
    // ดึงคอลัมน์ ID(A), Unit(B), Topic(C), Type(D)
    var units = {};
    var lessonData = [];
    if (lessonSheet && lessonSheet.getLastRow() > 1) {
        lessonData = lessonSheet.getRange(2, 1, lessonSheet.getLastRow() - 1, 4).getValues();
    }
    
    lessonData.forEach(r => {
        var unitName = String(r[1]).trim(); // Column B: Unit Name
        if (!units[unitName]) {
            // กำหนด Full Score จาก Map ที่สร้างไว้
            var preExamId = unitName + '_Pre-test';
            var postExamId = unitName + '_Post-test';
            var finalExamId = unitName + '_Final-test'; 
            
            units[unitName] = {
                name: unitName,
                totalContent: 0,
                completedContent: 0,
                preScore: '-',
                preFullScore: examFullScoreMap[preExamId] || 0, // 👈 คะแนนเต็ม Pre-test
                postScore: '-',
                postFullScore: examFullScoreMap[postExamId] || 0, // 👈 คะแนนเต็ม Post-test
                finalScore: '-',
                finalFullScore: examFullScoreMap[finalExamId] || 0, // คะแนนเต็ม Final-test
                percent: 0 // Progress percent
            };
        }
        // นับจำนวน Content 
        if (r[3] == 'content') {
            units[unitName].totalContent++;
        }
    });

    // 3. ดึงคะแนนของผู้ใช้และอัปเดต Scores & Progress (โค้ดเดิมที่ปรับปรุงเล็กน้อย)
    var userScores = [];
    if (scoreSheet && scoreSheet.getLastRow() > 1) {
        var scoreData = scoreSheet.getDataRange().getValues();
        userScores = scoreData.filter(r => r[0] == username);
    }
    
    var completedContentIds = {}; // ใช้สำหรับนับ Content ไม่ให้ซ้ำ

    userScores.forEach(r => {
        var examId = String(r[1]).trim(); 
        var score = r[2];
        var status = r[3];
        var unitName;

        // A. กรณีเป็นข้อสอบ
        if (examId.includes('_Pre-test')) {
            unitName = examId.replace('_Pre-test', '');
            if (units[unitName]) units[unitName].preScore = score;
        } else if (examId.includes('_Post-test')) {
            unitName = examId.replace('_Post-test', '');
            if (units[unitName]) units[unitName].postScore = score;
        } else if (examId.includes('Final-test') || examId.includes('Achievement')) {
            // กรณี Final Test
             for (var uKey in units) {
                if (examId.includes(uKey)) {
                    units[uKey].finalScore = score; 
                    break;
                }
            }
        }
        
        // B. กรณีเป็น Content
        if (status == 'completed') {
             var lessonMatch = lessonData.find(L => String(L[0]).trim() == examId); 
             if (lessonMatch) {
                unitName = String(lessonMatch[1]).trim(); 
                if (units[unitName] && lessonMatch[3] == 'content' && !completedContentIds[examId]) {
                    units[unitName].completedContent++;
                    completedContentIds[examId] = true;
                }
             }
        }
    });

    // 4. คำนวณเปอร์เซ็นต์ความก้าวหน้าเนื้อหา
    for (var u in units) {
        var obj = units[u];
        if (obj.totalContent > 0) {
            obj.percent = Math.round((obj.completedContent / obj.totalContent) * 100);
        } else if(obj.postScore !== '-' || obj.preScore !== '-'){
             obj.percent = 100;
        } else {
             obj.percent = 0;
        }
    }

    return units;
}

function saveExamScore(username, examId, score, fullScore) {
  var sheet = ss.getSheetByName('scores');
  var settingSheet = ss.getSheetByName('settings');
  
  // 1. คำนวณเปอร์เซ็นต์
  var percent = 0;
  if (fullScore > 0) {
    percent = Math.round((score / fullScore) * 100);
  }

  // 2. ตรวจสอบเงื่อนไขเกียรติบัตร (เป็น Final และ ผ่าน 70%)
  var certNo = ""; // ค่าเริ่มต้นว่าง
  var isFinal = String(examId).includes('Final') || String(examId).includes('FINAL_TEST');
  
  if (isFinal && percent >= 70) {
      // --- เริ่มกระบวนการออกเลขที่เกียรติบัตรอัตโนมัติ ---
      try {
        var currentRun = settingSheet.getRange("B2").getValue();
        var nextRun = parseInt(currentRun) + 1;
        
        // อัปเดตเลขรันใหม่
        settingSheet.getRange("B2").setValue(nextRun);
        
        // สร้างเลขที่เกียรติบัตร เช่น 00009/2568
        var year = new Date().getFullYear() + 543;
        certNo = String(nextRun).padStart(5, '0') + "/" + year;
      } catch (e) {
        // กรณีฉุกเฉิน (เช่น อ่าน setting ไม่ได้)
        certNo = "Error-" + new Date().getTime(); 
      }
  }

  // 3. บันทึกข้อมูลลง Sheet (ให้ครบทั้ง 7 คอลัมน์)
  // Username, LessonID, Score, Status, Timestamp, Percentage, CertificateNumber
  sheet.appendRow([
    username,  
    examId,  
    score, 
    'tested', 
    new Date(),
    percent,   // บันทึก %
    certNo     // บันทึกเลขที่เกียรติบัตร (ถ้าไม่ผ่านจะเป็นค่าว่าง)
  ]);
  
  return {status: true, certNo: certNo};
}

// --- ฟังก์ชันดึงสถิติ ---
function getVisitorStats() {
  var sheet = ss.getSheetByName('visitor_logs');
  if (!sheet) return { daily: 0, monthly: 0, total: 0 };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { daily: 0, monthly: 0, total: 0 };

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var now = new Date();
  var currentMonth = now.getMonth();
  var currentYear = now.getFullYear();
  var currentDateStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");

  var total = 0;
  var daily = 0;
  var monthly = 0;

  for (var i = 0; i < data.length; i++) {
    var rowDate = new Date(data[i][0]);
    var count = parseInt(data[i][1]) || 0; // อ่านค่าจาก Col B (Total)

    // ยอดรวมทั้งหมด
    total += count;

    // ตรวจสอบเดือนและปีเดียวกัน
    if (rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
      monthly += count;
      
      // ตรวจสอบว่าเป็น "วันนี้" หรือไม่
      var rowDateStr = Utilities.formatDate(rowDate, "GMT+7", "dd/MM/yyyy");
      if (rowDateStr === currentDateStr) {
        daily += count;
      }
    }
  }

  return {
    daily: daily,
    monthly: monthly,
    total: total
  };
}

// 1. ดึงข้อมูลรายรายการเพื่อนำไปแก้ไข
function getDataForEdit(type, id) {
  var sheetName = '';
  if (type === 'user') sheetName = 'user';
  else if (type === 'lesson') sheetName = 'lessons';
  else if (type === 'exam') sheetName = 'exams';
  
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  // ค้นหาแถวที่มี ID ตรงกัน
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(id)) {
      var row = data[i];
      
      // Return ข้อมูลตามประเภท
      if (type === 'user') {
        // ID, Name, User, Pass, Role
        // ต้องแกะ Name ออกเป็น Prefix, Fname, Lname (แบบคร่าวๆ)
        var nameParts = row[1].split(' ');
        var prefix = "นาย"; // default
        var fname = row[1];
        var lname = "";
        
        // ลองเดาคำนำหน้า (Logic ง่ายๆ)
        var prefixes = ["ด.ช.", "ด.ญ.", "นาย", "นาง", "น.ส."];
        for(var p of prefixes){
            if(row[1].startsWith(p)){
                prefix = p;
                var rest = row[1].substring(p.length);
                var names = rest.trim().split(' ');
                fname = names[0];
                lname = names.slice(1).join(' ');
                break;
            }
        }

        return {
          id: row[0],
          prefix: prefix,
          fname: fname,
          lname: lname,
          user: row[2],
          pass: row[3]
        };
      } 
      else if (type === 'lesson') {
        // ID, Unit, Topic, Type, Content, Link
        return {
          id: row[0],
          unit: row[1],
          topic: row[2],
          content: row[4],
          link: row[5]
        };
      } 
      else if (type === 'exam') {
        // ID, Unit, Cat, Type, Question, Choices, Answer
        // ส่ง choices ดิบๆ ไปเลย เดี๋ยว JS ไป parse เองตาม Type
        return {
          id: row[0],
          unit: row[1],
          cat: row[2],
          type: row[3],
          question: row[4],
          choices: row[5], 
          answer: row[6]
        };
      }
    }
  }
  return null;
}

// 2. ฟังก์ชันแก้ไขข้อมูล (แก้ไขให้รองรับทุกประเภท)
function updateData(type, form) {
  var sheetName = '';
  if (type === 'user') sheetName = 'user';
  else if (type === 'lesson') sheetName = 'lessons';
  else if (type === 'exam') sheetName = 'exams';
  
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var editId = form.edit_id;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(editId)) {
      var rowIdx = i + 1;
      
      if (type === 'user') {
         var fullName = form.reg_prefix + form.reg_fname + " " + form.reg_lname;
         sheet.getRange(rowIdx, 2, 1, 3).setValues([[fullName, form.reg_user, form.reg_pass]]); 
         if(form.reg_status) sheet.getRange(rowIdx, 6).setValue(form.reg_status);
      } 
      else if (type === 'lesson') {
        var content = form.content_desc || '';
        content = processBase64Images(content);
        
        var link = form.content_link || data[i][5];
        
        // *** NEW LOGIC: ถ้าลิงก์ว่าง ให้ดึงจาก HTML ***
        if (!link || link === "") {
            link = extractLinkFromHtmlOnly(content);
        }
        if (link) link = convertGoogleDriveToDirectUrl(link);
        
        var quizData = extractQuizDataFromHtml(content);
        
        // อัปเดต Col B-J (2-10) และ K (11-MediaType) 
        // ต้องระวังเรื่อง MediaType ให้ถูกต้อง
        var autoMediaType = detectMediaType(content, link);

        sheet.getRange(rowIdx, 2, 1, 10).setValues([[
            form.content_unit, form.content_topic, 'content',
            content, link,
            quizData.qType, quizData.question, quizData.choices, quizData.answer,
            autoMediaType
        ]]);
        clearAllCaches(); 
      }
      else if (type === 'exam') {
         var choices = "";
         if(form.exam_type == 'mcq' || form.exam_type == 'complex') {
            choices = [form.choice_1, form.choice_2, form.choice_3, form.choice_4].join('|');
         } else if(form.exam_type == 'tf') {
            choices = "True|False";
         } else {
            choices = form.choices || "";
         }
         
         var questionHtml = form.exam_question || "";
         questionHtml = processBase64Images(questionHtml);
         
         var mediaLink = extractMediaUrlFromHtml(questionHtml);
         if (!mediaLink) mediaLink = extractLinkFromHtmlOnly(questionHtml);
         if (mediaLink) mediaLink = convertGoogleDriveToDirectUrl(mediaLink);
         
         var autoMediaType = detectMediaType(questionHtml, mediaLink);
         
         // อัปเดต Col B-I
         sheet.getRange(rowIdx, 2, 1, 8).setValues([[
            form.exam_unit, form.exam_cat, form.exam_type,
            questionHtml, choices, form.exam_answer, mediaLink,
            autoMediaType
         ]]);
         
         try {
             var cache = CacheService.getScriptCache();
             cache.remove("exam_" + form.exam_unit + "_" + form.exam_cat);
         } catch(e){}
      }
      
      return { status: true, msg: "แก้ไขข้อมูลเรียบร้อยแล้ว" };
    }
  }
  return { status: false, msg: "ไม่พบข้อมูล ID นี้" };
}

// 3. ฟังก์ชันพิเศษสำหรับดึงรายชื่อข้อสอบในกลุ่ม (สำหรับ Dropdown ตอนกดแก้ไขข้อสอบ)
function getExamQuestionsInGroup(idListString) {
  var sheet = ss.getSheetByName('exams');
  var data = sheet.getDataRange().getValues();
  var ids = idListString.split(',');
  var questions = [];
  
  data.forEach(r => {
    if (ids.includes(String(r[0]))) {
       questions.push({id: r[0], question: r[4]});
    }
  });
  return questions;
}


// ========== ฟังก์ชัน Proxy สำหรับเล่นไฟล์เสียงจาก Google Drive ==========
function getAudioBase64(fileIdOrUrl) {
  try {
    var fileId = fileIdOrUrl;
    
    // ถ้าเป็น URL ให้ดึง ID ออกมา
    if (fileIdOrUrl.includes('drive.google.com')) {
      var match = fileIdOrUrl.match(/[-\w]{25,}/);
      if (match) {
        fileId = match[0];
      }
    }
    
    // ดึงไฟล์จาก Drive
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var mimeType = blob.getContentType();
    var base64Data = Utilities.base64Encode(blob.getBytes());
    
    return {
      status: 'ok',
      mimeType: mimeType,
      base64: base64Data,
      fileName: file.getName()
    };
    
  } catch (e) {
    Logger.log('Error in getAudioBase64: ' + e.toString());
    return {
      status: 'error',
      message: e.toString()
    };
  }
}

// ฟังก์ชันดึง File ID จาก URL
function extractFileIdFromUrl(url) {
  if (!url) return null;
  
  // Pattern: /file/d/FILE_ID/
  var match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) return match1[1];
  
  // Pattern: id=FILE_ID
  var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match2) return match2[1];
  
  // Pattern: เป็น ID โดยตรง (25+ characters)
  var match3 = url.match(/^([a-zA-Z0-9_-]{25,})$/);
  if (match3) return match3[1];
  
  return null;
}

// --- Code.gs ---

// ฟังก์ชันดึงรายชื่อบทเรียนจากชีต Settings (Column C)
function getUnitList() {
  var sheet = ss.getSheetByName('settings');
  var data = sheet.getDataRange().getValues();
  var units = [];
  
  // เริ่มวนลูปจากแถวที่ 2 (index 1) เพื่อข้ามหัวตาราง
  // สมมติว่าชื่อ Unit อยู่คอลัมน์ C (Index 2)
  for (var i = 1; i < data.length; i++) {
    // เช็คว่ามีข้อมูลในช่อง Unit ไหม
    if (data[i][2] && String(data[i][2]).trim() !== "") {
      units.push(String(data[i][2]).trim());
    }
  }
  
  return units;
}

//*** แดชบอร์ดข้อมูลผู้ลงทะเบียนและผู้ได้รับเกียรติบัตร ***
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('user');
  
  // 1. ดึงข้อมูลสถานภาพ (G2:K2)
  // G2: ผู้บริหาร, H2: ครู, I2: บุคลากรทางการศึกษา, J2: นักศึกษา, K2: บุคคลทั่วไป
  const statusValues = sheet.getRange('G2:K2').getValues()[0]; // [2, 3, 1, 2, 1]
  const statusLabels = ['ผู้บริหาร', 'ครู', 'บุคลากรฯ', 'นักศึกษา', 'บุคคลทั่วไป'];
  const totalRegistered = sheet.getRange('L2').getValue(); // รวมผู้ลงทะเบียน (L2)
  const totalCertificate = sheet.getRange('M2').getValue(); // ผู้ได้รับเกียรติบัตร (M2)

  // ----------------------------------------------------
  // 1.1 ประมวลผลข้อมูลสำหรับ Pie Chart (สถานภาพ)
  // ----------------------------------------------------
  const pieData = [];
  pieData.push(['สถานภาพ', 'จำนวนผู้ลงทะเบียน']); // Header
  
  // ตรวจสอบความถูกต้องของค่ารวม (เผื่อค่า L2 ผิดพลาด)
  const actualTotal = statusValues.reduce((sum, current) => sum + current, 0);

  for (let i = 0; i < statusValues.length; i++) {
    const count = statusValues[i];
    let percentage = (count / actualTotal) * 100;
    
    // สร้าง Label ที่รวมจำนวนและร้อยละ
    const labelWithCountAndPercent = `${statusLabels[i]}: ${count} คน | ${percentage.toFixed(2)}%`;
    
    // เก็บข้อมูลในรูปแบบที่ Google Charts Pie Chart ต้องการ
    pieData.push([labelWithCountAndPercent, count]);
  }

  // ----------------------------------------------------
  // 1.2 ประมวลผลข้อมูลสำหรับ Bar Chart (เปรียบเทียบ)
  // ----------------------------------------------------
  const barData = [
    // Header: เพิ่มคอลัมน์ 'style' สำหรับกำหนดสีแท่ง และ 'annotation' สำหรับตัวเลข
    ['รายการ', 'จำนวน', {role: 'style'}, {role: 'annotation'}], 
    
    // Data 1: ผู้ลงทะเบียนทั้งหมด (สีน้ำเงิน)
    ['ผู้ลงทะเบียนทั้งหมด', totalRegistered, '#007bff', totalRegistered.toString()],
    
    // Data 2: ผู้ได้รับเกียรติบัตร (สีแดง/เขียว)
    ['ผู้ได้รับเกียรติบัตร', totalCertificate, '#28a745', totalCertificate.toString()] // ใช้สีเขียวสวยงาม
    // หรือจะใช้สีแดงก็ได้: '#dc3545'
  ];
  
  return {
    pieChartData: pieData,
    barChartData: barData,
    totalRegistered: totalRegistered, // อาจใช้แสดงเป็นข้อความเสริมได้
    totalCertificate: totalCertificate
  };
}

function extractLinkFromHtmlOnly(html) {
  if (!html) return "";
  
  // 1. หาลิงก์ Google Drive / Docs / Video / Audio
  var regex = /(https?:\/\/(?:drive|docs)\.google\.com\/[^\s"']+)/;
  var match = html.match(regex);
  if (match) return match[1];

  // 2. หาลิงก์ Youtube
  var ytRegex = /(https?:\/\/(?:www\.)?(?:youtube\.com|youtu\.be)\/[^\s"']+)/;
  var ytMatch = html.match(ytRegex);
  if (ytMatch) return ytMatch[1];
  
  return "";
}

// --- Tool: ฟังก์ชันสำหรับกดรัน 1 ครั้ง เพื่อซ่อมข้อมูลเก่า (Migration) ---
// *** วิธีใช้: เลือกฟังก์ชันนี้ในตัวเลือกด้านบนแล้วกด "Run" (เรียกใช้) ***
function runFixExtractLinks() {
  var lessonSheet = ss.getSheetByName('lessons');
  var examSheet = ss.getSheetByName('exams');
  
  // 1. ซ่อมชีท Lessons
  var lData = lessonSheet.getDataRange().getValues();
  for (var i = 1; i < lData.length; i++) {
    var content = lData[i][4]; // Col E: ContentData
    var currentLink = lData[i][5]; // Col F: MediaLink
    
    // ถ้าช่องลิงก์ว่างเปล่า ให้พยายามดึงจากเนื้อหา
    if (!currentLink || currentLink === "") {
      var extracted = extractLinkFromHtmlOnly(content);
      if (extracted) {
        lessonSheet.getRange(i + 1, 6).setValue(extracted); // Set Col F
        // อัปเดต MediaType (Col K / Index 10) ด้วย
        var type = detectMediaType(content, extracted);
        lessonSheet.getRange(i + 1, 11).setValue(type);
      }
    }
  }

  // 2. ซ่อมชีท Exams
  var eData = examSheet.getDataRange().getValues();
  for (var j = 1; j < eData.length; j++) {
    var question = eData[j][4]; // Col E: Question
    var currentLink = eData[j][7]; // Col H: MediaLink
    
    // ถ้าช่องลิงก์ว่างเปล่า ให้พยายามดึงจากโจทย์
    if (!currentLink || currentLink === "") {
      var extracted = extractLinkFromHtmlOnly(question);
      if (extracted) {
        examSheet.getRange(j + 1, 8).setValue(extracted); // Set Col H
        // อัปเดต MediaType (Col I / Index 8) ด้วย
        var type = detectMediaType(question, extracted);
        examSheet.getRange(j + 1, 9).setValue(type);
      }
    }
  }
  
  // ล้าง Cache เพื่อให้หน้าเว็บเห็นการเปลี่ยนแปลง
  clearAllCaches();
  return "ดึงลิงก์จากเนื้อหามาใส่คอลัมน์ MediaLink เรียบร้อยแล้ว";
}

function runConsolidateVisitorLogs() {
  var sheet = ss.getSheetByName('visitor_logs');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // ไม่มีข้อมูล

  // ดึงข้อมูลทั้งหมดมาประมวลผล
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var groupedData = {};
  
  // วนลูปนับรวมยอดตามวันที่
  for (var i = 0; i < data.length; i++) {
    var rawDate = data[i][0];
    var count = data[i][1] ? parseInt(data[i][1]) : 1; // ถ้ามีเลขอยู่แล้วให้ใช้เลขนั้น ถ้าไม่มี(แบบเก่า) ให้นับ 1
    
    if (rawDate instanceof Date) {
      var dateKey = Utilities.formatDate(rawDate, "GMT+7", "yyyy-MM-dd"); // ใช้ Format สากลเพื่อการจัดกลุ่ม
      
      if (groupedData[dateKey]) {
        groupedData[dateKey] += count;
      } else {
        groupedData[dateKey] = count;
      }
    }
  }

  // เตรียมข้อมูลสำหรับเขียนกลับ
  var newData = [];
  // เรียงวันที่จากอดีต -> ปัจจุบัน
  var sortedKeys = Object.keys(groupedData).sort();
  
  for (var j = 0; j < sortedKeys.length; j++) {
    var k = sortedKeys[j];
    // แปลง String กลับเป็น Date Object (เวลา 00:00:00) เพื่อให้ Google Sheet เข้าใจว่าเป็นวันที่
    newData.push([new Date(k), groupedData[k]]);
  }

  // ล้างข้อมูลเก่าทั้งหมด แล้วเขียนข้อมูลใหม่ที่ยุบรวมแล้วลงไป
  sheet.getRange(2, 1, sheet.getLastRow(), 2).clearContent();
  
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 2).setValues(newData);
    // จัด Format วันที่ให้สวยงาม (dd/MM/yyyy)
    sheet.getRange(2, 1, newData.length, 1).setNumberFormat("dd/MM/yyyy");
  }
  
  return "รวมยอดผู้เข้าชมเรียบร้อยแล้ว";
}

// --- 2. แก้ไข recordVisit (เช็ควันปัจจุบันก่อนบันทึก) ---
function recordVisit() {
  var sheet = ss.getSheetByName('visitor_logs');
  if (!sheet) {
    sheet = ss.insertSheet('visitor_logs');
    sheet.appendRow(['TimeStamp', 'Total']); // สร้างหัวตารางถ้ายังไม่มี
  }
  
  var now = new Date();
  var todayStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
  
  var lastRow = sheet.getLastRow();
  
  // กรณีเพิ่งสร้างชีตใหม่ หรือยังไม่มีข้อมูล
  if (lastRow < 2) {
    sheet.appendRow([now, 1]);
    sheet.getRange(2, 1).setNumberFormat("dd/MM/yyyy");
    return;
  }
  
  // เช็ควันที่ของบรรทัดสุดท้าย
  var lastDateVal = sheet.getRange(lastRow, 1).getValue();
  var lastDateStr = "";
  if (lastDateVal instanceof Date) {
    lastDateStr = Utilities.formatDate(lastDateVal, "GMT+7", "dd/MM/yyyy");
  }
  
  // ถ้าวันที่ตรงกัน ให้บวกเพิ่มใน Col B
  if (lastDateStr === todayStr) {
    var currentCount = sheet.getRange(lastRow, 2).getValue();
    var newCount = (parseInt(currentCount) || 0) + 1;
    sheet.getRange(lastRow, 2).setValue(newCount);
  } else {
    // ถ้าวันที่ไม่ตรง (ขึ้นวันใหม่) ให้เพิ่มแถวใหม่
    sheet.appendRow([now, 1]);
    sheet.getRange(lastRow + 1, 1).setNumberFormat("dd/MM/yyyy");
  }
}

function runFixMediaTypeErrors() {
  var sheetsToFix = ['lessons', 'exams']; // ชื่อ Sheet ที่ต้องการแก้
  var log = [];

  sheetsToFix.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    var lastRow = sheet.getLastRow();
    
    // กำหนด Index คอลัมน์ให้ถูกตาม Sheet
    // lessons: Content=E(4), Link=F(5), Type=K(10)
    // exams:   Question=E(4), Link=H(7), Type=I(8)
    var colContent = 4;
    var colLink = (sheetName === 'lessons') ? 5 : 7;
    var colType = (sheetName === 'lessons') ? 10 : 8;

    for (var i = 1; i < data.length; i++) {
      var content = data[i][colContent];
      var link = data[i][colLink];
      var currentType = data[i][colType];
      
      // คำนวณประเภทใหม่
      var newType = detectMediaType(content, link);
      
      // ถ้าประเภทใหม่ไม่ตรงกับของเดิม และของเดิมผิด (เช่น เป็น audio แต่จริงๆ คือ image)
      if (newType !== '' && newType !== currentType) {
         // บันทึกค่าใหม่ลง Sheet
         sheet.getRange(i + 1, colType + 1).setValue(newType);
         log.push(sheetName + " Row " + (i+1) + ": " + currentType + " -> " + newType);
      }
    }
  });
  
  clearAllCaches();
  return log.length > 0 ? "แก้ไขเรียบร้อย: \n" + log.join("\n") : "ข้อมูลถูกต้องอยู่แล้ว ไม่มีการแก้ไข";
}
