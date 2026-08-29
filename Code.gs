const CONFIG = {
  SPREADSHEET_ID: '1ueua9ETipeU7bA_W3wQdHpF3ujLdZhFN1QyUy-Fb8S0',
  DRIVE_FOLDER_ID: '1aoOH3ILlHRqk0TQe7li-nEwUXzK1jOkn',
  ADMIN_EMAIL: 'info.documentsmonaco@gmail.com',
  SESSION_HOURS: 15,
  MAX_UPLOAD_MB: 10
};

const SHEETS = {
  USERS: 'Users', COURSES: 'Courses', LESSONS: 'Lessons', ENROLLMENTS: 'Enrollments',
  PROGRESS: 'Progress', ASSIGNMENTS: 'Assignments', SUBMISSIONS: 'Submissions',
  QUIZZES: 'Quizzes', QUIZ_RESULTS: 'QuizResults', FEEDBACK: 'Feedback', CERTIFICATES: 'Certificates'
};

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : 'health';
    if (action === 'health') return json_({ok:true, service:'IUM LMS API', version:'2.0.0'});
    return json_({ok:false,error:'Use POST for LMS actions.'});
  } catch (err) { return json_({ok:false,error:err.message}); }
}

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(d.action || '');
    let data;
    switch (action) {
      case 'setup': data = setupSheets_(); break;
      case 'register': data = register_(d); break;
      case 'login': data = login_(d); break;
      case 'logout': data = logout_(d); break;
      case 'me': data = me_(d); break;
      case 'studentCourses': data = studentCourses_(d); break;
      case 'courseLessons': data = courseLessons_(d); break;
      case 'saveProgress': data = saveProgress_(d); break;
      case 'submitAssignment': data = submitAssignment_(d); break;
      case 'submitQuiz': data = submitQuiz_(d); break;
      case 'studentDashboard': data = studentDashboard_(d); break;
      case 'myFeedback': data = myFeedback_(d); break;
      case 'adminOverview': data = adminOnly_(d, adminOverview_); break;
      case 'adminStudents': data = adminOnly_(d, adminStudents_); break;
      case 'approveUser': data = adminOnly_(d, approveUser_); break;
      case 'assignCourse': data = adminOnly_(d, assignCourse_); break;
      case 'adminCourses': data = adminOnly_(d, adminCourses_); break;
      case 'adminCourseLessons': data = adminOnly_(d, adminCourseLessons_); break;
      case 'saveCourse': data = adminOnly_(d, saveCourse_); break;
      case 'saveLesson': data = adminOnly_(d, saveLesson_); break;
      case 'deleteCourse': data = adminOnly_(d, deleteCourse_); break;
      case 'deleteLesson': data = adminOnly_(d, deleteLesson_); break;
      case 'adminProgress': data = adminOnly_(d, adminProgress_); break;
      case 'saveFeedback': data = staffOnly_(d, saveFeedback_); break;
      default: throw Error('Unknown action.');
    }
    return json_({ok:true,data:data});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)});
  }
}

function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ss_() { if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.indexOf('PASTE_') === 0) throw Error('Set SPREADSHEET_ID in Code.gs.'); return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function now_() { return new Date(); }

function setupSheets_() {
  const defs = {};
  defs[SHEETS.USERS] = ['Timestamp','User ID','Full Name','Email','Password Hash','Role','Status'];
  defs[SHEETS.COURSES] = ['Course ID','Course Code','Title','Description','Active','Instructor Email'];
  defs[SHEETS.LESSONS] = ['Lesson ID','Course ID','Lesson No','Title','Description','YouTube URL','PDF URL','Active'];
  defs[SHEETS.ENROLLMENTS] = ['Timestamp','Enrollment ID','Student ID','Course ID','Status','Assigned By'];
  defs[SHEETS.PROGRESS] = ['Timestamp','Student ID','Course ID','Lesson ID','Completed'];
  defs[SHEETS.ASSIGNMENTS] = ['Assignment ID','Course ID','Lesson ID','Title','Instructions','Due Date','Active'];
  defs[SHEETS.SUBMISSIONS] = ['Timestamp','Submission ID','Assignment ID','Student ID','Course ID','Lesson ID','File Name','Drive URL','Status','Score','Feedback'];
  defs[SHEETS.QUIZZES] = ['Quiz ID','Course ID','Lesson ID','Title','Questions JSON','Active'];
  defs[SHEETS.QUIZ_RESULTS] = ['Timestamp','Result ID','Quiz ID','Student ID','Course ID','Lesson ID','Score','Total'];
  defs[SHEETS.FEEDBACK] = ['Timestamp','Student ID','Course ID','Lesson ID','Feedback','Teacher'];
  defs[SHEETS.CERTIFICATES] = ['Timestamp','Certificate ID','Student ID','Course ID','Issued Date','Status'];
  const s = ss_();
  Object.keys(defs).forEach(n => { let sh=s.getSheetByName(n); if(!sh) sh=s.insertSheet(n); if(sh.getLastRow()===0) sh.appendRow(defs[n]); });
  seed_();
  return {ready:true,sheets:Object.keys(defs)};
}

function seed_() {
  const s=ss_();
  const c=s.getSheetByName(SHEETS.COURSES);
  if(c.getLastRow()<=1) c.appendRow(['BM101','BM101','Diploma in Business Management','Core business management programme',true,'']);
  const l=s.getSheetByName(SHEETS.LESSONS);
  if(l.getLastRow()<=1) l.getRange(2,1,4,8).setValues([
    ['BM101-L01','BM101',1,'Introduction to Business Management','Introduction to business management concepts','','',true],
    ['BM101-L02','BM101',2,'Principles of Management','Planning, organizing, leading and controlling','','',true],
    ['BM101-L03','BM101',3,'Business Environment','Internal and external business environment','','',true],
    ['BM101-L04','BM101',4,'Entrepreneurship Basics','Entrepreneurship and opportunity recognition','','',true]
  ]);
}

function rows_(name) { const sh=ss_().getSheetByName(name); return sh && sh.getLastRow()>1 ? sh.getDataRange().getValues().slice(1) : []; }
function append_(name,row){ ss_().getSheetByName(name).appendRow(row); }
function uid_(prefix){ return prefix+'-'+Utilities.getUuid().split('-')[0].toUpperCase(); }
function hash_(p){ return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(p),Utilities.Charset.UTF_8).map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join(''); }
function email_(x){ return String(x||'').trim().toLowerCase(); }
function active_(x){ return String(x).toLowerCase()==='true' || String(x).toLowerCase()==='active'; }
function safeEq_(a,b){ return String(a||'').trim().toLowerCase()===String(b||'').trim().toLowerCase(); }

function register_(d){
  setupSheets_();
  const name=String(d.fullName||'').trim(), email=email_(d.email), id=String(d.studentId||'').trim(), pass=String(d.password||'');
  if(!name||!email||!id||!pass) throw Error('Complete all registration fields.');
  if(pass.length<6) throw Error('Password must be at least 6 characters.');
  const users=rows_(SHEETS.USERS);
  if(users.some(r=>safeEq_(r[1],id)||safeEq_(r[3],email))) throw Error('Student ID or email already exists.');
  append_(SHEETS.USERS,[now_(),id,name,email,hash_(pass),'STUDENT','PENDING']);
  return {userId:id,fullName:name,email:email,status:'PENDING',message:'Registration submitted. Wait for administrator approval.'};
}

function login_(d){
  setupSheets_();
  const email=email_(d.email), pass=String(d.password||'');
  const r=rows_(SHEETS.USERS).find(x=>safeEq_(x[3],email) && String(x[4])===hash_(pass));
  if(!r) throw Error('Invalid email or password.');
  if(!safeEq_(r[6],'ACTIVE')) throw Error('Your account is pending or disabled. Contact the administrator.');
  const role=String(r[5]).toUpperCase(), token=Utilities.getUuid(), expires=Date.now()+CONFIG.SESSION_HOURS*3600000;
  PropertiesService.getScriptProperties().setProperty('SESSION_'+token,JSON.stringify({userId:String(r[1]),email:String(r[3]),role:role,expires:expires}));
  return {token:token,userId:String(r[1]),fullName:String(r[2]),email:String(r[3]),role:role,expires:expires};
}
function session_(token){
  if(!token) throw Error('Login required.');
  const key='SESSION_'+String(token), raw=PropertiesService.getScriptProperties().getProperty(key);
  if(!raw) throw Error('Session expired. Please login again.');
  const s=JSON.parse(raw);
  if(Date.now()>Number(s.expires)){ PropertiesService.getScriptProperties().deleteProperty(key); throw Error('Session expired. Please login again.'); }
  return s;
}
function logout_(d){ if(d.token) PropertiesService.getScriptProperties().deleteProperty('SESSION_'+String(d.token)); return true; }
function me_(d){ const s=session_(d.token); return s; }
function requireRole_(d,roles){ const s=session_(d.token); if(roles.indexOf(s.role)<0) throw Error('Access denied.'); return s; }
function adminOnly_(d,fn){ const s=requireRole_(d,['ADMIN']); return fn(d,s); }
function staffOnly_(d,fn){ const s=requireRole_(d,['ADMIN','INSTRUCTOR']); return fn(d,s); }

function enrollmentsFor_(studentId){ return rows_(SHEETS.ENROLLMENTS).filter(r=>safeEq_(r[2],studentId)&&safeEq_(r[4],'ACTIVE')); }
function enrolled_(studentId,courseId){ return enrollmentsFor_(studentId).some(r=>safeEq_(r[3],courseId)); }
function courseRow_(id){ return rows_(SHEETS.COURSES).find(r=>safeEq_(r[0],id)&&active_(r[4])); }
function lessonRow_(id){ return rows_(SHEETS.LESSONS).find(r=>safeEq_(r[0],id)&&active_(r[7])); }
function assertStudent_(d){ const s=requireRole_(d,['STUDENT']); if(!safeEq_(s.userId,d.studentId)) throw Error('Student identity mismatch.'); return s; }

function studentCourses_(d){
  const s=assertStudent_(d);
  const ids=enrollmentsFor_(s.userId).map(r=>String(r[3]).toLowerCase());
  return rows_(SHEETS.COURSES).filter(r=>ids.indexOf(String(r[0]).toLowerCase())>=0 && active_(r[4])).map(r=>({id:r[0],code:r[1],title:r[2],description:r[3],instructor:r[5]}));
}
function courseLessons_(d){
  const s=assertStudent_(d), cid=String(d.courseId||'');
  if(!enrolled_(s.userId,cid)) throw Error('You are not enrolled in this course.');
  const c=courseRow_(cid); if(!c) throw Error('Course not found.');
  const progress=rows_(SHEETS.PROGRESS).filter(r=>safeEq_(r[1],s.userId)&&safeEq_(r[2],cid)&&active_(r[4])).map(r=>String(r[3]));
  return {course:{id:c[0],code:c[1],title:c[2],description:c[3]},lessons:rows_(SHEETS.LESSONS).filter(r=>safeEq_(r[1],cid)&&active_(r[7])).sort((a,b)=>Number(a[2])-Number(b[2])).map(r=>({id:r[0],no:r[2],title:r[3],description:r[4],youtube:r[5],pdf:r[6],completed:progress.indexOf(String(r[0]))>=0}))};
}
function saveProgress_(d){
  const s=assertStudent_(d), cid=String(d.courseId||''), lid=String(d.lessonId||'');
  if(!enrolled_(s.userId,cid)) throw Error('You are not enrolled in this course.');
  const l=lessonRow_(lid); if(!l||!safeEq_(l[1],cid)) throw Error('Lesson does not belong to this course.');
  const sh=ss_().getSheetByName(SHEETS.PROGRESS), all=rows_(SHEETS.PROGRESS);
  if(!all.some(r=>safeEq_(r[1],s.userId)&&safeEq_(r[2],cid)&&safeEq_(r[3],lid))) append_(SHEETS.PROGRESS,[now_(),s.userId,cid,lid,true]);
  return {saved:true};
}

function submitAssignment_(d){
  const s=assertStudent_(d), aid=String(d.assignmentId||''), a=rows_(SHEETS.ASSIGNMENTS).find(r=>safeEq_(r[0],aid)&&active_(r[6]));
  if(!a) throw Error('Assignment not found.');
  if(!enrolled_(s.userId,a[1])) throw Error('You are not enrolled in this course.');
  if(!safeEq_(a[2],d.lessonId)) throw Error('Assignment lesson mismatch.');
  if(!d.fileData) throw Error('Choose a file.');
  if(String(d.fileData).length>CONFIG.MAX_UPLOAD_MB*1024*1024*1.4) throw Error('File is too large. Maximum '+CONFIG.MAX_UPLOAD_MB+' MB.');
  if(!CONFIG.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID.indexOf('PASTE_')===0) throw Error('Set DRIVE_FOLDER_ID in Code.gs.');
  const folder=DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const blob=Utilities.newBlob(Utilities.base64Decode(String(d.fileData)),String(d.mimeType||'application/octet-stream'),String(d.fileName||'assignment'));
  const file=folder.createFile(blob);
  append_(SHEETS.SUBMISSIONS,[now_(),uid_('SUB'),aid,s.userId,a[1],a[2],String(d.fileName||''),file.getUrl(),'SUBMITTED','', '']);
  return {submitted:true,url:file.getUrl()};
}
function submitQuiz_(d){
  const s=assertStudent_(d), q=rows_(SHEETS.QUIZZES).find(r=>safeEq_(r[0],d.quizId)&&active_(r[5]));
  if(!q) throw Error('Quiz not found.');
  if(!enrolled_(s.userId,q[1])) throw Error('You are not enrolled in this course.');
  if(!safeEq_(q[2],d.lessonId)) throw Error('Quiz lesson mismatch.');
  append_(SHEETS.QUIZ_RESULTS,[now_(),uid_('QUIZ'),q[0],s.userId,q[1],q[2],Number(d.score||0),Number(d.total||0)]);
  return {saved:true};
}
function studentDashboard_(d){
  const s=assertStudent_(d), cids=enrollmentsFor_(s.userId).map(r=>String(r[3]));
  const p=rows_(SHEETS.PROGRESS).filter(r=>safeEq_(r[1],s.userId)&&cids.indexOf(String(r[2]))>=0);
  const q=rows_(SHEETS.QUIZ_RESULTS).filter(r=>safeEq_(r[3],s.userId)&&cids.indexOf(String(r[4]))>=0);
  const sub=rows_(SHEETS.SUBMISSIONS).filter(r=>safeEq_(r[3],s.userId)&&cids.indexOf(String(r[4]))>=0);
  return {courses:cids.length,completedLessons:p.length,quizResults:q.map(r=>({quizId:r[2],courseId:r[4],lessonId:r[5],score:r[6],total:r[7]})),submissions:sub.length};
}
function myFeedback_(d){
  const s=assertStudent_(d), cid=String(d.courseId||'');
  if(!enrolled_(s.userId,cid)) throw Error('You are not enrolled in this course.');
  return rows_(SHEETS.FEEDBACK).filter(r=>safeEq_(r[1],s.userId)&&safeEq_(r[2],cid)).map(r=>({lessonId:r[3],feedback:r[4],teacher:r[5],timestamp:r[0]}));
}

function adminOverview_(){ return {students:rows_(SHEETS.USERS).filter(r=>safeEq_(r[5],'STUDENT')).length,instructors:rows_(SHEETS.USERS).filter(r=>safeEq_(r[5],'INSTRUCTOR')).length,courses:rows_(SHEETS.COURSES).filter(r=>active_(r[4])).length,lessons:rows_(SHEETS.LESSONS).filter(r=>active_(r[7])).length,submissions:rows_(SHEETS.SUBMISSIONS).length}; }
function adminStudents_(){ return rows_(SHEETS.USERS).map(r=>({id:r[1],name:r[2],email:r[3],role:r[5],status:r[6]})); }
function approveUser_(d){ const id=String(d.userId||''); const sh=ss_().getSheetByName(SHEETS.USERS), data=sh.getDataRange().getValues(); for(let i=1;i<data.length;i++){ if(safeEq_(data[i][1],id)){ sh.getRange(i+1,7).setValue(String(d.status||'ACTIVE').toUpperCase()); return true; }} throw Error('User not found.'); }
function assignCourse_(d,s){ const studentId=String(d.studentId||''), courseId=String(d.courseId||''); if(!rows_(SHEETS.USERS).some(r=>safeEq_(r[1],studentId)&&safeEq_(r[5],'STUDENT'))) throw Error('Student not found.'); if(!courseRow_(courseId)) throw Error('Course not found.'); if(enrolled_(studentId,courseId)) return true; append_(SHEETS.ENROLLMENTS,[now_(),uid_('ENR'),studentId,courseId,'ACTIVE',s.email]); return true; }

function adminCourseLessons_(d){
  const cid=String(d.courseId||'');
  const c=courseRow_(cid); if(!c) throw Error('Course not found.');
  return {course:{id:c[0],code:c[1],title:c[2]},lessons:rows_(SHEETS.LESSONS).filter(r=>safeEq_(r[1],cid)).sort((a,b)=>Number(a[2])-Number(b[2])).map(r=>({id:r[0],no:r[2],title:r[3],description:r[4],youtube:r[5],pdf:r[6],active:active_(r[7])}))};
}
function adminCourses_(){ return rows_(SHEETS.COURSES).map(r=>({id:r[0],code:r[1],title:r[2],description:r[3],active:r[4],instructor:r[5]})); }
function saveCourse_(d){ const id=String(d.id||'').trim(); if(!id||!d.title) throw Error('Course ID and title are required.'); if(courseRow_(id)) throw Error('Course ID already exists.'); append_(SHEETS.COURSES,[id,String(d.code||id),String(d.title),String(d.description||''),true,String(d.instructorEmail||'')]); return true; }
function saveLesson_(d){ const id=String(d.id||'').trim(), cid=String(d.courseId||'').trim(); if(!id||!cid||!d.title) throw Error('Lesson ID, Course ID and title are required.'); if(!courseRow_(cid)) throw Error('Course not found.'); if(lessonRow_(id)) throw Error('Lesson ID already exists.'); append_(SHEETS.LESSONS,[id,cid,Number(d.lessonNo||1),String(d.title),String(d.description||''),String(d.youtube||''),String(d.pdf||''),true]); return true; }
function deleteCourse_(d){ const sh=ss_().getSheetByName(SHEETS.COURSES), data=sh.getDataRange().getValues(); for(let i=1;i<data.length;i++) if(safeEq_(data[i][0],d.courseId)){sh.getRange(i+1,5).setValue(false);return true;} throw Error('Course not found.'); }
function deleteLesson_(d){ const sh=ss_().getSheetByName(SHEETS.LESSONS), data=sh.getDataRange().getValues(); for(let i=1;i<data.length;i++) if(safeEq_(data[i][0],d.lessonId)){sh.getRange(i+1,8).setValue(false);return true;} throw Error('Lesson not found.'); }
function adminProgress_(){ return rows_(SHEETS.PROGRESS).map(r=>({studentId:r[1],courseId:r[2],lessonId:r[3],completed:r[4],timestamp:r[0]})); }
function saveFeedback_(d,s){ if(!d.studentId||!d.courseId||!d.lessonId||!d.feedback) throw Error('Complete feedback fields.'); if(!enrolled_(d.studentId,d.courseId)) throw Error('Student is not enrolled in this course.'); append_(SHEETS.FEEDBACK,[now_(),d.studentId,d.courseId,d.lessonId,String(d.feedback),String(d.teacher||s.email)]); return true; }
