// const express = require('express');
// const multer = require('multer');
// const ExcelJS = require('exceljs');
// const fs = require('fs');
// const path = require('path');
// const cloudinary = require('cloudinary').v2;
// const streamifier = require('streamifier');
// const lockfile = require('proper-lockfile');
// require("dotenv").config();
// const cors = require("cors");

// const app = express();
// const PORT = process.env.PORT || 3000;

// function logMemoryUsage() {
//   const usage = process.memoryUsage();
//   console.log('📊 استهلاك الذاكرة:', {
//     rss: `${Math.round(usage.rss / 1024 / 1024)} MB`,
//     heapUsed: `${Math.round(usage.heapUsed / 1024 / 1024)} MB`,
//     heapTotal: `${Math.round(usage.heapTotal / 1024 / 1024)} MB`,
//     external: `${Math.round(usage.external / 1024 / 1024)} MB`
//   });
// }

// // تشغيل مراقب الذاكرة كل 30 ثانية
// setInterval(logMemoryUsage, 30000);

// cloudinary.config({
//   cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
//   api_key: process.env.CLOUDINARY_API_KEY,
//   api_secret: process.env.CLOUDINARY_API_SECRET,
// });

// // Middleware
// app.use(cors());
// app.use(express.json());
// app.use(express.urlencoded({ extended: true }));

// app.use((req, res, next) => {
//   res.on('finish', () => {
//     if (global.gc) {
//       global.gc();
//     }
//   });
//   next();
// });

// // Multer memory storage
// const upload = multer({
//   storage: multer.memoryStorage(),
//   limits: { fileSize: 100 * 1024 * 1024 }, // 100MB
//   fileFilter: (req, file, cb) => {
//     if (file.mimetype.startsWith('image/')) cb(null, true);
//     else cb(new Error('يُسمح بالصور فقط!'), false);
//   },
// });

// const excelFilePath = path.join(__dirname, 'spider_agency_applications.xlsx');

// const EXCEL_COLUMNS = [
//   { header: 'النوع', key: 'gender', width: 15 },
//   { header: 'الاسم الثلاثي', key: 'fullName', width: 25 },
//   { header: 'تاريخ الميلاد', key: 'birthDate', width: 20 },
//   { header: 'رقم التليفون', key: 'phoneNumber', width: 20 },
//   { header: 'الايميل', key: 'email', width: 25 },
//   { header: 'الجامعة', key: 'university', width: 20 },
//   { header: 'الكلية', key: 'college', width: 20 },
//   { header: 'تاريخ التخرج', key: 'graduationDate', width: 20 },
//   { header: 'العنوان', key: 'address', width: 30 },
//   { header: 'لديه خبرة في التسويق', key: 'hasMarketingExperience', width: 25 },
//   { header: 'الخبرة في التسويق', key: 'marketingExperience', width: 30 },
//   { header: 'أهداف التدريب', key: 'trainingGoals', width: 30 },
//   { header: 'فكرة المنتج', key: 'productIdea', width: 30 },
//   { header: 'رابط الصورة', key: 'imageUrl', width: 40 },
//   { header: 'تاريخ التقديم', key: 'submissionDate', width: 20 },
// ];

// async function recreateExcelFile() {
//   try {
//     console.log('🔄 إعادة إنشاء ملف Excel من البداية...');
    
//     if (fs.existsSync(excelFilePath)) {
//       const backupPath = excelFilePath.replace('.xlsx', `_backup_${Date.now()}.xlsx`);
//       fs.renameSync(excelFilePath, backupPath);
//       console.log(`📦 تم إنشاء نسخة احتياطية: ${backupPath}`);
//     }

//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet('Applications');
    
//     worksheet.columns = EXCEL_COLUMNS;
    
//     const headerRow = worksheet.getRow(1);
//     EXCEL_COLUMNS.forEach((col, index) => {
//       const cell = headerRow.getCell(index + 1);
//       cell.value = col.header;
//       cell.font = { bold: true };
//       cell.fill = {
//         type: 'pattern',
//         pattern: 'solid',
//         fgColor: { argb: 'FFE6E6FA' }
//       };
//     });
    
//     await workbook.xlsx.writeFile(excelFilePath);
//     console.log('✅ تم إنشاء ملف Excel جديد بنجاح');
    
//     return true;
//   } catch (error) {
//     console.error('❌ خطأ في إعادة إنشاء ملف Excel:', error);
//     throw error;
//   }
// }

// async function ensureExcelExists() {
//   try {
//     const dir = path.dirname(excelFilePath);
//     if (!fs.existsSync(dir)) {
//       fs.mkdirSync(dir, { recursive: true });
//       console.log('📁 تم إنشاء المجلد');
//     }

//     if (!fs.existsSync(excelFilePath)) {
//       await recreateExcelFile();
//       return;
//     }

//     try {
//       const stats = fs.statSync(excelFilePath);
//       console.log(`📊 حجم الملف: ${(stats.size / 1024).toFixed(2)} KB`);
      
//       if (stats.size > 10 * 1024 * 1024) {
//         console.warn('⚠️ الملف كبير جداً، قد يحتوي على مشاكل');
//         throw new Error('الملف كبير جداً');
//       }

//       const workbook = new ExcelJS.Workbook();
      
//       const stream = fs.createReadStream(excelFilePath);
//       await workbook.xlsx.read(stream);
      
//       const worksheet = workbook.getWorksheet('Applications');
//       if (!worksheet) {
//         throw new Error('ورقة العمل غير موجودة');
//       }

//       console.log(`✅ ملف Excel صالح، عدد الصفوف: ${worksheet.rowCount}`);
      
//     } catch (readError) {
//       console.error('❌ الملف موجود لكن فاسد:', readError.message);
//       await recreateExcelFile();
//     }

//   } catch (error) {
//     console.error('❌ خطأ في إعداد ملف Excel:', error);
//     throw error;
//   }
// }

// async function addApplicationToExcel(applicationData) {
//   let release = null;
  
//   try {
//     console.log('🔄 بدء إضافة البيانات لملف Excel...');
    
//     release = await lockfile.lock(excelFilePath, { 
//       retries: 5, 
//       stale: 30000,
//       update: 5000   
//     });
//     console.log('🔒 تم الحصول على قفل الملف');

//     if (!fs.existsSync(excelFilePath)) {
//       console.log('📑 الملف غير موجود، إنشاء ملف جديد...');
//       await recreateExcelFile();
//     }

//     const workbook = new ExcelJS.Workbook();
//     const readStream = fs.createReadStream(excelFilePath);
//     await workbook.xlsx.read(readStream);
    
//     const worksheet = workbook.getWorksheet('Applications');
//     if (!worksheet) {
//       throw new Error('ورقة العمل غير موجودة');
//     }

//     let lastRowWithData = 1; 
    
//     const maxRowsToCheck = Math.min(worksheet.rowCount + 10, 1000);
    
//     for (let i = 2; i <= maxRowsToCheck; i++) {
//       const row = worksheet.getRow(i);
//       if (row.getCell(1).value || row.getCell(2).value || row.getCell(3).value) {
//         lastRowWithData = i;
//       }
//     }

//     const newRowNumber = lastRowWithData + 1;
//     console.log(`📊 آخر صف يحتوي على بيانات: ${lastRowWithData}, الصف الجديد: ${newRowNumber}`);

//     const newRow = worksheet.getRow(newRowNumber);
    
//     const values = [
//       applicationData.gender,
//       applicationData.fullName,
//       applicationData.birthDate,
//       applicationData.phoneNumber,
//       applicationData.email,
//       applicationData.university,
//       applicationData.college,
//       applicationData.graduationDate,
//       applicationData.address,
//       applicationData.hasMarketingExperience,
//       applicationData.marketingExperience,
//       applicationData.trainingGoals,
//       applicationData.productIdea,
//       applicationData.imageUrl,
//       applicationData.submissionDate
//     ];

//     values.forEach((value, index) => {
//       newRow.getCell(index + 1).value = value;
//     });

//     newRow.commit();
//     console.log(`✅ تم إضافة الصف رقم: ${newRowNumber}`);

//     await workbook.xlsx.writeFile(excelFilePath);
//     console.log('✅ تم حفظ الملف بنجاح');

//     workbook.removeWorksheet('Applications');

//     const totalApplications = newRowNumber - 1;
    
//     return {
//       rowNumber: newRowNumber,
//       totalApplications
//     };

//   } catch (error) {
//     console.error('❌ خطأ في إضافة البيانات لملف Excel:', {
//       message: error.message,
//       stack: error.stack.substring(0, 500) 
//     });
    
//     if (error.message.includes('heap out of memory') || error.message.includes('allocation failed')) {
//       console.log('🔄 محاولة إعادة إنشاء الملف بسبب مشكلة في الذاكرة...');
//       try {
//         await recreateExcelFile();
//         return await addApplicationToExcel(applicationData);
//       } catch (recreateError) {
//         console.error('❌ فشل في إعادة إنشاء الملف:', recreateError.message);
//       }
//     }
    
//     throw error;
//   } finally {
//     if (release) {
//       try {
//         await release();
//         console.log('🔓 تم تحرير قفل الملف');
//       } catch (unlockError) {
//         console.error('❌ خطأ في تحرير القفل:', unlockError.message);
//       }
//     }
    
//     if (global.gc) {
//       global.gc();
//       console.log('🧹 تم تنظيف الذاكرة');
//     }
//   }
// }

// function uploadToCloudinary(fileBuffer) {
//   return new Promise((resolve, reject) => {
//     const stream = cloudinary.uploader.upload_stream(
//       { 
//         folder: 'spider-agency-applications',
//         resource_type: 'image'
//       },
//       (error, result) => {
//         if (error) {
//           console.error('❌ خطأ في رفع الصورة على Cloudinary:', error);
//           return reject(error);
//         }
//         resolve(result.secure_url);
//       }
//     );
//     streamifier.createReadStream(fileBuffer).pipe(stream);
//   });
// }

// // الصفحة الرئيسية
// app.get('/', (req, res) => {
//   res.json({ 
//     message: "🕷️ Spider Agency API",
//     status: "running",
//     endpoints: {
//       "POST /api/spider-agency/apply": "تقديم طلب جديد",
//       "GET /api/spider-agency/download": "تحميل ملف Excel",
//       "GET /api/spider-agency/stats": "إحصائيات الطلبات",
//       "GET /api/spider-agency/health": "فحص حالة النظام",
//       "POST /api/spider-agency/reset-excel": "إعادة إنشاء ملف Excel"
//     }
//   });
// });

// app.post('/api/spider-agency/apply', upload.single('image'), async (req, res) => {
//   try {
//     console.log('🔄 استلام طلب تقديم جديد...');
//     console.log('البيانات المستلمة:', { ...req.body, image: req.file ? 'موجود' : 'غير موجود' });

//     const {
//       gender,
//       fullName,
//       birthDate,
//       phoneNumber,
//       email,
//       university,
//       college,
//       graduationDate,
//       address,
//       hasMarketingExperience,
//       marketingExperience,
//       trainingGoals,
//       productIdea,
//     } = req.body;

//     // تحقق من الحقول المطلوبة
//     const requiredFields = {
//       gender: 'النوع',
//       fullName: 'الاسم الثلاثي',
//       birthDate: 'تاريخ الميلاد',
//       phoneNumber: 'رقم التليفون',
//       email: 'الايميل',
//       university: 'الجامعة',
//       college: 'الكلية',
//       graduationDate: 'تاريخ التخرج',
//       address: 'العنوان',
//       hasMarketingExperience: 'الخبرة في التسويق',
//       trainingGoals: 'أهداف التدريب',
//       productIdea: 'فكرة المنتج',
//     };

//     const missingFields = [];
//     for (const [field, arabicName] of Object.entries(requiredFields)) {
//       if (!req.body[field] || req.body[field].toString().trim() === '') {
//         missingFields.push(arabicName);
//       }
//     }

//     if (missingFields.length > 0) {
//       console.log('❌ حقول مفقودة:', missingFields);
//       return res.status(400).json({
//         success: false,
//         message: `الحقول التالية مطلوبة: ${missingFields.join(', ')}`,
//       });
//     }

//     if (!req.file) {
//       console.log('❌ الصورة مفقودة');
//       return res.status(400).json({
//         success: false,
//         message: 'الصورة الشخصية مطلوبة',
//       });
//     }

//     console.log('🔄 بدء رفع الصورة على Cloudinary...');
//     const imageUrl = await uploadToCloudinary(req.file.buffer);
//     console.log('✅ تم رفع الصورة بنجاح:', imageUrl);

//     // إعداد بيانات الطلب
//     const applicationData = {
//       gender: gender.trim(),
//       fullName: fullName.trim(),
//       birthDate: birthDate.trim(),
//       phoneNumber: phoneNumber.trim(),
//       email: email.trim(),
//       university: university.trim(),
//       college: college.trim(),
//       graduationDate: graduationDate.trim(),
//       address: address.trim(),
//       hasMarketingExperience: hasMarketingExperience.trim(),
//       marketingExperience: marketingExperience ? marketingExperience.trim() : 'لا توجد',
//       trainingGoals: trainingGoals.trim(),
//       productIdea: productIdea.trim(),
//       imageUrl,
//       submissionDate: new Date().toISOString().split('T')[0],
//     };

//     console.log('🔄 إضافة البيانات لملف Excel...');
    
//     const result = await addApplicationToExcel(applicationData);
//     console.log('✅ تم إضافة الطلب بنجاح لملف Excel:', result);

//     res.status(201).json({
//       success: true,
//       message: 'تم إرسال طلب التقديم بنجاح',
//       data: {
//         applicationId: Date.now(),
//         submissionDate: applicationData.submissionDate,
//         applicantName: fullName,
//         imageUrl,
//         excelRowNumber: result.rowNumber,
//         totalApplications: result.totalApplications
//       },
//     });
//   } catch (error) {
//     console.error('❌ خطأ في معالجة طلب التقديم:', {
//       message: error.message,
//       stack: error.stack,
//       name: error.name
//     });
//     res.status(500).json({
//       success: false,
//       message: 'حدث خطأ أثناء معالجة طلبك، يرجى المحاولة مرة أخرى',
//       error: process.env.NODE_ENV === 'development' ? error.message : undefined
//     });
//   }
// });

// app.get('/api/spider-agency/download', async (req, res) => {
//   try {
//     if (!fs.existsSync(excelFilePath)) {
//       return res.status(404).json({
//         success: false,
//         message: 'ملف Excel غير موجود',
//       });
//     }

//     const filename = `spider_agency_applications_${new Date().toISOString().split('T')[0]}.xlsx`;
    
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
//     res.setHeader('Content-Length', fs.statSync(excelFilePath).size);

//     const fileStream = fs.createReadStream(excelFilePath);
//     fileStream.pipe(res);

//     console.log('✅ تم تحميل ملف Excel بنجاح');
//   } catch (error) {
//     console.error('❌ خطأ في تحميل ملف Excel:', error);
//     res.status(500).json({
//       success: false,
//       message: 'حدث خطأ أثناء تحميل الملف',
//     });
//   }
// });

// app.get('/api/spider-agency/stats', async (req, res) => {
//   try {
//     if (!fs.existsSync(excelFilePath)) {
//       return res.json({
//         success: true,
//         data: {
//           totalApplications: 0,
//           fileExists: false,
//           message: 'لم يتم تقديم أي طلبات بعد'
//         }
//       });
//     }

//     // استخدام stream للقراءة
//     const workbook = new ExcelJS.Workbook();
//     const readStream = fs.createReadStream(excelFilePath);
//     await workbook.xlsx.read(readStream);
    
//     const worksheet = workbook.getWorksheet('Applications');
    
//     if (!worksheet) {
//       return res.json({
//         success: true,
//         data: {
//           totalApplications: 0,
//           fileExists: true,
//           message: 'الملف موجود لكن لا توجد طلبات'
//         }
//       });
//     }

//     let totalApplications = 0;
//     let maleCount = 0;
//     let femaleCount = 0;
//     let withExperience = 0;
    
//     // عد الصفوف التي تحتوي على بيانات فعلية
//     for (let i = 2; i <= worksheet.rowCount; i++) {
//       const row = worksheet.getRow(i);
//       const fullName = row.getCell(2).value;
      
//       if (fullName && fullName.toString().trim()) {
//         totalApplications++;
        
//         const gender = row.getCell(1).value;
//         const hasExp = row.getCell(10).value;
        
//         if (gender === 'ذكر') maleCount++;
//         else if (gender === 'أنثى' || gender === 'انثى') femaleCount++;
        
//         if (hasExp === 'نعم') withExperience++;
//       }
//     }

//     const fileStats = fs.statSync(excelFilePath);

//     // تنظيف الذاكرة
//     workbook.removeWorksheet('Applications');

//     res.json({
//       success: true,
//       data: {
//         totalApplications,
//         demographics: {
//           male: maleCount,
//           female: femaleCount
//         },
//         withMarketingExperience: withExperience,
//         fileExists: true,
//         fileSize: `${(fileStats.size / 1024).toFixed(2)} KB`,
//         lastModified: fileStats.mtime.toISOString().split('T')[0]
//       }
//     });

//   } catch (error) {
//     console.error('❌ خطأ في جلب الإحصائيات:', error);
//     res.status(500).json({
//       success: false,
//       message: 'حدث خطأ أثناء جلب الإحصائيات',
//     });
//   }
// });

// // إضافة endpoint لفحص حالة الذاكرة والنظام
// app.get('/api/spider-agency/health', (req, res) => {
//   const usage = process.memoryUsage();
//   const uptime = process.uptime();
  
//   res.json({
//     success: true,
//     data: {
//       uptime: `${Math.floor(uptime / 60)} دقيقة`,
//       memory: {
//         rss: `${Math.round(usage.rss / 1024 / 1024)} MB`,
//         heapUsed: `${Math.round(usage.heapUsed / 1024 / 1024)} MB`,
//         heapTotal: `${Math.round(usage.heapTotal / 1024 / 1024)} MB`
//       },
//       excelFileExists: fs.existsSync(excelFilePath),
//       excelFileSize: fs.existsSync(excelFilePath) ? `${(fs.statSync(excelFilePath).size / 1024).toFixed(2)} KB` : 'N/A',
//       timestamp: new Date().toISOString()
//     }
//   });
// });

// // إضافة endpoint لإعادة إنشاء ملف Excel في حالة الطوارئ
// app.post('/api/spider-agency/reset-excel', async (req, res) => {
//   try {
//     // يجب إضافة authentication هنا في الإنتاج
//     const { password } = req.body;
//     if (password !== process.env.ADMIN_PASSWORD && password !== 'spider_reset_2024') {
//       return res.status(401).json({
//         success: false,
//         message: 'غير مخول'
//       });
//     }

//     await recreateExcelFile();
    
//     res.json({
//       success: true,
//       message: 'تم إعادة إنشاء ملف Excel بنجاح'
//     });
//   } catch (error) {
//     console.error('❌ خطأ في إعادة إنشاء ملف Excel:', error);
//     res.status(500).json({
//       success: false,
//       message: 'حدث خطأ أثناء إعادة إنشاء الملف'
//     });
//   }
// });

// // Error handling middleware
// app.use((error, req, res, next) => {
//   if (error instanceof multer.MulterError) {
//     if (error.code === 'LIMIT_FILE_SIZE') {
//       return res.status(400).json({
//         success: false,
//         message: 'حجم الملف كبير جداً، الحد الأقصى 100MB',
//       });
//     }
//   }

//   console.error('❌ خطأ عام:', error);
//   res.status(500).json({
//     success: false,
//     message: error.message || 'حدث خطأ غير متوقع',
//   });
// });

// // 404 handler
// app.use('*', (req, res) => {
//   res.status(404).json({
//     success: false,
//     message: 'الصفحة المطلوبة غير موجودة',
//   });
// });

// // إنشاء ملف Excel عند بدء التشغيل
// ensureExcelExists().catch(console.error);

// // بدء تشغيل الخادم
// app.listen(PORT, () => {
//   console.log(`🕷️  Spider Agency API يعمل على المنفذ ${PORT}`);
//   console.log(`📊 ملف Excel: ${excelFilePath}`);
//   console.log(`☁️ الصور تُرفع على Cloudinary`);
//   console.log(`📊 مسار ملف Excel الكامل: ${path.resolve(excelFilePath)}`);
//   console.log(`🌐 الصفحة الرئيسية: http://localhost:${PORT}`);
//   console.log(`📥 تحميل Excel: http://localhost:${PORT}/api/spider-agency/download`);
//   console.log(`📈 الإحصائيات: http://localhost:${PORT}/api/spider-agency/stats`);
//   console.log(`🏥 فحص الحالة: http://localhost:${PORT}/api/spider-agency/health`);
  
//   // لوج الذاكرة الأولي
//   logMemoryUsage();
// });

// module.exports = app;



const express = require('express');
const multer = require('multer');
const { google } = require('googleapis');
const cloudinary = require('cloudinary').v2;
const streamifier = require('streamifier');
require("dotenv").config();
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

// Cloudinary Configuration
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Multer memory storage
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('يُسمح بالصور فقط!'), false);
  },
});

// Google Sheets Configuration
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;

function getGoogleSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      type: 'service_account',
      project_id: process.env.GOOGLE_PROJECT_ID,
      private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      client_id: process.env.GOOGLE_CLIENT_ID,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return google.sheets({ version: 'v4', auth });
}

async function initializeGoogleSheet() {
  try {
    const sheets = getGoogleSheetsClient();
    
    // التحقق من وجود الـ Sheet
    const response = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
    });

    // البحث عن ورقة "Applications"
    const sheet = response.data.sheets.find(s => s.properties.title === 'Applications');
    
    if (!sheet) {
      // إنشاء ورقة جديدة
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {
          requests: [{
            addSheet: {
              properties: {
                title: 'Applications',
              }
            }
          }]
        }
      });
    }

    // إضافة الهيدر إذا لم يكن موجود
    const headerCheck = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Applications!A1:O1',
    });

    if (!headerCheck.data.values || headerCheck.data.values.length === 0) {
      const headers = [
        'النوع',
        'الاسم الثلاثي',
        'تاريخ الميلاد',
        'رقم التليفون',
        'الايميل',
        'الجامعة',
        'الكلية',
        'تاريخ التخرج',
        'العنوان',
        'لديه خبرة في التسويق',
        'الخبرة في التسويق',
        'أهداف التدريب',
        'فكرة المنتج',
        'رابط الصورة',
        'تاريخ التقديم'
      ];

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Applications!A1:O1',
        valueInputOption: 'RAW',
        resource: {
          values: [headers],
        },
      });

      // تنسيق الهيدر
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: {
          requests: [{
            repeatCell: {
              range: {
                sheetId: sheet?.properties.sheetId || 0,
                startRowIndex: 0,
                endRowIndex: 1,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.9, green: 0.9, blue: 0.98 },
                  textFormat: { bold: true },
                }
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat)',
            }
          }]
        }
      });
    }

    console.log('✅ Google Sheet جاهز للاستخدام');
  } catch (error) {
    console.error('❌ خطأ في تهيئة Google Sheet:', error.message);
    throw error;
  }
}

async function addApplicationToGoogleSheet(applicationData) {
  try {
    console.log('🔄 بدء إضافة البيانات لـ Google Sheet...');
    
    const sheets = getGoogleSheetsClient();

    const values = [[
      applicationData.gender,
      applicationData.fullName,
      applicationData.birthDate,
      applicationData.phoneNumber,
      applicationData.email,
      applicationData.university,
      applicationData.college,
      applicationData.graduationDate,
      applicationData.address,
      applicationData.hasMarketingExperience,
      applicationData.marketingExperience,
      applicationData.trainingGoals,
      applicationData.productIdea,
      applicationData.imageUrl,
      applicationData.submissionDate
    ]];

    const response = await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Applications!A:O',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      resource: { values },
    });

    // حساب عدد الطلبات
    const allData = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Applications!A:A',
    });

    const totalApplications = (allData.data.values?.length || 1) - 1; // ناقص الهيدر

    console.log('✅ تم إضافة البيانات بنجاح لـ Google Sheet');

    return {
      rowNumber: totalApplications + 1,
      totalApplications,
      updatedRange: response.data.updates.updatedRange
    };

  } catch (error) {
    console.error('❌ خطأ في إضافة البيانات لـ Google Sheet:', error.message);
    throw error;
  }
}

function uploadToCloudinary(fileBuffer) {
  return new Promise((resolve, reject) => {
    const stream = cloudinary.uploader.upload_stream(
      { 
        folder: 'spider-agency-applications',
        resource_type: 'image'
      },
      (error, result) => {
        if (error) {
          console.error('❌ خطأ في رفع الصورة على Cloudinary:', error);
          return reject(error);
        }
        resolve(result.secure_url);
      }
    );
    streamifier.createReadStream(fileBuffer).pipe(stream);
  });
}

// الصفحة الرئيسية
app.get('/', (req, res) => {
  res.json({ 
    message: "🕷️ Spider Agency API",
    status: "running",
    storage: "Google Sheets",
    endpoints: {
      "POST /api/spider-agency/apply": "تقديم طلب جديد",
      "GET /api/spider-agency/stats": "إحصائيات الطلبات",
      "GET /api/spider-agency/sheet-url": "رابط Google Sheet",
    }
  });
});

app.post('/api/spider-agency/apply', upload.single('image'), async (req, res) => {
  try {
    console.log('🔄 استلام طلب تقديم جديد...');

    const {
      gender,
      fullName,
      birthDate,
      phoneNumber,
      email,
      university,
      college,
      graduationDate,
      address,
      hasMarketingExperience,
      marketingExperience,
      trainingGoals,
      productIdea,
    } = req.body;

    // تحقق من الحقول المطلوبة
    const requiredFields = {
      gender: 'النوع',
      fullName: 'الاسم الثلاثي',
      birthDate: 'تاريخ الميلاد',
      phoneNumber: 'رقم التليفون',
      email: 'الايميل',
      university: 'الجامعة',
      college: 'الكلية',
      graduationDate: 'تاريخ التخرج',
      address: 'العنوان',
      hasMarketingExperience: 'الخبرة في التسويق',
      trainingGoals: 'أهداف التدريب',
      productIdea: 'فكرة المنتج',
    };

    const missingFields = [];
    for (const [field, arabicName] of Object.entries(requiredFields)) {
      if (!req.body[field] || req.body[field].toString().trim() === '') {
        missingFields.push(arabicName);
      }
    }

    if (missingFields.length > 0) {
      return res.status(400).json({
        success: false,
        message: `الحقول التالية مطلوبة: ${missingFields.join(', ')}`,
      });
    }

    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'الصورة الشخصية مطلوبة',
      });
    }

    console.log('🔄 بدء رفع الصورة على Cloudinary...');
    const imageUrl = await uploadToCloudinary(req.file.buffer);
    console.log('✅ تم رفع الصورة بنجاح');

    const applicationData = {
      gender: gender.trim(),
      fullName: fullName.trim(),
      birthDate: birthDate.trim(),
      phoneNumber: phoneNumber.trim(),
      email: email.trim(),
      university: university.trim(),
      college: college.trim(),
      graduationDate: graduationDate.trim(),
      address: address.trim(),
      hasMarketingExperience: hasMarketingExperience.trim(),
      marketingExperience: marketingExperience ? marketingExperience.trim() : 'لا توجد',
      trainingGoals: trainingGoals.trim(),
      productIdea: productIdea.trim(),
      imageUrl,
      submissionDate: new Date().toISOString().split('T')[0],
    };

    const result = await addApplicationToGoogleSheet(applicationData);

    res.status(201).json({
      success: true,
      message: 'تم إرسال طلب التقديم بنجاح',
      data: {
        applicationId: Date.now(),
        submissionDate: applicationData.submissionDate,
        applicantName: fullName,
        imageUrl,
        rowNumber: result.rowNumber,
        totalApplications: result.totalApplications
      },
    });
  } catch (error) {
    console.error('❌ خطأ في معالجة طلب التقديم:', error);
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء معالجة طلبك، يرجى المحاولة مرة أخرى',
      error: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
});

app.get('/api/spider-agency/stats', async (req, res) => {
  try {
    const sheets = getGoogleSheetsClient();

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Applications!A2:J',
    });

    const rows = response.data.values || [];
    const totalApplications = rows.length;

    let maleCount = 0;
    let femaleCount = 0;
    let withExperience = 0;

    rows.forEach(row => {
      if (row[0] === 'ذكر') maleCount++;
      else if (row[0] === 'أنثى' || row[0] === 'انثى') femaleCount++;
      
      if (row[9] === 'نعم') withExperience++;
    });

    res.json({
      success: true,
      data: {
        totalApplications,
        demographics: {
          male: maleCount,
          female: femaleCount
        },
        withMarketingExperience: withExperience,
        storage: 'Google Sheets',
        sheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`
      }
    });

  } catch (error) {
    console.error('❌ خطأ في جلب الإحصائيات:', error);
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء جلب الإحصائيات',
    });
  }
});

app.get('/api/spider-agency/sheet-url', (req, res) => {
  res.json({
    success: true,
    sheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`,
    message: 'افتح الرابط لعرض جميع الطلبات'
  });
});

// Error handling
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        success: false,
        message: 'حجم الملف كبير جداً، الحد الأقصى 100MB',
      });
    }
  }

  console.error('❌ خطأ عام:', error);
  res.status(500).json({
    success: false,
    message: error.message || 'حدث خطأ غير متوقع',
  });
});

app.use('*', (req, res) => {
  res.status(404).json({
    success: false,
    message: 'الصفحة المطلوبة غير موجودة',
  });
});

// Initialize Google Sheet on startup (for local dev)
if (process.env.NODE_ENV !== 'production') {
  initializeGoogleSheet().catch(console.error);
}

// Vercel serverless handler
if (process.env.VERCEL) {
  module.exports = app;
} else {
  app.listen(PORT, () => {
    console.log(`🕷️ Spider Agency API يعمل على المنفذ ${PORT}`);
    console.log(`☁️ التخزين: Google Sheets`);
    console.log(`🌐 الصفحة الرئيسية: http://localhost:${PORT}`);
  });
}

module.exports = app;