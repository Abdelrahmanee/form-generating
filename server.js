
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

app.delete('/api/spider-agency/clear', async (req, res) => {
  try {
    const sheets = getGoogleSheetsClient();
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Applications!A2:Z', 
    });

    res.json({
      success: true,
      message: '✅ تم مسح جميع البيانات من Google Sheet (ما عدا العناوين).',
      sheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`
    });
  } catch (error) {
    console.error('❌ خطأ أثناء مسح البيانات:', error);
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء مسح البيانات من الجدول',
      error: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
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