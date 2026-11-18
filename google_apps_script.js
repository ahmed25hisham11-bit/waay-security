/**
 * ══════════════════════════════════════════════════════════════════════
 * كود ربط فورم "وعي" مع شيت جوجل إكسيل (Google Sheets Webhook)
 * ══════════════════════════════════════════════════════════════════════
 * 
 * خطوات التفعيل في 60 ثانية:
 * 1. افتح Google Sheets وأنشئ شيت جديد باسم: "طلبات استشارة وعي"
 * 2. من القائمة العلوية اضغط: Extensions (الإضافات) -> Apps Script
 * 3. امسح أي كود موجود هناك، والصق هذا الكود بالكامل.
 * 4. اضغط زر Deploy (نشر) بالأعلى -> واختر New deployment (نشر جديد).
 * 5. اضغط على أيقونة الترس بجانب "Select type" واختر Web app (تطبيق ويب).
 * 6. اضبط الإعدادات كالتالي:
 *    - Description: وعي - استلام الطلبات
 *    - Execute as: Me (حسابك)
 *    - Who has access: Anyone (أي شخص)  <-- خطوة هامة جداً!
 * 7. اضغط Deploy ووافق على الصلاحيات، ثم انسخ رابط (Web App URL).
 * 8. ضع الرابط في ملف الموقع في خانة: ENDPOINT_URL
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // إذا كان الشيت جديد وفارغ، نقوم بإنشاء الهيدر وتنسيقه تلقائياً
    if (sheet.getLastRow() === 0) {
      var headers = [
        "تاريخ ووقت الطلب",
        "اسم العميل",
        "رقم الهاتف / واتساب",
        "اسم الشركة أو المنشأة",
        "مجال عمل الشركة",
        "عدد الموظفين المتوقع",
        "حالة الطلب وملاحظات"
      ];
      sheet.appendRow(headers);
      
      // تنسيق شريط العناوين بلون هوية "وعي" الأخضر الفاخر
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#0A6E63");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      headerRange.setFontSize(11);
      
      sheet.setFrozenRows(1);
      sheet.setRightToLeft(true); // ضبط اتجاه الشيت من اليمين لليسار
    }
    
    // قراءة البيانات المرسلة من الموقع
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }
    
    var now = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd HH:mm:ss");
    var timestamp = data.timestamp || now;
    var name = data.name || "غير محدد";
    var phone = data.phone || "غير محدد";
    var company = data.company || "غير محدد";
    var sector = data.sector || "غير محدد";
    var count = data.count || "غير محدد";
    var status = data.status || "طلب جديد 🔴";
    
    // إضافة السطر الجديد في الإكسيل
    sheet.appendRow([
      timestamp,
      name,
      phone,
      company,
      sector,
      count,
      status
    ]);
    
    // محاذاة النص وتعديل عرض الأعمدة تلقائياً
    var newRowIndex = sheet.getLastRow();
    var newRowRange = sheet.getRange(newRowIndex, 1, 1, 7);
    newRowRange.setHorizontalAlignment("center");
    newRowRange.setFontFamily("Tajawal");
    
    // إرجاع رد نجاح
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "تم حفظ السجل في الإكسيل بنجاح",
      row: newRowIndex
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// اختبار تجريبي عند فتح الرابط في المتصفح
function doGet(e) {
  return ContentService.createTextOutput("✅ سيرفر استقبال طلبات وعي يعمل بنجاح! جاهز لاستقبال بيانات العملاء عبر POST.").setMimeType(ContentService.MimeType.TEXT);
}

