"""
══════════════════════════════════════════════════════════════════════
سيرفر بايثون المحلي لحفظ طلبات استشارة "وعي" في ملف إكسيل (.xlsx)
══════════════════════════════════════════════════════════════════════
طريقة التشغيل:
1. افتح التيرمينال واكتب:
   python backend.py
2. السيرفر هيشتغل تلقائياً على الرابط: http://localhost:5000
3. كل ما عميل يسجل في الموقع، هيتم إنشاء/تحديث ملف:
   "حجوزات_وعي_العملاء.xlsx" في نفس الفولدر تلقائياً!
══════════════════════════════════════════════════════════════════════
"""

import os
import json
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_FILE = os.path.join(os.path.dirname(__file__), "حجوزات_وعي_العملاء.xlsx")
PORT = 5000

def initialize_excel():
    """إنشاء وتنسيق ملف الإكسيل إذا لم يكن موجوداً من قبل"""
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "طلبات الاستشارة والجلسات"
        ws.views.sheetView[0].rightToLeft = True  # اتجاه من اليمين لليسار

        headers = [
            "تاريخ ووقت الطلب",
            "اسم العميل",
            "رقم الهاتف / واتساب",
            "اسم الشركة أو المنشأة",
            "مجال عمل الشركة",
            "عدد الموظفين المتوقع",
            "حالة الطلب وملاحظات"
        ]

        ws.append(headers)

        # تنسيق صف العناوين باللون الأخضر المميز لـ "وعي"
        header_fill = PatternFill(start_color="0A6E63", end_color="0A6E63", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # تعديل العرض الافتراضي للأعمدة
        col_widths = [22, 22, 20, 25, 24, 22, 20]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        ws.row_dimensions[1].height = 30
        wb.save(EXCEL_FILE)
        print(f"[*] تم إنشاء ملف الإكسيل بنجاح: {os.path.basename(EXCEL_FILE)}")

def append_to_excel(lead_data):
    """إضافة سجل عميل جديد في ملف الإكسيل"""
    initialize_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    now_str = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    row = [
        lead_data.get("timestamp", now_str),
        lead_data.get("name", "غير محدد"),
        lead_data.get("phone", "غير محدد"),
        lead_data.get("company", "غير محدد"),
        lead_data.get("sector", "غير محدد"),
        lead_data.get("count", "غير محدد"),
        lead_data.get("status", "طلب جديد 🔴")
    ]

    ws.append(row)
    new_row = ws.max_row
    ws.row_dimensions[new_row].height = 24

    # تنسيق خلايا السطر الجديد
    row_font = Font(name="Arial", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    for col_idx in range(1, len(row) + 1):
        cell = ws.cell(row=new_row, column=col_idx)
        cell.font = row_font
        cell.alignment = center_align
        cell.border = thin_border
        # تمييز السطور الزوجية بلون خلفية خفيف لسهولة القراءة
        if new_row % 2 == 0:
            cell.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    wb.save(EXCEL_FILE)
    print(f"\n[+] تم تسجيل عميل جديد بنجاح:")
    print(f"    - الاسم: {row[1]}")
    print(f"    - الشركة: {row[3]}")
    print(f"    - الهاتف: {row[2]}")
    print(f"    - الوقت: {row[0]}")
    print(f"    --> تم الحفظ في سطر رقم {new_row} بملف {os.path.basename(EXCEL_FILE)}")


class LeadRequestHandler(BaseHTTPRequestHandler):
    """معالج طلبات الويب والـ API"""

    def _set_cors_headers(self, status=200):
        self.send_response(status)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()

    def do_OPTIONS(self):
        """معالجة فحص الـ CORS المسبق (Pre-flight request)"""
        self._set_cors_headers(200)

    def do_GET(self):
        """صفحة فحص حالة السيرفر وتنزيل الملف"""
        self._set_cors_headers(200)
        status_info = {
            "service": "سيرفر استقبال طلبات استشارة وعي",
            "status": "online",
            "excel_file": os.path.basename(EXCEL_FILE),
            "file_exists": os.path.exists(EXCEL_FILE),
            "endpoints": {
                "POST /api/leads": "إرسال بيانات حجز عميل جديد"
            }
        }
        self.wfile.write(json.dumps(status_info, ensure_ascii=False, indent=2).encode('utf-8'))

    def do_POST(self):
        """استقبال بيانات حجز العميل"""
        if self.path == "/api/leads" or self.path == "/":
            content_length = int(self.headers.get("Content-Length", 0))
            post_body = self.rfile.read(content_length).decode("utf-8")

            try:
                data = json.loads(post_body)
                append_to_excel(data)

                self._set_cors_headers(200)
                response = {
                    "status": "success",
                    "message": "تم استلام الطلب وتسجيله في شيت الإكسيل بنجاح!",
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                self.wfile.write(json.dumps(response, ensure_ascii=False).encode("utf-8"))

            except Exception as e:
                print(f"[!] خطأ أثناء معالجة الطلب: {e}")
                self._set_cors_headers(400)
                error_response = {
                    "status": "error",
                    "message": f"حدث خطأ أثناء حفظ البيانات: {str(e)}"
                }
                self.wfile.write(json.dumps(error_response, ensure_ascii=False).encode("utf-8"))
        else:
            self._set_cors_headers(404)
            self.wfile.write(json.dumps({"error": "مسار غير موجود"}, ensure_ascii=False).encode("utf-8"))

def run():
    initialize_excel()
    server_address = ("", PORT)
    httpd = HTTPServer(server_address, LeadRequestHandler)
    print("=" * 65)
    print(f"🚀 سيرفر 'وعي' لحفظ الحجوزات في الإكسيل يعمل الآن على:")
    print(f"   --> http://localhost:{PORT}")
    print(f"   --> مسار الـ API: http://localhost:{PORT}/api/leads")
    print(f"   --> ملف الإكسيل: {os.path.abspath(EXCEL_FILE)}")
    print("=" * 65)
    print("بانتظار استقبال طلبات العملاء من الموقع... (اضغط Ctrl+C للإيقاف)")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nتم إيقاف السيرفر.")

if __name__ == "__main__":
    run()

