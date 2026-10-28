import datetime
from docx import Document

# 1. فتح ملف الوورد المطلوب بالاسم المحدد
doc = Document("أختان تحرير.docx")

# قاموس الأخطاء والتصحيحات (أضيفي هنا الكلمات الخطأ وما يقابلها من تصحيح)
corrections_dict = {
    "الى": "إلى",
    "الخطا": "الخطأ",
    # يمكنكِ إضافة المزيد من الكلمات هنا بنفس الطريقة: "الكلمة الخاطئة": "الكلمة الصحيحة",
}


# دالة تصحيح النص داخل المستند
def correct_text(text):
  # تقسيم النص إلى كلمات وفحص كل كلمة
  words = text.split()
  corrected_words = []
  for word in words:
    # إزالة أي علامات ترقيم ملتصقة بالكلمة مؤقتاً لتتم مطابقتها بدقة (اختياري)
    # والبحث في القاموس عما إذا كانت الكلمة خاطئة
    if word in corrections_dict:
      corrected_words.append(corrections_dict[word])
    else:
      corrected_words.append(word)

  return " ".join(corrected_words)


# بدلاً من ذلك، إذا أردتِ استبدال الجمل أو الكلمات مباشرة بدون تقسيم حاد:
def correct_text_direct(text):
  for wrong, correct in corrections_dict.items():
    text = text.replace(wrong, correct)
  return text


# المرور على فقرات الملف وتعديلها
for paragraph in doc.paragraphs:
  if paragraph.text:
    paragraph.text = correct_text_direct(paragraph.text)

# المرور على جداول الملف وتعديلها (إذا كان هناك جداول)
for table in doc.tables:
  for row in table.rows:
    for cell in row.cells:
      if cell.text:
        cell.text = correct_text_direct(cell.text)

# 2. إنشاء اسم ملف فريد للنتيجة لتجنب قفل الملف الأصلي
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f"corrected_text_{timestamp}.docx"

# حفظ الملف الجديد
doc.save(output_filename)
print(f"تم تصحيح وحفظ الملف بنجاح باسم: {output_filename}")