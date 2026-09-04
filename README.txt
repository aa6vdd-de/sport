الخطوة الثانية - Firebase Firestore

1) افتح Firebase Console > Databases & Storage > Firestore Database.
2) اختر Create database.
3) اختر موقع قاعدة البيانات المناسب.
4) أثناء التطوير فقط يمكنك اختيار Test mode.
5) بعد إنشاء Firestore افتح Rules وضع محتوى firestore.rules ثم Publish.
6) ارفع index.html الجديد إلى GitHub Pages بدل النسخة القديمة.

المجموعة المستخدمة في Firestore:
bookings

الحقول التي سيقرأها الموقع:
name
email
phone
activityName
participantsCount
category
venue
bookingMethod
date
time
status
createdAt

مهم:
القواعد الموجودة في firestore.rules مفتوحة للتجربة فقط وليست آمنة للنشر العام.
بعد أن ننتهي من الربط، الأفضل إضافة Firebase Authentication وقفل الكتابة على حساب الإدارة فقط.

الخطوة الثالثة لاحقاً:
Google Apps Script سيأخذ كل رد جديد من Google Form/Google Sheet ويضيفه تلقائياً إلى collection اسمها bookings.
كما سنربط قرار القبول/الرفض بإرسال البريد.
