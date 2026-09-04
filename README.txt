تحديث الإخفاء فقط

- زر "إخفاء" لا يحذف الطلب من Firebase ولا من Google Sheet.
- عند الضغط عليه يتم وضع hidden=true في Firestore.
- الموقع يستبعد الطلبات المخفية من الجدول ومن العدادات.
- الطلب يبقى محفوظاً للرجوع إليه لاحقاً.

طريقة التحديث:
1) استبدل index.html في GitHub.
2) استبدل Code.gs في Apps Script.
3) Deploy > Manage deployments > Edit > New version > Deploy.
