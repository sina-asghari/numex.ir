حتماً ✅
این curl برای گرفتن اطلاعات (getinfo) با API Key شما:

⸻

🔹 درخواست ساده (بدون فیلتر)

curl -X GET "https://api.numberland.ir/v2.php/?apikey=7143e4c5a8173ca572232dcc15773cbc&method=getinfo"


⸻

🔹 با فیلتر اپراتور

curl -X GET "https://api.numberland.ir/v2.php/?apikey=7143e4c5a8173ca572232dcc15773cbc&method=getinfo&operator=OPERATOR_CODE"


⸻

🔹 با فیلتر کشور

curl -X GET "https://api.numberland.ir/v2.php/?apikey=7143e4c5a8173ca572232dcc15773cbc&method=getinfo&country=COUNTRY_CODE"


⸻

🔹 با فیلتر سرویس

curl -X GET "https://api.numberland.ir/v2.php/?apikey=7143e4c5a8173ca572232dcc15773cbc&method=getinfo&service=SERVICE_CODE"


⸻

🔹 مثال کامل با چند فیلتر همزمان

curl -X GET "https://api.numberland.ir/v2.php/?apikey=7143e4c5a8173ca572232dcc15773cbc&method=getinfo&operator=1&country=8&service=1"


⸻

اگر بخواهی، می‌توانم همین را به صورت کد Node.js (TypeScript + Axios) هم برای بک‌اندت آماده کنم 👌