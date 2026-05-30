<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام مسابقات كنيسة أبي سيفين</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root { --coptic-maroon: #580F0F; --bg-cream: #FAF8F5; }
        body { font-family: sans-serif; background: var(--bg-cream); padding: 20px; }
        .container { max-width: 800px; margin: auto; background: white; padding: 20px; border-radius: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        .page { display: none; } .active { display: block; }
        input, select, button { width: 100%; padding: 12px; margin: 10px 0; border-radius: 8px; border: 1px solid #ddd; }
        button { background: var(--coptic-maroon); color: white; border: none; cursor: pointer; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ccc; padding: 10px; text-align: center; }
        .actions { display: flex; gap: 5px; justify-content: center; }
        .btn-small { padding: 5px 8px; font-size: 12px; }
    </style>
</head>
<body>

<div class="container">
    <div id="settingsPage" class="page active">
        <h2>إعدادات النظام (رفع ملفات الإكسيل)</h2>
        <input type="file" id="questionsExcel" onchange="processFile(this, 'questions')">
        <input type="file" id="chapterExcel" onchange="processFile(this, 'chapter')">
        <button onclick="switchPage('loginPage')">دخول الخدام</button>
    </div>

    <div id="loginPage" class="page">
        <h2>تسجيل دخول</h2>
        <select id="servantSelect"><option>-- اختر اسمك --</option></select>
        <input type="password" id="pass" placeholder="الكود السري">
        <button onclick="login()">دخول</button>
        <button style="background:#333" onclick="switchPage('adminPage')">لوحة التحكم</button>
    </div>

    <div id="adminPage" class="page">
        <h2>النتائج</h2>
        <table id="adminTable">
            <thead><tr><th>الاسم</th><th>الدرجة</th><th>إجراءات</th></tr></thead>
            <tbody id="adminBody"></tbody>
        </table>
        <button onclick="switchPage('settingsPage')">الرجوع للإعدادات</button>
    </div>
</div>

<script>
    // 1. البيانات الثابتة للخدام
    const SERVANTS = {
        "ابراهيم عادل": "1", "ابو الخير": "2", "ايريني سمير": "3", "ايريني عاطف": "4", "امال سامي": "5",
        "توني هاني": "6", "ثناء زكريا": "7", "جاكلين سعيد": "8", "جانيت ويصا": "9", "جمال حنا": "10",
        "جورج فايز": "11", "جيرمين زاهر": "12", "جيهان شحاتة": "13", "جيهان ماهر": "14", "رأفت خير": "15",
        "رامز امير": "16", "رامي سعيد": "17", "رؤوف خير": "18", "رمزي سعيد": "19", "سارة عماد": "20",
        "سامية جريس": "21", "ساندرا جورج": "22", "سماح حلمي": "23", "سماح نعيم": "24", "سمر جادالرب": "25",
        "سمير زكي": "26", "سناء سعيد": "27", "شريف شهدي": "28", "شكرالله نصر": "29", "شيري شهدي": "30",
        "صفاء سمير": "31", "صفاء عادلي": "32", "عادل فرج": "33", "عايدة معوض": "34", "عايده يعقوب": "35",
        "عدلي رمزي": "36", "عماد سمير": "37", "فادي سليمان": "38", "فادي كحيل": "39", "فيوليت فوزي": "40",
        "كمال كريم": "41", "كريستين نصر": "42", "كريم شكرالله": "43", "كيرلس اسامه": "44", "كيرلس فادي": "45",
        "كيرلس يوسف": "46", "ليلي وهيب": "47", "ماجدة نجيب": "48", "ماركو مجدي": "49", "ماري عادل": "50",
        "ماري وجيه": "51", "ماريان مجدي": "52", "ماريانا ادورد": "53", "مارينا اشرف": "54", "مريان عادل": "55",
        "مريم عماد": "56", "مريم مكرم": "57", "منى ميشيل": "58", "ميرفت يعقوب": "59", "ميرنا عماد": "60",
        "مينا اشرف": "61", "مينا مجدي": "62", "ناهد فاروق": "63", "نرجس خير": "64", "نرمين سليمان": "65",
        "نعمة عبدالسيد": "66", "هيلانة جورج": "67", "وائل ماهر": "68", "ورده ماهر": "69", "ياسر نبيه": "70"
    };

    // 2. تعبئة القائمة تلقائياً
    const select = document.getElementById('servantSelect');
    Object.keys(SERVANTS).sort().forEach(name => {
        let opt = document.createElement('option');
        opt.value = name; opt.innerHTML = name;
        select.appendChild(opt);
    });

    // 3. وظائف التنقل والإدارة
    function switchPage(id) {
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.getElementById(id).classList.add('active');
        if(id === 'adminPage') updateAdminTable();
    }

    function processFile(input, type) {
        const file = input.files[0];
        const reader = new FileReader();
        reader.onload = e => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            localStorage.setItem("data_" + type, JSON.stringify(json));
            alert("تم حفظ ملف " + type);
        };
        reader.readAsArrayBuffer(file);
    }

    function updateAdminTable() {
        let scores = JSON.parse(localStorage.getItem("scores")) || {};
        let body = document.getElementById('adminBody');
        body.innerHTML = "";
        for(let name in scores) {
            body.innerHTML += `<tr><td>${name}</td><td>${scores[name]}</td>
            <td class="actions">
                <button class="btn-small" onclick="editScore('${name}')">تعديل</button>
                <button class="btn-small" style="background:#b71c1c" onclick="deleteRecord('${name}')">مسح</button>
            </td></tr>`;
        }
    }

    function deleteRecord(name) {
        let scores = JSON.parse(localStorage.getItem("scores"));
        delete scores[name];
        localStorage.setItem("scores", JSON.stringify(scores));
        updateAdminTable();
    }
</script>
</body>
</html>
