/* Shalfa Ops bilingual interface (Arabic / English).
 * Display-only localization: never changes form values, IDs, statuses sent to APIs,
 * employee/school names, server responses or payroll calculations.
 */
(function () {
  'use strict';
  var KEY = 'shalfa-ui-language';
  var translations = {
    'شركة شلفا لإدارة المرافق':'Shalfa Facilities Management',
    'نظام حضور البوابين المركزي':'Central Gate Attendants Attendance System',
    'دخول شلفا':'Shalfa Login',
    'اسم المستخدم':'Username', 'كلمة المرور':'Password', 'دخول':'Sign in',
    'الإدارة: admin / 1234':'Administration: admin / 1234',
    'المنسق: الرقم الوظيفي / 1234':'Coordinator: Employee ID / 1234',
    'الرئيسية':'Home', 'الخريطة التفاعلية':'Interactive map',
    'الحضور اليومي':'Daily attendance','الحضور والانصراف':'Attendance & checkout',
    'تقارير الحضور':'Attendance reports','غرفة العمليات':'Operations center',
    'طلبات تصاريح التشغيل والصيانة':'Operations & maintenance permits',
    'البلاغات والحوادث':'Incidents & reports','المخالفات والإنذارات':'Violations & warnings',
    'الجولات الأمنية':'Security patrols','العهد والمفاتيح':'Assets & keys',
    'الورديات والعجز':'Shifts & shortages','تقييم الحراس':'Guard evaluations',
    'انتهاء المستندات':'Document expiry','المدارس والبوابون':'Schools & attendants',
    'ربط البوابين':'Assign attendants','إدارة المنسقين / المشرفين':'Coordinators / supervisors',
    'مركز جودة البيانات':'Data quality center','إدارة الإحداثيات':'Coordinate management',
    'تحديث بيانات المدارس':'Update school data','مسير الرواتب':'Payroll',
    'التقارير':'Reports','مركز الربط والمزامنة':'Linking & sync center',
    'لوحة العميل':'Client dashboard','إدارة المستخدمين':'User management',
    'إعدادات النظام':'System settings','سجل التدقيق':'Audit trail',
    'المهام':'Tasks','تطبيق المنسق':'Coordinator app','اسأل شلفا':'Ask Shalfa',
    'خروج':'Sign out','الصيانة والنظافة':'Maintenance & cleaning',
    'موظفو الصيانة والنظافة':'Maintenance & cleaning staff',
    'حصر الأصول':'Asset inventory','تسعير الأصول':'Asset pricing',
    'أوامر الصيانة':'Maintenance work orders','النظافة والجودة':'Cleaning quality',
    'الصيانة الوقائية':'Preventive maintenance',
    'لوحة التحكم الرئيسية':'Main dashboard',
    'حالة النظام':'System status','متصل ومستقر':'Connected and stable',
    'إعداد وإشراف عام':'General preparation & supervision',
    'المشروع':'Project','التاريخ':'Date','إصدار البرنامج':'Software version',
    'الهيكل الإداري':'Management structure',
    'مدير خدمات أعمال البوابين':'Gate services manager',
    'مدير قطاع الخدمات الأمنية':'Security sector manager',
    'المشرف العام':'General supervisor',
    'نسبة التغطية اليوم':'Today’s coverage',
    'الجولات اليوم':'Today’s patrols','المخالفات اليوم':'Today’s violations',
    'فرق داخل المدارس':'Teams inside schools','إجمالي البوابين':'Total attendants',
    'إجمالي المدارس':'Total schools','الخدمات الرئيسية':'Main services',
    'المؤشرات المباشرة':'Live metrics','المدارس والمواقع':'Schools & locations',
    'التسجيل اليومي':'Daily recording','GPS ونقاط التفتيش':'GPS & checkpoints',
    'دخول وخروج موثق':'Recorded entry and exit',
    'المتابعة والإغلاق':'Follow-up & closure',
    'إجراءات الحراس':'Guard actions',
    'التسليم والاستلام':'Handover & receipt',
    'المطلوب والمسكن':'Required & assigned',
    'الانضباط والأداء':'Discipline & performance',
    'الانتهاء والتنبيهات':'Expiry & alerts',
    'البيانات والبحث':'Data & search','التوزيع والأداء':'Distribution & performance',
    'تحديث Excel':'Excel update','تحديث البيانات الذكي':'Smart data update',
    'Excel والتقارير التشغيلية':'Excel & operational reports',
    'الحسابات والصلاحيات':'Accounts & permissions',
    'الهيكل والبرنامج':'Structure & system',
    'إنهاء الجلسة':'End session',
    'مدارس لم تسجل اليوم':'Schools not recorded today',
    'عرض التفاصيل':'View details',
    'بلاغات مفتوحة':'Open incidents','متابعة البلاغات':'Follow incidents',
    'تصاريح بانتظار الموافقة':'Permits pending approval',
    'فتح التصاريح':'Open permits',
    'جميع الحقوق محفوظة © 2026':'All rights reserved © 2026',
    'متابعة الحضور اليومية':'Daily attendance monitoring',
    'تاريخ المتابعة':'Monitoring date',
    'المدينة':'City','كل المدن':'All cities',
    'المنسق':'Coordinator','كل المنسقين':'All coordinators',
    'نوع المدارس':'School type','كل المدارس':'All schools',
    'الرئيسية فقط':'Main schools only',
    'الرئيسية والفرعية':'Main & branch schools',
    'الفرعية فقط':'Branch schools only','الملحقة فقط':'Annex schools only',
    'حالة الحضور':'Attendance status','كل الحالات':'All statuses',
    'تحديث المتابعة':'Refresh monitoring','تحديث':'Refresh',
    'تصدير Excel':'Export Excel','تصدير':'Export',
    'حاضر':'Present','غائب':'Absent','إجازة':'Leave',
    'بديل':'Replacement','لم يسجل':'Not recorded',
    'نسبة التغطية':'Coverage rate','نسبة التسجيل':'Recording rate',
    'تحتاج متابعة':'Needs follow-up',
    'استيراد ملف حضور Excel':'Import attendance Excel',
    'استيراد حضور Excel':'Import attendance Excel',
    'ملخص أداء المنسقين والمشرفين':'Coordinators & supervisors performance',
    'الإجمالي':'Total','آخر تسجيل:':'Last recorded:',
    'تفاصيل المدارس':'School details',
    'المدرسة':'School','البواب':'Attendant',
    'الرقم الوظيفي':'Employee ID','الجوال':'Mobile',
    'الحالة':'Status','وقت التسجيل':'Recording time',
    'إجراء سريع':'Quick action',
    'لا توجد نتائج مطابقة.':'No matching results.',
    'غير متوفر':'Unavailable','غير مرتبط':'Not assigned',
    'الإدارة':'Administration','المنسقون':'Coordinators',
    'المنسقين':'Coordinators','المشرفين':'Supervisors',
    'البوابون':'Attendants','المنسقون والمشرفون':'Coordinators & supervisors',
    'المنسقون / المشرفون':'Coordinators / supervisors',
    'اسم المدرسة':'School name','الرقم الوزاري':'Ministry ID',
    'اسم البواب':'Attendant name','اسم المنسق':'Coordinator name',
    'اسم المشرف':'Supervisor name','رقم الجوال':'Mobile number',
    'رقم الهوية':'National ID','رقم البطاقة':'ID number',
    'الرقم التسلسلي':'Serial number','الحي':'Neighborhood',
    'المنطقة':'Region','مكتب التعليم':'Education office',
    'المرحلة':'School stage','الجنس':'Gender',
    'بنين':'Boys','بنات':'Girls',
    'اسم الموظف':'Employee name','اسم المستخدم / الرقم الوظيفي':'Username / Employee ID',
    'الملاحظات':'Notes','ملاحظات':'Notes',
    'البحث':'Search','بحث':'Search','بحث سريع':'Quick search',
    'ابحث':'Search','إضافة':'Add','إضافة جديد':'Add new',
    'تعديل':'Edit','حذف':'Delete','حفظ':'Save',
    'حفظ التغييرات':'Save changes','إلغاء':'Cancel',
    'رجوع':'Back','العودة':'Back','التالي':'Next',
    'السابق':'Previous','عرض':'View','التفاصيل':'Details',
    'المزيد':'More','إغلاق':'Close','تأكيد':'Confirm',
    'اختيار':'Select','الكل':'All',
    'إرسال':'Send','اعتماد':'Approve','رفض':'Reject',
    'موافقة':'Approval','طباعة':'Print','تحميل':'Download',
    'استيراد':'Import','رفع ملف':'Upload file',
    'ملف Excel':'Excel file','اختيار ملف':'Choose file',
    'إعدادات البرنامج والهيكل الإداري':'System & management settings',
    'تعديل البيانات التي تظهر في الصفحة الرئيسية.':'Edit the details shown on the home page.',
    'بيانات البرنامج':'System information',
    'اسم الشركة':'Company name','اسم النظام':'System name',
    'اسم المشروع':'Project name','الإصدار':'Version',
    'حفظ الإعدادات':'Save settings',
    'التقارير التشغيلية':'Operational reports',
    'من تاريخ':'From date','إلى تاريخ':'To date',
    'الفترة':'Period','اليوم':'Today','هذا الشهر':'This month',
    'الملخص':'Summary','لوحة المعلومات':'Dashboard',
    'نوع المنشأة':'Facility type',
    'المستندات':'Documents',
    'العهد والمعدات':'Assets & equipment',
    'الدعم الفني':'Technical support',
    'لا توجد بيانات':'No data',
    'جار التحميل...':'Loading...',
    'جار التحميل':'Loading',
    'تم الحفظ بنجاح':'Saved successfully',
    'الاسم':'Name','الرقم':'Number','العدد':'Count',
    'الفصل':'Classroom','الفصول':'Classrooms',
    'العدد الإجمالي':'Grand total',
    'المستخدمون':'Users','صلاحيات المستخدم':'User permissions',
    'تسجيل الدخول':'Sign in','تسجيل الخروج':'Sign out',
    'إعداد: محمد هندي':'Prepared by: Mohammed Hendy'
  };
  var originalNodes = new WeakMap();
  var lastOutput = new WeakMap();
  var originalAttrs = new WeakMap();
  var language = 'ar', observer = null, scheduled = false;
  function storedLanguage() {
    try { return localStorage.getItem(KEY) === 'en' ? 'en' : 'ar'; }
    catch (_) { return 'ar'; }
  }
  function translate(source) {
    var m = /^(\s*)([\s\S]*?)(\s*)$/.exec(source || '');
    if (!m) return source;
    var phrase = m[2].replace(/\s+/g, ' ').trim();
    return m[1] + (translations[phrase] || m[2]) + m[3];
  }
  function skipElement(el) {
    if (!el || !el.closest) return true;
    return !!el.closest(
      'script,style,noscript,pre,code,textarea,svg,canvas,[data-no-translate],'+
      'table tbody td .school-name,table tbody td .guard-name'
    );
  }
  function translateText(node) {
    if (!node.parentElement || skipElement(node.parentElement)) return;
    var current = node.nodeValue;
    var previous = lastOutput.get(node);
    var original = originalNodes.get(node);
    if (original === undefined || (previous !== undefined && current !== previous)) {
      original = current;
      originalNodes.set(node, original);
    }
    var next = language === 'en' ? translate(original) : original;
    lastOutput.set(node, next);
    if (current !== next) node.nodeValue = next;
  }
  function translateElement(el) {
    if (skipElement(el)) return;
    var attributeNames = ['placeholder','title','aria-label'];
    var record = originalAttrs.get(el) || {};
    attributeNames.forEach(function(name) {
      if (!el.hasAttribute(name)) return;
      var current = el.getAttribute(name);
      var previous = record[name] && record[name].last;
      var original = record[name] && record[name].original;
      if (original === undefined || (previous !== undefined && current !== previous)) original = current;
      var next = language === 'en' ? translate(original) : original;
      record[name] = {original:original,last:next};
      if (current !== next) el.setAttribute(name, next);
    });
    originalAttrs.set(el, record);
  }
  function applyTranslations() {
    scheduled = false;
    document.documentElement.lang = language;
    document.documentElement.dir = language === 'en' ? 'ltr' : 'rtl';
    var all = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    var node;
    while ((node = all.nextNode())) translateText(node);
    document.querySelectorAll('[placeholder],[title],[aria-label]').forEach(translateElement);
    if (document.title) {
      document.documentElement.dataset.shalfaLanguage = language;
    }
    var select = document.getElementById('shalfa-language');
    if (select && select.value !== language) select.value = language;
  }
  function schedule() {
    if (scheduled) return;
    scheduled = true;
    Promise.resolve().then(applyTranslations);
  }
  function setLanguage(next) {
    language = next === 'en' ? 'en' : 'ar';
    try { localStorage.setItem(KEY, language); } catch (_) {}
    schedule();
  }
  function addStyles() {
    var css = document.createElement('style');
    css.id = 'shalfa-language-style';
    css.textContent =
      '.shalfa-language-control{display:flex;align-items:center;gap:7px;z-index:30;'+
      'padding:7px 9px;border:1px solid #cad7e7;border-radius:10px;'+
      'background:#fff;color:#082b60;font:600 13px Tahoma,Arial,sans-serif;direction:ltr}'+
      '.shalfa-language-control select{min-width:105px;border:1px solid #ccd6e1;'+
      'border-radius:7px;padding:6px;background:#fff;color:#082b60;'+
      'font:600 13px Tahoma,Arial,sans-serif;cursor:pointer}'+
      '.login-card .shalfa-language-control{justify-content:center;margin-bottom:16px}'+
      '.app-brand .shalfa-language-control{margin:8px auto;max-width:200px}'+
      '.shalfa-language-floating{position:fixed;top:12px;inset-inline-end:12px;'+
      'box-shadow:0 3px 16px #081c3533;z-index:1000}'+
      'html[dir=ltr] .app-sidebar{text-align:left}'+
      'html[dir=ltr] .app-nav a{text-align:left}'+
      '@media(max-width:900px){.shalfa-language-floating{top:60px}}';
    document.head.appendChild(css);
  }
  function addControl() {
    if (document.getElementById('shalfa-language')) return;
    var control = document.createElement('div');
    control.className = 'shalfa-language-control';
    control.setAttribute('data-no-translate','');
    var icon = document.createElement('span');
    icon.textContent = '🌐';
    var select = document.createElement('select');
    select.id = 'shalfa-language';
    select.setAttribute('aria-label','Language / اللغة');
    [['ar','العربية'],['en','English']].forEach(function(entry) {
      var option = document.createElement('option');
      option.value = entry[0];
      option.textContent = entry[1];
      select.appendChild(option);
    });
    select.value = language;
    select.addEventListener('change', function () { setLanguage(select.value); });
    control.appendChild(icon);
    control.appendChild(select);
    var login = document.querySelector('.login-card');
    var brand = document.querySelector('.app-sidebar .app-brand');
    var topbar = document.querySelector('.approved-topbar .topbar-user');
    if (login) login.insertBefore(control, login.firstChild);
    else if (topbar) topbar.insertBefore(control, topbar.firstChild);
    else if (brand) brand.appendChild(control);
    else {
      control.classList.add('shalfa-language-floating');
      document.body.appendChild(control);
    }
  }
  function start() {
    language = storedLanguage();
    addStyles();
    addControl();
    applyTranslations();
    observer = new MutationObserver(function(mutations) {
      // Dynamically inserted alerts, tables, cards and dialogs are localized.
      if (mutations.some(function(m) {
        return m.type === 'characterData' || m.type === 'childList';
      })) schedule();
    });
    observer.observe(document.body,{subtree:true,childList:true,characterData:true});
    window.ShalfaLanguage = {get:function(){return language;},set:setLanguage};
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded',start);
  else start();
})();