# دليل تطبيقات الويب التقدمية (PWA)

> مرجع شامل لأساسيات PWA، مكوناتها، وآليات العمل دون اتصال بالإنترنت.

---

## 1. مقدمة

**تطبيقات الويب التقدمية** (Progressive Web Apps) هي تطبيقات ويب حديثة تُقدّم تجربة مستخدم قريبة من التطبيقات الأصلية (Native Apps)، مع الحفاظ على مرونة الويب وسهولة الوصول.

### 1.1 الخصائص الأساسية

| الخاصية | الوصف |
|---------|-------|
| **Fast** | استجابة فورية وتحميل سريع |
| **Reliable** | تعمل حتى بدون اتصال بالإنترنت |
| **Engaging** | تجربة غامرة تشبه التطبيقات الأصلية |
| **Secure** | تعمل فقط عبر HTTPS |
| **Installable** | يمكن إضافتها للشاشة الرئيسية |

---

## 2. Service Worker

### 2.1 التعريف

**Service Worker** هو ملف JavaScript يعمل في **خيط خلفية** (background thread) منفصل عن صفحة الويب الرئيسية، ويعمل كوسيط (Proxy) بين التطبيق والشبكة.

### 2.2 دورة الحياة

1. **التسجيل** (Registration): إخبار المتصفح بوجود Service Worker.
2. **التثبيت** (Install): تخزين الموارد الأساسية في الكاش.
3. **التفعيل** (Activate): حذف الكاش القديم.
4. **الاعتراض** (Fetch): التحكم في الطلبات الشبكية.

### 2.3 مثال عملي

```javascript
const CACHE_NAME = 'pwa-cache-v1';

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(['/', '/index.html']))
      .then(() => self.skipWaiting())
  );
});
```

**ملاحظة مهمة:** استخدام المتغير `caches.open()` لإدارة الكاش يتطلب HTTPS.

---

## 3. حلول التخزين

| الميزة | Cache Storage | IndexedDB | LocalStorage |
|--------|---------------|-----------|--------------|
| **نوع البيانات** | Request/Response | Objects, Blobs | String |
| **السعة** | كبيرة جدًا | كبيرة جدًا | صغيرة (5-10MB) |
| **المزامنة** | Async | Async | Sync |
| **الاستخدام** | الموارد الثابتة | بيانات معقدة | إعدادات بسيطة |

### 3.1 نصيحة أمنية

> لا تحفظ كلمات المرور أو بيانات بطاقات الائتمان في `LocalStorage` — استخدم Cookies مع `HttpOnly` و `Secure`.

---

## 4. Push Notifications

### 4.1 التدفق الكامل

1. طلب الإذن من المستخدم.
2. إنشاء اشتراك فريد (PushSubscription).
3. تخزين الاشتراك في الخادم.
4. إرسال الإشعار من الخادم.
5. عرض الإشعار عبر Service Worker.

### 4.2 الأكواد الجوهرية

```javascript
self.addEventListener('push', event => {
  const data = event.data.json();
  const options = {
    body: data.body,
    icon: 'images/icon.png'
  };
  event.waitUntil(
    self.registration.showNotification(data.title, options)
  );
});
```

---

## 5. مقارنة سريعة: PWA vs Native

| المعيار | PWA | تطبيق أصلي |
|---------|-----|-----------|
| **التثبيت** | من المتصفح مباشرة | من متجر التطبيقات |
| **التحديث** | تلقائي | يدوي أو تلقائي |
| **الحجم** | صغير | كبير |
| **التكلفة** | منخفضة | مرتفعة |

---

## 6. نقاط يجب حفظها

- ✅ تعريف PWA وخصائصه الخمس.
- ✅ دورة حياة Service Worker (Registration, Install, Activate, Fetch).
- ✅ استراتيجيات التخزين المؤقت الخمس.
- ✅ خصائص Web App Manifest الأساسية.
- ✅ الفرق بين Cache Storage و IndexedDB.
- ✅ HTTPS شرط أساسي لتشغيل PWA.

---

## 7. روابط مفيدة

- [MDN Web Docs](https://developer.mozilla.org/)
- [web.dev](https://web.dev/)
- [PWA Builder](https://www.pwabuilder.com/)

---

*نهاية الملخص — بالتوفيق في الاختبار!*