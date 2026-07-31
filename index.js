const puppeteer = require('puppeteer');
const fs = require('fs');
const ExcelJS = require('exceljs');

// ==================== 1. رابط الموقع الرئيسي ====================
const HOME_URL = 'https://eg1xbet.com/ar';

// ==================== 2. قائمة الحسابات والأسماء ====================
const ACCOUNTS = [
  { username: "AbaAbasel3eldao800@yahoo.com", password: "qw12345678", name: "faiyum" },
  { username: "Wa2ilAdel012@gmail.com", password: "qw12345678", name: "Adel" },
  { username: "Rady5Reda22s@yahoo.com", password: "qw12345678", name: "Reda" },
  { username: "Ra2oofA7med77A@yahoo.com", password: "qw12345678", name: "A7med" },
  { username: "MadyTamer808@gmail.com", password: "qw12345678", name: "Tamer" },
  { username: "A7medWagdy707@yahoo.com", password: "qw12345678", name: "Wagdy" },
  { username: "Ah7medSamr144@yahoo.com", password: "qw12345678", name: "Samr" },
  { username: "A7meSa3eed515@gmail.com", password: "qw12345678", name: "Sa3eed" },
  { username: "EhabSameer337@gmail.com", password: "qw12345678", name: "Sameer" },
  { username: "Sa3edEhabj404@gmail.com", password: "qw12345678", name: "Ehabj" },
  { username: "QamrrKanell127@yahoo.com", password: "qw12345678", name: "Kanell" },
  { username: "A7medMoneer325@gmail.com", password: "qw12345678", name: "Moneer" },
  { username: "RadwaMarouo22002@gmail.com", password: "qw12345678", name: "Marouo" },
  { username: "KamalZidanZ123@yahoo.com", password: "qw12345678", name: "ZidanZ" },
  { username: "A7medSabrriiyT257@yahoo.com", password: "qw12345678", name: "SabrriiyT" },
  { username: "FayzaaWaleedWF555@gmail.com", password: "qw12345678", name: "WaleedWF" },
  { username: "Ra7ma7Rami17@yahoo.com", password: "qw12345678", name: "Rami" },
  { username: "Ehab7anaa667@yahoo.com", password: "qw12345678", name: "7anaa" },
  { username: "Mo7senbihMomtaz101@yahoo.com", password: "qw12345678", name: "Momtaz" },
  { username: "FaouziaFaroukk90@yahoo.com", password: "qw12345678", name: "Faroukk" },
  { username: "GehanElsadat832@yahoo.com", password: "qw12345678", name: "Elsadat" },
  { username: "MonaZaky237@yahoo.com", password: "qw12345678", name: "Zaky" },
  { username: "Hanaa2magdyy66@gmail.com", password: "qw12345678", name: "2magdyy" },
  { username: "Khloodkhaleed555@yahoo.com", password: "qw12345678", name: "khaleed" },
  { username: "SamerrHadyelsa3ed34@yahoo.com", password: "qw12345678", name: "Hadyelsa3ed" },
  { username: "TamerTanra8o8@yahoo.com", password: "qw12345678", name: "Tanra" },
  { username: "DaniloEmad57@yahoo.com", password: "qw12345678", name: "Emad" },
  { username: "Galal4okry404@yahoo.com", password: "qw12345678", name: "4okry" },
  { username: "AshrafMaDkour666@gmail.com", password: "qw12345678", name: "MaDkour" },
  { username: "Sa3eed4Shady444@gmail.com", password: "qw12345678", name: "Shady" },
  { username: "Amr3BarraKat8@gmail.com", password: "qw12345678", name: "BarraKat" },
  { username: "FaredSadeteik1987@yahoo.com", password: "qw12345678", name: "Sadeteik" },
  { username: "SamyaGabeer777@yahoo.com", password: "qw12345678", name: "Gabeer" },
  { username: "GaberEbRaheem619@gmail.com", password: "qw12345678", name: "EbRaheem" },
  { username: "RehamFadl909@yahoo.com", password: "qw12345678", name: "Fadl" },
  { username: "Moh7amedMo7ey12@yahoo.com", password: "qw12345678", name: "Mo7ey" },
  { username: "Bahaa2SoulTann1@yahho.com", password: "qw12345678", name: "SoulTann" },
  { username: "TamerHosny011@yahoo.com", password: "qw12345678", name: "Hosny" },
  { username: "Sa3dZa8loul1919@yahoo.com", password: "qw12345678", name: "Za8loul" },
  { username: "MoustfaKamell11@yahoo.com", password: "qw12345678", name: "Kamell" },
  { username: "Hassanshehatta4@yahoo.com", password: "qw12345678", name: "shehatta" },
  { username: "HaderMaGDiiy5@yahoo.com", password: "qw12345678", name: "MaGDiiy" },
  { username: "GamalGendy643@yahoo.com", password: "qw12345678", name: "Gendy" },
  { username: "SamaARGaMall33@yahoo.com", password: "qw12345678", name: "ARGaMall" },
  { username: "A7medHelmy771@yahoo.com", password: "qw12345678", name: "Helmy" },
  { username: "FahmyRef3aat505@yahoo.com", password: "qw12345678", name: "Ref3aat" },
  { username: "KareemSzhalaby170@yahoo.com", password: "qw12345678", name: "Szhalaby" },

  // الحسابات المضافة
  { username: "samehmohamed1000@yahoo.com", password: "Sameh112000", name: "alex" },
  { username: "samehmohamed2000@yahoo.com", password: "Sameh112000", name: "borg el arab" },
  { username: "samehmohamed3000@yahoo.com", password: "Sameh112000", name: "al mandarah" },
  { username: "samehmohamed40000@yahoo.com", password: "Sameh112000", name: "marsa matruh" }
];

const SEEN_PHONE_NUMBERS = new Set();

// ==================== 3. محددات العناصر ====================
const OPEN_LOGIN_BTN_SELECTOR = '#app > div.layout-content.layout-content--theme-primary--40 > div.layout-content__header > header > div.header-top.header__top > div.user-control-dashboard--padding-right.user-control-dashboard.header-top__controls--outside.header-top__controls > div:nth-child(5) > div > div > div > div > button';
const USERNAME_SELECTOR = '#username';
const PASSWORD_SELECTOR = '#username-password';
const SUBMIT_BTN_SELECTOR = '#app > div.layout-content.layout-content--theme-primary--40 > div.layout-content__header > header > div.header-top.header__top > div.user-control-dashboard--padding-right.user-control-dashboard.header-top__controls--outside.header-top__controls > div:nth-child(5) > div > div > div > div.ui-inline-dropdown__content > div > div > div > div > form > button';

const NAME_INPUT_SELECTOR = '#__WELCOME_APP__ > div.default-layout-container > div > div > section > main > div > div > div > div > div.user-verify-common-type__form.user-verify-common-type-form > div.ui-input-base-default--theme-default.ui-input-base-default--size-m.ui-input-base-default.ui-input.user-verify-common-type-form__input.ui-input-base.ui-input.user-verify-common-type-form__input > div > div.ui-input-base-default__content > input';

const DEPOSIT_BTN_SELECTOR = '#app > div.layout-content.layout-content--theme-primary--40 > div.layout-content__header > header > div.header-top.header__top > div.user-control-dashboard--padding-right.user-control-dashboard.header-top__controls--outside.header-top__controls > div:nth-child(4) > a';
const VODAFONE_CASH_SELECTOR = '#vodafone_1';
const PHONE_NUMBER_SELECTOR = '#payment_modal_container > div.payment_modal_body > form > div:nth-child(2) > div > span.modal-message-address';

const EXCEL_FILE_PATH = './accounts_phones.xlsx';

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// 📸 دالة صياغة اسم ملف الصورة بناءً على رقم الهاتف والطلب
function getFormattedFileName(phoneNumber, isDuplicate = false) {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  
  // تنظيف اسم الملف واستخدام رقم الهاتف كعنوان رئيسي للصورة
  const cleanPhone = phoneNumber ? phoneNumber.replace(/[^0-9]/g, '') : 'UNKNOWN_PHONE';
  const dupTag = isDuplicate ? '_DUPLICATED' : '';

  return `${cleanPhone}_${year}-${month}-${day}_${hours}-${minutes}${dupTag}.jpg`;
}

// 📊 دالة الحفظ في Excel مع توضيح حالة التكرار
async function appendToExcel(phoneNumber, isDuplicate) {
  const workbook = new ExcelJS.Workbook();
  let worksheet;

  if (fs.existsSync(EXCEL_FILE_PATH)) {
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);
    worksheet = workbook.getWorksheet('Results');
  } else {
    worksheet = workbook.addWorksheet('Results');
    worksheet.columns = [
      { header: 'رقم الموبايل (Phone)', key: 'phone', width: 25 },
      { header: 'حالة الرقم (Status)', key: 'status', width: 20 }
    ];
  }

  worksheet.addRow({
    phone: phoneNumber || 'لم يتم العثور على الرقم',
    status: isDuplicate ? 'DUPLICATED' : 'NEW'
  });

  await workbook.xlsx.writeFile(EXCEL_FILE_PATH);
}

// 🔍 دالة فحص واستخراج رقم الموبايل عبر كل الصفحات والأطر المفتوحة
async function extractPhoneNumber(browser) {
  const pages = await browser.pages();
  
  for (let i = pages.length - 1; i >= 0; i--) {
    const page = pages[i];
    
    // البحث داخل الأطر (Frames)
    for (const frame of page.frames()) {
      try {
        await frame.waitForSelector(PHONE_NUMBER_SELECTOR, { timeout: 1500 }).catch(() => {});
        const phoneText = await frame.evaluate((sel) => {
          const el = document.querySelector(sel);
          return el ? (el.innerText || el.textContent || el.value || '').trim() : null;
        }, PHONE_NUMBER_SELECTOR);

        if (phoneText && phoneText.length > 5) {
          return phoneText;
        }
      } catch (e) {}
    }

    // البحث في الصفحة نفسها
    try {
      await page.waitForSelector(PHONE_NUMBER_SELECTOR, { timeout: 1500 }).catch(() => {});
      const phoneText = await page.evaluate((sel) => {
        const el = document.querySelector(sel);
        return el ? (el.innerText || el.textContent || el.value || '').trim() : null;
      }, PHONE_NUMBER_SELECTOR);

      if (phoneText && phoneText.length > 5) {
        return phoneText;
      }
    } catch (e) {}
  }

  return null;
}

async function processAccount(account, index, total) {
  console.log(`\n[${index + 1}/${total}] 🔑 جاري معالجة: ${account.username}`);

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: {
      width: 1920,
      height: 1080,
      deviceScaleFactor: 2
    },
    args: ['--start-maximized']
  });

  try {
    const pages = await browser.pages();
    let page = pages[0];

    await page.goto(HOME_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });

    await page.waitForSelector(OPEN_LOGIN_BTN_SELECTOR, { visible: true, timeout: 30000 });
    await page.click(OPEN_LOGIN_BTN_SELECTOR);
    await delay(1000);

    // كتابة الإيميل والباسورد
    await page.waitForSelector(USERNAME_SELECTOR, { visible: true, timeout: 30000 });
    await page.click(USERNAME_SELECTOR);
    await delay(300);
    await page.type(USERNAME_SELECTOR, account.username, { delay: 50 });

    await page.click(PASSWORD_SELECTOR);
    await delay(300);
    await page.type(PASSWORD_SELECTOR, account.password, { delay: 50 });
    await delay(500);

    // تسجيل الدخول
    try {
      await page.click(SUBMIT_BTN_SELECTOR);
    } catch (e) {
      await page.keyboard.press('Enter');
    }

    await delay(3000);

    // 🔍 فحص ظهور صفحة الاسم بعد تسجيل الدخول
    try {
      const nameInput = await page.waitForSelector(NAME_INPUT_SELECTOR, { timeout: 4000 });
      if (nameInput) {
        console.log(`👤 ظهرت صفحة الاسم! جاري إدخال: ${account.name}`);
        await nameInput.click();
        await delay(300);
        await nameInput.type(account.name, { delay: 50 });
        await delay(500);
        await page.keyboard.press('Enter');
        await delay(3000);
      }
    } catch (err) {}

    // الضغط على زر الإيداع
    await page.waitForSelector(DEPOSIT_BTN_SELECTOR, { visible: true, timeout: 30000 });
    await page.click(DEPOSIT_BTN_SELECTOR);

    await delay(6000);

    const allPages = await browser.pages();
    page = allPages[allPages.length - 1];

    let clicked = false;
    for (const frame of page.frames()) {
      try {
        const el = await frame.$(VODAFONE_CASH_SELECTOR);
        if (el) {
          await el.scrollIntoViewIfNeeded();
          await el.click();
          clicked = true;
          break;
        }
      } catch (e) {}
    }

    if (!clicked) {
      await page.waitForSelector(VODAFONE_CASH_SELECTOR, { visible: true, timeout: 15000 }).catch(() => {});
      await page.evaluate((sel) => {
        const btn = document.querySelector(sel);
        if (btn) {
          btn.scrollIntoView();
          btn.click();
        }
      }, VODAFONE_CASH_SELECTOR);
    }

    await delay(8000);

    // 1️⃣ استخراج رقم الموبايل بالبحث في كافة المتصفح
    const phoneNumber = await extractPhoneNumber(browser);
    let isDuplicate = false;

    if (phoneNumber) {
      // تنظيف الرقم من أي رموز أو مسافات خفية للمقارنة الدقيقة
      const cleanPhone = phoneNumber.replace(/[^0-9]/g, '');
      if (SEEN_PHONE_NUMBERS.has(cleanPhone)) {
        isDuplicate = true;
        console.warn(`⚠️ تنبيه: الرقم (${phoneNumber}) مكرر!`);
      } else {
        SEEN_PHONE_NUMBERS.add(cleanPhone);
      }
    }

    // 2️⃣ التقاط الصورة وتسميتها برقم الهاتف وبوسم DUPLICATED إن وجد
    const finalPages = await browser.pages();
    const activePage = finalPages[finalPages.length - 1];
    const fileName = getFormattedFileName(phoneNumber, isDuplicate);

    await activePage.screenshot({
      path: fileName,
      type: 'jpeg',
      quality: 100,
      fullPage: false
    });

    console.log(`📞 رقم الموبايل المستخرج: ${phoneNumber || 'لم يتم العثور عليه'}`);
    console.log(`📸 تم التقاط اللقطة باسم: ${fileName}`);

    // 3️⃣ الحفظ الفعلي لرقم الموبايل والتكرار في Excel
    await appendToExcel(phoneNumber, isDuplicate);
    console.log(`📝 تم حفظ بيانات الرقم في ملف Excel بنجاح.`);

  } catch (error) {
    console.log(`❌ حدث خطأ مع الحساب (${account.username}): ${error.message}`);
  } finally {
    await browser.close();
  }
}

(async () => {
  console.log(`🚀 بدء تنفيذ العملية لـ ${ACCOUNTS.length} حساب...`);

  for (let i = 0; i < ACCOUNTS.length; i++) {
    await processAccount(ACCOUNTS[i], i, ACCOUNTS.length);
    await delay(2000);
  }

  console.log(`\n🎉 اكتملت العملية تماماً لجميع الحسابات وتوقف السكربت!`);
})();