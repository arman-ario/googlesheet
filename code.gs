const CONFIG = {
  BASE_URL: "https://abighermez.com/wp-json/wc/v3",
  CONSUMER_KEY: "ck_60b093fe5c5a816993f9ed11a7ac32bf580d75bc",
  CONSUMER_SECRET: "cs_cddb9a82b9e5cd27f2be1404c9b2ae082bf8b941",
  MAX_PRODUCTS: 100,
  SHEET_NAME: 'Products',
  DRIVE_CONFIG: {
    PDF_FOLDER_NAME: 'Product Lists PDF'
  }
};

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('مدیریت محصولات')
    .addItem('دریافت محصولات', 'startProcess')
    .addItem('فعال‌سازی به‌روزرسانی خودکار', 'setupAllAutomatedTasks')
    .addSeparator()
    .addItem('مرتب‌سازی بر اساس دسته‌بندی و نام', 'sortProductsByCategoyAndName')
    .addItem('حذف هدر کالاهای متغیر', 'removeVariableProductHeaders')
    .addSeparator()
    .addItem('تنظیم نمای چاپ A4', 'setupPrintView')
    .addItem('دریافت خروجی PDF', 'exportToPDF')
    .addToUi();
}

function deleteTriggers(functionName = null) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (!functionName || trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function getInStockProductsCount() {
  const url = `${CONFIG.BASE_URL}/products?` +
    `consumer_key=${CONFIG.CONSUMER_KEY}&` +
    `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
    `per_page=1&` +
    `stock_status=instock&` +
    `status=publish`;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      muteHttpExceptions: true,
      timeout: 30000
    });

    if (response.getResponseCode() === 200) {
      const headers = response.getAllHeaders();
      return parseInt(headers['x-wp-total'] || headers['X-WP-Total']) || 0;
    }
    return 0;
  } catch (error) {
    Logger.log(`Error getting total products: ${error}`);
    throw error;
  }
}

function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  
  sheet.clear();
  
  const headers = [
  "ID",
  "دسته‌بندی",
  "نام محصول",
  "قیمت نقدی",
  "قیمت چکی",
  "قیمت با اسنپ‌پی",
  "قیمت مصرف‌کننده",
  "موجودی",          // این خط جدید است
  "تصویر",
  "رنگ‌های موجود",
  "برند",
  "وضعیت موجودی",
  "نوع"
];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4a90e2');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  sheet.setColumnWidths(1, headers.length, 120);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 150);
}

function startProcess() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const totalProducts = getInStockProductsCount();
    
    if (totalProducts > 0) {
      const result = ui.alert(
        'شروع ایمپورت محصولات',
        `تعداد کل محصولات موجود: ${totalProducts}\n\nآیا می‌خواهید فرآیند ایمپورت را شروع کنید؟`,
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        deleteTriggers('importProducts');
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.deleteAllProperties();
        scriptProperties.setProperties({
          'totalProducts': totalProducts.toString(),
          'currentPage': '1',
          'processedProducts': '0',
          'retryCount': '0'
        });
        
        initializeSheet();
        importProducts();
      }
    } else {
      ui.alert('خطا', 'هیچ محصول موجودی یافت نشد.', ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('خطا', 'خطا در دریافت اطلاعات محصولات: ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('Error in startProcess: ' + error.toString());
  }
}

function getProductColors(product) {
  if (!product.attributes) return '';
  const colorAttr = product.attributes.find(attr => 
    ['رنگ', 'color', 'rang'].includes(attr.name.toLowerCase())
  );
  return colorAttr ? (colorAttr.options || []).join(', ') : '';
}

function getProductBrand(product) {
  if (!product.attributes) return '';
  const brandAttr = product.attributes.find(attr => 
    ['برند', 'brand'].includes(attr.name.toLowerCase())
  );
  return brandAttr ? (brandAttr.options || [])[0] || '' : '';
}

function importProducts() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentPage = parseInt(scriptProperties.getProperty('currentPage') || '1');
  const totalProducts = parseInt(scriptProperties.getProperty('totalProducts'));
  const processedProducts = parseInt(scriptProperties.getProperty('processedProducts') || '0');
  const retryCount = parseInt(scriptProperties.getProperty('retryCount') || '0');
  
  const url = `${CONFIG.BASE_URL}/products?` +
    `consumer_key=${CONFIG.CONSUMER_KEY}&` +
    `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
    `per_page=${CONFIG.MAX_PRODUCTS}&` +
    `page=${currentPage}&` +
    `stock_status=instock&` +
    `status=publish&` +
    `orderby=id&` +
    `order=asc`;
  
  try {
    const response = UrlFetchApp.fetch(url, { 
      muteHttpExceptions: true,
      timeout: 30000
    });
    
    if (response.getResponseCode() === 200) {
      const products = JSON.parse(response.getContentText());
      
      if (products && products.length > 0) {
        writeProductsToSheet(products);
        
        const newProcessedProducts = processedProducts + products.length;
        const progress = Math.round((newProcessedProducts / totalProducts) * 100);
        
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `پردازش ${newProcessedProducts} از ${totalProducts} محصول (${progress}%)`
        );
        
        scriptProperties.setProperties({
          'currentPage': (currentPage + 1).toString(),
          'processedProducts': newProcessedProducts.toString(),
          'retryCount': '0'
        });
        
        if (newProcessedProducts < totalProducts) {
          Utilities.sleep(5000);
          ScriptApp.newTrigger('importProducts')
            .timeBased()
            .after(60000)
            .create();
        } else {
          completedImport(newProcessedProducts);
        }
      } else {
        completedImport(processedProducts);
      }
    } else {
      handleError(currentPage, retryCount);
    }
  } catch (error) {
    handleError(currentPage, retryCount);
  }
}

function writeProductsToSheet(products) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  
  const existingIDs = new Set();
  if (sheet.getLastRow() > 1) {
    const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    existingData.forEach(row => existingIDs.add(row[0].toString()));
  }
  
  let allProductsData = [];
  
  products.forEach(product => {
    if (!existingIDs.has(product.id.toString())) {
      const regularPrice = parseFloat(product.regular_price || product.price || 0);
      const checkPrice = regularPrice;
      const cashPrice = Math.round(checkPrice * 0.95);
      const snappayPrice = Math.round(checkPrice * 1.05);
      const consumerPrice = Math.round(checkPrice * 1.3);

      allProductsData.push([
  product.id,
  (product.categories || []).map(cat => cat.name).join(', '),
  product.name,
  cashPrice,
  checkPrice,
  snappayPrice,
  consumerPrice,
  product.stock_quantity || 0,     // این خط جدید است
  product.images.length ? `=IMAGE("${product.images[0].src}")` : '',
  getProductColors(product),
  getProductBrand(product),
  product.stock_status,
  product.type
]);

      if (product.type === 'variable' && product.variations && product.variations.length > 0) {
        try {
          const variationUrl = `${CONFIG.BASE_URL}/products/${product.id}/variations?` +
            `consumer_key=${CONFIG.CONSUMER_KEY}&` +
            `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
            `per_page=100&` +
            `stock_status=instock`;
          
          const variationResponse = UrlFetchApp.fetch(variationUrl, { 
            muteHttpExceptions: true,
            timeout: 30000
          });
          
          if (variationResponse.getResponseCode() === 200) {
            const variations = JSON.parse(variationResponse.getContentText());
            variations.forEach(variation => {
              if (variation.stock_status === 'instock') {
                const varRegularPrice = parseFloat(variation.regular_price || variation.price || 0);
                const varCheckPrice = varRegularPrice;
                const varCashPrice = Math.round(varCheckPrice * 0.95);
                const varSnappayPrice = Math.round(varCheckPrice * 1.05);
                const varConsumerPrice = Math.round(varCheckPrice * 1.3);

                const variationAttributes = variation.attributes
                  .map(attr => attr.option)
                  .filter(option => option)
                  .join(' - ');
                
                allProductsData.push([
  variation.id,
  (product.categories || []).map(cat => cat.name).join(', '),
  `${product.name} (${variationAttributes})`,
  varCashPrice,
  varCheckPrice,
  varSnappayPrice,
  varConsumerPrice,
  variation.stock_quantity || 0,    // این خط جدید است
  product.images.length ? `=IMAGE("${product.images[0].src}")` : '',
  variationAttributes,
  getProductBrand(product),
  variation.stock_status,
  'variation'
]);
              }
            });
          }
        } catch (error) {
          Logger.log(`Error fetching variations for product ${product.id}: ${error}`);
        }
      }
    }
  });
  
  if (allProductsData.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(startRow, 1, allProductsData.length, allProductsData[0].length);
    range.setValues(allProductsData);
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    const fullRange = sheet.getRange(1, 1, lastRow, lastColumn);
    
    fullRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 150);
    sheet.setColumnWidth(9, 150);
    sheet.setColumnWidth(10, 100);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 80);

    fullRange.setVerticalAlignment('middle');
    sheet.getRange(1, 1, lastRow, lastColumn).setHorizontalAlignment('center');
    sheet.getRange(1, 3, lastRow, 1).setHorizontalAlignment('right');
    sheet.getRange(2, 4, lastRow - 1, 4).setNumberFormat('#,##0');

    sheet.setRowHeight(1, 35);
    for (let i = startRow; i <= lastRow; i++) {
      sheet.setRowHeight(i, 150);
    }

    fullRange.setFontFamily('B Nazanin')
             .setFontSize(11)
             .setFontWeight('bold');
    
    sheet.getRange(1, 1, 1, lastColumn)
         .setFontSize(12)
         .setFontWeight('bold');

    sheet.getRange(1, 3, lastRow, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    sheet.getRange(1, lastColumn + 1).setValue('آخرین به‌روزرسانی:')
         .setFontFamily('B Nazanin')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
    
    sheet.getRange(1, lastColumn + 2).setValue('2025-02-15 14:31:12')
         .setFontFamily('B Nazanin')
         .setHorizontalAlignment('center');

    sheet.getRange(1, lastColumn + 3).setValue('User:')
         .setFontFamily('B Nazanin')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
    
    sheet.getRange(1, lastColumn + 4).setValue('s-arman-m-j')
         .setFontFamily('B Nazanin')
         .setHorizontalAlignment('center');
  }
}
function handleError(currentPage, retryCount) {
  const scriptProperties = PropertiesService.getScriptProperties();
  if (retryCount < 3) {
    scriptProperties.setProperty('retryCount', (retryCount + 1).toString());
    ScriptApp.newTrigger('importProducts')
      .timeBased()
      .after(120000)
      .create();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `خطا در پردازش صفحه ${currentPage}. تلاش مجدد در 2 دقیقه دیگر...`
    );
  } else {
    deleteTriggers('importProducts');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'خطا در پردازش. لطفاً دوباره تلاش کنید.'
    );
  }
}

function completedImport(totalProcessed) {
  deleteTriggers('importProducts');
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `ایمپورت کامل شد. ${totalProcessed} محصول با موفقیت پردازش شد.`
  );
}

function removeVariableProductHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  try {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      ui.alert('خطا', 'داده‌ای برای پردازش وجود ندارد.', ui.ButtonSet.OK);
      return;
    }

    // خواندن تمام داده‌ها
    const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const typeColumnIndex = 12; // ستون نوع محصول (M)
    let rowsToDelete = [];
    let previousWasVariable = false;

    // پیدا کردن ردیف‌های هدر محصولات متغیر
    for (let i = 1; i < data.length; i++) {
      const currentType = data[i][typeColumnIndex];
      
      if (currentType === 'variable') {
        rowsToDelete.push(i + 1);
        previousWasVariable = true;
      } else {
        previousWasVariable = false;
      }
    }

    // حذف ردیف‌ها از آخر به اول
    if (rowsToDelete.length > 0) {
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
      }

      // به‌روزرسانی اطلاعات
      sheet.getRange(1, lastColumn + 1).setValue('Last Update:')
           .setFontFamily('B Nazanin')
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
      
      sheet.getRange(1, lastColumn + 2).setValue('2025-02-25 09:55:10')
           .setFontFamily('B Nazanin')
           .setHorizontalAlignment('center');

      sheet.getRange(1, lastColumn + 3).setValue('User:')
           .setFontFamily('B Nazanin')
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
      
      sheet.getRange(1, lastColumn + 4).setValue('arman-ario')
           .setFontFamily('B Nazanin')
           .setHorizontalAlignment('center');

      ui.alert(
        'عملیات موفق',
        `تعداد ${rowsToDelete.length} ردیف هدر محصول متغیر با موفقیت حذف شد.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'اطلاعات',
        'هیچ هدر محصول متغیری برای حذف یافت نشد.',
        ui.ButtonSet.OK
      );
    }

  } catch (error) {
    Logger.log('Error in removeVariableProductHeaders: ' + error.toString());
    ui.alert(
      'خطا',
      'خطایی در حذف هدر محصولات متغیر رخ داد. لطفاً دوباره تلاش کنید.',
      ui.ButtonSet.OK
    );

    // ارسال ایمیل خطا
    MailApp.sendEmail({
      to: "arman-m-j@gmail.com",
      subject: "خطا در حذف هدر محصولات متغیر",
      body: `خطای زیر در هنگام حذف هدر محصولات متغیر رخ داد:\n\nزمان خطا: 2025-02-25 09:55:10\nکاربر: arman-ario\nخطا: ${error.toString()}\n\nلطفاً بررسی کنید و در صورت نیاز عملیات را به صورت دستی انجام دهید.`
    });
  }
}

function sortProductsByCategoyAndName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('خطا', 'شیت مورد نظر یافت نشد!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('توجه', 'داده‌ای برای مرتب‌سازی وجود ندارد!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    range.sort([
      {column: 2, ascending: true},
      {column: 3, ascending: true}
    ]);

    SpreadsheetApp.getActiveSpreadsheet().toast('مرتب‌سازی با موفقیت انجام شد.');

  } catch (error) {
    Logger.log('Error in sorting: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('خطا در مرتب‌سازی. لطفاً دوباره تلاش کنید.');
  }
}

function setupPrintView() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  
  try {
    const totalColumns = sheet.getMaxColumns();
    for (let i = 1; i <= totalColumns; i++) {
      // فقط ستون‌های 3 (نام), 4 (قیمت نقدی), 5 (قیمت چکی), 6 (قیمت اسنپ‌پی) و 9 (عکس) نمایش داده شوند
      if (![3,4,5,6,9].includes(i)) {
        sheet.hideColumn(sheet.getRange(1, i, 1, 1));
      }
    }

    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(9, 150); // تنظیم عرض ستون عکس

    sheet.setFrozenRows(1);
    sheet.setPageView();
  } catch (error) {
    Logger.log('Error in setupPrintView: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('خطا در تنظیم نمای چاپ. لطفاً دوباره تلاش کنید.');
  }
}

function exportToPDF() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  try {
    setupPrintView();
    
    const timestamp = Utilities.formatDate(new Date(), "Asia/Tehran", "yyyy-MM-dd_HH-mm-ss");
    const pdfFileName = `Products_List_${timestamp}.pdf`;

    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
                'exportFormat=pdf&' +
                'format=pdf&' +
                'size=A4&' +
                'portrait=true&' +
                'fitw=true&' +
                'sheetnames=false&' +
                'printtitle=false&' +
                'pagenumbers=false&' +
                'gridlines=false&' +
                'fzr=false&' +
                'top_margin=0.25&' +
                'bottom_margin=0.25&' +
                'left_margin=0.25&' +
                'right_margin=0.25&' +
                `gid=${sheet.getSheetId()}`;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    const pdfBlob = response.getBlob().setName(pdfFileName);
    const pdfFolder = getOrCreateFolder(CONFIG.DRIVE_CONFIG.PDF_FOLDER_NAME);
    const pdfFile = pdfFolder.createFile(pdfBlob);
    const downloadUrl = pdfFile.getUrl();

    const htmlOutput = HtmlService
      .createHtmlOutput(`
        <p dir="rtl" style="font-family: 'B Nazanin'; text-align: center;">
          فایل PDF با موفقیت ایجاد شد.<br><br>
          <a href="${downloadUrl}" target="_blank" style="text-decoration: none;">
            <button style="padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer;">
              دانلود فایل PDF
            </button>
          </a>
        </p>
        <p dir="rtl" style="font-size: 12px; color: #666; text-align: center;">
          فایل در گوگل درایو شما نیز ذخیره شده است
        </p>
      `)
      .setWidth(300)
      .setHeight(150);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'دانلود PDF');

  } catch (error) {
    Logger.log('Error in exportToPDF: ' + error.toString());
    ui.alert('خطا', 'خطای زیر رخ داد:\n\n' + error.toString(), ui.ButtonSet.OK);
  }
}

// فانکشن‌های مربوط به به‌روزرسانی خودکار
function startProcess(isAutomatic = false) {
  try {
    // دریافت تعداد کل محصولات (موجود و ناموجود)
    const url = `${CONFIG.BASE_URL}/products?` +
      `consumer_key=${CONFIG.CONSUMER_KEY}&` +
      `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
      `per_page=1`;
    
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      muteHttpExceptions: true,
      timeout: 30000
    });

    let stats = {
      totalProducts: 0,
      totalInStock: 0,
      totalVariable: 0,
      variableInStock: 0,
      variableHeaders: 0
    };

    if (response.getResponseCode() === 200) {
      const headers = response.getAllHeaders();
      stats.totalProducts = parseInt(headers['x-wp-total'] || headers['X-WP-Total']) || 0;

      // دریافت تعداد کل محصولات موجود
      const inStockUrl = `${CONFIG.BASE_URL}/products?` +
        `consumer_key=${CONFIG.CONSUMER_KEY}&` +
        `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
        `per_page=1&` +
        `stock_status=instock`;

      const inStockResponse = UrlFetchApp.fetch(inStockUrl, {
        method: "GET",
        muteHttpExceptions: true,
        timeout: 30000
      });

      if (inStockResponse.getResponseCode() === 200) {
        const inStockHeaders = inStockResponse.getAllHeaders();
        stats.totalInStock = parseInt(inStockHeaders['x-wp-total'] || inStockHeaders['X-WP-Total']) || 0;
      }

      // دریافت تعداد کل محصولات متغیر
      const variableUrl = `${CONFIG.BASE_URL}/products?` +
        `consumer_key=${CONFIG.CONSUMER_KEY}&` +
        `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
        `per_page=1&` +
        `type=variable`;

      const variableResponse = UrlFetchApp.fetch(variableUrl, {
        method: "GET",
        muteHttpExceptions: true,
        timeout: 30000
      });

      if (variableResponse.getResponseCode() === 200) {
        const variableHeaders = variableResponse.getAllHeaders();
        stats.totalVariable = parseInt(variableHeaders['x-wp-total'] || variableHeaders['X-WP-Total']) || 0;
      }

      // دریافت تعداد محصولات متغیر موجود
      const variableInStockUrl = `${CONFIG.BASE_URL}/products?` +
        `consumer_key=${CONFIG.CONSUMER_KEY}&` +
        `consumer_secret=${CONFIG.CONSUMER_SECRET}&` +
        `per_page=1&` +
        `type=variable&` +
        `stock_status=instock`;

      const variableInStockResponse = UrlFetchApp.fetch(variableInStockUrl, {
        method: "GET",
        muteHttpExceptions: true,
        timeout: 30000
      });

      if (variableInStockResponse.getResponseCode() === 200) {
        const varInStockHeaders = variableInStockResponse.getAllHeaders();
        stats.variableInStock = parseInt(varInStockHeaders['x-wp-total'] || varInStockHeaders['X-WP-Total']) || 0;
        // تعداد هدرهای متغیر موجود برابر با تعداد محصولات متغیر موجود است
        stats.variableHeaders = stats.variableInStock;
      }
    }
    
    if (stats.totalInStock > 0) {
      const currentDate = Utilities.formatDate(new Date(), "UTC", "YYYY-MM-DD HH:MM:SS");
      const currentUser = 's-arman-m-j';
      
      if (!isAutomatic) {
        const ui = SpreadsheetApp.getUi();
        const message = 
          `Current Date and Time (UTC): ${currentDate}\n` +
          `Current User's Login: ${currentUser}\n\n` +
          `آمار محصولات:\n` +
          `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
          `📦 تعداد کل محصولات: ${stats.totalProducts}\n` +
          `✅ تعداد کل محصولات موجود: ${stats.totalInStock}\n` +
          `🔄 تعداد کل محصولات متغیر: ${stats.totalVariable}\n` +
          `📍 تعداد محصولات متغیر موجود: ${stats.variableInStock}\n` +
          `🔰 تعداد هدرهای متغیر موجود: ${stats.variableHeaders}\n\n` +
          `آیا می‌خواهید فرآیند ایمپورت را شروع کنید؟`;

        const result = ui.alert('شروع ایمپورت محصولات', message, ui.ButtonSet.YES_NO);
        if (result !== ui.Button.YES) return;
      }
      
      deleteTriggers('importProducts');
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.deleteAllProperties();
      scriptProperties.setProperties({
        'totalProducts': stats.totalInStock.toString(),
        'currentPage': '1',
        'processedProducts': '0',
        'retryCount': '0',
        'startTime': currentDate,
        'startUser': currentUser,
        'totalStats': JSON.stringify(stats)
      });
      
      initializeSheet();
      importProducts();

      // نمایش پیام آمار در لاگ
      Logger.log(`
        Import started at: ${currentDate}
        User: ${currentUser}
        Total Products: ${stats.totalProducts}
        Total In-Stock Products: ${stats.totalInStock}
        Total Variable Products: ${stats.totalVariable}
        Variable In-Stock Products: ${stats.variableInStock}
        Variable Headers: ${stats.variableHeaders}
      `);
      
      if (!isAutomatic) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `شروع دریافت ${stats.totalInStock} محصول موجود...`
        );
      }
    } else {
      if (!isAutomatic) {
        SpreadsheetApp.getUi().alert('خطا', 'هیچ محصول موجودی یافت نشد.', SpreadsheetApp.getUi().ButtonSet.OK);
      }
      Logger.log('No products found to import');
    }
  } catch (error) {
    if (!isAutomatic) {
      SpreadsheetApp.getUi().alert('خطا', 'خطا در دریافت اطلاعات محصولات: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
    Logger.log('Error in startProcess: ' + error.toString());
  }
}

function completedImport(totalProcessed) {
  deleteTriggers('importProducts');
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = scriptProperties.getProperty('startTime');
  const startUser = scriptProperties.getProperty('startUser');
  const statsString = scriptProperties.getProperty('totalStats');
  let stats = { 
    totalProducts: 0, 
    totalInStock: 0, 
    totalVariable: 0, 
    variableInStock: 0, 
    variableHeaders: 0 
  };
  
  try {
    stats = JSON.parse(statsString);
  } catch (e) {
    Logger.log('Error parsing stats: ' + e.toString());
  }

  const currentDate = Utilities.formatDate(new Date(), "UTC", "YYYY-MM-DD HH:MM:SS");
  
  const message = 
    `Current Date and Time (UTC): ${currentDate}\n` +
    `Current User's Login: ${startUser}\n\n` +
    `فرآیند ایمپورت با موفقیت به پایان رسید.\n\n` +
    `شروع: ${startTime}\n` +
    `پایان: ${currentDate}\n\n` +
    `آمار نهایی:\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `📦 تعداد کل محصولات: ${stats.totalProducts}\n` +
    `✅ تعداد کل محصولات موجود: ${stats.totalInStock}\n` +
    `🔄 تعداد کل محصولات متغیر: ${stats.totalVariable}\n` +
    `📍 تعداد محصولات متغیر موجود: ${stats.variableInStock}\n` +
    `🔰 تعداد هدرهای متغیر موجود: ${stats.variableHeaders}\n` +
    `📥 محصولات پردازش شده: ${totalProcessed}`;

  SpreadsheetApp.getActiveSpreadsheet().toast(message, '✅ تکمیل فرآیند', -1);
  scriptProperties.deleteAllProperties();
}

function exportToPDFAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  try {
    setupPrintView();
    
    const timestamp = Utilities.formatDate(new Date(), "Asia/Tehran", "yyyy-MM-dd_HH-mm");
    const pdfFileName = `Products_List_${timestamp}.pdf`;

    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
                'exportFormat=pdf&' +
                'format=pdf&' +
                'size=A4&' +
                'portrait=true&' +
                'fitw=true&' +
                'sheetnames=false&' +
                'printtitle=false&' +
                'pagenumbers=false&' +
                'gridlines=false&' +
                'fzr=false&' +
                'top_margin=0.25&' +
                'bottom_margin=0.25&' +
                'left_margin=0.25&' +
                'right_margin=0.25&' +
                `gid=${sheet.getSheetId()}`;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    const pdfBlob = response.getBlob().setName(pdfFileName);
    const pdfFolder = getOrCreateFolder(CONFIG.DRIVE_CONFIG.PDF_FOLDER_NAME);
    pdfFolder.createFile(pdfBlob);

    Logger.log(`PDF file created automatically: ${pdfFileName}`);
  } catch (error) {
    Logger.log('Error in exportToPDFAuto: ' + error.toString());
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}
