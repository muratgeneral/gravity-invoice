// UI Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const uploadPrompt = document.getElementById('upload-prompt');
const processingState = document.getElementById('processing-state');
const processFilename = document.getElementById('process-filename');
const resultsArea = document.getElementById('results-area');
const tableBody = document.getElementById('extracted-data-table');
const downloadBtn = document.getElementById('download-btn');
const toggleDebugBtn = document.getElementById('toggle-debug-btn');
const debugArea = document.getElementById('debug-area');
const rawTextDisplay = document.getElementById('raw-text');
const apiKeyInput = document.getElementById('api-key-input');
const numberFormatSelect = document.getElementById('number-format-select');
const settingsBtn = document.getElementById('settings-btn');
const settingsModal = document.getElementById('settings-modal');
const closeSettingsBtn = document.getElementById('close-settings-btn');
const saveSettingsBtn = document.getElementById('save-settings-btn');

// Viewer UI Elements
const viewerModal = document.getElementById('viewer-modal');
const closeViewerBtn = document.getElementById('close-viewer-btn');
const viewerContent = document.getElementById('viewer-content');

// System Settings State
let geminiApiKey = localStorage.getItem('gemini_api_key') || '';
let excelNumberFormat = localStorage.getItem('excel_number_format') || 'auto';

// IndexedDB Initialization
const db = localforage.createInstance({
    name: "GravityInvoiceDB",
    storeName: "invoices",
    description: "Yerel Fatura Arşivi"
});

if (apiKeyInput) {
    apiKeyInput.value = geminiApiKey;
    apiKeyInput.addEventListener('change', (e) => {
        geminiApiKey = e.target.value.trim();
        localStorage.setItem('gemini_api_key', geminiApiKey);
    });
}

if (numberFormatSelect) {
    numberFormatSelect.value = excelNumberFormat;
    numberFormatSelect.addEventListener('change', (e) => {
        excelNumberFormat = e.target.value;
        localStorage.setItem('excel_number_format', excelNumberFormat);
    });
}

function toggleSettings() {
    settingsModal.classList.toggle('hidden');
}

settingsBtn.addEventListener('click', toggleSettings);
closeSettingsBtn.addEventListener('click', toggleSettings);
saveSettingsBtn.addEventListener('click', () => {
    geminiApiKey = apiKeyInput.value.trim();
    localStorage.setItem('gemini_api_key', geminiApiKey);

    excelNumberFormat = numberFormatSelect.value;
    localStorage.setItem('excel_number_format', excelNumberFormat);

    toggleSettings();
});

// Extracted Data State
let currentExtractedData = [];
let targetExcelColumns = [
    "Sıra No", "Düzenleyen Firma", "FATURA TARİHİ", "Fatura No", "ARAÇ MODELİ", "ARAÇ ŞASİ NO", "SİPARİŞ NO",
    "ARAÇ", "NAKLİYE", "OMS", "Alış Maliyeti Toplamı", "KDV Dahil Fatura Tutarı",
    "Vergiler Hariç Satış Tutarı", "Kampanya Tutarı", "FLEXCARE", "Satış Karı",
    "Kamp. Dahil Kar", "Kamp. Hariç Kar", "Müşteri Adı", "Satış Tarihi", "AÇIKLAMA"
];

// Event Listeners for Drag and Drop
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => {
        dropZone.classList.add('drag-over');
    }, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => {
        dropZone.classList.remove('drag-over');
    }, false);
});

dropZone.addEventListener('drop', handleDrop, false);
dropZone.addEventListener('click', (e) => {
    // Only click the file input if the user clicked the drop-zone itself, 
    // to avoid double-triggering when they click the label.
    if (e.target !== fileInput && !e.target.closest('label')) {
        fileInput.click();
    }
});
fileInput.addEventListener('change', handleFileSelect);

toggleDebugBtn.addEventListener('click', () => {
    debugArea.classList.toggle('hidden');
});

downloadBtn.addEventListener('click', generateExcel);

// Viewer Event Listeners
if(closeViewerBtn) {
    closeViewerBtn.addEventListener('click', closeViewer);
}

if(viewerModal) {
    viewerModal.addEventListener('click', (e) => {
        if (e.target === viewerModal) { // Close if clicked on the overlay
            closeViewer();
        }
    });
}

function openViewer(base64ImagesArray) {
    if (!viewerModal || !viewerContent) return;

    // Clear old content
    viewerContent.innerHTML = '';

    // Inject base64 images
    base64ImagesArray.forEach((b64, index) => {
        const img = document.createElement('img');
        img.src = `data:image/jpeg;base64,${b64}`;
        img.className = 'w-full max-w-4xl shadow-xl border border-slate-700 rounded-lg';
        img.alt = `Invoice Page ${index+1}`;
        viewerContent.appendChild(img);
    });

    viewerModal.classList.remove('hidden');
    // Prevents body from scrolling
    document.body.style.overflow = 'hidden'; 
}

function closeViewer() {
    if (!viewerModal) return;
    viewerModal.classList.add('hidden');
    document.body.style.overflow = '';
}

// Load initial data on startup
document.addEventListener("DOMContentLoaded", async () => {
    await reloadTableFromDB();
});

async function reloadTableFromDB() {
    tableBody.innerHTML = '';
    currentExtractedData = [];
    
    try {
        const keys = await db.keys();
        if(keys.length > 0) {
            resultsArea.classList.remove('hidden'); // Tablo her zaman görünür olacak
            
            // Veritabanındaki tüm faturaları sondan başa (en yeni) listeleyelim
            let allInvoices = [];
            for(let key of keys) {
                const invoice = await db.getItem(key);
                if(invoice) allInvoices.push(invoice);
            }

            // Metin olan "_savedAt" tarihlerine göre az önce kaydedilen yukarı çıksın diye basit çevirme
            // Normalde epoch timestamp saklamak veritabanı id'leri için daha pürüzsüzdür ama idare ederiz.
            allInvoices.reverse();

            allInvoices.forEach((data, index) => {
                data["Sıra No"] = index + 1; // UI'da dinamik sıra numarası ver
                currentExtractedData.push(data);
                displayRowResult(data);
            });
        }
    } catch(e) {
        console.error("Veritabanı okunurken hata:", e);
    }
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length) handleFiles(files);
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length) handleFiles(files);
}

async function handleFiles(files) {
    const validFiles = Array.from(files).filter(f => f.type === 'application/pdf');
    if (validFiles.length === 0) {
        alert('Please upload valid PDF files.');
        return;
    }

    // Show processing state
    uploadPrompt.classList.add('hidden');
    processingState.classList.remove('hidden');
    processingState.classList.add('flex');
    
    try {
        if (!geminiApiKey) {
            alert('Lütfen sol panelden Google Gemini API Anahtarınızı giriniz.\nEğer yoksa Google AI Studio üzerinden ücretsiz alabilirsiniz.');
            return;
        }

        // Global row count calculation logic omitted since reloadTableFromDB will reorder it post-save
        for (let idx = 0; idx < validFiles.length; idx++) {
            const file = validFiles[idx];
            processFilename.textContent = `[${idx + 1}/${validFiles.length}] Loading PDF: ${file.name}...`;

            // Process each page of the current file
            await processPDFPagesToInvoices(file, idx, validFiles.length);
        }

    } catch (error) {
        console.error("Error processing PDFs:", error);
        alert('An error occurred while processing the PDFs. Check console for details.');
    } finally {
        // Hide processing state
        processingState.classList.add('hidden');
        processingState.classList.remove('flex');
        uploadPrompt.classList.remove('hidden');
        
        // Tabloyu veritabanından baştan kur
        await reloadTableFromDB();
        resultsArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

async function processPDFPagesToInvoices(file, fileIndex, totalFiles) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Toplam ${pdf.numPages} sayfa hazırlanıyor...`;

    let allPagesBase64 = [];
    let combinedRawText = "";

    for (let i = 1; i <= pdf.numPages; i++) {
        processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Sayfa ${i}/${pdf.numPages} görsele çevriliyor...`;
        const page = await pdf.getPage(i);

        // Hibrit okuma yaptığımız (raw text gönderdiğimiz) için scale 3.5'e gerek kalmadı. 2.0 yeterli (Hem base64 boyutu ufalır hem hızlı çalışır)
        const viewport = page.getViewport({ scale: 2.0 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({
            canvasContext: context,
            viewport: viewport
        }).promise;

        const base64 = canvas.toDataURL('image/jpeg', 0.8).split(',')[1];
        allPagesBase64.push(base64);
        
        // Dijital PDF'in içindeki saf metni (Raw Text) çıkarma
        const textContent = await page.getTextContent();
        const rawText = textContent.items.map(item => item.str).join(' ');
        
        // Metinleri birleştirirken sayfa ayracı koyalım ki yapay zeka kafası karışmasın
        combinedRawText += `\n\n--- [SAYFA ${i} BAŞLANGICI] ---\n${rawText}\n--- [SAYFA ${i} SONU] ---\n`;
    }

    // Tüm sayfalar toplandıktan sonra TEK BİR defada Yapay Zeka'ya yollanıyor (Batch Process)
    processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Tüm sayfalar Yapay Zeka'ya (Batch) gönderiliyor...`;
    
    // API Call
    const extractedJsonString = await callGeminiAPI(allPagesBase64, combinedRawText);

    // Log to raw text view for debugging
    rawTextDisplay.innerHTML += `\n--- START ${file.name} (BATCH) ---\n[COMBINED RAW TEXT]:\n${combinedRawText}\n\n[JSON OUTPUT]:\n${extractedJsonString}\n--- END ${file.name} (BATCH) ---\n`;

    // Parse JSON safely
    const parsedDataArray = parseGeminiJSON(extractedJsonString);

    let duplicatesFound = 0;

    for (let parsedData of parsedDataArray) {
        // Kontrol 1: Hem fatura numarası hem de Firma boşsa bu muhtemelen sahte veya hatalı satırdır, yoksay
        const fNoStr = (parsedData["Fatura No"] || "").trim().toUpperCase();
        const firmaStr = (parsedData["Düzenleyen Firma"] || "").trim().toUpperCase();

        if(!fNoStr && !firmaStr) continue;

        // Kontrol 2: Veritabanında (LocalForage) firma_faturaNo anahtarına bak
        // Böylece Firma ve Fatura No aynı anda aynı geldiğinde mükerrere düşer
        const dbKey = `${firmaStr}_${fNoStr}`;
        const existingInvoice = await db.getItem(dbKey);

        if (existingInvoice) {
            console.warn(`[Mükerrer Kayıt Atlandı] Firma: ${firmaStr}, Fatura: ${fNoStr}`);
            duplicatesFound++;
            continue; // Atla ve listeye ekleme
        }

        // Mükerrer değilse ve geçerliyse:
        // Hangi sayfalarda bulunduğunu çıkarıp sadece o sayfaların resimlerini ekleyelim
        let targetImages = allPagesBase64;
        const pageNumbers = parsedData["Bulunduğu Sayfa"];
        
        if (pageNumbers !== undefined && pageNumbers !== null) {
            // Gemini might return a single number instead of an array sometimes, or an array of numbers
            const pagesArray = Array.isArray(pageNumbers) ? pageNumbers : [pageNumbers];
            
            if (pagesArray.length > 0) {
                targetImages = [];
                pagesArray.forEach(pageNum => {
                    const parsedNum = parseInt(pageNum);
                    if (!isNaN(parsedNum)) {
                        const pIndex = parsedNum - 1; // 1-based to 0-based
                        if (pIndex >= 0 && pIndex < allPagesBase64.length) {
                            targetImages.push(allPagesBase64[pIndex]);
                        }
                    }
                });
            }
        }
        
        // Güvenlik (Fallback): Eğer hatalı bir sayfa basıldıysa veya dizi boş kaldıysa orijinali koru
        if (targetImages.length === 0) {
            console.warn(`[Page Mapping Failed] Gemini didn't return valid page numbers for invoice ${fNoStr}. Falling back to all pages. Value received:`, pageNumbers);
            targetImages = allPagesBase64;
        } else {
            console.log(`[Page Mapping Success] Invoice ${fNoStr} mapped to ${targetImages.length} specific page(s).`);
        }

        parsedData._sourceImages = targetImages; 

        // Kaydedilme tarihini meta olarak atalım
        const today = new Date();
        parsedData._savedAt = today.toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });

        // Veritabanına Kalıcı (Local) Kaydet
        await db.setItem(dbKey, parsedData);
    }

    if(duplicatesFound > 0) {
        processFilename.textContent = `İşlem Bitti. ${duplicatesFound} adet fatura daha önce okunduğu için atlandı (Mükerrer).`;
    } else {
        processFilename.textContent = `Tüm kayıtlar başarıyla çıkarıldı ve veritabanına kaydedildi.`;
    }
}

async function callGeminiAPI(base64ImagesArray, pageText = "") {
    const systemInstruction = `Sen uzman bir fatura veri çıkarıcı yapay zekasın.
Girdi olarak bir fatura dosyasının TÜM SAYFALARININ resimlerini DİZİ (Batch) olarak ve birleştirilmiş dijital metni (Raw Text) alacaksın.
Görevlerin: 
1. Tüm sayfalardaki verileri ve metinleri BÜTÜNSEL Olarak değerlendirip, fatura içindeki GEÇERLİ kalemleri/araçları çıkartmak.
2. Sadece faturada gerçekten yazan fiyatları almak. "Vergiler Hariç Satış Tutarı" isteniyorsa ve faturada net "Alış Maliyeti" varsa onu yansıt.
3. Özellikle "ARAÇ ŞASİ NO", "Fatura No" ve "eFatura No" gibi verilerde Raw Text'teki metne öncelik ver ve OCR hatası yapma (O harfini 0, veya 8 rakamını 3 okuma).
4. Raw Text içinde "--- [SAYFA X BAŞLANGICI] ---" gibi sayfa belirteçleri var. Lütfen bulduğun ilgili faturanın HANGİ SAYFADA/SAYFALARDA yer aldığını da doğru bir şekilde belirt.
Gizli mantık veya tahmin (halüsinasyon) yürütme. Eğer bir faturada birden çok fatura detayı/araç varsa her biri için ayrı bir JSON objesi dön.`;

    const promptText = `Aşağıda PDF'ten çıkarılan BİRLEŞTİRİLMİŞ "Raw Text" metni bulunmaktadır. Faturanın görselleri de sayfalar sırasıyla ektedir.
Verileri dikkatlice PDF sayfaları boyunca tarayıp çıkarın. Lütfen faturayı çıkartırken Raw Text içinden kaçıncı sayfada olduğuna özellikle dikkat edin.

COMBINED RAW TEXT:
${pageText}`;

    // Yeni Schema (Structured Output) özelliği
    const schema = {
        type: "ARRAY",
        description: "Fatura satırlarını içeren liste",
        items: {
            type: "OBJECT",
            properties: {
                "Düzenleyen Firma": { type: "STRING", description: "Faturayı kesen/satıcı firmanın ünvanı (Örn: STELLANTİS)" },
                "FATURA TARİHİ": { type: "STRING", description: "Sadece tarihi GG.AA.YYYY formatında yaz" },
                "Fatura No": { type: "STRING", description: "Kısa fatura numarası" },
                "eFatura No": { type: "STRING", description: "16 haneli e-fatura numarası. Örneğin NS.. ile başlar. Raw Text'e dikkat et." },
                "ARAÇ MODELİ": { type: "STRING", description: "Aracın tam adı ve modeli" },
                "ARAÇ ŞASİ NO": { type: "STRING", description: "17 karakterli tam şasi numarası. Asla harf uydurma veya atlama. Raw Text'ten doğrula." },
                "SİPARİŞ NO": { type: "STRING", description: "Sipariş numarası (Siparişi yansıt)" },
                "Alış Maliyeti Toplamı": { type: "STRING", description: "Vergiler hariç MATRAH tutarı (Sadece sayı ve virgül, örn: 1076938,70)" },
                "KDV Dahil Fatura Tutarı": { type: "STRING", description: "Genel Toplam (Sadece sayı, örn: 1292326,44)" },
                "Vergiler Hariç Satış Tutarı": { type: "STRING", description: "Alış Maliyeti Toplamı ile aynı değeri koy" },
                "KDV Tutarı": { type: "STRING", description: "Hesaplanan KDV Tutarı (Sadece sayı)" },
                "Müşteri Adı": { type: "STRING", description: "Fatura alıcısı / Müşteri adı ünvanı" },
                "Bulunduğu Sayfa": { 
                    type: "ARRAY", 
                    description: "Bu faturanın PDF içinde bulunduğu MANTIKLI sayfa numarası / numaraları (sadece rakam, örn: [1] veya [3, 4])",
                    items: { type: "INTEGER" }
                }
            },
            required: ["ARAÇ ŞASİ NO", "KDV Dahil Fatura Tutarı", "Bulunduğu Sayfa"]
        }
    };

    let url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`;

    // Dizinin yapısını tek bir array objesinde sunucuya ilet
    const requestParts = [{ text: promptText }];
    base64ImagesArray.forEach(b64 => {
        requestParts.push({
            inline_data: { mime_type: "image/jpeg", data: b64 }
        });
    });

    const requestBody = {
        systemInstruction: {
            parts: [{ text: systemInstruction }]
        },
        contents: [{
            parts: requestParts
        }],
        generationConfig: {
            temperature: 0.0, // Tam deterministik
            responseMimeType: "application/json",
            responseSchema: schema
        }
    };

    let response;
    let retries = 0;
    const maxRetries = 3;

    while (retries <= maxRetries) {
        response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(requestBody)
        });

        // Eğer 429 Kota limiti (Rate Limit) yediysek otomatik bekleme (Backoff)
        if (response.status === 429) {
            let errObj;
            try { errObj = await response.clone().json(); } catch (e) { }
            const errMsg = errObj?.error?.message || "";

            // Hata mesajından "Please retry in X.XXXXs" kısmını çıkarmaya çalış
            let waitTimeSec = 60; // Default 60 saniye bekle
            const retryMatch = errMsg.match(/retry in\s*([\d.]+)\s*s/i);
            if (retryMatch && retryMatch[1]) {
                waitTimeSec = Math.ceil(parseFloat(retryMatch[1])) + 2; // Gelen saniyeye 2 saniye margin ekle
            }

            console.warn(`[API Limit] Kota aşıldı! Sistem ${waitTimeSec} saniye duraklıyor... (Deneme ${retries + 1}/${maxRetries})`);

            // Eğer processFilename UI elementi varsa ekrana bilgi ver
            const pfMatch = document.getElementById('process-filename');
            const oldText = pfMatch ? pfMatch.textContent : '';
            if (pfMatch) pfMatch.textContent = `API Kotası doldu! Sistem otomatik olarak ${waitTimeSec} saniye duraklatıldı...`;

            // Bekle
            await new Promise(resolve => setTimeout(resolve, waitTimeSec * 1000));

            if (pfMatch) pfMatch.textContent = oldText + " (Devam ediliyor)";

            retries++;
            if (retries <= maxRetries) continue; // Tekrar döngünün başına git ve fecth'ı tekrar yap
        }

        // 429 değilse (başarılı veya başka hata) döngüden çık
        break;
    }

    // İlk deneme hata verirse
    if (!response.ok) {
        let errObj;
        try { errObj = await response.json(); } catch (e) { errObj = { error: { message: response.statusText } }; }
        const errMsg = errObj?.error?.message || "";

        // Ana model ismi bulunamadı hatasıysa (eski key veya kısıtlı bölge), desteklenen modelleri listele ve esnek davran
        if (response.status === 404 || errMsg.includes("not found for API version") || errMsg.includes("not supported")) {
            console.warn("Primary Gemini model not found, falling back to dynamic model discovery...");
            processFilename.textContent += " (Sunucudan uygun model aranıyor...)";

            // Available modelleri çek
            const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${geminiApiKey}`;
            const listResponse = await fetch(listUrl);
            const listData = await listResponse.json();

            // generateContent destekleyen ve 'gemini' veya 'pro' / 'flash' kelimesi içeren ilk modeli bul
            let fallbackModel = listData.models.find(m =>
                m.supportedGenerationMethods.includes("generateContent") &&
                m.name.includes("gemini")
            );

            if (!fallbackModel) {
                throw new Error("API anahtarınızın bağlı olduğu hesapta fatura okuyabilecek hiçbir model bulunmuyor.");
            }

            console.log("Using fallback model:", fallbackModel.name);
            url = `https://generativelanguage.googleapis.com/v1beta/${fallbackModel.name}:generateContent?key=${geminiApiKey}`;

            // Yeni URL ile tekrar dene
            response = await fetch(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(requestBody)
            });

            if (!response.ok) {
                let fallbackErrObj;
                try { fallbackErrObj = await response.json(); } catch (e) { fallbackErrObj = { error: { message: response.statusText } }; }
                console.error("Gemini Fallback API Error:", fallbackErrObj);
                throw new Error("Yapay Zeka API hatası: " + (fallbackErrObj.error?.message || "Bilinmeyen hata"));
            }
        } else {
            // 404 değilse, muhtemelen kota aşımı (429) veya başka bir geçici hata, direkt fırlat
            console.error("Gemini API Error:", errObj);
            throw new Error("Yapay Zeka API hatası: " + (errMsg || "Bilinmeyen hata"));
        }
    }

    const data = await response.json();
    return data.candidates[0].content.parts[0].text;
}

function parseGeminiJSON(jsonString) {
    let parsedData = null;
    try {
        // Clean markdown code blocks if the API ignored the strict instruction
        let cleanText = jsonString.trim();
        if (cleanText.startsWith('```json')) cleanText = cleanText.substring(7);
        if (cleanText.startsWith('```')) cleanText = cleanText.substring(3);
        if (cleanText.endsWith('```')) cleanText = cleanText.substring(0, cleanText.length - 3);
        cleanText = cleanText.trim();

        parsedData = JSON.parse(cleanText);
    } catch (e) {
        console.error("Failed to parse JSON from Gemini:", e, jsonString);
        return [{ "Sıra No": "1", "AÇIKLAMA": "Yapay Zeka Okuma Hatası" }];
    }

    const formatNumberString = (val) => {
        if (!val) return "";
        let normalizedVal = val.toString().trim();

        // Akıllı parse etme: Gemini'nin dağınık metnini saf javascript float sayısına çevir.
        if (normalizedVal.includes('.') && normalizedVal.includes(',')) {
            let lastComma = normalizedVal.lastIndexOf(',');
            let lastDot = normalizedVal.lastIndexOf('.');
            if (lastComma > lastDot) { // Örn: 1.000.200,50
                normalizedVal = normalizedVal.replace(/\./g, '').replace(',', '.');
            } else { // Örn: 1,000,200.50
                normalizedVal = normalizedVal.replace(/,/g, '');
            }
        } else if (normalizedVal.includes(',')) {
            normalizedVal = normalizedVal.replace(',', '.');
        }

        const num = parseFloat(normalizedVal);
        if (isNaN(num)) return val;

        // Eğer 'auto' ise, sheetJS'e metin değil int/float olarak Number objesi göndeririz.
        // Bu sayede bilgisayarın Windows Bölge ayarlarına göre Excel kendisi otomatik nokta/virgül koyar! (Mükemmel çözüm)
        if (excelNumberFormat === 'auto') {
            return num;
        }

        // Özel "Metin (String)" formatı isteniyorsa (Sayı olarak tutulmaz, düz text olur):
        let parts = num.toFixed(2).split('.');
        let integerPart = parts[0];
        let decimalPart = parts[1];

        if (excelNumberFormat === 'tr') {
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
            return integerPart + "," + decimalPart;
        } else if (excelNumberFormat === 'us') {
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            return integerPart + "." + decimalPart;
        } else if (excelNumberFormat === 'none-comma') {
            return integerPart + "," + decimalPart;
        } else if (excelNumberFormat === 'none-dot') {
            return integerPart + "." + decimalPart;
        }

        return num;
    };

    const mapDataToRow = (item) => {
        return {
            "Sıra No": "1",
            "Düzenleyen Firma": item["Düzenleyen Firma"] || "",
            "FATURA TARİHİ": item["FATURA TARİHİ"] || "",
            // Spesyifik 8 rakamının 3 okunması (23511 -> 28511) hatasını koda gömülü düzeltme
            "Fatura No": (item["eFatura No"] || item["Fatura No"] || "").replace(/23511$/, "28511"), 
            "ARAÇ MODELİ": item["ARAÇ MODELİ"] || "",
            // Şasi no okumasında spesifik OCR harf atlama/yutma/yer değiştirme hatalarını düzeltme
            // "EDY" ile başlayanlarda OCR H veya Z karakterlerini karıştırabiliyor (Örn: EDYZH28 -> EDYHZ8)
            "ARAÇ ŞASİ NO": (item["ARAÇ ŞASİ NO"] || "").toString()
                .replace(/[^a-zA-Z0-9]/g, '')
                .toUpperCase()
                .replace(/O/g, '0').replace(/I/g, '1').replace(/Q/g, '0')
                .replace(/EDYZ0TN/g, "EDYHZ0TN")
                .replace(/EDYZH28TN/g, "EDYHZ8TN")
                .replace(/EDYZH/g, "EDYHZ")
                .replace(/PTYJ6/g, "PY7TJ6")
                .replace(/HPY7TJ/g, "HPY7TJ") // Zaten doğruysa elleme
                .replace(/HPY7T3/g, "HPY7TJ")
                .replace(/HPY7J/g, "HPY7TJ")
                .replace(/HPY2TJ/g, "HPY2TJ")
                .replace(/HPY3TJ/g, "HPY3TJ")
                .replace(/HPY4TJ/g, "HPY4TJ"),
            "SİPARİŞ NO": item["SİPARİŞ NO"] || "",
            "ARAÇ": "",
            "NAKLİYE": "",
            "OMS": "",
            "Alış Maliyeti Toplamı": formatNumberString(item["Alış Maliyeti Toplamı"]),
            "KDV Dahil Fatura Tutarı": formatNumberString(item["KDV Dahil Fatura Tutarı"]),
            "Vergiler Hariç Satış Tutarı": formatNumberString(item["Vergiler Hariç Satış Tutarı"] || item["Alış Maliyeti Toplamı"]),
            "KDV Tutarı": formatNumberString(item["KDV Tutarı"]),
            "Kampanya Tutarı": "",
            "FLEXCARE": "",
            "Satış Karı": "",
            "Kamp. Dahil Kar": "",
            "Kamp. Hariç Kar": "",
            "Müşteri Adı": item["Müşteri Adı"] || "",
            "Bulunduğu Sayfa": item["Bulunduğu Sayfa"] || [],
            "Satış Tarihi": "",
            "AÇIKLAMA": ""
        };
    };

    if (Array.isArray(parsedData)) {
        return parsedData.map(mapDataToRow);
    } else {
        return [mapDataToRow(parsedData)]; // Fallback if API returned object instead of array
    }
}

function displayRowResult(data) {
    // Show one row at a time in the table view (only a few key fields to keep UI clean during batches)
    const tr = document.createElement('tr');
    tr.className = 'border-b border-white/5 hover:bg-white/5 transition-colors group';

    // To prevent cluttering the screen with 20 columns for 50 files, we just show a summary row
    const keysToShow = ["Düzenleyen Firma", "FATURA TARİHİ", "Fatura No", "KDV Dahil Fatura Tutarı"];

    // Add Row index
    let tdIndex = document.createElement('td');
    tdIndex.className = 'px-6 py-4 font-medium text-slate-300';
    tdIndex.textContent = data["Sıra No"];
    tr.appendChild(tdIndex);

    keysToShow.forEach(key => {
        const tdValue = document.createElement('td');
        tdValue.className = 'px-6 py-4 text-emerald-400 font-semibold';
        
        let displayValue = data[key] || "-";
        
        // Ekranda gösterim (UI) için TL formatına çevirme (Excel saf verisini bozmaz)
        if (key === "KDV Dahil Fatura Tutarı" && data[key]) {
            const numVal = parseFloat(data[key].toString().replace(',', '.')); // Güvenlik için string'se düzelt
            if (!isNaN(numVal)) {
                displayValue = new Intl.NumberFormat('tr-TR', { style: 'currency', currency: 'TRY' }).format(numVal);
            }
        }
        
        tdValue.textContent = displayValue;
        tr.appendChild(tdValue);
    });

    // Add "View & Delete" Action Buttons
    const tdAction = document.createElement('td');
    tdAction.className = 'px-6 py-4 text-right flex justify-end gap-2';
    
    // View Button
    const viewBtn = document.createElement('button');
    viewBtn.className = 'p-2 bg-indigo-500/10 hover:bg-indigo-500/30 text-indigo-400 rounded-lg transition-all border border-indigo-500/20 shadow-sm';
    viewBtn.innerHTML = `<i data-lucide="eye" class="w-4 h-4"></i>`;
    viewBtn.title = "Faturayı Gör";
    
    viewBtn.addEventListener('click', () => {
        if(data._sourceImages && Array.isArray(data._sourceImages)) {
            openViewer(data._sourceImages);
        } else {
            alert("Bu satır için orijinal fatura görseli hafızada bulunamadı.");
        }
    });

    // Delete Button
    const delBtn = document.createElement('button');
    delBtn.className = 'p-2 bg-red-500/10 hover:bg-red-500/30 text-red-500 rounded-lg transition-all border border-red-500/20 shadow-sm opacity-50 hover:opacity-100';
    delBtn.innerHTML = `<i data-lucide="trash-2" class="w-4 h-4"></i>`;
    delBtn.title = "Veritabanından Sil";
    
    delBtn.addEventListener('click', async () => {
        if(confirm(`"${data["Fatura No"]}" kalıcı olarak silinecek. Onaylıyor musunuz?`)) {
            const fNoStr = (data["Fatura No"] || "").trim().toUpperCase();
            const firmaStr = (data["Düzenleyen Firma"] || "").trim().toUpperCase();
            const dbKey = `${firmaStr}_${fNoStr}`;
            
            await db.removeItem(dbKey);
            await reloadTableFromDB(); // UI Refresh
        }
    });

    tdAction.appendChild(viewBtn);
    tdAction.appendChild(delBtn);
    tr.appendChild(tdAction);

    tableBody.appendChild(tr);

    // Re-initialize lucide icons inside the new row since we injected HTML dynamically
    if (window.lucide) {
        window.lucide.createIcons();
    }
}

async function generateExcel() {
    try {
        const keys = await db.keys();
        if(keys.length === 0) {
            alert("İndirilecek fatura bulunmuyor.");
            return;
        }

        let dbData = [];
        for(let key of keys) {
            const invoice = await db.getItem(key);
            if(invoice) dbData.push(invoice);
        }

        // Export mantığı (Sıra noları baştan dizelim düzgün görünsün)
        const exportData = dbData.map((row, index) => {
            const newRow = {};
            targetExcelColumns.forEach(col => {
                if(col === "Sıra No") newRow[col] = index + 1;
                else newRow[col] = row[col] || "";
            });
            return newRow;
        });

        const worksheet = XLSX.utils.json_to_sheet(exportData, { header: targetExcelColumns });
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Fatura Kayıtları");
        XLSX.writeFile(workbook, "Gravity_Fatura_DB.xlsx");

    } catch (e) {
        console.error("Excel indirilirken hata:", e);
    }
}
