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

// System Settings State
let geminiApiKey = localStorage.getItem('gemini_api_key') || '';
let excelNumberFormat = localStorage.getItem('excel_number_format') || 'auto';

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
    "Sıra No", "FATURA TARİHİ", "Fatura No", "ARAÇ MODELİ", "ARAÇ ŞASİ NO", "SİPARİŞ NO",
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
    resultsArea.classList.add('hidden');

    // Clear previous extracts if any, so we only download what we just uploaded
    currentExtractedData = [];
    tableBody.innerHTML = '';
    rawTextDisplay.innerHTML = '';

    // Global row count tracking across multiple files and multiple pages
    let globalRowIndex = 1;

    try {
        if (!geminiApiKey) {
            alert('Lütfen sol panelden Google Gemini API Anahtarınızı giriniz.\nEğer yoksa Google AI Studio üzerinden ücretsiz alabilirsiniz.');
            return;
        }

        for (let idx = 0; idx < validFiles.length; idx++) {
            const file = validFiles[idx];
            processFilename.textContent = `[${idx + 1}/${validFiles.length}] Loading PDF: ${file.name}...`;

            // Process each page of the current file
            globalRowIndex = await processPDFPagesToInvoices(file, idx, validFiles.length, globalRowIndex);
        }

    } catch (error) {
        console.error("Error processing PDFs:", error);
        alert('An error occurred while processing the PDFs. Check console for details.');
    } finally {
        // Hide processing state
        processingState.classList.add('hidden');
        processingState.classList.remove('flex');
        uploadPrompt.classList.remove('hidden');
        if (currentExtractedData.length > 0) {
            resultsArea.classList.remove('hidden');
            resultsArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }
}

async function processPDFPagesToInvoices(file, fileIndex, totalFiles, currentRowIndex) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    // Helper sleep function for free tier API rate limits
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    // Her dosya arasında minik bir bekleme (5 sn) kotayı rahatlatır
    if (fileIndex > 0) {
        processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Güvenlik molası (API Kotası)...`;
        await sleep(5000);
    }

    processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Toplam ${pdf.numPages} sayfa hazırlanıyor...`;

    // Bütün sayfaları tek bir dizide toplayacağız
    const base64Images = [];

    for (let i = 1; i <= pdf.numPages; i++) {
        processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Sayfa ${i}/${pdf.numPages} görsele çevriliyor...`;
        const page = await pdf.getPage(i);

        const viewport = page.getViewport({ scale: 2.0 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({
            canvasContext: context,
            viewport: viewport
        }).promise;

        base64Images.push(canvas.toDataURL('image/jpeg', 0.8).split(',')[1]);
    }

    // Call Gemini API once for ALL pages in this PDF
    processFilename.textContent = `[Dosya ${fileIndex + 1}/${totalFiles}] - Tüm ${pdf.numPages} sayfa tek seferde Yapay Zeka'ya gönderiliyor...`;
    const extractedJsonString = await callGeminiAPI(base64Images);

    // Log to raw text view for debugging
    rawTextDisplay.innerHTML += `\n--- START ${file.name} (Tüm Sayfalar Toplu) ---\n${extractedJsonString}\n--- END ${file.name} ---\n`;

    // Parse JSON safely (şimdi bize bir Dizi(Array) gelecek)
    const parsedDataArray = parseGeminiJSON(extractedJsonString);

    parsedDataArray.forEach(parsedData => {
        parsedData["Sıra No"] = currentRowIndex;
        currentExtractedData.push(parsedData);
        displayRowResult(parsedData);
        currentRowIndex++;
    });

    return currentRowIndex;
}

async function callGeminiAPI(base64ImagesArray) {
    const prompt = `Sen uzman bir fatura veri çıkarıcı yapay zekasın.
Girdi olarak tek bir seferde BİRDEN FAZLA fatura sayfası görüntüsü alacaksın.
Senin görevin her bir görsel sayfasındaki verileri tek tek okuyup, bunları bir JSON JSON DİZİSİ (Array) içerisinde döndürmek.
Eksiksiz ve birebir aynı formatta çıkararak cevap vermelisin. 
Hiçbir ekstra metin, açıklama veya markdown satırı ekleme; DOĞRUDAN geçerli bir JSON DİZİSİ (Array) objesi olarak cevap ver. Asla \`\`\`json blokları kullanma.
Beklenen Çıktı Formatı (Sayfa sayısı kaç ise, dizi içinde o kadar obje olmalı):
[
  {
    "FATURA TARİHİ": "Sadece tarihi GG.AA.YYYY yaz (Ödeme tarihini alma)",
    "Fatura No": "Kısa fatura numarası (varsa)",
    "eFatura No": "Genellikle 3 harf (INS, E53 vb.) ve 13 rakamdan oluşan 16 haneli tam kod",
    "ARAÇ MODELİ": "Aracın sadece tam adı ve modeli (Motor no, gümrük kodu veya ETTN şifrelerini alma)",
    "ARAÇ ŞASİ NO": "17 haneli şasi numarası (VR7, vs ile başlar. I ve O yerine 1 ve 0 lara dikkat et)",
    "SİPARİŞ NO": "Sipariş No (C202... tarzı numaralar)",
    "Alış Maliyeti Toplamı": "Vergiler hariç MATRAH tutarı (Sadece sayı, noktasız ve küsürat virgüllü: 1076938,70 gibi)",
    "KDV Dahil Fatura Tutarı": "Genel toplam/Fatura Tutarı (Sadece sayı, örn: 1292326,44)",
    "Vergiler Hariç Satış Tutarı": "Alış Maliyeti Toplamı ile aynı değeri koy",
    "KDV Tutarı": "%20 KDV tutarının rakamı (Sadece sayı)",
    "Müşteri Adı": "Fatura Alıcısı kısmındaki müşteri ünvanı/adı"
  }
]`;

    let url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`;

    // Dizinin yapısını tek bir array objesinde sunucuya ilet
    const requestParts = [{ text: prompt }];
    base64ImagesArray.forEach(b64 => {
        requestParts.push({
            inline_data: { mime_type: "image/jpeg", data: b64 }
        });
    });

    const requestBody = {
        contents: [{
            parts: requestParts
        }],
        generationConfig: {
            temperature: 0.1, // Deterministic
            response_mime_type: "application/json" // Force JSON array 
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
            "FATURA TARİHİ": item["FATURA TARİHİ"] || "",
            "Fatura No": item["eFatura No"] || item["Fatura No"] || "", // Prioritize 16 digit one
            "ARAÇ MODELİ": item["ARAÇ MODELİ"] || "",
            "ARAÇ ŞASİ NO": item["ARAÇ ŞASİ NO"] || "",
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
    tr.className = 'border-b border-white/5 hover:bg-white/5 transition-colors';

    // To prevent cluttering the screen with 20 columns for 50 files, we just show a summary row
    const keysToShow = ["Fatura No", "ARAÇ ŞASİ NO", "KDV Dahil Fatura Tutarı"];

    // Add Row index
    let tdIndex = document.createElement('td');
    tdIndex.className = 'px-6 py-4 font-medium text-slate-300';
    tdIndex.textContent = data["Sıra No"];
    tr.appendChild(tdIndex);

    keysToShow.forEach(key => {
        const tdValue = document.createElement('td');
        tdValue.className = 'px-6 py-4 text-emerald-400 font-semibold';
        tdValue.textContent = data[key] || "-";
        tr.appendChild(tdValue);
    });

    tableBody.appendChild(tr);
}

function generateExcel() {
    if (!currentExtractedData || currentExtractedData.length === 0) return;

    // Ensure all target columns exist in the row even if empty
    const exportData = currentExtractedData.map(row => {
        const newRow = {};
        targetExcelColumns.forEach(col => {
            newRow[col] = row[col] || "";
        });
        return newRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData, { header: targetExcelColumns });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Invoice Data");

    XLSX.writeFile(workbook, "Parsed_Invoice_Gravity.xlsx");
}
