/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import { GoogleGenAI, Type } from "@google/genai";
import * as XLSX from 'xlsx';
import { uploadDataToGoogleSheet } from "./google-api-handler.ts";

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

// --- DOM Elements (Authentication) ---
const authView = document.getElementById('auth-view') as HTMLElement;
const appView = document.getElementById('app-view') as HTMLElement;
const loginForm = document.getElementById('login-form') as HTMLFormElement;
const usernameInput = document.getElementById('username') as HTMLInputElement;
const passwordInput = document.getElementById('password') as HTMLInputElement;
const loginError = document.getElementById('login-error') as HTMLElement;
const logoutBtn = document.getElementById('logout-btn') as HTMLElement;
const authJoinLink = document.getElementById('auth-join-link') as HTMLElement;

// --- DOM Elements (Audit Page) ---
const pageAudit = document.getElementById('page-audit') as HTMLElement;
const navAudit = document.getElementById('nav-audit') as HTMLElement;
const dropZone = document.getElementById('drop-zone') as HTMLElement;
const fileInput = document.getElementById('file-input') as HTMLInputElement;
const fileNameDisplay = document.getElementById('file-name') as HTMLElement;
const analyzeButton = document.getElementById('analyze-button') as HTMLButtonElement;
const uploadContainer = document.getElementById('upload-container') as HTMLElement;
const loadingContainer = document.getElementById('loading-container') as HTMLElement;
const resultsContainer = document.getElementById('results-container') as HTMLElement;
const loadingMessage = document.getElementById('loading-message') as HTMLElement;

// --- DOM Elements (Verification Page) ---
const pageVerify = document.getElementById('page-verify') as HTMLElement;
const navVerify = document.getElementById('nav-verify') as HTMLElement;
const verifyDropZone = document.getElementById('verify-drop-zone') as HTMLElement;
const verifyFileInput = document.getElementById('verify-file-input') as HTMLInputElement;
const verifyFileNameDisplay = document.getElementById('verify-file-name') as HTMLElement;
const verifyButton = document.getElementById('verify-button') as HTMLButtonElement;
const verifyUploadContainer = document.getElementById('verify-upload-container') as HTMLElement;
const verifyLoadingContainer = document.getElementById('verify-loading-container') as HTMLElement;
const verifyResultsContainer = document.getElementById('verify-results-container') as HTMLElement;
const verifyLoadingMessage = document.getElementById('verify-loading-message') as HTMLElement;

// --- DOM Elements (Static Pages) ---
const pageRequest = document.getElementById('page-request') as HTMLElement;
const navRequest = document.getElementById('nav-request') as HTMLElement;
const openRequestFormBtn = document.getElementById('open-request-form-btn') as HTMLElement;

// --- App State ---
let selectedFiles: File[] = [];
let selectedVerifyFiles: File[] = [];
const allPages = [pageAudit, pageVerify, pageRequest];
const allNavLinks = [navAudit, navVerify, navRequest];
let evaluatorUsers: {username: string, password: string}[] = [];
let clientUsers: {username: string, password: string}[] = [];

// --- Mappings and Schema for AI (Audit) ---
const RESULTS_MAP = {
  reportNumber: 'رقم تقرير',
  evaluatorName: 'اسم المقيم',
  membershipNumber: 'رقم العضوية',
  membershipCategory: 'فئة العضوية',
  evaluatorEmail: 'البريد الإلكتروني للمقيم',
  ownerName: 'اسم مالك العقار',
  ownerId: 'رقم هوية المالك',
  evaluationPurpose: 'الغرض من التقييم',
  reportType: 'نوع التقرير',
  evaluationMethod: 'أسلوب التقييم',
  propertyType: 'نوع العقار',
  ownershipType: 'نوع الملكية',
  propertyArea: 'مساحة العقار (م2)',
  pricePerMeter: 'قيمة المتر',
  marketValue: 'القيمة السوقية للعقار (ريال سعودي)',
  marketValueWritten: 'القيمة السوقية للعقار (كتابة)',
  region: 'المنطقة',
  propertyCity: 'مدينة العقار',
  propertyDistrict: 'الحي',
  planNumber: 'رقم المخطط',
  deedNumber: 'رقم الصك',
  inspectionDate: 'تاريخ المعاينة',
  valuationDate: 'تاريخ التقييم (بالميلادي)',
  reportIssueDate: 'إصدار التقرير',
  propertyCoordinates: 'إحداثيات العقار',
  restrictions: 'القيود',
};
const IMAGE_DATA_MAP = {
  deedImage: 'صورة الصك',
  membershipImage: 'صورة العضوية',
  propertyExteriorImage: 'صورة للعقار خارجية',
  propertyInteriorImage: 'صورة للعقار داخلية',
  siteAerialImage: 'صورة جوية للموقع',
  buildingPermitImage: 'صورة رخصة البناء',
  assignmentLetterImage: 'خطاب تكليف بالتقييم',
};
const SINGLE_PROPERTY_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    evaluatorName: { type: Type.STRING, description: "اسم المقيّم الكامل. استخرج الاسم فقط بدون أي ألقاب (مثل: أ./، م./) أو مناصب." },
    membershipNumber: { type: Type.STRING, description: 'رقم عضوية المقيّم' },
    membershipCategory: { type: Type.STRING, description: "فئة عضوية المقيّم. يجب أن تكون واحدة من: 'أساسي'، 'أساسي زميل'، 'شريك'، 'طالب منتسب'" },
    evaluatorEmail: { type: Type.STRING, description: 'البريد الإلكتروني للمقيّم' },
    propertyType: { type: Type.STRING, description: "نوع العقار. يجب أن يكون واحداً من: 'سكني'، 'تجاري'، 'زراعي'، 'سكني/تجاري'" },
    propertyArea: { type: Type.NUMBER, description: 'مساحة العقار بالأرقام الإنجليزية فقط (0-9). لا تقم بتضمين الوحدة "م2".' },
    propertyCity: { type: Type.STRING, description: 'مدينة العقار' },
    propertyDistrict: { type: Type.STRING, description: 'حي العقار' },
    planNumber: { type: Type.STRING, description: 'رقم المخطط' },
    propertyCoordinates: { type: Type.STRING, description: "ضع رابط خرائط جوجل (Google Maps) الكامل لموقع العقار. مثال: 'https://www.google.com/maps?q=24.7136,46.6753'" },
    ownerName: { type: Type.STRING, description: 'اسم مالك العقار. استخرج الاسم فقط بدون أي ألقاب أو مناصب.' },
    evaluationPurpose: { type: Type.STRING, description: "الغرض من التقييم. يجب أن يكون واحداً من: 'التمويل'، 'الشراء'، 'البيع'، 'التصفيه'، 'الدمج'، 'الاستحواذ'، 'الميراث'، 'حل النزاعات'، 'القرض العقاري'" },
    marketValue: { type: Type.NUMBER, description: 'القيمة السوقية الإجمالية للعقار بالأرقام الإنجليزية فقط (0-9). لا تقم بتضمين العملة.' },
    deedNumber: { type: Type.STRING, description: 'رقم صك الملكية' },
    restrictions: { type: Type.STRING, description: "أي قيود أو شروط مفروضة على العقار. إذا لم تكن هناك قيود، أرجع النص الحرفي 'لا توجد قيود'. وإلا، لخصها في جملة واحدة بحد أقصى 10 كلمات." },
    deedImage: { type: Type.STRING, description: "صف بإيجاز محتوى صورة صك الملكية. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'."},
    membershipImage: { type: Type.STRING, description: "صف بإيجاز محتوى صورة شهادة عضوية المقيّم. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'."},
    propertyExteriorImage: { type: Type.STRING, description: "صف بإيجاز محتوى الصورة الفوتوغرافية الخارجية للعقار. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'." },
    propertyInteriorImage: { type: Type.STRING, description: "صف بإيجاز محتوى الصورة الفوتوغرافية الداخلية للعقار. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'." },
    siteAerialImage: { type: Type.STRING, description: "صف بإيجاز محتوى الصورة الجوية لموقع العقار. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'." },
    ownershipType: { type: Type.STRING, description: "نوع ملكية العقار. يجب أن تكون الإجابة واحدة من: 'ملكية خاصة'، 'حكومي'." },
    region: { type: Type.STRING, description: "المنطقة الإدارية التي يقع فيها العقار. يجب أن تكون واحدة من مناطق المملكة الـ 13 (مثال: 'الرياض', 'مكة المكرمة')." },
    ownerId: { type: Type.STRING, description: "رقم هوية المالك (الهوية الوطنية أو السجل التجاري). يجب أن يكون أرقام إنجليزية فقط." },
    deedDate: { type: Type.STRING, description: "تاريخ إصدار صك الملكية. أعده بصيغة 'DD/MM/YYYY'." },
    reportType: { type: Type.STRING, description: "نوع تقرير التقييم. يجب أن تكون الإجابة واحدة من: 'تقرير مفصل'، 'ملخص تنفيذي'." },
    evaluationMethod: { type: Type.STRING, description: "الأسلوب المستخدم في تقييم العقار. يجب أن تكون الإجابة واحدة من: 'طريقة السوق'، 'طريقة الدخل'، 'طريقة التكلفة'." },
    inspectionDate: { type: Type.STRING, description: "تاريخ معاينة العقار من قبل المقيّم. أعده بصيغة 'DD/MM/YYYY'." },
    reportIssueDate: { type: Type.STRING, description: "تاريخ إصدار التقرير النهائي. أعده بصيغة 'DD/MM/YYYY'." },
    valuationDate: { type: Type.STRING, description: "تاريخ التقييم الفعلي للعقار. أعده بصيغة 'DD/MM/YYYY'." },
    pricePerMeter: { type: Type.NUMBER, description: "القيمة السوقية للمتر المربع بالأرقام الإنجليزية فقط (0-9)." },
    marketValueWritten: { type: Type.STRING, description: "القيمة السوقية الإجمالية للعقار مكتوبة بالأحرف كما هي في التقرير." },
    buildingPermitImage: { type: Type.STRING, description: "صف بإيجاز محتوى صورة رخصة البناء. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'." },
    assignmentLetterImage: { type: Type.STRING, description: "صف بإيجاز محتوى صورة خطاب التكليف بالتقييم. إذا لم يتم العثور عليها، أرجع النص الحرفي 'غير موجود'." },
  },
  required: [...Object.keys(RESULTS_MAP).filter(k => k !== 'reportNumber'), ...Object.keys(IMAGE_DATA_MAP)]
};
const RESPONSE_SCHEMA = { type: Type.ARRAY, items: SINGLE_PROPERTY_SCHEMA };

// --- Mappings and Schema for AI (Verification) ---
const PROFESSIONAL_REQUIREMENTS = {
    valuerIdentity: "هوية وصفة المقيم وتوقيعه",
    valuationDate: "تاريخ التقييم وتاريخ المعاينة",
    reportObjective: "الهدف من التقرير ونطاق العمل",
    clientIdentity: "هوية العميل والجهات المستخدمة للتقرير",
    propertyRights: "الحقوق العقارية موضوع التقييم",
    propertyDescription: "وصف تفصيلي للعقار ومكوناته",
    analysisMethod: "أسلوب التقييم المستخدم والأساس المنطقي",
    finalValue: "القيمة النهائية للعقار والتوصيات",
};
const REGULATORY_REQUIREMENTS = {
    complianceStatement: "إقرار الالتزام بمعايير التقييم الدولية",
    independenceStatement: "إقرار عدم وجود مصلحة شخصية للمقيم",
    taqeemStandards: "الإشارة إلى الالتزام بأنظمة الهيئة السعودية للمقيمين",
    highestBestUse: "تحليل أعلى وأفضل استخدام للعقار",
    marketAnalysis: "تحليل السوق العقاري ذي الصلة",
    deedInfo: "بيانات الصك أو الوثيقة الرسمية للعقار",
    propertyBoundaries: "حدود وأطوال العقار وموقعه",
    assumptions: "الافتراضات والظروف المقيدة للتقييم",
};
const verificationSchemaProperties = {};
[...Object.keys(PROFESSIONAL_REQUIREMENTS), ...Object.keys(REGULATORY_REQUIREMENTS)].forEach(key => {
    verificationSchemaProperties[key] = {
        type: Type.STRING,
        description: `هل هذا البند موجود؟ أجب بـ "موجود" أو "غير موجود" فقط.`
    };
});
const SINGLE_VERIFICATION_SCHEMA = {
    type: Type.OBJECT,
    properties: verificationSchemaProperties,
    required: Object.keys(verificationSchemaProperties)
};
const VERIFICATION_RESPONSE_SCHEMA = { type: Type.ARRAY, items: SINGLE_VERIFICATION_SCHEMA };


// --- Authentication Logic ---
function showAuthView() {
    authView.classList.remove('hidden');
    appView.classList.add('hidden');
}

function showAppView() {
    authView.classList.add('hidden');
    appView.classList.remove('hidden');
    const userRole = sessionStorage.getItem('userRole');

    if (userRole === 'evaluator') {
        navAudit.classList.remove('hidden');
        navVerify.classList.remove('hidden');
        navRequest.classList.remove('hidden');
        showPage(pageAudit); // Default to audit page
    } else if (userRole === 'client') {
        navAudit.classList.add('hidden');
        navVerify.classList.remove('hidden');
        navRequest.classList.add('hidden');
        showPage(pageVerify); // Default to verify page
    } else {
        // Fallback: If no role, log out
        handleLogout();
    }
}

function handleLogin(event: Event) {
    event.preventDefault();
    const username = usernameInput.value;
    const password = passwordInput.value;

    const isEvaluator = evaluatorUsers.find(user => user.username === username && user.password === password);
    const isClient = clientUsers.find(user => user.username === username && user.password === password);

    if (isEvaluator) {
        sessionStorage.setItem('isLoggedIn', 'true');
        sessionStorage.setItem('userRole', 'evaluator');
        showAppView();
        loginError.classList.add('hidden');
        loginForm.reset();
    } else if (isClient) {
        sessionStorage.setItem('isLoggedIn', 'true');
        sessionStorage.setItem('userRole', 'client');
        showAppView();
        loginError.classList.add('hidden');
        loginForm.reset();
    } else {
        loginError.classList.remove('hidden');
    }
}

function handleLogout() {
    sessionStorage.removeItem('isLoggedIn');
    sessionStorage.removeItem('userRole');
    showAuthView();
}

async function initAuth() {
    try {
        const [evaluatorsResponse, clientsResponse] = await Promise.all([
            fetch('./Evaluators.json'),
            fetch('./Clients.json')
        ]);

        if (!evaluatorsResponse.ok) {
            throw new Error(`HTTP error loading Evaluators.json! status: ${evaluatorsResponse.status}`);
        }
        if (!clientsResponse.ok) {
            throw new Error(`HTTP error loading Clients.json! status: ${clientsResponse.status}`);
        }

        evaluatorUsers = await evaluatorsResponse.json();
        clientUsers = await clientsResponse.json();
    } catch (error) {
        console.error("Could not load user credentials:", error);
        loginError.textContent = "خطأ في تحميل بيانات الاعتماد. لا يمكن تسجيل الدخول.";
        loginError.classList.remove('hidden');
        (document.getElementById('login-button') as HTMLButtonElement).disabled = true;
        return;
    }

    loginForm.addEventListener('submit', handleLogin);
    logoutBtn.addEventListener('click', (e) => {
      e.preventDefault();
      handleLogout();
    });
    authJoinLink.addEventListener('click', (e) => {
        e.preventDefault();
        openJoinForm();
    });

    if (sessionStorage.getItem('isLoggedIn') === 'true') {
        showAppView();
    } else {
        showAuthView();
    }
}

// --- Navigation Handling ---
function showPage(pageToShow: HTMLElement) {
    allPages.forEach(page => page.classList.toggle('hidden', page !== pageToShow));
    const linkId = `nav-${pageToShow.id.split('-')[1]}`;
    allNavLinks.forEach(link => link.classList.toggle('active', link.id === linkId));
    window.scrollTo(0, 0);
}
navAudit.addEventListener('click', (e) => { e.preventDefault(); showPage(pageAudit); });
navVerify.addEventListener('click', (e) => { e.preventDefault(); showPage(pageVerify); });
navRequest.addEventListener('click', (e) => { e.preventDefault(); showPage(pageRequest); });


// --- Form Handlers ---
function openJoinForm() {
    window.open('https://docs.google.com/forms/d/17xmxX50GNl2wpcL77krNnWvCT04rJnf9fkQo9oWj2QE/prefill', '_blank');
}
function openRequestForm() {
    window.open('https://docs.google.com/forms/d/e/1FAIpQLSfx3sQi4_ISGhePJBggfEyyHW5kWOcYngDEyBi4tTRFhdLhzg/viewform?usp=header', '_blank');
}
openRequestFormBtn.addEventListener('click', openRequestForm);


// --- General Helper Functions ---
function delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function generateReportNumber(fileIndex: number, reportIndex: number): string {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    return `VT-${year}${month}${day}-${hours}${minutes}${seconds}-${fileIndex + 1}-${reportIndex + 1}`;
}

function fileToGenerativePart(file: File) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
            const result = reader.result as string;
            const base64Data = result.split('base64,')[1];
            resolve({
                inlineData: {
                    data: base64Data,
                    mimeType: file.type,
                },
            });
        };
        reader.onerror = (err) => reject(err);
    });
}
function filterPDFFiles(files: FileList | null): File[] {
    if (!files) return [];
    return Array.from(files).filter((file: File) => {
        if (file.type !== 'application/pdf') {
            alert(`الملف '${file.name}' ليس من نوع PDF وسيتم تجاهله.`);
            return false;
        }
        return true;
    });
}


// --- File Handling and Analysis (AUDIT) ---
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.classList.remove('dragover'); handleFiles(e.dataTransfer?.files); });
fileInput.addEventListener('change', (e: Event) => handleFiles((e.target as HTMLInputElement).files));
analyzeButton.addEventListener('click', () => { if (selectedFiles.length > 0) startAnalysis(selectedFiles); });

function handleFiles(files: FileList | null) {
    selectedFiles = filterPDFFiles(files);
    if (selectedFiles.length > 0) {
        fileNameDisplay.textContent = `${selectedFiles.length} ملفات تم اختيارها`;
        analyzeButton.disabled = false;
    } else {
        fileNameDisplay.textContent = '';
        analyzeButton.disabled = true;
    }
    resultsContainer.innerHTML = '';
    resultsContainer.classList.add('hidden');
}

type PageType = 'audit' | 'verify';
function setUIState(state: 'upload' | 'loading' | 'results', page: PageType) {
    const containers = {
        audit: { upload: uploadContainer, loading: loadingContainer, results: resultsContainer },
        verify: { upload: verifyUploadContainer, loading: verifyLoadingContainer, results: verifyResultsContainer }
    };
    const target = containers[page];
    target.upload.classList.toggle('hidden', state !== 'upload');
    target.loading.classList.toggle('hidden', state !== 'loading');
    target.results.classList.toggle('hidden', state !== 'results');
}

interface AnalysisResult {
    file: File;
    reports: any[];
    error?: string;
}

async function startAnalysis(files: File[]) {
    setUIState('loading', 'audit');
    resultsContainer.innerHTML = '';
    const allAnalysisResults: AnalysisResult[] = [];

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        loadingMessage.textContent = `جاري تحليل الملف ${i + 1} من ${files.length}: ${file.name}...`;
        
        // Always wait before an API call to avoid rate limiting.
        await delay(2000); // 2-second delay

        try {
            const result = await analyzeSinglePdf(file);

            if (result.reports && Array.isArray(result.reports)) {
                result.reports.forEach((report, j) => {
                    report.reportNumber = generateReportNumber(i, j);
                });
            }

            allAnalysisResults.push({ file, ...result });
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            console.error(`Error analyzing file ${file.name}:`, error);
            allAnalysisResults.push({ file, reports: [], error: errorMessage });
        }
    }

    renderAllResults(allAnalysisResults);
    setUIState('results', 'audit');
}

async function analyzeSinglePdf(file: File) {
    const filePart = await fileToGenerativePart(file);
    const promptText = `أنت خبير في تحليل تقارير التقييم العقاري باللغة العربية. مهمتك هي تحليل ملف PDF المرفق بدقة.
قاعدة أساسية: قبل أي تحليل، تحقق إذا كان الملف المرفق هو "تقرير تقييم عقاري". إذا لم يكن كذلك، أرجع مصفوفة فارغة [] مباشرة.
إذا كان تقريراً صالحاً، فاستخرج جميع تقارير العقارات الموجودة فيه. يجب أن يكون تحليلك متسقًا وحتميًا؛ لنفس الملف، يجب أن تُرجع دائمًا نفس النتيجة تمامًا. أرجع النتائج كمصفوفة (array) من كائنات JSON، حيث يمثل كل كائن تقريراً واحداً.
    
    التزم بالقواعد التالية بدقة:
    - **الأسماء:** بالنسبة لأسماء الأشخاص (مثل المقيّم أو المالك)، استخرج الاسم الكامل فقط بدون أي ألقاب (مثل السيد/، المهندس/) أو مناصب وظيفية.
    - **الأرقام:** استخدم الأرقام العربية الغربية (0-9) لجميع القيم الرقمية.
    - **مساحة العقار والقيمة السوقية:** استخرج القيمة الرقمية فقط، بدون أي وحدات أو عملات.
    - **إحداثيات العقار:** قدم رابط خرائط جوجل (Google Maps) كامل للموقع.
    - **نوع العقار:** يجب أن يكون واحداً من: 'سكني'، 'تجاري'، 'زراعي'، 'سكني/تجاري'.
    - **الغرض من التقييم:** يجب أن يكون واحداً من: 'التمويل'، 'الشراء'، 'البيع'، 'التصفيه'، 'الدمج'، 'الاستحواذ'، 'الميراث'، 'حل النزاعات'، 'القرض العقاري'.
    - **فئة العضوية:** يجب أن تكون واحدة من: 'أساسي'، 'أساسي زميل'، 'شريك'، 'طالب منتسب'.
    - **القيود:** إذا لم تكن هناك أي قيود مذكورة، أو ذكر النص صراحة عدم وجودها (مثل "لا يوجد"، "لا قيود")، فأرجع النص الحرفي "لا توجد قيود". وإلا، لخص القيود في جملة واحدة بحد أقصى 10 كلمات.
    - **القيم غير الموجودة:** إذا كانت معلومة غير موجودة، أرجع قيمة 'غير موجود' حرفيًا.
    - **الصور:** صف محتوى كل صورة بإيجاز. إذا لم تجد صورة، أرجع 'غير موجود'.`;

    const aiResponse = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: [
            {
                parts: [
                    { text: promptText },
                    filePart
                ]
            }
        ],
        config: {
            responseMimeType: "application/json",
            responseSchema: RESPONSE_SCHEMA,
            temperature: 0,
        },
    });
    const resultJson = JSON.parse(aiResponse.text.trim());
    if (!Array.isArray(resultJson)) throw new Error("AI response is not an array.");
    return { reports: resultJson };
}


// --- UI Rendering (AUDIT) ---
function renderAllResults(results: AnalysisResult[]) {
    resultsContainer.innerHTML = '';
    const allReports = results.flatMap(r => r.reports);
    const mainTitle = document.createElement('h2');
    mainTitle.textContent = `نتائج التحليل لـ ${results.length} ملفات`;
    resultsContainer.appendChild(mainTitle);

    if (allReports.length > 0) {
        const actionBar = document.createElement('div');
        actionBar.className = 'action-bar';

        // Download Excel Button
        const downloadButton = document.createElement('button');
        downloadButton.textContent = `تنزيل كل التقارير (${allReports.length}) كـ Excel`;
        downloadButton.className = 'btn-primary';
        downloadButton.onclick = () => downloadReportAsXLSX(allReports);
        actionBar.appendChild(downloadButton);

        // Export to Google Sheets Button
        const sheetsButton = document.createElement('button');
        sheetsButton.innerHTML = `تصدير إلى<br>Google Sheets`;
        sheetsButton.className = 'btn-primary';
        sheetsButton.onclick = () => {
            const statusElement = document.getElementById('google-sheets-status');
            if (!statusElement) return;

            statusElement.textContent = 'جاري التصدير إلى Google Sheets...';
            statusElement.className = 'google-sheets-status progress';
            sheetsButton.disabled = true;

            const allHeadersMap = { ...RESULTS_MAP, ...IMAGE_DATA_MAP };
            const requirementKeys = Object.keys(allHeadersMap);
            const totalRequirements = requirementKeys.length;

            allReports.forEach(report => {
                if (report.requirementsMet === undefined) {
                    const metCount = requirementKeys.reduce((acc, key) => {
                        const value = report[key];
                        return (value && value !== 'غير موجود') ? acc + 1 : acc;
                    }, 0);
                    report.requirementsMet = `${metCount}/${totalRequirements}`;
                }
            });

            const headersForSheet = { ...allHeadersMap, requirementsMet: 'إتمام المتطلبات' };

            uploadDataToGoogleSheet(allReports, headersForSheet).then(message => {
                statusElement.textContent = message;
                if (message.startsWith('فشل')) {
                    statusElement.className = 'google-sheets-status error';
                } else {
                    statusElement.className = 'google-sheets-status success';
                }
            }).finally(() => {
                sheetsButton.disabled = false;
            });
        };
        actionBar.appendChild(sheetsButton);

        resultsContainer.appendChild(actionBar);

        const sheetsStatus = document.createElement('p');
        sheetsStatus.id = 'google-sheets-status';
        sheetsStatus.className = 'google-sheets-status hidden';
        resultsContainer.appendChild(sheetsStatus);
    }
    results.forEach(result => {
        const fileResultContainer = document.createElement('div');
        fileResultContainer.className = 'file-result-container';
        const fileTitle = document.createElement('h3');
        fileTitle.className = 'file-result-title';
        fileTitle.textContent = `نتائج الملف: ${result.file.name}`;
        fileResultContainer.appendChild(fileTitle);
        if (result.error) {
            const errorP = document.createElement('p');
            errorP.className = 'error-message';
            errorP.textContent = `فشل تحليل الملف: ${result.error}`;
            fileResultContainer.appendChild(errorP);
        } else {
            renderReportsForFile(fileResultContainer, result.reports);
        }
        resultsContainer.appendChild(fileResultContainer);
    });
}

function renderReportsForFile(container: HTMLElement, reports: any[]) {
    if (reports.length === 0) {
        const noReportsP = document.createElement('p');
        noReportsP.className = 'error-message';
        noReportsP.textContent = 'الملف الذي تم رفعه ليس تقرير تقييم عقاري صالح. الرجاء رفع ملف تقرير تقييم عقاري فقط.';
        container.appendChild(noReportsP);
        return;
    }

    const allHeadersMap = { ...RESULTS_MAP, ...IMAGE_DATA_MAP };
    const requirementKeys = Object.keys(allHeadersMap);
    const totalRequirements = requirementKeys.length;

    reports.forEach((reportData, index) => {
        const accordion = document.createElement('details');
        accordion.className = 'report-accordion';
        if (index === 0) accordion.open = true;

        const summary = document.createElement('summary');
        summary.textContent = `تقرير ${index + 1}: ${reportData.propertyType || 'عقار'} في ${reportData.propertyCity || 'مدينة غير محددة'}`;
        accordion.appendChild(summary);

        const content = document.createElement('div');
        content.className = 'report-content';

        const metCount = requirementKeys.reduce((acc, key) => {
            const value = reportData[key];
            return (value && String(value).trim() !== 'غير موجود') ? acc + 1 : acc;
        }, 0);
        const percentage = totalRequirements > 0 ? (metCount / totalRequirements) * 100 : 0;

        const requirementsSummary = document.createElement('div');
        requirementsSummary.className = 'requirements-summary';
        requirementsSummary.innerHTML = `
            <div class="requirements-label">
                <strong>إتمام المتطلبات</strong>
                <span>${metCount} / ${totalRequirements}</span>
            </div>
            <div class="progress-bar">
                <div class="progress-bar-fill" style="width: ${percentage}%;"></div>
            </div>
        `;
        content.appendChild(requirementsSummary);

        const infoTitle = document.createElement('h4');
        infoTitle.textContent = 'نتائج التدقيق';
        content.appendChild(infoTitle);

        const mainGrid = document.createElement('div');
        mainGrid.className = 'results-grid';
        
        for (const key in RESULTS_MAP) {
            if (Object.prototype.hasOwnProperty.call(RESULTS_MAP, key)) {
                const item = document.createElement('div');
                item.className = 'result-item';
                const value = reportData[key] || 'غير موجود';
                
                let valueText = value;
                if (typeof value === 'number') {
                    valueText = value.toLocaleString('en-US');
                }

                let valueContent;
                if (value === 'غير موجود') {
                    valueContent = `<span class="value-not-found">${valueText}</span>`;
                } else if (key === 'propertyCoordinates' && typeof value === 'string' && value.startsWith('http')) {
                     valueContent = `<a href="${value}" target="_blank" rel="noopener noreferrer">عرض على الخريطة</a>`;
                } else {
                    valueContent = `<div>${valueText}</div>`;
                }

                if (key === 'membershipNumber' && value !== 'غير موجود') {
                    valueContent += `
                        <div class="verification-links">
                            <a href="https://taqeem.gov.sa/en/authority-members?sector=real_estate&search[name_or_membership]=${value}" target="_blank" title="التحقق كعضو فرد">تحقق (فرد)</a>
                            <a href="https://taqeem.gov.sa/en/facilities?sector=real_estate&search[name_or_membership]=${value}" target="_blank" title="التحقق كمنشأة">تحقق (منشأة)</a>
                        </div>
                    `;
                }

                item.innerHTML = `
                    <div class="label">${RESULTS_MAP[key]}</div>
                    <div class="value">${valueContent}</div>
                `;
                mainGrid.appendChild(item);
            }
        }
        content.appendChild(mainGrid);

        const imagesTitle = document.createElement('h4');
        imagesTitle.textContent = 'المستندات المصورة';
        content.appendChild(imagesTitle);

        const imageGrid = document.createElement('div');
        imageGrid.className = 'document-image-grid';
        
        for (const key in IMAGE_DATA_MAP) {
            if (Object.prototype.hasOwnProperty.call(IMAGE_DATA_MAP, key)) {
                const value = reportData[key]; // This will be the description or 'غير موجود'
                const item = document.createElement('div');
                item.className = 'document-image-item';

                const label = document.createElement('div');
                label.className = 'label';
                label.textContent = IMAGE_DATA_MAP[key];
                item.appendChild(label);
                
                const statusDiv = document.createElement('div');

                if (value && value.trim() !== 'غير موجود') {
                    statusDiv.className = 'image-status-found';
                    
                    const statusText = document.createElement('span');
                    statusText.textContent = 'موجود';
                    statusDiv.appendChild(statusText);

                    const descriptionP = document.createElement('p');
                    descriptionP.className = 'image-description';
                    // Adding quotes for clarity
                    descriptionP.textContent = `"${value}"`;
                    statusDiv.appendChild(descriptionP);
                } else {
                    statusDiv.className = 'image-status-not-found';
                    statusDiv.textContent = 'غير موجود';
                }
                item.appendChild(statusDiv);
                
                imageGrid.appendChild(item);
            }
        }
        content.appendChild(imageGrid);

        accordion.appendChild(content);
        container.appendChild(accordion);
    });
}

// --- Data Export (AUDIT) ---
function downloadReportAsXLSX(reports: any[]) {
    if (reports.length === 0) return;

    const allHeadersMap = { ...RESULTS_MAP, ...IMAGE_DATA_MAP };
    const requirementKeys = Object.keys(allHeadersMap);
    const totalRequirements = requirementKeys.length;
    const excelHeadersWithCount = { ...allHeadersMap, requirementsMet: 'إتمام المتطلبات' };
    const orderedKeys = Object.keys(excelHeadersWithCount);
    
    const headerRow = orderedKeys.map(key => excelHeadersWithCount[key]);
    const sheetData = [headerRow];

    const greenFill = { fgColor: { rgb: "C6EFCE" } };
    const redFill = { fgColor: { rgb: "FFC7CE" } };

    reports.forEach(report => {
        const metCount = requirementKeys.reduce((acc, key) => {
            const value = report[key];
            return (value && value !== 'غير موجود') ? acc + 1 : acc;
        }, 0);

        const row = orderedKeys.map(key => {
            if (key === 'requirementsMet') {
                return `${metCount}/${totalRequirements}`;
            }
            if (key === 'propertyCoordinates') {
                const url = report[key];
                if (url && typeof url === 'string' && url.startsWith('http')) {
                    return { v: "رابط الخريطة", l: { Target: url, Tooltip: url } };
                }
                return url || 'غير موجود';
            }
            // For marketValue and propertyArea, ensure they are numbers
            if ((key === 'marketValue' || key === 'propertyArea') && report[key]) {
                const num = Number(report[key]);
                return isNaN(num) ? report[key] : num;
            }
            return report[key] || 'غير موجود';
        });
        sheetData.push(row);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Apply styles to "requirementsMet" column
    const requirementsColIndex = orderedKeys.indexOf('requirementsMet');
    if (requirementsColIndex > -1) {
        reports.forEach((report, index) => {
            const metCount = requirementKeys.reduce((acc, key) => {
                const value = report[key];
                return (value && value !== 'غير موجود') ? acc + 1 : acc;
            }, 0);
            const isComplete = metCount === totalRequirements;
            const cellAddress = XLSX.utils.encode_cell({ r: index + 1, c: requirementsColIndex });
            if (worksheet[cellAddress]) {
                worksheet[cellAddress].s = { fill: isComplete ? greenFill : redFill };
            }
        });
    }

    // Set column widths
    const colWidths = headerRow.map((header, i) => {
        let maxLength = header.length;
        for (let j = 1; j < sheetData.length; j++) {
            const cellValue = sheetData[j][i];
            const cellText = cellValue?.v || cellValue; // Handle hyperlink objects
            const cellLength = cellText ? String(cellText).length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        }
        return { wch: maxLength + 2 };
    });
    worksheet['!cols'] = colWidths;

    if (!worksheet['!props']) worksheet['!props'] = {};
    worksheet['!props'].RTL = true;
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'تقارير التدقيق');
    XLSX.writeFile(workbook, "valuation_audit_report.xlsx", { bookType: 'xlsx', type: 'binary' });
}


// --- File Handling and Analysis (VERIFICATION) ---
verifyDropZone.addEventListener('click', () => verifyFileInput.click());
verifyDropZone.addEventListener('dragover', (e) => { e.preventDefault(); verifyDropZone.classList.add('dragover'); });
verifyDropZone.addEventListener('dragleave', () => verifyDropZone.classList.remove('dragover'));
verifyDropZone.addEventListener('drop', (e) => { e.preventDefault(); verifyDropZone.classList.remove('dragover'); handleVerifyFiles(e.dataTransfer?.files); });
verifyFileInput.addEventListener('change', (e: Event) => handleVerifyFiles((e.target as HTMLInputElement).files));
verifyButton.addEventListener('click', () => { if (selectedVerifyFiles.length > 0) startVerificationAnalysis(selectedVerifyFiles); });

function handleVerifyFiles(files: FileList | null) {
    selectedVerifyFiles = filterPDFFiles(files);
    if (selectedVerifyFiles.length > 0) {
        verifyFileNameDisplay.textContent = `${selectedVerifyFiles.length} ملفات تم اختيارها`;
        verifyButton.disabled = false;
    } else {
        verifyFileNameDisplay.textContent = '';
        verifyButton.disabled = true;
    }
    verifyResultsContainer.innerHTML = '';
    verifyResultsContainer.classList.add('hidden');
}

interface VerificationResult {
    file: File;
    reports: any[];
    error?: string;
}

async function startVerificationAnalysis(files: File[]) {
    setUIState('loading', 'verify');
    verifyResultsContainer.innerHTML = '';
    const allVerificationResults: VerificationResult[] = [];

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        verifyLoadingMessage.textContent = `جاري التحقق من الملف ${i + 1} من ${files.length}: ${file.name}...`;
        
        // Always wait before an API call to avoid rate limiting.
        await delay(2000); // 2-second delay

        try {
            const result = await analyzeSinglePdfForVerification(file);
            allVerificationResults.push({ file, ...result });
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            console.error(`Error verifying file ${file.name}:`, error);
            allVerificationResults.push({ file, reports: [], error: errorMessage });
        }
    }

    renderAllVerificationResults(allVerificationResults);
    setUIState('results', 'verify');
}

async function analyzeSinglePdfForVerification(file: File) {
    const filePart = await fileToGenerativePart(file);
    const promptText = `أنت خبير محترف في تدقيق تقارير التقييم العقاري في المملكة العربية السعودية ومتخصص في معايير الهيئة السعودية للمقيمين المعتمدين (تقييم).
قاعدة أساسية: قبل أي تحليل، تحقق إذا كان الملف المرفق هو "تقرير تقييم عقاري". إذا لم يكن كذلك، أرجع مصفوفة فارغة [] مباشرة.
إذا كان تقريراً صالحاً، فمهمتك هي تحليل ملف PDF المرفق والتحقق من وجود المتطلبات المهنية والنظامية. لكل متطلب في الهيكل المطلوب، تحقق من وجوده في التقرير وأجب بـ "موجود" أو "غير موجود" فقط. لا تقدم أي شروحات إضافية. كن دقيقاً جداً. أرجع النتائج كمصفوفة (array) من كائنات JSON، حيث يمثل كل كائن تقريراً واحداً تم العثور عليه في الملف.`;

    const aiResponse = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: [{ parts: [{ text: promptText }, filePart] }],
        config: {
            responseMimeType: "application/json",
            responseSchema: VERIFICATION_RESPONSE_SCHEMA,
            temperature: 0,
        },
    });

    const resultJson = JSON.parse(aiResponse.text.trim());
    if (!Array.isArray(resultJson)) throw new Error("AI response is not an array.");
    return { reports: resultJson };
}

// --- UI Rendering (VERIFICATION) ---
function renderAllVerificationResults(results: VerificationResult[]) {
    verifyResultsContainer.innerHTML = '';
    const mainTitle = document.createElement('h2');
    mainTitle.textContent = `نتائج التحقق لـ ${results.length} ملفات`;
    verifyResultsContainer.appendChild(mainTitle);
    
    results.forEach(result => {
        const fileResultContainer = document.createElement('div');
        fileResultContainer.className = 'file-result-container';
        
        const fileTitle = document.createElement('h3');
        fileTitle.className = 'file-result-title';
        fileTitle.textContent = `نتائج التحقق للملف: ${result.file.name}`;
        fileResultContainer.appendChild(fileTitle);
        
        if (result.error) {
            const errorP = document.createElement('p');
            errorP.className = 'error-message';
            errorP.textContent = `فشل التحقق من الملف: ${result.error}`;
            fileResultContainer.appendChild(errorP);
        } else {
            renderVerificationReportsForFile(fileResultContainer, result.reports);
        }
        
        verifyResultsContainer.appendChild(fileResultContainer);
    });
}

function renderVerificationReportsForFile(container: HTMLElement, reports: any[]) {
    if (reports.length === 0) {
        const noReportsP = document.createElement('p');
        noReportsP.className = 'error-message';
        noReportsP.textContent = 'الملف الذي تم رفعه ليس تقرير تقييم عقاري صالح. الرجاء رفع ملف تقرير تقييم عقاري فقط.';
        container.appendChild(noReportsP);
        return;
    }

    const presentIcon = `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M20 6L9 17L4 12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
    const notPresentIcon = `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M18 6L6 18" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M6 6L18 18" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

    reports.forEach((reportData, index) => {
        const accordion = document.createElement('details');
        accordion.className = 'report-accordion';
        if (index === 0) accordion.open = true;

        const summary = document.createElement('summary');
        summary.textContent = `تقرير التحقق ${index + 1}`;
        accordion.appendChild(summary);

        const content = document.createElement('div');
        content.className = 'report-content';

        // Professional Requirements
        const profTitle = document.createElement('h4');
        profTitle.className = 'verification-section-title';
        profTitle.textContent = 'المتطلبات المهنية';
        content.appendChild(profTitle);
        const profList = document.createElement('ul');
        profList.className = 'verification-checklist';
        for(const key in PROFESSIONAL_REQUIREMENTS) {
            const li = document.createElement('li');
            li.className = 'verification-item';
            const status = reportData[key] === 'موجود' ? 'present' : 'not-present';
            li.innerHTML = `
                <span class="verification-text">${PROFESSIONAL_REQUIREMENTS[key]}</span>
                <span class="verification-status status-${status}">${status === 'present' ? presentIcon : notPresentIcon}</span>
            `;
            profList.appendChild(li);
        }
        content.appendChild(profList);
        
        // Regulatory Requirements
        const regTitle = document.createElement('h4');
        regTitle.className = 'verification-section-title';
        regTitle.textContent = 'المتطلبات النظامية';
        content.appendChild(regTitle);
        const regList = document.createElement('ul');
        regList.className = 'verification-checklist';
        for(const key in REGULATORY_REQUIREMENTS) {
            const li = document.createElement('li');
            li.className = 'verification-item';
            const status = reportData[key] === 'موجود' ? 'present' : 'not-present';
            li.innerHTML = `
                <span class="verification-text">${REGULATORY_REQUIREMENTS[key]}</span>
                <span class="verification-status status-${status}">${status === 'present' ? presentIcon : notPresentIcon}</span>
            `;
            regList.appendChild(li);
        }
        content.appendChild(regList);

        accordion.appendChild(content);
        container.appendChild(accordion);
    });
}

// --- Data Export (VERIFICATION) ---
function downloadVerificationAsXLSX(reports: any[]) {
    if (reports.length === 0) return;

    const allHeadersMap = { ...PROFESSIONAL_REQUIREMENTS, ...REGULATORY_REQUIREMENTS };
    const dataForSheet = reports.map(report => {
        const row = {
            'اسم الملف': report.fileName,
            'رقم التقرير': report.reportIndex,
        };
        for(const key in allHeadersMap) {
            const header = allHeadersMap[key];
            row[header] = report[key] || 'غير موجود';
        }
        return row;
    });

    const headers = ['اسم الملف', 'رقم التقرير', ...Object.values(allHeadersMap)];
    const worksheet = XLSX.utils.json_to_sheet(dataForSheet, { header: headers });

    const colWidths = headers.map(header => ({ wch: Math.max(header.length, 20) }));
    worksheet['!cols'] = colWidths;
    if (!worksheet['!props']) worksheet['!props'] = {};
    worksheet['!props'].RTL = true;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'نتائج التحقق');
    XLSX.writeFile(workbook, "verification_checklist_report.xlsx", { bookType: 'xlsx', type: 'binary' });
}

// --- Initialize App ---
document.addEventListener('DOMContentLoaded', initAuth);