import * as React from 'react';
import styles from './GmApprovalDashboard.module.scss';
import { IGmApprovalDashboardProps } from './IGmApprovalDashboardProps';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export type LanguageCode = 'en' | 'ar';
export type PdfSourceMode = 'upload' | 'create';
export type PdfRole = 'officeManager' | 'hod' | 'secretary';

export interface IRequestItem {
  Id: number;
  Title: string;
  RequestNo: string;
  RequestDate: string;
  Department: string;
  RequestDetails: string;
  Status: string;
  OfficeManagerComment: string;
  HODComment: string;
  GMComment: string;
  PDFFileUrl: any;
  PDFFileJson: string;
  SignedPDFUrl: any;
  SharedFolderPath: string;
  GMApprovalDetected: boolean;
  GMApprovalConfirmed: boolean;
  Attachments: boolean;
  AttachmentFiles: any;
  Author: any;
}

export interface IGmApprovalDashboardState {
  language: LanguageCode;
  loading: boolean;
  saving: boolean;
  error: string;
  message: string;

  currentUserId: number;
  currentUserTitle: string;
  groups: string[];

  isSecretary: boolean;
  isOfficeManager: boolean;
  isHod: boolean;
  isServiceAccount: boolean;

  requests: IRequestItem[];
  selectedRequest: IRequestItem;
  showNewRequestForm: boolean;

  newTitle: string;
  newDepartment: string;
  newDetails: string;
  newPdfFile: any;
  newPdfMode: PdfSourceMode;
  showNewRequestEditor: boolean;
  newRequestStatus: string;

  officeManagerComment: string;
  hodComment: string;
  gmComment: string;
  secretaryComment: string;
  documentSourceMode: PdfSourceMode;
  rolePdfFile: any;
  activeDocumentRole: string;
  showOfficeManagerEditor: boolean;
}

interface IPdfUploadResult {
  role: string;
  documentType: string;
  fileName: string;
  originalFileName: string;
  contentType: string;
  size: number;
  serverRelativeUrl: string;
  url: string;
  attachmentUrl: string;
}

interface IPdfDocumentRecord {
  role: string;
  documentType: string;
  fileName: string;
  originalFileName: string;
  contentType: string;
  size: number;
  uploadedAt: string;
  libraryUrl: string;
  serverRelativeUrl: string;
  attachmentUrl: string;
}

interface ITranslatedText {
  en: string;
  ar: string;
}

interface IWorkflowComment {
  id: string;
  createdAt: string;
  authorId: number;
  authorName: string;
  role: string;
  kind: string;
  fromStatus: string;
  toStatus: string;
  text: string;
  legacy: boolean;
}

export default class GmApprovalDashboard extends React.Component<IGmApprovalDashboardProps, IGmApprovalDashboardState> {

  private buildVersion: string = '0.0.5';
  private gmSiteUrl: string = 'http://spse26h/GM';
  private requestsListName: string = 'GM Requests';
  private documentsLibraryName: string = 'GM Approval Documents';
  private sharedFolderRootPath: string = '\\\\SPSE26H\\GMApprovalShare';
  private pdfJsonFieldName: string = 'PDFFileJson';
  private odataJsonHeader: string = 'application/json;odata=verbose';
  private odataVersion: string = '3.0';
  private requestsListItemEntityTypeFullName: string = '';
  private documentsLibraryListItemEntityTypeFullName: string = '';
  private pdfFileInput: HTMLInputElement;
  private rolePdfFileInput: HTMLInputElement;
  private newRequestEditor: HTMLDivElement;
  private officeManagerEditor: HTMLDivElement;
  private newRequestDocumentHtml: string = '';
  private officeManagerDocumentHtml: string = '';
  private selectedRequestRefreshTimer: number = 0;
  private besafeLogoUrl: string = require('../assets/besafe-logo.svg');
  private translations: { [key: string]: ITranslatedText } = {
    appTitle: { en: 'General Manager Approval Dashboard', ar: 'لوحة موافقات المدير العام' },
    english: { en: 'English', ar: 'English' },
    arabic: { en: 'Arabic', ar: 'العربية' },
    currentUser: { en: 'Current user:', ar: 'المستخدم الحالي:' },
    groups: { en: 'Groups:', ar: 'المجموعات:' },
    groupSecretaries: { en: 'GM Secretaries', ar: 'سكرتارية المدير العام' },
    groupOfficeManagers: { en: 'GM Office Managers', ar: 'مديرو مكتب المدير العام' },
    groupHods: { en: 'GM HODs', ar: 'رؤساء الأقسام' },
    groupServiceAccounts: { en: 'GM Service Accounts', ar: 'حسابات الخدمة' },
    loadingRequests: { en: 'Loading requests...', ar: 'جاري تحميل الطلبات...' },
    noAccess: { en: 'You are not a member of GM Secretaries, GM Office Managers, or GM HODs. Please contact SharePoint administrator.', ar: 'أنت لست عضوا في مجموعات سكرتارية المدير العام أو مديري مكتب المدير العام أو رؤساء الأقسام. يرجى التواصل مع مسؤول SharePoint.' },
    footer: { en: 'Besafe SharePoint Team', ar: 'فريق Besafe SharePoint' },
    appSubtitle: { en: 'A clear, secure workspace for request review and approval', ar: 'مساحة عمل واضحة وآمنة لمراجعة الطلبات واعتمادها' },

    newRequest: { en: 'New Request', ar: 'طلب جديد' },
    newRequestEyebrow: { en: 'Request workspace', ar: 'مساحة عمل الطلب' },
    newRequestDescription: { en: 'Add the request information and supporting PDF, then choose the next approval step.', ar: 'أضف بيانات الطلب وملف PDF الداعم، ثم اختر خطوة الاعتماد التالية.' },
    title: { en: 'Title', ar: 'العنوان' },
    department: { en: 'Department', ar: 'القسم' },
    selectDepartment: { en: 'Select department', ar: 'اختر القسم' },
    requestDetails: { en: 'Request Details', ar: 'تفاصيل الطلب' },
    pdfAttachment: { en: 'PDF Attachment', ar: 'مرفق PDF' },
    pdfHelp: { en: 'Stored in GM Approval Documents and the request Attachments', ar: 'سيتم حفظه في مكتبة مستندات موافقات المدير العام ومرفقات الطلب' },
    choosePdf: { en: 'Choose supporting PDF', ar: 'اختر ملف PDF الداعم' },
    replacePdf: { en: 'Replace PDF', ar: 'استبدال ملف PDF' },
    pdfFolderHelp: { en: 'The file will be stored inside the folder named with the request reference.', ar: 'سيتم حفظ الملف داخل مجلد يحمل رقم مرجع الطلب.' },
    submitToOfficeManager: { en: 'Submit to Office Manager', ar: 'إرسال إلى مدير المكتب' },
    submitting: { en: 'Submitting...', ar: 'جاري الإرسال...' },
    cancel: { en: 'Cancel', ar: 'إلغاء' },
    routeTo: { en: 'Route To', ar: 'توجيه إلى' },
    routeToOfficeManager: { en: 'Office Manager', ar: 'مدير المكتب' },
    routeToHod: { en: 'HOD', ar: 'رئيس القسم' },
    routeToGmSignature: { en: 'GM Signature', ar: 'توقيع المدير العام' },

    myRequests: { en: 'My Requests', ar: 'طلباتي' },
    officeManagerQueue: { en: 'Office Manager Queue', ar: 'قائمة مدير المكتب' },
    hodQueue: { en: 'HOD Queue', ar: 'قائمة رئيس القسم' },
    secretaryQueue: { en: 'Secretary Queue', ar: 'قائمة السكرتارية' },
    noRequestsFound: { en: 'No requests found.', ar: 'لا توجد طلبات.' },
    requestNo: { en: 'Request No', ar: 'رقم الطلب' },
    status: { en: 'Status', ar: 'الحالة' },
    createdBy: { en: 'Created By', ar: 'أنشئ بواسطة' },
    pdf: { en: 'PDF', ar: 'PDF' },
    open: { en: 'Open', ar: 'فتح' },

    requestDetailsTitle: { en: 'Request Details', ar: 'تفاصيل الطلب' },
    requestOverview: { en: 'Request overview', ar: 'نظرة عامة على الطلب' },
    requestOverviewDescription: { en: 'Review the request, its documents, and the available workflow actions.', ar: 'راجع الطلب ومستنداته وإجراءات سير العمل المتاحة.' },
    id: { en: 'ID:', ar: 'المعرف:' },
    requestNoLabel: { en: 'Request No:', ar: 'رقم الطلب:' },
    titleLabel: { en: 'Title:', ar: 'العنوان:' },
    departmentLabel: { en: 'Department:', ar: 'القسم:' },
    statusLabel: { en: 'Status:', ar: 'الحالة:' },
    createdByLabel: { en: 'Created By:', ar: 'أنشئ بواسطة:' },
    gmSharedFolderPath: { en: 'GM Shared Folder Path:', ar: 'مسار مجلد المدير العام المشترك:' },
    attachedPdf: { en: 'Attached PDF:', ar: 'ملف PDF المرفق:' },
    openPdf: { en: 'Open PDF', ar: 'فتح PDF' },
    signedPdf: { en: 'Signed PDF:', ar: 'ملف PDF الموقّع:' },
    openSignedPdf: { en: 'Open Signed PDF', ar: 'فتح ملف PDF الموقّع' },
    documents: { en: 'Request documents', ar: 'مستندات الطلب' },
    requestPdf: { en: 'Supporting PDF', ar: 'ملف PDF الداعم' },
    officeManagerPdf: { en: 'Office Manager PDF', ar: 'ملف PDF الخاص بمدير المكتب' },
    workflowPdf: { en: 'Workflow PDF', ar: 'ملف سير العمل' },
    originalRequestPdf: { en: 'Original request PDF', ar: 'ملف الطلب الأصلي' },
    hodPdf: { en: 'HOD PDF', ar: 'ملف رئيس القسم' },
    secretaryPdf: { en: 'Secretary PDF', ar: 'ملف السكرتارية' },
    signedPdfDocument: { en: 'Signed PDF', ar: 'ملف PDF الموقّع' },
    storedInRequestFolder: { en: 'Stored in the request folder', ar: 'محفوظ داخل مجلد الطلب' },
    openDocument: { en: 'Open document', ar: 'فتح المستند' },
    noDocuments: { en: 'No PDF documents have been added yet.', ar: 'لم تتم إضافة مستندات PDF بعد.' },
    closeDetails: { en: 'Close Details', ar: 'إغلاق التفاصيل' },

    officeManagerActions: { en: 'Office Manager Actions', ar: 'إجراءات مدير المكتب' },
    officeManagerWorkspace: { en: 'Office Manager workspace', ar: 'مساحة عمل مدير المكتب' },
    officeManagerWorkspaceDescription: { en: 'Prepare the Office Manager document, add your comment, and route the request.', ar: 'جهّز مستند مدير المكتب، وأضف تعليقك، ثم وجّه الطلب.' },
    officeManagerComment: { en: 'Office Manager Comment', ar: 'تعليق مدير المكتب' },
    officeManagerDocument: { en: 'Office Manager document', ar: 'مستند مدير المكتب' },
    officeManagerDocumentDescription: { en: 'Create a formatted document in a familiar Word-style editor and attach it as PDF.', ar: 'أنشئ مستنداً منسقاً في محرر يشبه Word وأرفقه بصيغة PDF.' },
    createOfficeManagerDocument: { en: 'Create Office Manager PDF', ar: 'إنشاء ملف PDF لمدير المكتب' },
    editOfficeManagerDocument: { en: 'Create replacement Office Manager PDF', ar: 'إنشاء ملف PDF بديل لمدير المكتب' },
    pdfSource: { en: 'Document source', ar: 'مصدر المستند' },
    uploadPdf: { en: 'Upload PDF', ar: 'رفع ملف PDF' },
    createWithWord: { en: 'Create in editor', ar: 'إنشاء في المحرر' },
    attachSelectedPdf: { en: 'Attach selected PDF', ar: 'إرفاق ملف PDF المحدد' },
    noPdfSelected: { en: 'Choose a PDF before attaching it.', ar: 'اختر ملف PDF قبل إرفاقه.' },
    documentFrozenAfterSignature: { en: 'The workflow document is locked after GM signature.', ar: 'تم قفل مستند سير العمل بعد توقيع المدير العام.' },
    editorTitle: { en: 'Office Manager document editor', ar: 'محرر مستند مدير المكتب' },
    editorSubtitle: { en: 'Format the page, then create and attach the final PDF.', ar: 'نسّق الصفحة، ثم أنشئ ملف PDF النهائي وأرفقه.' },
    paragraph: { en: 'Paragraph', ar: 'فقرة' },
    heading: { en: 'Heading', ar: 'عنوان' },
    bold: { en: 'Bold', ar: 'عريض' },
    italic: { en: 'Italic', ar: 'مائل' },
    underline: { en: 'Underline', ar: 'تحته خط' },
    bullets: { en: 'Bullets', ar: 'تعداد نقطي' },
    alignLeft: { en: 'Align left', ar: 'محاذاة لليسار' },
    alignCenter: { en: 'Align center', ar: 'توسيط' },
    alignRight: { en: 'Align right', ar: 'محاذاة لليمين' },
    undo: { en: 'Undo', ar: 'تراجع' },
    redo: { en: 'Redo', ar: 'إعادة' },
    createAttachPdf: { en: 'Create & attach PDF', ar: 'إنشاء وإرفاق PDF' },
    prepareRequestPdf: { en: 'Create PDF for this request', ar: 'إنشاء PDF لهذا الطلب' },
    requestPdfReady: { en: 'The generated PDF is ready to submit with the request.', ar: 'ملف PDF الذي تم إنشاؤه جاهز للإرسال مع الطلب.' },
    generatingPdf: { en: 'Creating and uploading PDF...', ar: 'جاري إنشاء ملف PDF ورفعه...' },
    closeEditor: { en: 'Close editor', ar: 'إغلاق المحرر' },
    hodComment: { en: 'HOD Comment', ar: 'تعليق رئيس القسم' },
    gmCommentNotes: { en: 'GM Comment / Notes', ar: 'تعليق / ملاحظات المدير العام' },
    putOnHold: { en: 'Put On Hold', ar: 'وضع الطلب قيد التعليق' },
    sendToHod: { en: 'Send to HOD', ar: 'إرسال إلى رئيس القسم' },
    sendToGmSignature: { en: 'Send to GM Signature', ar: 'إرسال لتوقيع المدير العام' },
    reject: { en: 'Reject', ar: 'رفض' },
    returnToPendingOfficeManager: { en: 'Return to Pending Office Manager', ar: 'إرجاع إلى مدير المكتب' },
    confirmGmApproval: { en: 'Confirm GM Approval', ar: 'تأكيد موافقة المدير العام' },
    gmReject: { en: 'GM Reject', ar: 'رفض المدير العام' },
    requestSecretaryInformation: { en: 'Request information from Secretary', ar: 'طلب معلومات من السكرتارية' },
    requestHodInformation: { en: 'Request information from HOD', ar: 'طلب معلومات من رئيس القسم' },
    secretaryActions: { en: 'Secretary response', ar: 'رد السكرتارية' },
    secretaryComment: { en: 'Secretary comment', ar: 'تعليق السكرتارية' },
    submitInformation: { en: 'Send information to Office Manager', ar: 'إرسال المعلومات إلى مدير المكتب' },
    conversationTitle: { en: 'Workflow conversation', ar: 'محادثة سير العمل' },
    conversationDescription: { en: 'Previous messages are preserved in chronological order.', ar: 'يتم الاحتفاظ بالرسائل السابقة بالترتيب الزمني.' },
    noComments: { en: 'No comments have been added yet.', ar: 'لم تتم إضافة تعليقات بعد.' },
    legacyComment: { en: 'Previous comment', ar: 'تعليق سابق' },
    roleOfficeManager: { en: 'Office Manager', ar: 'مدير المكتب' },
    roleHod: { en: 'HOD', ar: 'رئيس القسم' },
    roleSecretary: { en: 'Secretary', ar: 'السكرتارية' },
    roleGm: { en: 'General Manager', ar: 'المدير العام' },
    closeRequest: { en: 'Close Request', ar: 'إغلاق الطلب' },
    hodActions: { en: 'HOD Actions', ar: 'إجراءات رئيس القسم' },
    returnToOfficeManager: { en: 'Return to Office Manager', ar: 'إرجاع إلى مدير المكتب' },

    errorRequiredFields: { en: 'Title, department, and request details are required.', ar: 'العنوان والقسم وتفاصيل الطلب مطلوبة.' },
    errorPdfOnly: { en: 'Only PDF files can be attached.', ar: 'يمكن إرفاق ملفات PDF فقط.' },
    errorNoItemId: { en: 'SharePoint created the request, but did not return the new item ID.', ar: 'تم إنشاء الطلب في SharePoint ولكن لم يتم إرجاع معرف العنصر الجديد.' },
    errorUnknown: { en: 'Unknown error.', ar: 'خطأ غير معروف.' },
    errorReadPdf: { en: 'Could not read the selected PDF file.', ar: 'تعذر قراءة ملف PDF المحدد.' },
    errorPdfJson: { en: 'PDF uploaded, but JSON could not be saved. Create a multiple lines of text column with internal name "{0}". {1}', ar: 'تم رفع ملف PDF، ولكن تعذر حفظ JSON. قم بإنشاء عمود نص متعدد الأسطر بالاسم الداخلي "{0}". {1}' },
    errorLibraryUpload: { en: 'Document library upload failed. {0}', ar: 'فشل رفع الملف إلى مكتبة المستندات. {0}' },
    errorLibraryMetadata: { en: 'Document library metadata update failed. {0}', ar: 'فشل تحديث بيانات مكتبة المستندات. {0}' },
    errorListAttachment: { en: 'Request list attachment upload failed. {0}', ar: 'فشل رفع المرفق إلى قائمة الطلبات. {0}' },
    errorDocumentsLibraryMissing: { en: 'Could not find the GM Approval Documents library. {0}', ar: 'تعذر العثور على مكتبة مستندات موافقات المدير العام. {0}' },
    errorRequestFolder: { en: 'Could not create or open the request document folder. {0}', ar: 'تعذر إنشاء مجلد مستندات الطلب أو فتحه. {0}' },
    errorOfficeManagerPdf: { en: 'The Office Manager PDF could not be created or attached. {0}', ar: 'تعذر إنشاء ملف PDF الخاص بمدير المكتب أو إرفاقه. {0}' },
    errorEditorUnavailable: { en: 'The document editor is not available. Close it and try again.', ar: 'محرر المستند غير متاح. أغلقه ثم حاول مرة أخرى.' },
    errorCommentRequired: { en: 'Add a comment before this action.', ar: 'أضف تعليقاً قبل تنفيذ هذا الإجراء.' },
    errorPrimaryPdfRequired: { en: 'Add a workflow PDF or original request PDF before sending this request for GM signature.', ar: 'أضف ملف PDF لسير العمل أو ملف الطلب الأصلي قبل إرسال الطلب لتوقيع المدير العام.' },
    errorRoleDocumentLocked: { en: 'This request is no longer pending for your role, so its document cannot be replaced.', ar: 'لم يعد هذا الطلب بانتظار دورك، لذلك لا يمكن استبدال مستنده.' },
    errorStatusChanged: { en: 'This request was updated by another user. Reopen it and try again.', ar: 'تم تحديث هذا الطلب بواسطة مستخدم آخر. أعد فتحه ثم حاول مرة أخرى.' },
    messageUploadingPdf: { en: 'Request created. Uploading PDF to the document library and list attachment...', ar: 'تم إنشاء الطلب. جاري رفع ملف PDF إلى مكتبة المستندات ومرفقات الطلب...' },
    messageCreated: { en: 'Request created successfully: {0}', ar: 'تم إنشاء الطلب بنجاح: {0}' },
    messagePdfStepFailed: { en: 'Request {0} was created, but the PDF step failed: ', ar: 'تم إنشاء الطلب {0}، ولكن فشلت خطوة PDF: ' },
    messageUpdatedTo: { en: 'Request updated to: {0}', ar: 'تم تحديث الطلب إلى: {0}' },
    messageReturnedToOfficeManager: { en: 'Request returned to Office Manager.', ar: 'تم إرجاع الطلب إلى مدير المكتب.' },
    messageGmApprovalConfirmed: { en: 'GM approval confirmed.', ar: 'تم تأكيد موافقة المدير العام.' },
    messageGmRejected: { en: 'The request was rejected by GM.', ar: 'تم رفض الطلب من المدير العام.' },
    messageInformationRequested: { en: 'Information request sent to: {0}', ar: 'تم إرسال طلب المعلومات إلى: {0}' },
    messageInformationSubmitted: { en: 'Information sent to Office Manager.', ar: 'تم إرسال المعلومات إلى مدير المكتب.' },
    messageRequestClosed: { en: 'Request closed.', ar: 'تم إغلاق الطلب.' },
    messageOfficeManagerPdfCreated: { en: 'Office Manager PDF created and attached successfully: {0}', ar: 'تم إنشاء ملف PDF الخاص بمدير المكتب وإرفاقه بنجاح: {0}' },
    messageRolePdfAttached: { en: '{0} attached successfully: {1}', ar: 'تم إرفاق {0} بنجاح: {1}' },

    statusPendingOfficeManager: { en: 'Pending Office Manager', ar: 'بانتظار مدير المكتب' },
    statusOnHold: { en: 'On Hold', ar: 'قيد التعليق' },
    statusReturnedFromHod: { en: 'Returned from HOD', ar: 'مُعاد من رئيس القسم' },
    statusPendingGmSignature: { en: 'Pending GM Signature', ar: 'بانتظار توقيع المدير العام' },
    statusGmSignedPendingOfficeManagerConfirmation: { en: 'GM Signed Pending Office Manager Confirmation', ar: 'تم توقيع المدير العام وبانتظار تأكيد مدير المكتب' },
    statusSentToHod: { en: 'Sent to HOD', ar: 'مرسل إلى رئيس القسم' },
    statusRejected: { en: 'Rejected', ar: 'مرفوض' },
    statusApprovedByGm: { en: 'Approved by GM', ar: 'معتمد من المدير العام' },
    statusPendingSecretaryInformation: { en: 'Pending Secretary Information', ar: 'بانتظار معلومات السكرتارية' },
    statusPendingHodInformation: { en: 'Pending HOD Information', ar: 'بانتظار معلومات رئيس القسم' },
    statusClosed: { en: 'Closed', ar: 'مغلق' },

    departmentNetworkDevelopmentAndMaintenance: { en: 'Network Development and Maintenance', ar: 'تطوير وصيانة الشبكات' },
    departmentInformationTechnology: { en: 'Information Technology', ar: 'تقنية المعلومات' },
    departmentOperations: { en: 'Operations', ar: 'العمليات' },
    departmentFinance: { en: 'Finance', ar: 'المالية' },
    departmentHumanResources: { en: 'Human Resources', ar: 'الموارد البشرية' },
    departmentProcurement: { en: 'Procurement', ar: 'المشتريات' },
    departmentAdministration: { en: 'Administration', ar: 'الإدارة' },
    departmentSales: { en: 'Sales', ar: 'المبيعات' },
    departmentCustomerService: { en: 'Customer Service', ar: 'خدمة العملاء' }
  };
  private departmentOptions: string[] = [
    'Network Development and Maintenance',
    'Information Technology',
    'Operations',
    'Finance',
    'Human Resources',
    'Procurement',
    'Administration',
    'Sales',
    'Customer Service'
  ];

  constructor(props: IGmApprovalDashboardProps) {
    super(props);

    this.state = {
      language: this.getDefaultLanguage(),
      loading: true,
      saving: false,
      error: '',
      message: '',

      currentUserId: 0,
      currentUserTitle: '',
      groups: [],

      isSecretary: false,
      isOfficeManager: false,
      isHod: false,
      isServiceAccount: false,

      requests: [],
      selectedRequest: null,
      showNewRequestForm: false,

      newTitle: '',
      newDepartment: '',
      newDetails: '',
      newPdfFile: null,
      newPdfMode: 'upload',
      showNewRequestEditor: false,
      newRequestStatus: '',

      officeManagerComment: '',
      hodComment: '',
      gmComment: '',
      secretaryComment: '',
      documentSourceMode: 'create',
      rolePdfFile: null,
      activeDocumentRole: '',
      showOfficeManagerEditor: false
    };
  }

  private getDefaultLanguage(): LanguageCode {
    return 'ar';
  }

  private setLanguage(language: LanguageCode): void {
    this.setState({
      language: language
    });
  }

  private t(key: string): string {
    const translation: ITranslatedText = this.translations[key];

    if (!translation) {
      return key;
    }

    return translation[this.state.language] || translation.en;
  }

  private formatText(key: string, firstValue?: string, secondValue?: string): string {
    return this.t(key)
      .replace('{0}', firstValue || '')
      .replace('{1}', secondValue || '');
  }

  private getRootClassName(): string {
    if (this.state.language === 'ar') {
      return styles.gmApprovalDashboard + ' ' + styles.rtl;
    }

    return styles.gmApprovalDashboard;
  }

  private getStatusText(status: string): string {
    let key: string = '';

    switch (status) {
      case 'Pending Office Manager':
        key = 'statusPendingOfficeManager';
        break;
      case 'On Hold':
        key = 'statusOnHold';
        break;
      case 'Returned from HOD':
        key = 'statusReturnedFromHod';
        break;
      case 'Pending GM Signature':
        key = 'statusPendingGmSignature';
        break;
      case 'GM Signed Pending Office Manager Confirmation':
        key = 'statusGmSignedPendingOfficeManagerConfirmation';
        break;
      case 'Sent to HOD':
        key = 'statusSentToHod';
        break;
      case 'Rejected':
        key = 'statusRejected';
        break;
      case 'Approved by GM':
        key = 'statusApprovedByGm';
        break;
      case 'Pending Secretary Information':
        key = 'statusPendingSecretaryInformation';
        break;
      case 'Pending HOD Information':
        key = 'statusPendingHodInformation';
        break;
      case 'Closed':
        key = 'statusClosed';
        break;
      default:
        key = '';
        break;
    }

    const translation: ITranslatedText = this.translations[key];
    return translation ? this.t(key) : status;
  }

  private getStatusClassName(status: string): string {
    let variantClass: string = styles.statusNeutral;

    if (status === 'Rejected') {
      variantClass = styles.statusDanger;
    } else if (status === 'Approved by GM' || status === 'Closed') {
      variantClass = styles.statusSuccess;
    } else if (
      status === 'Pending Office Manager' ||
      status === 'Pending GM Signature' ||
      status === 'GM Signed Pending Office Manager Confirmation' ||
      status === 'Pending Secretary Information' ||
      status === 'Pending HOD Information' ||
      status === 'On Hold'
    ) {
      variantClass = styles.statusPending;
    }

    return styles.statusBadge + ' ' + variantClass;
  }

  private renderStatusBadge(status: string): JSX.Element {
    return <span className={this.getStatusClassName(status)}>{this.getStatusText(status)}</span>;
  }

  private getDepartmentText(department: string): string {
    let key: string = '';

    switch (department) {
      case 'Network Development and Maintenance':
        key = 'departmentNetworkDevelopmentAndMaintenance';
        break;
      case 'Information Technology':
        key = 'departmentInformationTechnology';
        break;
      case 'Operations':
        key = 'departmentOperations';
        break;
      case 'Finance':
        key = 'departmentFinance';
        break;
      case 'Human Resources':
        key = 'departmentHumanResources';
        break;
      case 'Procurement':
        key = 'departmentProcurement';
        break;
      case 'Administration':
        key = 'departmentAdministration';
        break;
      case 'Sales':
        key = 'departmentSales';
        break;
      case 'Customer Service':
        key = 'departmentCustomerService';
        break;
      default:
        key = '';
        break;
    }

    const translation: ITranslatedText = this.translations[key];
    return translation ? this.t(key) : department;
  }

  private getGroupText(groupName: string): string {
    let key: string = '';

    switch (groupName) {
      case 'GM Secretaries':
        key = 'groupSecretaries';
        break;
      case 'GM Office Managers':
        key = 'groupOfficeManagers';
        break;
      case 'GM HODs':
        key = 'groupHods';
        break;
      case 'GM Service Accounts':
        key = 'groupServiceAccounts';
        break;
      default:
        key = '';
        break;
    }

    return key ? this.t(key) : groupName;
  }

  private getGroupsText(): string {
    const names: string[] = [];

    for (let i = 0; i < this.state.groups.length; i++) {
      names.push(this.getGroupText(this.state.groups[i]));
    }

    return names.join(this.state.language === 'ar' ? '، ' : ', ');
  }

  public componentDidMount(): void {
    this.selectedRequestRefreshTimer = window.setInterval(() => this.refreshSelectedRequest(), 15000);

    this.loadSecurity()
      .then(() => this.loadRequests())
      .catch((error: any) => {
        this.setState({
          loading: false,
          error: this.getErrorMessage(error)
        });
      });
  }

  public componentWillUnmount(): void {
    if (this.selectedRequestRefreshTimer) {
      window.clearInterval(this.selectedRequestRefreshTimer);
      this.selectedRequestRefreshTimer = 0;
    }
  }

  private refreshSelectedRequest(): void {
    const selectedItem: IRequestItem = this.state.selectedRequest;
    if (!selectedItem || this.state.saving) {
      return;
    }

    this.getRequestById(selectedItem.Id)
      .then((freshItem: IRequestItem) => {
        if (!this.state.selectedRequest || this.state.selectedRequest.Id !== freshItem.Id) {
          return;
        }

        this.setState({
          selectedRequest: freshItem,
          requests: this.replaceRequestInCollection(freshItem.Id, freshItem)
        });
      })
      .catch(() => {
        return;
      });
  }

  private get webUrl(): string {
    const contextWebUrl: string = this.props.context.pageContext.web.absoluteUrl;

    if (contextWebUrl && contextWebUrl.indexOf('localhost') < 0 && contextWebUrl.indexOf('127.0.0.1') < 0) {
      return contextWebUrl;
    }

    return this.gmSiteUrl;
  }

  private getJsonRequestOptions(): ISPHttpClientOptions {
    return {
      headers: {
        'Accept': this.odataJsonHeader,
        'OData-Version': this.odataVersion
      }
    };
  }

  private getJsonPostOptions(values: any): ISPHttpClientOptions {
    return {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': this.odataJsonHeader,
        'OData-Version': this.odataVersion
      },
      body: JSON.stringify(values)
    };
  }

  private getJsonMergeOptions(values: any, etag?: string): ISPHttpClientOptions {
    return {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': this.odataJsonHeader,
        'IF-MATCH': etag || '*',
        'X-HTTP-Method': 'MERGE',
        'OData-Version': this.odataVersion
      },
      body: JSON.stringify(values)
    };
  }

  private addODataMetadata(values: any, entityTypeName: string): any {
    const result: any = {};

    if (entityTypeName) {
      result.__metadata = {
        type: entityTypeName
      };
    }

    for (const key in values) {
      if (values.hasOwnProperty(key)) {
        result[key] = values[key];
      }
    }

    return result;
  }

  private getRequestsListItemEntityTypeFullName(): Promise<string> {
    if (this.requestsListItemEntityTypeFullName) {
      return Promise.resolve(this.requestsListItemEntityTypeFullName);
    }

    const url: string = this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.escapeODataString(this.requestsListName) + "')" +
      "?$select=ListItemEntityTypeFullName";

    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }

        return response.json();
      })
      .then((data: any) => {
        const listData: any = data.d ? data.d : data;
        this.requestsListItemEntityTypeFullName = listData.ListItemEntityTypeFullName || '';
        return this.requestsListItemEntityTypeFullName;
      });
  }

  private getDocumentsLibraryListItemEntityTypeFullName(): Promise<string> {
    if (this.documentsLibraryListItemEntityTypeFullName) {
      return Promise.resolve(this.documentsLibraryListItemEntityTypeFullName);
    }

    const url: string = this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.escapeODataString(this.documentsLibraryName) + "')" +
      "?$select=ListItemEntityTypeFullName";

    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }

        return response.json();
      })
      .then((data: any) => {
        const listData: any = data.d ? data.d : data;
        this.documentsLibraryListItemEntityTypeFullName = listData.ListItemEntityTypeFullName || '';
        return this.documentsLibraryListItemEntityTypeFullName;
      });
  }

  private getListItemsUrl(): string {
    return this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.requestsListName + "')/items" +
      this.getListItemQuery() +
      "&$orderby=Id desc" +
      "&$top=200";
  }

  private getListItemQuery(): string {
    return "?$select=Id,Title,RequestNo,RequestDate,Department,RequestDetails,Status," +
      "OfficeManagerComment,HODComment,GMComment,PDFFileUrl,PDFFileJson,SignedPDFUrl,SharedFolderPath," +
      "GMApprovalDetected,GMApprovalConfirmed,Attachments,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl,Author/Id,Author/Title" +
      "&$expand=Author,AttachmentFiles";
  }

  private getRequestById(itemId: number): Promise<IRequestItem> {
    const url: string = this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.requestsListName + "')/items(" + itemId + ")" +
      this.getListItemQuery();

    let responseEtag: string = '';

    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }

        responseEtag = response.headers ? (response.headers.get('ETag') || '') : '';
        return response.json();
      })
      .then((data: any) => {
        const item: any = data.d ? data.d : data;
        item.__etag = responseEtag || (item.__metadata && item.__metadata.etag) || '';
        return item as IRequestItem;
      });
  }

  private loadSecurity(): Promise<void> {
    const userUrl: string = this.webUrl + "/_api/web/currentuser?$select=Id,Title,LoginName,Email";
    const groupsUrl: string = this.webUrl + "/_api/web/currentuser/groups?$select=Title";

    return Promise.all([
      this.props.context.spHttpClient.get(userUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions()),
      this.props.context.spHttpClient.get(groupsUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
    ])
      .then((responses: SPHttpClientResponse[]) => {
        return Promise.all([
          responses[0].json(),
          responses[1].json()
        ]);
      })
      .then((results: any[]) => {
        const userData: any = results[0];
        const groupsData: any = results[1];

        const user: any = userData.d ? userData.d : userData;
        const groupResults: any[] = groupsData.value || (groupsData.d && groupsData.d.results) || [];

        const groups: string[] = [];
        for (let i = 0; i < groupResults.length; i++) {
          groups.push(groupResults[i].Title);
        }

        this.setState({
          currentUserId: user.Id,
          currentUserTitle: user.Title,
          groups: groups,

          isSecretary: groups.indexOf('GM Secretaries') >= 0,
          isOfficeManager: groups.indexOf('GM Office Managers') >= 0,
          isHod: groups.indexOf('GM HODs') >= 0 || groups.indexOf('GM HOD') >= 0,
          isServiceAccount: groups.indexOf('GM Service Accounts') >= 0
        });
      });
  }

  private loadRequests(clearNotifications?: boolean): Promise<void> {
    const nextState: any = {
      loading: true
    };

    if (clearNotifications !== false) {
      nextState.error = '';
      nextState.message = '';
    }

    this.setState(nextState);

    return this.props.context.spHttpClient
      .get(this.getListItemsUrl(), SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }

        return response.json();
      })
      .then((data: any) => {
        const items: IRequestItem[] = data.value || (data.d && data.d.results) || [];

        this.setState({
          requests: items,
          loading: false
        });
      });
  }

  private createRequest(): void {
    if (!this.state.newTitle || !this.state.newDepartment || !this.state.newDetails) {
      this.setState({
        error: this.t('errorRequiredFields')
      });
      return;
    }

    if (this.state.newPdfFile && !this.isPdfFile(this.state.newPdfFile)) {
      this.setState({
        error: this.t('errorPdfOnly')
      });
      return;
    }

    if (this.getNewRequestStatus() === 'Pending GM Signature' && !this.state.newPdfFile) {
      this.setState({ error: this.t('errorPrimaryPdfRequired') });
      return;
    }

    const requestDate: Date = new Date();
    const requestNo: string = this.buildRequestNo(this.state.newTitle, requestDate);
    const initialStatus: string = this.getNewRequestStatus();

    const body: any = {
      Title: this.state.newTitle,
      RequestNo: requestNo,
      Department: this.state.newDepartment,
      RequestDetails: this.state.newDetails,
      RequestDate: requestDate.toISOString(),
      Status: initialStatus,
      GMApprovalDetected: false,
      GMApprovalConfirmed: false
    };

    this.applySharedFolderPathForStatus(body, requestNo, initialStatus);

    this.setState({
      saving: true,
      error: '',
      message: ''
    });

    const url: string = this.webUrl + "/_api/web/lists/GetByTitle('" + this.requestsListName + "')/items";
    const pdfFile: any = this.state.newPdfFile;
    let createdItemId: number = 0;

    this.getRequestsListItemEntityTypeFullName()
      .then((entityTypeName: string) => {
        const createBody: any = this.addODataMetadata(body, entityTypeName);
        const options: ISPHttpClientOptions = this.getJsonPostOptions(createBody);

        return this.props.context.spHttpClient
          .post(url, SPHttpClient.configurations.v1, options);
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }

        return response.json();
      })
      .then((createdItem: any) => {
        createdItemId = this.getItemId(createdItem);

        if (!createdItemId) {
          throw new Error(this.t('errorNoItemId'));
        }

        if (!pdfFile) {
          return;
        }

        this.setState({
          message: this.t('messageUploadingPdf')
        });

        return this.uploadPdfToLibraryAndListItem(
          pdfFile,
          requestNo,
          createdItemId,
          this.state.newPdfMode === 'create' ? 'request-draft.pdf' : ''
        )
          .then((result: IPdfUploadResult) => {
            return this.updatePdfMetadata(createdItemId, result, requestNo, '', true);
          });
      })
      .then(() => {
        this.clearNewRequestForm();

        this.setState({
          saving: false,
          message: this.formatText('messageCreated', requestNo)
        });

        return this.loadRequests(false);
      })
      .catch((error: any) => {
        const prefix: string = createdItemId ? this.formatText('messagePdfStepFailed', requestNo) : '';
        this.setState({
          saving: false,
          error: prefix + this.getErrorMessage(error)
        });

        if (createdItemId) {
          this.loadRequests(false);
        }
      });
  }

  private updateItem(itemId: number, values: any, etag?: string): Promise<void> {
    const url: string = this.webUrl + "/_api/web/lists/GetByTitle('" + this.requestsListName + "')/items(" + itemId + ")";

    return this.getRequestsListItemEntityTypeFullName()
      .then((entityTypeName: string) => {
        const updateBody: any = this.addODataMetadata(values, entityTypeName);
        const options: ISPHttpClientOptions = this.getJsonMergeOptions(updateBody, etag);

        return this.props.context.spHttpClient
          .post(url, SPHttpClient.configurations.v1, options);
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }
      });
  }

  private selectRequest(item: IRequestItem): void {
    this.officeManagerDocumentHtml = '';
    this.newRequestDocumentHtml = '';
    this.setState({
      selectedRequest: item,
      showNewRequestForm: false,
      showOfficeManagerEditor: false,
      officeManagerComment: '',
      hodComment: '',
      gmComment: '',
      secretaryComment: '',
      documentSourceMode: 'create',
      rolePdfFile: null,
      activeDocumentRole: '',
      error: '',
      message: ''
    }, () => {
      this.getRequestById(item.Id)
        .then((freshItem: IRequestItem) => {
          if (!this.state.selectedRequest || this.state.selectedRequest.Id !== item.Id) {
            return;
          }

          this.setState({
            selectedRequest: freshItem,
            requests: this.replaceRequestInCollection(item.Id, freshItem)
          });
        })
        .catch(() => {
          return;
        });
    });
  }

  private officeManagerAction(status: string): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    if (status === 'Pending GM Signature' && !this.hasPrimaryWorkflowPdf(item)) {
      this.setState({ error: this.t('errorPrimaryPdfRequired'), message: '' });
      return;
    }

    const updateValues: any = {};
    const comment: string = this.state.officeManagerComment;

    if (comment) {
      updateValues.OfficeManagerComment = comment;
    }

    this.runWorkflowUpdate(
      item,
      status,
      'officeManager',
      'workflow',
      comment,
      updateValues,
      this.formatText('messageUpdatedTo', this.getStatusText(status)),
      status === 'Rejected'
    );
  }

  private hodReturnToOfficeManager(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    const responseToInformationRequest: boolean = item.Status === 'Pending HOD Information';
    const nextStatus: string = responseToInformationRequest
      ? 'GM Signed Pending Office Manager Confirmation'
      : 'Returned from HOD';

    this.runWorkflowUpdate(
      item,
      nextStatus,
      'hod',
      responseToInformationRequest ? 'informationResponse' : 'workflow',
      this.state.hodComment,
      { HODComment: this.state.hodComment },
      responseToInformationRequest ? this.t('messageInformationSubmitted') : this.t('messageReturnedToOfficeManager'),
      true
    );
  }

  private confirmGmApproval(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    const updateValues: any = {
      GMApprovalConfirmed: true
    };

    if (this.state.gmComment) {
      updateValues.GMComment = this.state.gmComment;
    }

    this.runWorkflowUpdate(
      item,
      'Approved by GM',
      'officeManager',
      'gmApprovalConfirmed',
      this.combineOfficeManagerAndGmNotes(),
      updateValues,
      this.t('messageGmApprovalConfirmed'),
      false
    );
  }

  private gmSignedOfficeManagerAction(status: string): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    const comment: string = this.combineOfficeManagerAndGmNotes();
    const updateValues: any = {
      GMApprovalConfirmed: false
    };

    if (this.state.officeManagerComment) {
      updateValues.OfficeManagerComment = this.state.officeManagerComment;
    }

    if (this.state.gmComment) {
      updateValues.GMComment = this.state.gmComment;
    }

    let successMessage: string = this.t('messageGmRejected');
    let kind: string = 'gmRejected';

    if (status === 'Pending Secretary Information') {
      successMessage = this.formatText('messageInformationRequested', this.t('roleSecretary'));
      kind = 'informationRequest';
    } else if (status === 'Pending HOD Information') {
      successMessage = this.formatText('messageInformationRequested', this.t('roleHod'));
      kind = 'informationRequest';
    }

    this.runWorkflowUpdate(
      item,
      status,
      'officeManager',
      kind,
      comment,
      updateValues,
      successMessage,
      true
    );
  }

  private combineOfficeManagerAndGmNotes(): string {
    const officeManagerComment: string = (this.state.officeManagerComment || '').replace(/^\s+|\s+$/g, '');
    const gmComment: string = (this.state.gmComment || '').replace(/^\s+|\s+$/g, '');

    if (officeManagerComment && gmComment) {
      return officeManagerComment + '\n' + this.t('gmCommentNotes') + ': ' + gmComment;
    }

    return officeManagerComment || gmComment;
  }

  private secretaryReturnToOfficeManager(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    this.runWorkflowUpdate(
      item,
      'GM Signed Pending Office Manager Confirmation',
      'secretary',
      'informationResponse',
      this.state.secretaryComment,
      {},
      this.t('messageInformationSubmitted'),
      true
    );
  }

  private closeRequest(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    const updateValues: any = {};
    if (this.state.officeManagerComment) {
      updateValues.OfficeManagerComment = this.state.officeManagerComment;
    }

    this.runWorkflowUpdate(
      item,
      'Closed',
      'officeManager',
      'closed',
      this.state.officeManagerComment,
      updateValues,
      this.t('messageRequestClosed'),
      false
    );
  }

  private runUpdate(itemId: number, values: any, successMessage: string): void {
    this.setState({
      saving: true,
      error: '',
      message: ''
    });

    this.updateItem(itemId, values)
      .then(() => {
        this.setState({
          saving: false,
          message: successMessage,
          selectedRequest: null
        });

        return this.loadRequests(false);
      })
      .catch((error: any) => {
        this.setState({
          saving: false,
          error: this.getErrorMessage(error)
        });
      });
  }

  private runWorkflowUpdate(
    item: IRequestItem,
    nextStatus: string,
    role: string,
    kind: string,
    comment: string,
    values: any,
    successMessage: string,
    commentRequired: boolean
  ): void {
    const cleanComment: string = (comment || '').replace(/^\s+|\s+$/g, '');

    if (commentRequired && !cleanComment) {
      this.setState({ error: this.t('errorCommentRequired'), message: '' });
      return;
    }

    this.setState({ saving: true, error: '', message: '' });

    this.getRequestById(item.Id)
      .then((freshItem: IRequestItem) => {
        if (freshItem.Status !== item.Status) {
          throw new Error(this.t('errorStatusChanged'));
        }

        const updateValues: any = {};
        for (const key in values) {
          if (values.hasOwnProperty(key)) {
            updateValues[key] = values[key];
          }
        }

        updateValues.Status = nextStatus;
        this.applySharedFolderPathForStatus(updateValues, freshItem.RequestNo, nextStatus);

        if (cleanComment) {
          updateValues[this.pdfJsonFieldName] = this.appendWorkflowComment(
            freshItem,
            role,
            kind,
            nextStatus,
            cleanComment
          );
        }

        return this.updateItem(freshItem.Id, updateValues, (freshItem as any).__etag || '');
      })
      .then(() => {
        this.setState({
          saving: false,
          message: successMessage,
          selectedRequest: null,
          showOfficeManagerEditor: false,
          rolePdfFile: null,
          activeDocumentRole: ''
        });

        return this.loadRequests(false);
      })
      .catch((error: any) => {
        this.setState({
          saving: false,
          error: this.getErrorMessage(error),
          message: ''
        });
      });
  }

  private getMyRequests(): IRequestItem[] {
    const items: IRequestItem[] = [];
    for (let i = 0; i < this.state.requests.length; i++) {
      const item: IRequestItem = this.state.requests[i];
      if (item.Author && item.Author.Id === this.state.currentUserId) {
        items.push(item);
      }
    }
    return items;
  }

  private getOfficeManagerRequests(): IRequestItem[] {
    const allowedStatuses: string[] = [
      'Pending Office Manager',
      'On Hold',
      'Returned from HOD',
      'Pending GM Signature',
      'GM Signed Pending Office Manager Confirmation',
      'Pending Secretary Information',
      'Pending HOD Information',
      'Approved by GM'
    ];

    const items: IRequestItem[] = [];
    for (let i = 0; i < this.state.requests.length; i++) {
      const item: IRequestItem = this.state.requests[i];
      if (allowedStatuses.indexOf(item.Status) >= 0) {
        items.push(item);
      }
    }
    return items;
  }

  private getHodRequests(): IRequestItem[] {
    const items: IRequestItem[] = [];
    for (let i = 0; i < this.state.requests.length; i++) {
      const item: IRequestItem = this.state.requests[i];
      if (item.Status === 'Sent to HOD' || item.Status === 'Pending HOD Information') {
        items.push(item);
      }
    }
    return items;
  }

  private getSecretaryRequests(): IRequestItem[] {
    const items: IRequestItem[] = [];
    for (let i = 0; i < this.state.requests.length; i++) {
      if (this.state.requests[i].Status === 'Pending Secretary Information') {
        items.push(this.state.requests[i]);
      }
    }
    return items;
  }

  private canCreateRequest(): boolean {
    return this.state.isSecretary || this.state.isOfficeManager || this.state.isHod;
  }

  private getDefaultNewRequestStatus(): string {
    if (this.state.isOfficeManager) {
      return 'Sent to HOD';
    }

    return 'Pending Office Manager';
  }

  private getNewRequestStatus(): string {
    if (this.state.isOfficeManager) {
      return this.state.newRequestStatus || 'Sent to HOD';
    }

    return 'Pending Office Manager';
  }

  private getNewRequestSubmitText(): string {
    const status: string = this.getNewRequestStatus();

    if (status === 'Sent to HOD') {
      return this.t('sendToHod');
    }

    if (status === 'Pending GM Signature') {
      return this.t('sendToGmSignature');
    }

    return this.t('submitToOfficeManager');
  }

  private getSharedFolderNameForStatus(status: string): string {
    switch (status) {
      case 'Pending GM Signature':
        return 'Pending';
      case 'GM Signed Pending Office Manager Confirmation':
      case 'Pending Secretary Information':
      case 'Pending HOD Information':
      case 'Approved by GM':
        return 'Signed';
      case 'Closed':
      case 'Rejected':
        return 'Archive';
      default:
        return '';
    }
  }

  private getSharedFolderPathForStatus(requestNo: string, status: string): string {
    const folderName: string = this.getSharedFolderNameForStatus(status);

    if (!folderName || !requestNo) {
      return '';
    }

    return this.sharedFolderRootPath + '\\' + folderName + '\\' + requestNo + '.pdf';
  }

  private applySharedFolderPathForStatus(values: any, requestNo: string, status: string): void {
    const sharedFolderPath: string = this.getSharedFolderPathForStatus(requestNo, status);

    if (sharedFolderPath) {
      values.SharedFolderPath = sharedFolderPath;
    }
  }

  private showNewRequestForm(): void {
    this.setState({
      showNewRequestForm: true,
      selectedRequest: null,
      newRequestStatus: this.getDefaultNewRequestStatus(),
      error: '',
      message: ''
    });
  }

  private cancelNewRequest(): void {
    this.clearNewRequestForm();
  }

  private getItemId(item: any): number {
    if (!item) {
      return 0;
    }

    const data: any = item.d ? item.d : item;
    return data.Id || data.ID || 0;
  }

  private buildRequestNo(title: string, date: Date): string {
    const titlePart: string = this.sanitizeForRequestNo(title);
    const datePart: string = this.formatDateForRequestNo(date);
    const generatedNumber: string = this.pad(Math.floor(Math.random() * 9000) + 1000, 4);

    return titlePart + '-' + datePart + '-' + generatedNumber;
  }

  private sanitizeForRequestNo(value: string): string {
    let result: string = (value || '').toUpperCase();
    result = result.replace(/[^A-Z0-9]+/g, '-');
    result = result.replace(/^-+/, '').replace(/-+$/, '');

    if (!result) {
      result = 'REQUEST';
    }

    if (result.length > 40) {
      result = result.substr(0, 40).replace(/-+$/, '');
    }

    return result;
  }

  private formatDateForRequestNo(date: Date): string {
    const y: string = date.getFullYear().toString();
    const m: string = this.pad(date.getMonth() + 1, 2);
    const d: string = this.pad(date.getDate(), 2);
    return y + m + d;
  }

  private pad(value: number, length: number): string {
    let result: string = value.toString();
    while (result.length < length) {
      result = '0' + result;
    }
    return result;
  }

  private onPdfFileChange(event: any): void {
    const files: any = event.target.files;
    const file: any = files && files.length > 0 ? files[0] : null;

    if (file && !this.isPdfFile(file)) {
      if (this.pdfFileInput) {
        this.pdfFileInput.value = '';
      }

      this.setState({
        newPdfFile: null,
        error: this.t('errorPdfOnly')
      });
      return;
    }

    this.setState({
      newPdfFile: file,
      error: ''
    });
  }

  private setNewRequestPdfMode(mode: PdfSourceMode): void {
    if (this.pdfFileInput) {
      this.pdfFileInput.value = '';
    }

    this.newRequestDocumentHtml = '';
    this.setState({
      newPdfMode: mode,
      newPdfFile: null,
      showNewRequestEditor: mode === 'create',
      error: '',
      message: ''
    }, () => {
      if (mode === 'create') {
        this.newRequestDocumentHtml = this.buildNewRequestDocumentTemplate();
        if (this.newRequestEditor) {
          this.newRequestEditor.innerHTML = this.newRequestDocumentHtml;
        }
      }
    });
  }

  private buildNewRequestDocumentTemplate(): string {
    const title: string = this.escapeHtml(this.state.newTitle || this.t('newRequest'));
    const department: string = this.escapeHtml(this.getDepartmentText(this.state.newDepartment || ''));
    const details: string = this.escapeHtml(this.state.newDetails || '').replace(/\r?\n/g, '<br />');

    return '' +
      '<div>' +
        '<p style="margin:0 0 8px;color:#107c41;font-size:12px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;">' +
          this.escapeHtml(this.t('requestPdf')) +
        '</p>' +
        '<h1 style="margin:0 0 20px;color:#102a43;font-size:28px;line-height:1.25;">' + title + '</h1>' +
        '<p style="margin:0 0 24px;padding:10px;background:#f4f7f6;font-weight:700;">' + department + '</p>' +
        '<h2 style="margin:0 0 8px;color:#102a43;font-size:18px;">' + this.escapeHtml(this.t('requestDetails')) + '</h2>' +
        '<p style="min-height:180px;margin:0;line-height:1.65;">' + (details || '<br />') + '</p>' +
      '</div>';
  }

  private applyNewRequestEditorCommand(event: any, command: string, value?: string): void {
    if (event && event.preventDefault) {
      event.preventDefault();
    }

    if (!this.newRequestEditor) {
      return;
    }

    this.newRequestEditor.focus();
    document.execCommand(command, false, value || null);
    this.newRequestDocumentHtml = this.newRequestEditor.innerHTML;
    this.invalidateGeneratedNewRequestPdf();
  }

  private updateNewRequestField(fieldName: string, value: string): void {
    const values: any = {};
    values[fieldName] = value;

    if (this.state.newPdfMode === 'create') {
      values.newPdfFile = null;
      values.message = '';
    }

    this.setState(values);
  }

  private onNewRequestEditorInput(): void {
    if (this.newRequestEditor) {
      this.newRequestDocumentHtml = this.newRequestEditor.innerHTML;
    }

    this.invalidateGeneratedNewRequestPdf();
  }

  private invalidateGeneratedNewRequestPdf(): void {
    if (this.state.newPdfMode === 'create' && this.state.newPdfFile) {
      this.setState({ newPdfFile: null, message: '' });
    }
  }

  private createNewRequestPdf(): void {
    if (!this.newRequestEditor) {
      this.setState({ error: this.t('errorEditorUnavailable') });
      return;
    }

    this.newRequestDocumentHtml = this.newRequestEditor.innerHTML;
    this.setState({ saving: true, newPdfFile: null, error: '', message: this.t('generatingPdf') });

    this.createPdfBlob(this.newRequestEditor, 'request-draft.pdf')
      .then((blob: Blob) => {
        this.setState({
          saving: false,
          newPdfFile: blob,
          message: this.t('requestPdfReady')
        });
      })
      .catch((error: any) => {
        this.setState({
          saving: false,
          error: this.formatText('errorOfficeManagerPdf', this.getErrorMessage(error)),
          message: ''
        });
      });
  }

  private isPdfFile(file: any): boolean {
    if (!file) {
      return false;
    }

    const fileName: string = (file.name || '').toLowerCase();
    return file.type === 'application/pdf' || fileName.lastIndexOf('.pdf') === fileName.length - 4;
  }

  private uploadPdfToLibraryAndListItem(
    file: any,
    requestNo: string,
    itemId: number,
    suppliedOriginalFileName?: string
  ): Promise<IPdfUploadResult> {
    const originalFileName: string = suppliedOriginalFileName || file.name || 'request.pdf';
    const fileName: string = this.buildPdfFileName(requestNo, originalFileName);
    const documentType: string = 'Request PDF';

    return this.readFileAsArrayBuffer(file)
      .then((fileContent: ArrayBuffer) => {
        return this.uploadPdfFileToLibrary(fileName, fileContent, requestNo, 'request', documentType)
          .then((libraryResult: IPdfUploadResult) => {
            return this.updateLibraryFileMetadata(libraryResult.serverRelativeUrl, fileName, requestNo, itemId, documentType)
              .then(() => {
                return this.attachPdfFileToRequestItem(itemId, fileName, fileContent);
              })
              .then((attachmentUrl: string) => {
                libraryResult.originalFileName = originalFileName;
                libraryResult.contentType = file.type || 'application/pdf';
                libraryResult.size = file.size || 0;
                libraryResult.attachmentUrl = attachmentUrl;

                return libraryResult;
              });
          });
      });
  }

  private uploadPdfFileToLibrary(
    fileName: string,
    fileContent: ArrayBuffer,
    requestNo: string,
    role: string,
    documentType: string
  ): Promise<IPdfUploadResult> {
    const folderName: string = this.sanitizeSharePointFolderName(requestNo);

    return this.getRequestDocumentFolderServerRelativeUrl(requestNo)
      .then(() => {
        const uploadUrl: string = this.getDocumentsLibraryFolderApiUrl(folderName) +
          "/Files/Add(url='" + this.encodeODataUrlValue(fileName) + "',overwrite=true)" +
          "?$select=Name,ServerRelativeUrl";

        const options: ISPHttpClientOptions = {
          headers: {
            'Accept': this.odataJsonHeader,
            'Content-Type': 'application/octet-stream',
            'OData-Version': this.odataVersion
          },
          body: fileContent
        };

        return this.props.context.spHttpClient
          .post(uploadUrl, SPHttpClient.configurations.v1, options)
          .then((response: SPHttpClientResponse) => {
            if (!response.ok) {
              return response.text().then((text: string) => {
                throw this.createSharePointHttpError('Upload PDF to request folder', response, text, uploadUrl);
              });
            }

            return response.json();
          })
          .then((data: any) => {
            const uploadedFile: any = data.d ? data.d : data;
            const serverRelativeUrl: string = uploadedFile && uploadedFile.ServerRelativeUrl;

            if (!serverRelativeUrl) {
              throw new Error(
                'Upload PDF to request folder succeeded, but SharePoint did not return ServerRelativeUrl. ' +
                'Request: ' + uploadUrl
              );
            }

            return {
              role: role,
              documentType: documentType,
              fileName: fileName,
              originalFileName: '',
              contentType: 'application/pdf',
              size: 0,
              serverRelativeUrl: serverRelativeUrl,
              url: this.getAbsoluteUrl(serverRelativeUrl),
              attachmentUrl: ''
            };
          });
      })
      .catch((error: any) => {
        const message: string = this.getErrorMessage(error);

        if (message.indexOf(this.t('errorRequestFolder').replace('{0}', '')) === 0) {
          throw error;
        }

        throw new Error(this.formatText('errorLibraryUpload', message));
      });
  }

  private updateLibraryFileMetadata(
    serverRelativeUrl: string,
    fileName: string,
    requestNo: string,
    itemId: number,
    documentType: string
  ): Promise<void> {
    const requiredValues: any = {
      Title: fileName,
      RequestNo: requestNo,
      RequestItemId: itemId
    };

    const valuesWithDocumentType: any = this.copyValues(requiredValues);
    valuesWithDocumentType.DocumentType = documentType;

    return this.mergeLibraryFileMetadata(serverRelativeUrl, valuesWithDocumentType)
      .catch(() => {
        return this.mergeLibraryFileMetadata(serverRelativeUrl, requiredValues);
      });
  }

  private mergeLibraryFileMetadata(serverRelativeUrl: string, values: any): Promise<void> {
    const url: string = this.webUrl +
      "/_api/web/GetFileByServerRelativeUrl('" + this.escapeODataString(serverRelativeUrl) + "')" +
      "/ListItemAllFields";

    return this.getDocumentsLibraryListItemEntityTypeFullName()
      .then((entityTypeName: string) => {
        const updateBody: any = this.addODataMetadata(values, entityTypeName);
        const options: ISPHttpClientOptions = this.getJsonMergeOptions(updateBody);

        return this.props.context.spHttpClient
          .post(url, SPHttpClient.configurations.v1, options);
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(this.formatText('errorLibraryMetadata', text));
          });
        }
      });
  }

  private copyValues(values: any): any {
    const result: any = {};

    for (const key in values) {
      if (values.hasOwnProperty(key)) {
        result[key] = values[key];
      }
    }

    return result;
  }

  private replacePdfFileOnRequestItem(itemId: number, fileName: string, fileContent: ArrayBuffer): Promise<string> {
    const attachmentFileName: string = this.sanitizeSharePointAttachmentFileName(fileName);
    const collectionUrl: string = this.getRequestAttachmentCollectionApiUrl(itemId);
    const listUrl: string = collectionUrl + '?$select=FileName,ServerRelativeUrl';

    return this.props.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw this.createSharePointHttpError('List request attachments', response, text, listUrl);
          });
        }

        return response.json();
      })
      .then((data: any) => {
        const collection: any = data.d ? data.d : data;
        const attachments: any[] = collection.results || collection.value || [];
        let existingAttachment: any = null;

        for (let i: number = 0; i < attachments.length; i++) {
          const attachment: any = attachments[i];
          const existingFileName: string = attachment && attachment.FileName
            ? attachment.FileName.toLowerCase()
            : '';

          if (existingFileName === attachmentFileName.toLowerCase()) {
            existingAttachment = attachment;
            break;
          }
        }

        if (!existingAttachment) {
          return this.attachPdfFileToRequestItem(itemId, attachmentFileName, fileContent);
        }

        return this.updatePdfFileOnRequestItem(itemId, attachmentFileName, fileContent)
          .then(() => {
            return existingAttachment.ServerRelativeUrl
              ? this.getAbsoluteUrl(existingAttachment.ServerRelativeUrl)
              : '';
          });
      })
      .catch((error: any) => {
        const message: string = this.getErrorMessage(error);
        const localizedPrefix: string = this.t('errorListAttachment').replace('{0}', '');

        if (message.indexOf(localizedPrefix) === 0) {
          throw error;
        }

        throw new Error(this.formatText('errorListAttachment', message));
      });
  }

  private attachPdfFileToRequestItem(itemId: number, fileName: string, fileContent: ArrayBuffer): Promise<string> {
    const attachmentFileName: string = this.sanitizeSharePointAttachmentFileName(fileName);
    const attachmentUrl: string = this.getRequestAttachmentCollectionApiUrl(itemId) +
      "/add(FileName='" + this.encodeODataUrlValue(attachmentFileName) + "')";

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': 'application/pdf',
        'OData-Version': this.odataVersion
      },
      body: fileContent
    };

    return this.props.context.spHttpClient
      .post(attachmentUrl, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            const error: Error = this.createSharePointHttpError(
              'Add request list attachment',
              response,
              text,
              attachmentUrl
            );
            throw new Error(this.formatText('errorListAttachment', error.message));
          });
        }

        return response.text().then((text: string) => {
          if (!text) {
            return '';
          }

          try {
            const data: any = JSON.parse(text);
            const attachment: any = data.d ? data.d : data;

            if (attachment && attachment.ServerRelativeUrl) {
              return this.getAbsoluteUrl(attachment.ServerRelativeUrl);
            }
          } catch (e) {
            return '';
          }

          return '';
        });
      });
  }

  private updatePdfFileOnRequestItem(itemId: number, fileName: string, fileContent: ArrayBuffer): Promise<void> {
    const updateUrl: string = this.getRequestAttachmentCollectionApiUrl(itemId) +
      "('" + this.encodeODataUrlValue(fileName) + "')/$value";
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': 'application/pdf',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'PUT',
        'OData-Version': this.odataVersion
      },
      body: fileContent
    };

    return this.props.context.spHttpClient
      .post(updateUrl, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw this.createSharePointHttpError('Update request list attachment', response, text, updateUrl);
          });
        }
      });
  }

  private getRequestAttachmentCollectionApiUrl(itemId: number): string {
    return this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.encodeODataUrlValue(this.requestsListName) + "')" +
      '/items(' + itemId + ')/AttachmentFiles';
  }

  private sanitizeSharePointAttachmentFileName(fileName: string): string {
    let baseName: string = (fileName || 'attachment.pdf').replace(/\.pdf$/i, '');
    baseName = baseName.replace(/[^A-Za-z0-9_-]+/g, '-');
    baseName = baseName.replace(/-+/g, '-').replace(/^[-_]+/, '').replace(/[-_]+$/, '');

    if (!baseName) {
      baseName = 'attachment';
    }

    if (baseName.length > 116) {
      baseName = baseName.substring(0, 116).replace(/[-_]+$/, '');
    }

    return (baseName || 'attachment') + '.pdf';
  }

  private updatePdfMetadata(
    itemId: number,
    uploadResult: IPdfUploadResult,
    requestNo: string,
    existingPdfJson: string,
    updatePrimaryUrl: boolean,
    etag?: string
  ): Promise<string> {
    const pdfJson: string = this.buildPdfFileJson(uploadResult, requestNo, itemId, existingPdfJson);

    return this.updatePdfJson(itemId, pdfJson, etag)
      .then(() => {
        if (!updatePrimaryUrl) {
          return;
        }

        return this.updatePdfFileUrl(itemId, uploadResult.url, uploadResult.fileName);
      })
      .then(() => pdfJson);
  }

  private updatePdfJson(itemId: number, pdfJson: string, etag?: string): Promise<void> {
    const values: any = {};
    values[this.pdfJsonFieldName] = pdfJson;

    return this.updateItem(itemId, values, etag)
      .catch((error: any) => {
        throw new Error(this.formatText('errorPdfJson', this.pdfJsonFieldName, this.getErrorMessage(error)));
      });
  }

  private updatePdfFileUrl(itemId: number, pdfUrl: string, fileName: string): Promise<void> {
    return this.updateItem(itemId, {
      PDFFileUrl: {
        Url: pdfUrl,
        Description: fileName
      }
    }).catch(() => {
      return this.updateItem(itemId, {
        PDFFileUrl: pdfUrl
      });
    });
  }

  private buildPdfFileJson(
    uploadResult: IPdfUploadResult,
    requestNo: string,
    itemId: number,
    existingPdfJson: string
  ): string {
    const documents: IPdfDocumentRecord[] = this.getPdfDocumentsFromJson(existingPdfJson);
    const conversation: IWorkflowComment[] = this.getWorkflowConversationFromJson(existingPdfJson);
    const uploadedDocument: IPdfDocumentRecord = {
      role: uploadResult.role,
      documentType: uploadResult.documentType,
      fileName: uploadResult.fileName,
      originalFileName: uploadResult.originalFileName,
      contentType: uploadResult.contentType,
      size: uploadResult.size,
      uploadedAt: new Date().toISOString(),
      libraryUrl: uploadResult.url,
      serverRelativeUrl: uploadResult.serverRelativeUrl,
      attachmentUrl: uploadResult.attachmentUrl
    };
    let replaced: boolean = false;

    for (let i = 0; i < documents.length; i++) {
      if (documents[i].role === uploadResult.role) {
        documents[i] = uploadedDocument;
        replaced = true;
        break;
      }
    }

    if (!replaced) {
      documents.push(uploadedDocument);
    }

    return JSON.stringify({
      schemaVersion: 3,
      type: 'GM_REQUEST_PDF_COLLECTION',
      requestItemId: itemId,
      requestNo: requestNo,
      documents: documents,
      conversation: conversation
    });
  }

  private appendWorkflowComment(
    item: IRequestItem,
    role: string,
    kind: string,
    nextStatus: string,
    text: string
  ): string {
    let conversation: IWorkflowComment[] = this.getWorkflowConversationFromJson(item.PDFFileJson || '');

    if (conversation.length === 0) {
      conversation = this.getLegacyWorkflowComments(item);
    }

    conversation.push({
      id: this.buildWorkflowCommentId(),
      createdAt: new Date().toISOString(),
      authorId: this.state.currentUserId,
      authorName: this.state.currentUserTitle,
      role: role,
      kind: kind,
      fromStatus: item.Status || '',
      toStatus: nextStatus || '',
      text: text,
      legacy: false
    });

    return JSON.stringify({
      schemaVersion: 3,
      type: 'GM_REQUEST_PDF_COLLECTION',
      requestItemId: item.Id,
      requestNo: item.RequestNo,
      documents: this.getPdfDocuments(item),
      conversation: conversation
    });
  }

  private buildWorkflowCommentId(): string {
    return 'comment-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000000);
  }

  private getWorkflowConversationFromJson(pdfJson: string): IWorkflowComment[] {
    if (!pdfJson) {
      return [];
    }

    try {
      const parsed: any = JSON.parse(pdfJson);
      const rawConversation: any[] = parsed && parsed.conversation;
      const conversation: IWorkflowComment[] = [];

      if (!rawConversation || rawConversation.length === undefined) {
        return conversation;
      }

      for (let i = 0; i < rawConversation.length; i++) {
        const value: any = rawConversation[i] || {};
        if (!value.text) {
          continue;
        }

        conversation.push({
          id: value.id || ('legacy-json-' + i),
          createdAt: value.createdAt || '',
          authorId: value.authorId || 0,
          authorName: value.authorName || '',
          role: value.role || '',
          kind: value.kind || 'comment',
          fromStatus: value.fromStatus || '',
          toStatus: value.toStatus || '',
          text: value.text,
          legacy: value.legacy === true
        });
      }

      return conversation;
    } catch (e) {
      return [];
    }
  }

  private getLegacyWorkflowComments(item: IRequestItem): IWorkflowComment[] {
    const comments: IWorkflowComment[] = [];
    this.addLegacyWorkflowComment(comments, item.OfficeManagerComment, 'officeManager', 'legacy-office-manager');
    this.addLegacyWorkflowComment(comments, item.HODComment, 'hod', 'legacy-hod');
    this.addLegacyWorkflowComment(comments, item.GMComment, 'gm', 'legacy-gm');
    return comments;
  }

  private addLegacyWorkflowComment(
    comments: IWorkflowComment[],
    text: string,
    role: string,
    id: string
  ): void {
    const cleanText: string = (text || '').replace(/^\s+|\s+$/g, '');
    if (!cleanText) {
      return;
    }

    comments.push({
      id: id,
      createdAt: '',
      authorId: 0,
      authorName: this.getWorkflowRoleText(role),
      role: role,
      kind: 'legacy',
      fromStatus: '',
      toStatus: '',
      text: cleanText,
      legacy: true
    });
  }

  private getWorkflowConversation(item: IRequestItem): IWorkflowComment[] {
    const conversation: IWorkflowComment[] = this.getWorkflowConversationFromJson(item.PDFFileJson || '');
    return conversation.length > 0 ? conversation : this.getLegacyWorkflowComments(item);
  }

  private getWorkflowRoleText(role: string): string {
    switch (role) {
      case 'officeManager':
        return this.t('roleOfficeManager');
      case 'hod':
        return this.t('roleHod');
      case 'secretary':
        return this.t('roleSecretary');
      case 'gm':
        return this.t('roleGm');
      default:
        return role || this.t('legacyComment');
    }
  }

  private getPdfDocumentsFromJson(pdfJson: string): IPdfDocumentRecord[] {
    if (!pdfJson) {
      return [];
    }

    try {
      const parsed: any = JSON.parse(pdfJson);
      const rawDocuments: any[] = parsed && (parsed.documents || parsed.files);
      const documents: IPdfDocumentRecord[] = [];

      if (rawDocuments && rawDocuments.length !== undefined) {
        for (let i = 0; i < rawDocuments.length; i++) {
          documents.push(this.normalizePdfDocument(rawDocuments[i], ''));
        }
        return documents;
      }

      if (parsed && (parsed.fileName || parsed.libraryUrl || parsed.serverRelativeUrl)) {
        documents.push(this.normalizePdfDocument(parsed, 'request'));
      }

      return documents;
    } catch (e) {
      return [];
    }
  }

  private normalizePdfDocument(value: any, defaultRole: string): IPdfDocumentRecord {
    const documentType: string = value.documentType || value.DocumentType || '';
    let role: string = value.role || value.documentRole || defaultRole;

    if (!role) {
      const lowerDocumentType: string = documentType.toLowerCase();
      if (lowerDocumentType.indexOf('workflow') >= 0 || lowerDocumentType.indexOf('office manager') >= 0) {
        role = 'officeManager';
      } else if (lowerDocumentType.indexOf('secretary') >= 0) {
        role = 'secretary';
      } else if (lowerDocumentType.indexOf('hod') >= 0) {
        role = 'hod';
      } else {
        role = 'request';
      }
    }

    let fallbackDocumentType: string = 'Request PDF';
    if (role === 'officeManager') {
      fallbackDocumentType = 'Workflow PDF';
    } else if (role === 'hod') {
      fallbackDocumentType = 'HOD PDF';
    } else if (role === 'secretary') {
      fallbackDocumentType = 'Secretary PDF';
    }

    return {
      role: role,
      documentType: documentType || fallbackDocumentType,
      fileName: value.fileName || '',
      originalFileName: value.originalFileName || value.fileName || '',
      contentType: value.contentType || 'application/pdf',
      size: value.size || 0,
      uploadedAt: value.uploadedAt || '',
      libraryUrl: value.libraryUrl || value.url || '',
      serverRelativeUrl: value.serverRelativeUrl || '',
      attachmentUrl: value.attachmentUrl || ''
    };
  }

  private readFileAsArrayBuffer(file: any): Promise<ArrayBuffer> {
    return new Promise<ArrayBuffer>((resolve: (value: ArrayBuffer) => void, reject: (reason: any) => void) => {
      const reader: FileReader = new FileReader();

      reader.onload = (event: any) => {
        resolve(event.target.result);
      };

      reader.onerror = () => {
        reject(new Error(this.t('errorReadPdf')));
      };

      reader.readAsArrayBuffer(file);
    });
  }

  private getRequestDocumentFolderServerRelativeUrl(requestNo: string): Promise<string> {
    const folderName: string = this.sanitizeSharePointFolderName(requestNo);
    const folderApiUrl: string = this.getDocumentsLibraryFolderApiUrl(folderName);
    const getFolderUrl: string = folderApiUrl + '?$select=ServerRelativeUrl';

    return this.props.context.spHttpClient
      .get(getFolderUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return this.readFolderServerRelativeUrl(response);
        }

        if (response.status !== 404) {
          return response.text().then((text: string) => {
            throw this.createSharePointHttpError('Open request folder', response, text, getFolderUrl);
          });
        }

        return this.createRequestDocumentFolder(folderName, getFolderUrl);
      })
      .catch((error: any) => {
        throw new Error(this.formatText('errorRequestFolder', this.getErrorMessage(error)));
      });
  }

  private createRequestDocumentFolder(folderName: string, getFolderUrl: string): Promise<string> {
    const createFolderUrl: string = this.getDocumentsLibraryRootFolderApiUrl() +
      "/Folders/Add(url='" + this.encodeODataUrlValue(folderName) + "')" +
      "?$select=ServerRelativeUrl";
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': this.odataJsonHeader,
        'OData-Version': this.odataVersion
      }
    };

    return this.props.context.spHttpClient
      .post(createFolderUrl, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return this.readFolderServerRelativeUrl(response)
            .catch(() => this.getVerifiedFolderServerRelativeUrl(getFolderUrl));
        }

        return response.text().then((text: string) => {
          return this.getVerifiedFolderServerRelativeUrl(getFolderUrl)
            .catch(() => {
              throw this.createSharePointHttpError('Create request folder', response, text, createFolderUrl);
            });
        });
      });
  }

  private getVerifiedFolderServerRelativeUrl(getFolderUrl: string): Promise<string> {
    return this.props.context.spHttpClient
      .get(getFolderUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw this.createSharePointHttpError('Verify request folder', response, text, getFolderUrl);
          });
        }

        return this.readFolderServerRelativeUrl(response);
      });
  }

  private getDocumentsLibraryRootFolderApiUrl(): string {
    return this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.encodeODataUrlValue(this.documentsLibraryName) + "')" +
      '/RootFolder';
  }

  private getDocumentsLibraryFolderApiUrl(folderName: string): string {
    return this.getDocumentsLibraryRootFolderApiUrl() +
      "/Folders/GetByUrl('" + this.encodeODataUrlValue(folderName) + "')";
  }

  private encodeODataUrlValue(value: string): string {
    const escapedValue: string = this.escapeODataString(value);

    return encodeURIComponent(escapedValue)
      .replace(/!/g, '%21')
      .replace(/'/g, '%27')
      .replace(/\(/g, '%28')
      .replace(/\)/g, '%29')
      .replace(/\*/g, '%2A');
  }

  private createSharePointHttpError(
    operation: string,
    response: SPHttpClientResponse,
    responseText: string,
    requestUrl: string
  ): Error {
    let detail: string = responseText || 'SharePoint returned an empty response.';

    try {
      const parsed: any = JSON.parse(responseText);
      const sharePointError: any = parsed.error || parsed['odata.error'];

      if (sharePointError) {
        const code: string = sharePointError.code || '';
        const errorMessage: any = sharePointError.message;
        const message: string = errorMessage && errorMessage.value
          ? errorMessage.value
          : (typeof errorMessage === 'string' ? errorMessage : '');
        detail = (code ? code + ': ' : '') + (message || detail);
      }
    } catch (e) {
      detail = responseText || detail;
    }

    return new Error(
      operation + ' failed (HTTP ' + response.status +
      (response.statusText ? ' ' + response.statusText : '') + '). ' +
      detail + ' Request: ' + requestUrl
    );
  }

  private readFolderServerRelativeUrl(response: SPHttpClientResponse): Promise<string> {
    return response.json()
      .then((data: any) => {
        const folder: any = data.d ? data.d : data;

        if (!folder || !folder.ServerRelativeUrl) {
          throw new Error('SharePoint did not return a valid folder ServerRelativeUrl.');
        }

        return folder.ServerRelativeUrl;
      });
  }

  private sanitizeSharePointFolderName(requestNo: string): string {
    let result: string = (requestNo || 'REQUEST').replace(/[\\\/:*?"<>|#%]+/g, '-');
    result = result.replace(/^-+/, '').replace(/-+$/, '');
    return result || 'REQUEST';
  }

  private buildPdfFileName(requestNo: string, originalFileName: string): string {
    let cleanName: string = (originalFileName || 'request.pdf').replace(/\.pdf$/i, '');
    cleanName = cleanName.replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+/, '').replace(/-+$/, '');

    if (!cleanName) {
      cleanName = 'request';
    }

    return requestNo + '-' + cleanName + '.pdf';
  }

  private escapeODataString(value: string): string {
    return (value || '').replace(/'/g, "''");
  }

  private getAbsoluteUrl(serverRelativeUrl: string): string {
    if (!serverRelativeUrl) {
      return '';
    }

    if (serverRelativeUrl.indexOf('http://') === 0 || serverRelativeUrl.indexOf('https://') === 0) {
      return serverRelativeUrl;
    }

    const protocolIndex: number = this.webUrl.indexOf('://');
    const hostStartIndex: number = protocolIndex >= 0 ? protocolIndex + 3 : 0;
    const firstPathSlashIndex: number = this.webUrl.indexOf('/', hostStartIndex);
    const hostUrl: string = firstPathSlashIndex >= 0 ? this.webUrl.substring(0, firstPathSlashIndex) : this.webUrl;

    return encodeURI(hostUrl + serverRelativeUrl);
  }

  private getRequestPdfUrl(item: IRequestItem): string {
    const requestDocument: IPdfDocumentRecord = this.getPdfDocumentByRole(item, 'request');

    if (requestDocument) {
      const documentUrl: string = this.getPdfDocumentUrl(requestDocument);
      if (documentUrl) {
        return documentUrl;
      }
    }

    const fieldUrl: string = this.getUrlValue(item.PDFFileUrl);

    if (fieldUrl) {
      return fieldUrl;
    }

    return this.getFirstAttachmentUrl(item);
  }

  private getOfficeManagerPdfUrl(item: IRequestItem): string {
    const documentRecord: IPdfDocumentRecord = this.getPdfDocumentByRole(item, 'officeManager');
    return documentRecord ? this.getPdfDocumentUrl(documentRecord) : '';
  }

  private getSignedPdfUrl(item: IRequestItem): string {
    if (!item) {
      return '';
    }

    const signedUrl: string = this.getUrlValue(item.SignedPDFUrl);
    if (signedUrl) {
      return signedUrl;
    }

    const signedStatus: boolean =
      item.Status === 'GM Signed Pending Office Manager Confirmation' ||
      item.Status === 'Pending Secretary Information' ||
      item.Status === 'Pending HOD Information' ||
      item.Status === 'Approved by GM' ||
      item.Status === 'Closed' ||
      item.GMApprovalDetected === true;

    if (!signedStatus) {
      return '';
    }

    return this.getUrlValue(item.PDFFileUrl) || this.getOfficeManagerPdfUrl(item) || this.getRequestPdfUrl(item);
  }

  private hasPrimaryWorkflowPdf(item: IRequestItem): boolean {
    return !!(this.getOfficeManagerPdfUrl(item) || this.getRequestPdfUrl(item));
  }

  private getPdfDocuments(item: IRequestItem): IPdfDocumentRecord[] {
    if (!item) {
      return [];
    }

    const documents: IPdfDocumentRecord[] = this.getPdfDocumentsFromJson(item.PDFFileJson || '');
    let hasRequestDocument: boolean = false;

    for (let i = 0; i < documents.length; i++) {
      if (documents[i].role === 'request') {
        hasRequestDocument = true;
        break;
      }
    }

    let legacyUrl: string = this.getFirstAttachmentUrl(item);
    if (!legacyUrl && !this.getPdfDocumentByRoleFromCollection(documents, 'officeManager')) {
      legacyUrl = this.getUrlValue(item.PDFFileUrl);
    }
    if (!hasRequestDocument && legacyUrl) {
      documents.unshift({
        role: 'request',
        documentType: 'Request PDF',
        fileName: this.getFileNameFromUrl(legacyUrl),
        originalFileName: '',
        contentType: 'application/pdf',
        size: 0,
        uploadedAt: '',
        libraryUrl: legacyUrl,
        serverRelativeUrl: '',
        attachmentUrl: ''
      });
    }

    return documents;
  }

  private getPdfDocumentByRoleFromCollection(
    documents: IPdfDocumentRecord[],
    role: string
  ): IPdfDocumentRecord {
    for (let i = 0; i < documents.length; i++) {
      if (documents[i].role === role) {
        return documents[i];
      }
    }

    return null;
  }

  private getPdfDocumentByRole(item: IRequestItem, role: string): IPdfDocumentRecord {
    const documents: IPdfDocumentRecord[] = this.getPdfDocuments(item);
    return this.getPdfDocumentByRoleFromCollection(documents, role);
  }

  private getPdfDocumentUrl(documentRecord: IPdfDocumentRecord): string {
    if (!documentRecord) {
      return '';
    }

    if (documentRecord.libraryUrl) {
      return documentRecord.libraryUrl;
    }

    if (documentRecord.serverRelativeUrl) {
      return this.getAbsoluteUrl(documentRecord.serverRelativeUrl);
    }

    return documentRecord.attachmentUrl || '';
  }

  private getFileNameFromUrl(url: string): string {
    if (!url) {
      return '';
    }

    const cleanUrl: string = url.split('?')[0];
    const segments: string[] = cleanUrl.split('/');

    try {
      return decodeURIComponent(segments[segments.length - 1] || '');
    } catch (e) {
      return segments[segments.length - 1] || '';
    }
  }

  private getFirstAttachmentUrl(item: IRequestItem): string {
    if (!item || !item.AttachmentFiles) {
      return '';
    }

    const attachmentFiles: any = item.AttachmentFiles;
    const results: any[] = attachmentFiles.results || attachmentFiles.value || attachmentFiles;

    if (!results || results.length === 0) {
      return '';
    }

    for (let i = 0; i < results.length; i++) {
      const attachment: any = results[i];
      const attachmentFileName: string = (attachment.FileName || '').toLowerCase();
      const serverRelativeUrl: string = attachment.ServerRelativeUrl || '';
      const workflowFileName: string = this.sanitizeSharePointAttachmentFileName((item.RequestNo || '') + '.pdf').toLowerCase();
      const isRoleAttachment: boolean =
        attachmentFileName === workflowFileName ||
        attachmentFileName.indexOf('officemanager') === 0 ||
        attachmentFileName.indexOf('hod-') === 0 ||
        attachmentFileName.indexOf('secretary-') === 0;

      if (serverRelativeUrl && !isRoleAttachment) {
        return this.getAbsoluteUrl(serverRelativeUrl);
      }
    }

    return '';
  }

  private openRoleDocumentEditor(item: IRequestItem, role: PdfRole): void {
    this.officeManagerDocumentHtml = this.buildRoleDocumentTemplate(item, role);
    this.setState({
      showOfficeManagerEditor: true,
      activeDocumentRole: role,
      documentSourceMode: 'create',
      error: '',
      message: ''
    }, () => {
      if (this.officeManagerEditor) {
        this.officeManagerEditor.innerHTML = this.officeManagerDocumentHtml;
        this.officeManagerEditor.focus();
      }
    });
  }

  private closeOfficeManagerDocumentEditor(): void {
    this.setState({
      showOfficeManagerEditor: false
    });
  }

  private buildRoleDocumentTemplate(item: IRequestItem, role: PdfRole): string {
    const title: string = this.escapeHtml(item.Title || '');
    const requestNo: string = this.escapeHtml(item.RequestNo || '');
    const department: string = this.escapeHtml(this.getDepartmentText(item.Department || ''));
    const createdBy: string = this.escapeHtml(item.Author ? item.Author.Title : '');
    const requestDetails: string = this.escapeHtml(item.RequestDetails || '').replace(/\r?\n/g, '<br />');
    const roleComment: string = this.escapeHtml(this.getRoleComment(role)).replace(/\r?\n/g, '<br />');
    const documentTitle: string = this.getRoleDocumentTitle(role);
    const commentTitle: string = role === 'officeManager'
      ? this.t('officeManagerComment')
      : (role === 'hod' ? this.t('hodComment') : this.t('secretaryComment'));

    return '' +
      '<div>' +
        '<p style="margin:0 0 8px;color:#107c41;font-size:12px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;">' +
          this.escapeHtml(documentTitle) +
        '</p>' +
        '<h1 style="margin:0 0 20px;color:#102a43;font-size:28px;line-height:1.25;">' + title + '</h1>' +
        '<table style="width:100%;margin:0 0 24px;border-collapse:collapse;font-size:13px;">' +
          '<tbody>' +
            '<tr>' +
              '<td style="width:32%;padding:9px;border:1px solid #d9e2ec;background:#f4f7f6;font-weight:700;">' + this.escapeHtml(this.t('requestNo')) + '</td>' +
              '<td style="padding:9px;border:1px solid #d9e2ec;">' + requestNo + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding:9px;border:1px solid #d9e2ec;background:#f4f7f6;font-weight:700;">' + this.escapeHtml(this.t('department')) + '</td>' +
              '<td style="padding:9px;border:1px solid #d9e2ec;">' + department + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding:9px;border:1px solid #d9e2ec;background:#f4f7f6;font-weight:700;">' + this.escapeHtml(this.t('createdBy')) + '</td>' +
              '<td style="padding:9px;border:1px solid #d9e2ec;">' + createdBy + '</td>' +
            '</tr>' +
          '</tbody>' +
        '</table>' +
        '<h2 style="margin:0 0 8px;color:#102a43;font-size:18px;">' + this.escapeHtml(this.t('requestDetails')) + '</h2>' +
        '<p style="min-height:54px;margin:0 0 22px;line-height:1.65;">' + (requestDetails || '<br />') + '</p>' +
        '<h2 style="margin:0 0 8px;color:#102a43;font-size:18px;">' + this.escapeHtml(commentTitle) + '</h2>' +
        '<p style="min-height:92px;margin:0 0 22px;line-height:1.65;">' + (roleComment || '<br />') + '</p>' +
      '</div>';
  }

  private getRoleComment(role: PdfRole): string {
    if (role === 'hod') {
      return this.state.hodComment || '';
    }

    if (role === 'secretary') {
      return this.state.secretaryComment || '';
    }

    return this.state.officeManagerComment || '';
  }

  private getRoleDocumentTitle(role: string): string {
    if (role === 'officeManager') {
      return this.t('workflowPdf');
    }

    if (role === 'hod') {
      return this.t('hodPdf');
    }

    if (role === 'secretary') {
      return this.t('secretaryPdf');
    }

    return this.t('originalRequestPdf');
  }

  private applyEditorCommand(event: any, command: string, value?: string): void {
    if (event && event.preventDefault) {
      event.preventDefault();
    }

    if (!this.officeManagerEditor) {
      return;
    }

    this.officeManagerEditor.focus();
    document.execCommand(command, false, value || null);
    this.officeManagerDocumentHtml = this.officeManagerEditor.innerHTML;
  }

  private onEditorInput(): void {
    if (this.officeManagerEditor) {
      this.officeManagerDocumentHtml = this.officeManagerEditor.innerHTML;
    }
  }

  private createAndAttachRolePdf(role: PdfRole): void {
    const item: IRequestItem = this.state.selectedRequest;

    if (!item || !this.officeManagerEditor) {
      this.setState({ error: this.t('errorEditorUnavailable') });
      return;
    }

    const fileName: string = this.buildRolePdfFileName(item.RequestNo, role);
    const editorElement: HTMLDivElement = this.officeManagerEditor;
    this.officeManagerDocumentHtml = editorElement.innerHTML;

    this.setState({
      saving: true,
      error: '',
      message: this.t('generatingPdf')
    });

    this.createPdfBlob(editorElement, fileName)
      .then((pdfBlob: Blob) => {
        return this.readFileAsArrayBuffer(pdfBlob)
          .then((fileContent: ArrayBuffer) => {
            return this.saveRolePdfContent(
              item,
              role,
              fileContent,
              fileName,
              'application/pdf',
              pdfBlob.size
            );
          });
      })
      .catch((error: any) => {
        this.setState({
          saving: false,
          error: this.formatText('errorOfficeManagerPdf', this.getErrorMessage(error)),
          message: ''
        });
      });
  }

  private onRolePdfFileChange(event: any, role: PdfRole): void {
    const files: any = event.target.files;
    const file: any = files && files.length > 0 ? files[0] : null;

    if (file && !this.isPdfFile(file)) {
      if (this.rolePdfFileInput) {
        this.rolePdfFileInput.value = '';
      }

      this.setState({ rolePdfFile: null, error: this.t('errorPdfOnly') });
      return;
    }

    this.setState({
      rolePdfFile: file,
      activeDocumentRole: role,
      error: ''
    });
  }

  private attachSelectedRolePdf(role: PdfRole): void {
    const item: IRequestItem = this.state.selectedRequest;
    const file: any = this.state.rolePdfFile;

    if (!item || !file) {
      this.setState({ error: this.t('noPdfSelected'), message: '' });
      return;
    }

    this.setState({ saving: true, error: '', message: this.t('messageUploadingPdf') });

    this.readFileAsArrayBuffer(file)
      .then((fileContent: ArrayBuffer) => {
        return this.saveRolePdfContent(
          item,
          role,
          fileContent,
          file.name || this.buildRolePdfFileName(item.RequestNo, role),
          file.type || 'application/pdf',
          file.size || fileContent.byteLength
        );
      })
      .catch((error: any) => {
        this.setState({
          saving: false,
          error: this.formatText('errorOfficeManagerPdf', this.getErrorMessage(error)),
          message: ''
        });
      });
  }

  private saveRolePdfContent(
    item: IRequestItem,
    role: PdfRole,
    fileContent: ArrayBuffer,
    originalFileName: string,
    contentType: string,
    size: number
  ): Promise<void> {
    const fileName: string = this.buildRolePdfFileName(item.RequestNo, role);
    const documentType: string = this.getRoleDocumentType(role);
    let savedUploadResult: IPdfUploadResult = null;
    let freshItem: IRequestItem = item;

    return this.getRequestById(item.Id)
      .then((latestItem: IRequestItem) => {
        if (!this.isRoleDocumentStatusAllowed(role, latestItem.Status)) {
          throw new Error(this.t('errorRoleDocumentLocked'));
        }

        freshItem = latestItem;
        return this.uploadPdfFileToLibrary(fileName, fileContent, item.RequestNo, role, documentType);
      })
      .then((uploadResult: IPdfUploadResult) => {
        savedUploadResult = uploadResult;
        return this.updateLibraryFileMetadata(
          uploadResult.serverRelativeUrl,
          fileName,
          item.RequestNo,
          item.Id,
          documentType
        )
          .then(() => this.replacePdfFileOnRequestItem(item.Id, fileName, fileContent))
          .then((attachmentUrl: string) => {
            uploadResult.originalFileName = originalFileName || fileName;
            uploadResult.contentType = contentType || 'application/pdf';
            uploadResult.size = size || fileContent.byteLength;
            uploadResult.attachmentUrl = attachmentUrl;
            return this.getRequestById(item.Id);
          });
      })
      .then((latestItem: IRequestItem) => {
        if (!this.isRoleDocumentStatusAllowed(role, latestItem.Status)) {
          throw new Error(this.t('errorRoleDocumentLocked'));
        }

        freshItem = latestItem;
        return this.updatePdfMetadata(
          item.Id,
          savedUploadResult,
          item.RequestNo,
          this.buildPdfJsonSnapshot(latestItem),
          role === 'officeManager',
          (latestItem as any).__etag || ''
        );
      })
      .then((pdfJson: string) => {
        const updatedItem: IRequestItem = this.copyRequestItem(freshItem);
        updatedItem.PDFFileJson = pdfJson;
        if (role === 'officeManager') {
          updatedItem.PDFFileUrl = {
            Url: savedUploadResult.url,
            Description: savedUploadResult.fileName
          };
        }

        if (this.rolePdfFileInput) {
          this.rolePdfFileInput.value = '';
        }

        this.setState({
          saving: false,
          message: this.formatText('messageRolePdfAttached', this.getRoleDocumentTitle(role), fileName),
          selectedRequest: updatedItem,
          requests: this.replaceRequestInCollection(item.Id, updatedItem),
          showOfficeManagerEditor: false,
          rolePdfFile: null,
          activeDocumentRole: ''
        });
      });
  }

  private isRoleDocumentStatusAllowed(role: PdfRole, status: string): boolean {
    if (role === 'officeManager') {
      return status === 'Pending Office Manager' || status === 'On Hold' || status === 'Returned from HOD';
    }

    if (role === 'hod') {
      return status === 'Sent to HOD' || status === 'Pending HOD Information';
    }

    return status === 'Pending Secretary Information';
  }

  private getRoleDocumentType(role: PdfRole): string {
    if (role === 'officeManager') {
      return 'Workflow PDF';
    }

    return role === 'hod' ? 'HOD PDF' : 'Secretary PDF';
  }

  private createPdfBlob(element: HTMLElement, fileName: string): Promise<Blob> {
    const html2PdfModule: any = require('html2pdf.js');
    const html2Pdf: any = html2PdfModule.default || html2PdfModule;
    const captureElement: HTMLElement = element.cloneNode(true) as HTMLElement;
    captureElement.removeAttribute('contenteditable');
    captureElement.removeAttribute('role');
    captureElement.style.width = '100%';
    captureElement.style.minHeight = '277mm';
    captureElement.style.margin = '0';
    captureElement.style.padding = '48px';
    captureElement.style.boxSizing = 'border-box';
    captureElement.style.boxShadow = 'none';
    captureElement.style.background = '#ffffff';

    const options: any = {
      margin: [10, 10, 10, 10],
      filename: fileName,
      image: {
        type: 'jpeg',
        quality: 0.98
      },
      html2canvas: {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      },
      jsPDF: {
        unit: 'mm',
        format: 'a4',
        orientation: 'portrait'
      },
      pagebreak: {
        mode: ['css', 'legacy']
      }
    };

    return html2Pdf()
      .set(options)
      .from(captureElement)
      .toPdf()
      .outputPdf('blob')
      .then((blob: Blob) => blob)
      .catch((error: any) => {
        const overlay: Element = document.querySelector('.html2pdf__overlay');
        if (overlay && overlay.parentNode) {
          overlay.parentNode.removeChild(overlay);
        }
        throw error;
      });
  }

  private buildOfficeManagerPdfFileName(requestNo: string): string {
    return this.buildRolePdfFileName(requestNo, 'officeManager');
  }

  private buildRolePdfFileName(requestNo: string, role: PdfRole): string {
    const cleanRequestNo: string = this.sanitizeSharePointFolderName(requestNo);
    let prefix: string = '';

    if (role === 'hod') {
      prefix = 'HOD-';
    } else if (role === 'secretary') {
      prefix = 'Secretary-';
    }

    return this.sanitizeSharePointAttachmentFileName(prefix + cleanRequestNo + '.pdf');
  }

  private buildPdfJsonSnapshot(item: IRequestItem): string {
    return JSON.stringify({
      schemaVersion: 3,
      type: 'GM_REQUEST_PDF_COLLECTION',
      requestItemId: item.Id,
      requestNo: item.RequestNo,
      documents: this.getPdfDocuments(item),
      conversation: this.getWorkflowConversationFromJson(item.PDFFileJson || '')
    });
  }

  private copyRequestItem(item: IRequestItem): IRequestItem {
    const result: any = {};

    for (const key in item) {
      if (item.hasOwnProperty(key)) {
        result[key] = item[key];
      }
    }

    return result as IRequestItem;
  }

  private replaceRequestInCollection(itemId: number, replacement: IRequestItem): IRequestItem[] {
    const result: IRequestItem[] = this.state.requests.slice(0);

    for (let i = 0; i < result.length; i++) {
      if (result[i].Id === itemId) {
        result[i] = replacement;
        break;
      }
    }

    return result;
  }

  private escapeHtml(value: string): string {
    return (value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private formatFileSize(size: number): string {
    if (!size) {
      return '';
    }

    if (size < 1024 * 1024) {
      return Math.max(1, Math.round(size / 1024)) + ' KB';
    }

    return (Math.round((size / (1024 * 1024)) * 10) / 10) + ' MB';
  }

  private clearNewRequestForm(): void {
    if (this.pdfFileInput) {
      this.pdfFileInput.value = '';
    }

    this.setState({
      showNewRequestForm: false,
      newTitle: '',
      newDepartment: '',
      newDetails: '',
      newPdfFile: null,
      newPdfMode: 'upload',
      showNewRequestEditor: false,
      newRequestStatus: ''
    });
  }

  private getErrorMessage(error: any): string {
    let message: string = '';

    if (!error) {
      return this.t('errorUnknown');
    }

    if (error.message) {
      message = error.message;
    } else {
      message = error.toString();
    }

    return this.extractSharePointError(message);
  }

  private extractSharePointError(message: string): string {
    try {
      const parsed: any = JSON.parse(message);

      if (parsed.error && parsed.error.message) {
        if (parsed.error.message.value) {
          return parsed.error.message.value;
        }

        if (typeof parsed.error.message === 'string') {
          return parsed.error.message;
        }
      }

      if (parsed['odata.error'] && parsed['odata.error'].message && parsed['odata.error'].message.value) {
        return parsed['odata.error'].message.value;
      }
    } catch (e) {
      return message;
    }

    return message;
  }

  private getUrlValue(value: any): string {
    if (!value) {
      return '';
    }

    if (typeof value === 'string') {
      return value;
    }

    if (value.Url) {
      return value.Url;
    }

    return '';
  }

  public render(): React.ReactElement<IGmApprovalDashboardProps> {
    const isRequestOpen: boolean = !!this.state.selectedRequest;

    return (
      <div
        className={this.getRootClassName()}
        dir={this.state.language === 'ar' ? 'rtl' : 'ltr'}
        lang={this.state.language}
        aria-busy={this.state.saving || this.state.loading}
      >
        <div className={styles.header}>
          <div className={styles.headerTop}>
            <div className={styles.brandBlock}>
              <img className={styles.logoImage} src={this.besafeLogoUrl} alt="Besafe" />
              <div className={styles.headerText}>
                <h2>{this.t('appTitle')}</h2>
                <div className={styles.userMeta}>
                  <div>{this.t('appSubtitle')}</div>
                  <div>{this.t('currentUser')} <strong><span dir="auto">{this.state.currentUserTitle}</span></strong></div>
                  <div>{this.t('groups')} <span dir="auto">{this.getGroupsText()}</span></div>
                </div>
              </div>
            </div>
            <div className={styles.languageSwitcher}>
              <button
                type="button"
                className={this.state.language === 'en' ? styles.languageButtonActive : styles.languageButton}
                onClick={() => this.setLanguage('en')}
                aria-pressed={this.state.language === 'en'}
              >
                {this.t('english')}
              </button>
              <button
                type="button"
                className={this.state.language === 'ar' ? styles.languageButtonActive : styles.languageButton}
                onClick={() => this.setLanguage('ar')}
                aria-pressed={this.state.language === 'ar'}
              >
                {this.t('arabic')}
              </button>
            </div>
          </div>
        </div>

        <div className={styles.content}>
          {this.state.error &&
            <div className={styles.error} role="alert">{this.state.error}</div>
          }

          {this.state.message &&
            <div className={styles.success} role="status" aria-live="polite">{this.state.message}</div>
          }

          {this.state.loading &&
            <div className={styles.loadingCard} role="status" aria-live="polite">{this.t('loadingRequests')}</div>
          }

          {!this.state.loading && !isRequestOpen && this.renderAccessMessage()}

          {!this.state.loading && !isRequestOpen && this.canCreateRequest() && this.renderNewRequestLauncher()}

          {!this.state.loading && !isRequestOpen && this.canCreateRequest() && this.renderNewRequestForm()}

          {!this.state.loading && !isRequestOpen && this.canCreateRequest() && this.renderRequestTable(this.t('myRequests'), this.getMyRequests())}

          {!this.state.loading && !isRequestOpen && this.state.isOfficeManager && this.renderRequestTable(this.t('officeManagerQueue'), this.getOfficeManagerRequests())}

          {!this.state.loading && !isRequestOpen && this.state.isHod && this.renderRequestTable(this.t('hodQueue'), this.getHodRequests())}

          {!this.state.loading && !isRequestOpen && this.state.isSecretary && this.renderRequestTable(this.t('secretaryQueue'), this.getSecretaryRequests())}

          {!this.state.loading && this.renderDetailsPanel()}

          <div className={styles.footer}>
            {this.t('footer')} · v{this.buildVersion}
          </div>
        </div>
      </div>
    );
  }

  private renderAccessMessage(): JSX.Element {
    if (this.state.error) {
      return null;
    }

    if (this.state.isSecretary || this.state.isOfficeManager || this.state.isHod || this.state.isServiceAccount) {
      return null;
    }

    return (
      <div className={styles.warning} role="status">
        {this.t('noAccess')}
      </div>
    );
  }

  private renderNewRequestLauncher(): JSX.Element {
    if (this.state.showNewRequestForm) {
      return null;
    }

    return (
      <div className={styles.section + ' ' + styles.launcherSection}>
        <div className={styles.sectionTitleBlock}>
          <div className={styles.sectionEyebrow}>{this.t('newRequestEyebrow')}</div>
          <h3>{this.t('newRequest')}</h3>
          <div className={styles.sectionDescription}>{this.t('newRequestDescription')}</div>
        </div>
        <button
          type="button"
          className={styles.newRequestButton}
          disabled={this.state.saving}
          onClick={() => this.showNewRequestForm()}
        >
          {this.t('newRequest')}
        </button>
      </div>
    );
  }

  private renderNewRequestRouting(): JSX.Element {
    if (this.state.isOfficeManager) {
      return (
        <div className={styles.formRow}>
          <label htmlFor="gm-new-request-route">{this.t('routeTo')}</label>
          <select
            id="gm-new-request-route"
            value={this.state.newRequestStatus || 'Sent to HOD'}
            onChange={(e: any) => this.setState({ newRequestStatus: e.target.value })}
          >
            <option value="Sent to HOD">{this.t('routeToHod')}</option>
            <option value="Pending GM Signature">{this.t('routeToGmSignature')}</option>
          </select>
        </div>
      );
    }

    if (this.state.isHod) {
      return (
        <div className={styles.infoBox}>
          {this.t('routeTo')}: {this.t('routeToOfficeManager')}
        </div>
      );
    }

    return null;
  }

  private renderNewRequestForm(): JSX.Element {
    if (!this.state.showNewRequestForm) {
      return null;
    }

    return (
      <div className={styles.section}>
        <div className={styles.sectionHeader}>
          <div className={styles.sectionTitleBlock}>
            <div className={styles.sectionEyebrow}>{this.t('newRequestEyebrow')}</div>
            <h3>{this.t('newRequest')}</h3>
            <div className={styles.sectionDescription}>{this.t('newRequestDescription')}</div>
          </div>
        </div>

        <div className={styles.formGrid}>
          <div className={styles.formRow}>
            <label htmlFor="gm-new-request-title">{this.t('title')}</label>
            <input
              id="gm-new-request-title"
              type="text"
              value={this.state.newTitle}
              disabled={this.state.saving}
              onChange={(e: any) => this.updateNewRequestField('newTitle', e.target.value)}
            />
          </div>

          <div className={styles.formRow}>
            <label htmlFor="gm-new-request-department">{this.t('department')}</label>
            <select
              id="gm-new-request-department"
              value={this.state.newDepartment}
              disabled={this.state.saving}
              onChange={(e: any) => this.updateNewRequestField('newDepartment', e.target.value)}
            >
              <option value="">{this.t('selectDepartment')}</option>
              {this.departmentOptions.map((department: string) => {
                return (
                  <option key={department} value={department}>{this.getDepartmentText(department)}</option>
                );
              })}
            </select>
          </div>

          <div className={styles.formRowWide}>
            <label htmlFor="gm-new-request-details">{this.t('requestDetails')}</label>
            <textarea
              id="gm-new-request-details"
              value={this.state.newDetails}
              disabled={this.state.saving}
              onChange={(e: any) => this.updateNewRequestField('newDetails', e.target.value)}
            />
          </div>

          {this.renderNewRequestRouting()}

          <div className={styles.formRowWide}>
            <label>{this.t('pdfAttachment')}</label>
            <div className={styles.sourceToggle} role="group" aria-label={this.t('pdfSource')}>
              <button
                type="button"
                className={this.state.newPdfMode === 'upload' ? styles.sourceButtonActive : styles.sourceButton}
                disabled={this.state.saving}
                aria-pressed={this.state.newPdfMode === 'upload'}
                onClick={() => this.setNewRequestPdfMode('upload')}
              >
                {this.t('uploadPdf')}
              </button>
              <button
                type="button"
                className={this.state.newPdfMode === 'create' ? styles.sourceButtonActive : styles.sourceButton}
                disabled={this.state.saving}
                aria-pressed={this.state.newPdfMode === 'create'}
                onClick={() => this.setNewRequestPdfMode('create')}
              >
                {this.t('createWithWord')}
              </button>
            </div>

            {this.state.newPdfMode === 'upload' &&
              <div>
                <input
                  id="gm-new-request-pdf"
                  className={styles.hiddenFileInput}
                  type="file"
                  accept="application/pdf,.pdf"
                  ref={(input: HTMLInputElement) => { this.pdfFileInput = input; }}
                  onChange={(e: any) => this.onPdfFileChange(e)}
                />
                <label className={styles.filePicker} htmlFor="gm-new-request-pdf">
                  <span className={styles.filePickerIcon}>PDF</span>
                  <span className={styles.filePickerBody}>
                    <strong>{this.state.newPdfFile ? this.t('replacePdf') : this.t('choosePdf')}</strong>
                    <span>{this.t('pdfFolderHelp')}</span>
                  </span>
                </label>
              </div>
            }

            {this.state.newPdfMode === 'create' &&
              <div className={styles.editorShell}>
                <div className={styles.editorHeader}>
                  <span className={styles.wordIcon}>W</span>
                  <span className={styles.editorHeaderText}>{this.t('requestPdf')}</span>
                </div>
                <div className={styles.editorToolbar} role="toolbar" aria-label={this.t('requestPdf')}>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'bold')}><strong>B</strong></button>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'italic')}><em>I</em></button>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'underline')}><span style={{ textDecoration: 'underline' }}>U</span></button>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'insertUnorderedList')}>•</button>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'justifyLeft')}>⇤</button>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'justifyCenter')}>↔</button>
                  <button type="button" className={styles.toolButton} disabled={this.state.saving} onMouseDown={(e: any) => this.applyNewRequestEditorCommand(e, 'justifyRight')}>⇥</button>
                </div>
                <div
                  className={styles.documentPage}
                  ref={(element: HTMLDivElement) => { this.newRequestEditor = element; }}
                  contentEditable={!this.state.saving}
                  suppressContentEditableWarning={true}
                  role="textbox"
                  aria-multiline={true}
                  dir={this.state.language === 'ar' ? 'rtl' : 'ltr'}
                  onInput={() => this.onNewRequestEditorInput()}
                  dangerouslySetInnerHTML={{ __html: this.newRequestDocumentHtml }}
                />
                <div className={styles.editorActions}>
                  <button type="button" className={styles.buttonPrimary} disabled={this.state.saving} onClick={() => this.createNewRequestPdf()}>
                    {this.state.saving ? this.t('generatingPdf') : this.t('prepareRequestPdf')}
                  </button>
                </div>
              </div>
            }

            {this.state.newPdfFile &&
              <div className={styles.selectedFile}>
                <span className={styles.fileBadge}>PDF</span>
                <span>
                  {this.state.newPdfFile.name || 'request-draft.pdf'}
                  {this.formatFileSize(this.state.newPdfFile.size) && ' · ' + this.formatFileSize(this.state.newPdfFile.size)}
                </span>
              </div>
            }
            <div className={styles.helpText}>{this.t('pdfHelp')}</div>
          </div>
        </div>

        <div className={styles.formActions}>
          <button
            type="button"
            className={styles.buttonPrimary}
            disabled={this.state.saving}
            onClick={() => this.createRequest()}
          >
            {this.state.saving ? this.t('submitting') : this.getNewRequestSubmitText()}
          </button>

          <button
            type="button"
            className={styles.buttonSecondary}
            disabled={this.state.saving}
            onClick={() => this.cancelNewRequest()}
          >
            {this.t('cancel')}
          </button>
        </div>
      </div>
    );
  }

  private renderRequestTable(title: string, items: IRequestItem[]): JSX.Element {
    return (
      <div className={styles.section}>
        <div className={styles.sectionHeader}>
          <div className={styles.sectionTitleBlock}>
            <div className={styles.sectionEyebrow}>{this.t('requestOverview')}</div>
            <h3>{title}</h3>
          </div>
        </div>

        {items.length === 0 &&
          <div className={styles.emptyState}>{this.t('noRequestsFound')}</div>
        }

        {items.length > 0 &&
          <div className={styles.tableViewport} role="region" aria-label={title} tabIndex={0}>
            <table className={styles.table} aria-label={title}>
              <thead>
                <tr>
                  <th scope="col">{this.t('requestNo')}</th>
                  <th scope="col">{this.t('title')}</th>
                  <th scope="col">{this.t('department')}</th>
                  <th scope="col">{this.t('status')}</th>
                  <th scope="col">{this.t('createdBy')}</th>
                  <th scope="col">{this.t('pdf')}</th>
                  <th scope="col">{this.t('open')}</th>
                </tr>
              </thead>
              <tbody>
                {items.map((item: IRequestItem) => {
                  return (
                    <tr key={item.Id}>
                      <td><span className={styles.ltrValue} dir="ltr">{item.RequestNo}</span></td>
                      <td><span dir="auto">{item.Title}</span></td>
                      <td>{this.getDepartmentText(item.Department)}</td>
                      <td>{this.renderStatusBadge(item.Status)}</td>
                      <td><span dir="auto">{item.Author ? item.Author.Title : ''}</span></td>
                      <td>{this.renderRequestTablePdfLinks(item)}</td>
                      <td>
                        <button
                          type="button"
                          className={styles.openButton}
                          onClick={() => this.selectRequest(item)}
                        >
                          {this.t('open')}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        }
      </div>
    );
  }

  private renderRequestTablePdfLinks(item: IRequestItem): JSX.Element {
    const signedUrl: string = this.getSignedPdfUrl(item);
    const requestUrl: string = this.getRequestPdfUrl(item);
    const workflowUrl: string = this.getOfficeManagerPdfUrl(item);
    const showRequest: boolean = !!requestUrl && (!signedUrl || !this.areSameDocumentUrls(requestUrl, signedUrl));
    const showWorkflow: boolean = !!workflowUrl && !signedUrl && (!requestUrl || !this.areSameDocumentUrls(workflowUrl, requestUrl));

    return (
      <span>
        {signedUrl &&
          <a href={signedUrl} target="_blank" rel="noopener noreferrer">{this.t('signedPdfDocument')}</a>
        }
        {showRequest &&
          <span>
            {signedUrl && ' · '}
            <a href={requestUrl} target="_blank" rel="noopener noreferrer">{this.t('originalRequestPdf')}</a>
          </span>
        }
        {showWorkflow &&
          <span>
            {requestUrl && ' · '}
            <a href={workflowUrl} target="_blank" rel="noopener noreferrer">{this.t('workflowPdf')}</a>
          </span>
        }
      </span>
    );
  }

  private renderDetailsPanel(): JSX.Element {
    const item: IRequestItem = this.state.selectedRequest;

    if (!item) {
      return null;
    }

    return (
      <div className={styles.section}>
        <div className={styles.detailsHero}>
          <div className={styles.sectionTitleBlock}>
            <div className={styles.sectionEyebrow}>{this.t('requestOverview')}</div>
            <h3 className={styles.detailsTitle}><span dir="auto">{item.Title || this.t('requestDetailsTitle')}</span></h3>
            <div className={styles.sectionDescription}>{this.t('requestOverviewDescription')}</div>
          </div>
          {this.renderStatusBadge(item.Status)}
        </div>

        <div className={styles.metadataGrid}>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('id')}</div>
            <div className={styles.metadataValue}><span className={styles.ltrValue} dir="ltr">{item.Id}</span></div>
          </div>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('requestNoLabel')}</div>
            <div className={styles.metadataValue}><span className={styles.ltrValue} dir="ltr">{item.RequestNo}</span></div>
          </div>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('departmentLabel')}</div>
            <div className={styles.metadataValue}>{this.getDepartmentText(item.Department)}</div>
          </div>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('createdByLabel')}</div>
            <div className={styles.metadataValue}><span dir="auto">{item.Author ? item.Author.Title : ''}</span></div>
          </div>
        </div>

        <div className={styles.formRowWide}>
          <label htmlFor="gm-request-details-readonly">{this.t('requestDetails')}</label>
          <textarea id="gm-request-details-readonly" dir="auto" value={item.RequestDetails || ''} readOnly={true} />
        </div>

        {item.SharedFolderPath &&
          <div className={styles.infoBox}>
            {this.t('gmSharedFolderPath')} <span className={styles.ltrValue} dir="ltr">{item.SharedFolderPath}</span>
          </div>
        }

        <div className={styles.sectionHeader}>
          <div className={styles.sectionTitleBlock}>
            <div className={styles.sectionEyebrow}>{this.t('documents')}</div>
            <h3>{this.t('documents')}</h3>
          </div>
        </div>

        {this.renderRequestDocuments(item)}

        {this.renderWorkflowConversation(item)}

        {this.renderOfficeManagerActions(item)}
        {this.renderHodActions(item)}
        {this.renderSecretaryActions(item)}

        <div className={styles.formActions}>
          <button
            type="button"
            className={styles.buttonSecondary}
            disabled={this.state.saving}
            onClick={() => this.setState({ selectedRequest: null, showOfficeManagerEditor: false })}
          >
            {this.t('closeDetails')}
          </button>
        </div>
      </div>
    );
  }

  private renderRequestDocuments(item: IRequestItem): JSX.Element {
    const allDocuments: IPdfDocumentRecord[] = this.sortPdfDocuments(this.getPdfDocuments(item));
    const signedPdfUrl: string = this.getSignedPdfUrl(item);
    const documents: IPdfDocumentRecord[] = [];

    for (let i = 0; i < allDocuments.length; i++) {
      const recordUrl: string = this.getPdfDocumentUrl(allDocuments[i]);
      if (!signedPdfUrl || !this.areSameDocumentUrls(recordUrl, signedPdfUrl)) {
        documents.push(allDocuments[i]);
      }
    }

    if (documents.length === 0 && !signedPdfUrl) {
      return <div className={styles.emptyState}>{this.t('noDocuments')}</div>;
    }

    return (
      <div className={styles.documentGrid}>
        {signedPdfUrl &&
          <div className={styles.documentCard}>
            <div className={styles.documentIcon}>PDF</div>
            <div className={styles.documentInfo}>
              <div className={styles.documentTitle}>{this.t('signedPdfDocument')}</div>
              <div className={styles.documentName}>{this.getFileNameFromUrl(signedPdfUrl) || this.t('signedPdfDocument')}</div>
            </div>
            <a
              className={styles.documentLink}
              href={signedPdfUrl}
              target="_blank"
              rel="noopener noreferrer"
            >
              {this.t('openDocument')}
            </a>
          </div>
        }

        {documents.map((documentRecord: IPdfDocumentRecord, index: number) => {
          const documentUrl: string = this.getPdfDocumentUrl(documentRecord);
          const documentTitle: string = this.getRoleDocumentTitle(documentRecord.role);
          const fileName: string = documentRecord.fileName || this.getFileNameFromUrl(documentUrl) || documentTitle;

          return (
            <div className={styles.documentCard} key={documentRecord.role + '-' + index}>
              <div className={styles.documentIcon}>PDF</div>
              <div className={styles.documentInfo}>
                <div className={styles.documentTitle}>{documentTitle}</div>
                <div className={styles.documentName}>{fileName}</div>
                <div className={styles.documentMeta}>
                  {this.t('storedInRequestFolder')}
                  {this.formatFileSize(documentRecord.size) && ' · ' + this.formatFileSize(documentRecord.size)}
                </div>
              </div>
              {documentUrl &&
                <a
                  className={styles.documentLink}
                  href={documentUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  {this.t('openDocument')}
                </a>
              }
            </div>
          );
        })}
      </div>
    );
  }

  private sortPdfDocuments(documents: IPdfDocumentRecord[]): IPdfDocumentRecord[] {
    const result: IPdfDocumentRecord[] = documents.slice(0);
    const weights: any = {
      officeManager: 0,
      request: 1,
      hod: 2,
      secretary: 3
    };

    result.sort((left: IPdfDocumentRecord, right: IPdfDocumentRecord) => {
      const leftWeight: number = weights[left.role] === undefined ? 9 : weights[left.role];
      const rightWeight: number = weights[right.role] === undefined ? 9 : weights[right.role];
      return leftWeight - rightWeight;
    });

    return result;
  }

  private areSameDocumentUrls(left: string, right: string): boolean {
    if (!left || !right) {
      return false;
    }

    try {
      return decodeURI(left).split('?')[0].toLowerCase() === decodeURI(right).split('?')[0].toLowerCase();
    } catch (e) {
      return left.split('?')[0].toLowerCase() === right.split('?')[0].toLowerCase();
    }
  }

  private renderWorkflowConversation(item: IRequestItem): JSX.Element {
    const comments: IWorkflowComment[] = this.getWorkflowConversation(item);

    return (
      <div className={styles.conversationPanel}>
        <div className={styles.conversationHeader}>
          <h4>{this.t('conversationTitle')}</h4>
          <div>{this.t('conversationDescription')}</div>
        </div>

        {comments.length === 0 &&
          <div className={styles.emptyState}>{this.t('noComments')}</div>
        }

        {comments.length > 0 &&
          <div className={styles.conversationList}>
            {comments.map((comment: IWorkflowComment, index: number) => {
              return (
                <div className={styles.conversationMessage} key={comment.id || ('comment-' + index)}>
                  <div className={styles.conversationMeta}>
                    <strong><span dir="auto">{comment.authorName || this.getWorkflowRoleText(comment.role)}</span></strong>
                    <span>{this.getWorkflowRoleText(comment.role)}</span>
                    {comment.createdAt && <span className={styles.ltrValue} dir="ltr">{this.formatWorkflowCommentDate(comment.createdAt)}</span>}
                    {comment.legacy && <span>{this.t('legacyComment')}</span>}
                  </div>
                  <div className={styles.conversationText} dir="auto">{comment.text}</div>
                  {comment.fromStatus && comment.toStatus &&
                    <div className={styles.conversationTransition}>
                      {this.getStatusText(comment.fromStatus)} → {this.getStatusText(comment.toStatus)}
                    </div>
                  }
                </div>
              );
            })}
          </div>
        }
      </div>
    );
  }

  private formatWorkflowCommentDate(value: string): string {
    const date: Date = new Date(value);
    if (isNaN(date.getTime())) {
      return value;
    }

    return date.toLocaleString(this.state.language === 'ar' ? 'ar-EG' : 'en-US');
  }

  private renderOfficeManagerActions(item: IRequestItem): JSX.Element {
    const allowedStatuses: string[] = [
      'Pending Office Manager',
      'On Hold',
      'Returned from HOD',
      'GM Signed Pending Office Manager Confirmation',
      'Approved by GM'
    ];

    if (!this.state.isOfficeManager || allowedStatuses.indexOf(item.Status) < 0) {
      return null;
    }

    return (
      <div className={styles.actionPanel}>
        <div className={styles.actionPanelHeader}>
          <div className={styles.sectionTitleBlock}>
            <div className={styles.sectionEyebrow}>{this.t('officeManagerWorkspace')}</div>
            <h4>{this.t('officeManagerActions')}</h4>
            <div className={styles.sectionDescription}>{this.t('officeManagerWorkspaceDescription')}</div>
          </div>
          {this.renderStatusBadge(item.Status)}
        </div>

        <div className={styles.actionPanelBody}>
          <div className={styles.workspaceGrid}>
            <div className={styles.officeDocumentArea}>
              <div className={styles.officeDocumentIntro}>
                <strong>{this.t('officeManagerDocument')}</strong>
                {this.t('officeManagerDocumentDescription')}
              </div>
              {this.renderOfficeManagerDocumentEditor(item)}
            </div>

            <div>
              <div className={styles.formRow}>
                <label htmlFor="gm-office-manager-comment">{this.t('officeManagerComment')}</label>
                <textarea
                  id="gm-office-manager-comment"
                  value={this.state.officeManagerComment}
                  onChange={(e: any) => this.setState({ officeManagerComment: e.target.value })}
                />
              </div>

              {item.Status === 'Returned from HOD' &&
                <div className={styles.formRow}>
                  <label htmlFor="gm-hod-comment-readonly">{this.t('hodComment')}</label>
                  <textarea id="gm-hod-comment-readonly" value={item.HODComment || ''} readOnly={true} />
                </div>
              }

              {item.Status === 'GM Signed Pending Office Manager Confirmation' &&
                <div className={styles.formRow}>
                  <label htmlFor="gm-general-manager-comment">{this.t('gmCommentNotes')}</label>
                  <textarea
                    id="gm-general-manager-comment"
                    value={this.state.gmComment}
                    onChange={(e: any) => this.setState({ gmComment: e.target.value })}
                  />
                </div>
              }
            </div>
          </div>
        </div>

        {item.Status === 'Pending Office Manager' &&
          <div className={styles.actionFooter}>
            <button type="button" className={styles.buttonWarning} disabled={this.state.saving} onClick={() => this.officeManagerAction('On Hold')}>{this.t('putOnHold')}</button>
            <button type="button" className={styles.buttonSecondary} disabled={this.state.saving} onClick={() => this.officeManagerAction('Sent to HOD')}>{this.t('sendToHod')}</button>
            <button type="button" className={styles.buttonPrimary} disabled={this.state.saving} onClick={() => this.officeManagerAction('Pending GM Signature')}>{this.t('sendToGmSignature')}</button>
            <button type="button" className={styles.buttonDanger} disabled={this.state.saving} onClick={() => this.officeManagerAction('Rejected')}>{this.t('reject')}</button>
          </div>
        }

        {item.Status === 'On Hold' &&
          <div className={styles.actionFooter}>
            <button type="button" className={styles.buttonPrimary} disabled={this.state.saving} onClick={() => this.officeManagerAction('Pending Office Manager')}>{this.t('returnToPendingOfficeManager')}</button>
            <button type="button" className={styles.buttonDanger} disabled={this.state.saving} onClick={() => this.officeManagerAction('Rejected')}>{this.t('reject')}</button>
          </div>
        }

        {item.Status === 'Returned from HOD' &&
          <div className={styles.actionFooter}>
            <button type="button" className={styles.buttonPrimary} disabled={this.state.saving} onClick={() => this.officeManagerAction('Pending GM Signature')}>{this.t('sendToGmSignature')}</button>
            <button type="button" className={styles.buttonDanger} disabled={this.state.saving} onClick={() => this.officeManagerAction('Rejected')}>{this.t('reject')}</button>
          </div>
        }

        {item.Status === 'GM Signed Pending Office Manager Confirmation' &&
          <div className={styles.actionFooter}>
            <button type="button" className={styles.buttonPrimary} disabled={this.state.saving} onClick={() => this.confirmGmApproval()}>{this.t('confirmGmApproval')}</button>
            <button type="button" className={styles.buttonDanger} disabled={this.state.saving} onClick={() => this.gmSignedOfficeManagerAction('Rejected')}>{this.t('gmReject')}</button>
            <button type="button" className={styles.buttonSecondary} disabled={this.state.saving} onClick={() => this.gmSignedOfficeManagerAction('Pending Secretary Information')}>{this.t('requestSecretaryInformation')}</button>
            <button type="button" className={styles.buttonSecondary} disabled={this.state.saving} onClick={() => this.gmSignedOfficeManagerAction('Pending HOD Information')}>{this.t('requestHodInformation')}</button>
          </div>
        }

        {item.Status === 'Approved by GM' &&
          <div className={styles.actionFooter}>
            <button type="button" className={styles.buttonPrimary} disabled={this.state.saving} onClick={() => this.closeRequest()}>{this.t('closeRequest')}</button>
          </div>
        }
      </div>
    );
  }

  private renderOfficeManagerDocumentEditor(item: IRequestItem): JSX.Element {
    const canEdit: boolean =
      item.Status === 'Pending Office Manager' ||
      item.Status === 'On Hold' ||
      item.Status === 'Returned from HOD';

    return this.renderRoleDocumentComposer(item, 'officeManager', !canEdit);
  }

  private setRoleDocumentSource(role: PdfRole, mode: PdfSourceMode): void {
    if (this.rolePdfFileInput) {
      this.rolePdfFileInput.value = '';
    }

    this.officeManagerDocumentHtml = '';
    this.setState({
      documentSourceMode: mode,
      rolePdfFile: null,
      activeDocumentRole: role,
      showOfficeManagerEditor: false,
      error: '',
      message: ''
    });
  }

  private renderRoleDocumentComposer(item: IRequestItem, role: PdfRole, locked?: boolean): JSX.Element {
    if (locked) {
      return (
        <div className={styles.infoBox}>{this.t('documentFrozenAfterSignature')}</div>
      );
    }

    const isActiveRole: boolean = !this.state.activeDocumentRole || this.state.activeDocumentRole === role;
    const sourceMode: PdfSourceMode = isActiveRole ? this.state.documentSourceMode : 'create';
    const editorIsOpen: boolean = this.state.showOfficeManagerEditor && this.state.activeDocumentRole === role;
    const fileInputId: string = 'gm-role-pdf-' + role;

    return (
      <div>
        <div className={styles.sourceToggle} role="group" aria-label={this.t('pdfSource')}>
          <button
            type="button"
            className={sourceMode === 'upload' ? styles.sourceButtonActive : styles.sourceButton}
            disabled={this.state.saving}
            aria-pressed={sourceMode === 'upload'}
            onClick={() => this.setRoleDocumentSource(role, 'upload')}
          >
            {this.t('uploadPdf')}
          </button>
          <button
            type="button"
            className={sourceMode === 'create' ? styles.sourceButtonActive : styles.sourceButton}
            disabled={this.state.saving}
            aria-pressed={sourceMode === 'create'}
            onClick={() => this.setRoleDocumentSource(role, 'create')}
          >
            {this.t('createWithWord')}
          </button>
        </div>

        {sourceMode === 'upload' &&
          <div className={styles.roleUploadPanel}>
            <input
              id={fileInputId}
              className={styles.hiddenFileInput}
              type="file"
              accept="application/pdf,.pdf"
              ref={(input: HTMLInputElement) => { this.rolePdfFileInput = input; }}
              onChange={(e: any) => this.onRolePdfFileChange(e, role)}
            />
            <label className={styles.filePicker} htmlFor={fileInputId}>
              <span className={styles.filePickerIcon}>PDF</span>
              <span className={styles.filePickerBody}>
                <strong>{this.state.rolePdfFile && this.state.activeDocumentRole === role ? this.t('replacePdf') : this.t('uploadPdf')}</strong>
                <span>{this.buildRolePdfFileName(item.RequestNo, role)}</span>
              </span>
            </label>
            {this.state.rolePdfFile && this.state.activeDocumentRole === role &&
              <div className={styles.selectedFile}>
                <span className={styles.fileBadge}>PDF</span>
                <span>{this.state.rolePdfFile.name}</span>
              </div>
            }
            <button
              type="button"
              className={styles.buttonPrimary}
              disabled={this.state.saving || !this.state.rolePdfFile || this.state.activeDocumentRole !== role}
              onClick={() => this.attachSelectedRolePdf(role)}
            >
              {this.t('attachSelectedPdf')}
            </button>
          </div>
        }

        {sourceMode === 'create' && this.renderRoleDocumentEditor(item, role, editorIsOpen)}
      </div>
    );
  }

  private renderRoleDocumentEditor(item: IRequestItem, role: PdfRole, editorIsOpen: boolean): JSX.Element {
    if (!editorIsOpen) {
      return (
        <button
          type="button"
          className={styles.wordLaunchButton}
          disabled={this.state.saving}
          onClick={() => this.openRoleDocumentEditor(item, role)}
        >
          <span className={styles.wordIcon}>W</span>
          <span>{this.t('createWithWord')} · {this.getRoleDocumentTitle(role)}</span>
        </button>
      );
    }

    return (
      <div className={styles.editorShell}>
        <div className={styles.editorHeader}>
          <span className={styles.wordIcon}>W</span>
          <span className={styles.editorHeaderText}>
            {this.getRoleDocumentTitle(role)} · {this.buildRolePdfFileName(item.RequestNo, role)}
          </span>
        </div>

        <div className={styles.editorToolbar} role="toolbar" aria-label={this.t('editorTitle')}>
          <select
            className={styles.toolSelect}
            defaultValue="P"
            title={this.t('paragraph')}
            aria-label={this.t('paragraph')}
            disabled={this.state.saving}
            onChange={(e: any) => this.applyEditorCommand(e, 'formatBlock', e.target.value)}
          >
            <option value="P">{this.t('paragraph')}</option>
            <option value="H2">{this.t('heading')}</option>
          </select>
          <button type="button" className={styles.toolButton} title={this.t('bold')} aria-label={this.t('bold')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'bold')}><strong>B</strong></button>
          <button type="button" className={styles.toolButton} title={this.t('italic')} aria-label={this.t('italic')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'italic')}><em>I</em></button>
          <button type="button" className={styles.toolButton} title={this.t('underline')} aria-label={this.t('underline')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'underline')}><span style={{ textDecoration: 'underline' }}>U</span></button>
          <button type="button" className={styles.toolButton} title={this.t('bullets')} aria-label={this.t('bullets')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'insertUnorderedList')}>•</button>
          <button type="button" className={styles.toolButton} title={this.t('alignLeft')} aria-label={this.t('alignLeft')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'justifyLeft')}>⇤</button>
          <button type="button" className={styles.toolButton} title={this.t('alignCenter')} aria-label={this.t('alignCenter')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'justifyCenter')}>↔</button>
          <button type="button" className={styles.toolButton} title={this.t('alignRight')} aria-label={this.t('alignRight')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'justifyRight')}>⇥</button>
          <button type="button" className={styles.toolButton} title={this.t('undo')} aria-label={this.t('undo')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'undo')}>↶</button>
          <button type="button" className={styles.toolButton} title={this.t('redo')} aria-label={this.t('redo')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'redo')}>↷</button>
        </div>

        <div
          className={styles.documentPage}
          ref={(element: HTMLDivElement) => { this.officeManagerEditor = element; }}
          contentEditable={!this.state.saving}
          suppressContentEditableWarning={true}
          role="textbox"
          aria-multiline={true}
          aria-label={this.t('editorTitle')}
          dir={this.state.language === 'ar' ? 'rtl' : 'ltr'}
          spellCheck={true}
          onInput={() => this.onEditorInput()}
          dangerouslySetInnerHTML={{ __html: this.officeManagerDocumentHtml }}
        />

        <div className={styles.editorActions}>
          <button
            type="button"
            className={styles.buttonPrimary}
            disabled={this.state.saving}
            onClick={() => this.createAndAttachRolePdf(role)}
          >
            {this.state.saving ? this.t('generatingPdf') : this.t('createAttachPdf')}
          </button>
          <button
            type="button"
            className={styles.buttonSecondary}
            disabled={this.state.saving}
            onClick={() => this.closeOfficeManagerDocumentEditor()}
          >
            {this.t('closeEditor')}
          </button>
        </div>
      </div>
    );
  }

  private renderHodActions(item: IRequestItem): JSX.Element {
    if (!this.state.isHod || (item.Status !== 'Sent to HOD' && item.Status !== 'Pending HOD Information')) {
      return null;
    }

    return (
      <div className={styles.actionPanel}>
        <div className={styles.actionPanelHeader}>
          <h4>{this.t('hodActions')}</h4>
          {this.renderStatusBadge(item.Status)}
        </div>

        <div className={styles.actionPanelBody}>
          <div className={styles.workspaceGrid}>
            <div className={styles.officeDocumentArea}>
              <div className={styles.officeDocumentIntro}>
                <strong>{this.t('hodPdf')}</strong>
                {this.t('officeManagerDocumentDescription')}
              </div>
              {this.renderRoleDocumentComposer(item, 'hod')}
            </div>

            <div>
              <div className={styles.formRow}>
                <label htmlFor="gm-hod-comment">{this.t('hodComment')}</label>
                <textarea
                  id="gm-hod-comment"
                  value={this.state.hodComment}
                  onChange={(e: any) => this.setState({ hodComment: e.target.value })}
                />
              </div>
            </div>
          </div>
        </div>

        <div className={styles.actionFooter}>
          <button
            type="button"
            className={styles.buttonPrimary}
            disabled={this.state.saving}
            onClick={() => this.hodReturnToOfficeManager()}
          >
            {item.Status === 'Pending HOD Information' ? this.t('submitInformation') : this.t('returnToOfficeManager')}
          </button>
        </div>
      </div>
    );
  }

  private renderSecretaryActions(item: IRequestItem): JSX.Element {
    if (!this.state.isSecretary || item.Status !== 'Pending Secretary Information') {
      return null;
    }

    return (
      <div className={styles.actionPanel}>
        <div className={styles.actionPanelHeader}>
          <h4>{this.t('secretaryActions')}</h4>
          {this.renderStatusBadge(item.Status)}
        </div>

        <div className={styles.actionPanelBody}>
          <div className={styles.workspaceGrid}>
            <div className={styles.officeDocumentArea}>
              <div className={styles.officeDocumentIntro}>
                <strong>{this.t('secretaryPdf')}</strong>
                {this.t('officeManagerDocumentDescription')}
              </div>
              {this.renderRoleDocumentComposer(item, 'secretary')}
            </div>

            <div className={styles.formRow}>
              <label htmlFor="gm-secretary-comment">{this.t('secretaryComment')}</label>
              <textarea
                id="gm-secretary-comment"
                value={this.state.secretaryComment}
                onChange={(e: any) => this.setState({ secretaryComment: e.target.value })}
              />
            </div>
          </div>
        </div>

        <div className={styles.actionFooter}>
          <button
            type="button"
            className={styles.buttonPrimary}
            disabled={this.state.saving}
            onClick={() => this.secretaryReturnToOfficeManager()}
          >
            {this.t('submitInformation')}
          </button>
        </div>
      </div>
    );
  }
}
