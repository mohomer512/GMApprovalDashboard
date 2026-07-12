import * as React from 'react';
import styles from './GmApprovalDashboard.module.scss';
import { IGmApprovalDashboardProps } from './IGmApprovalDashboardProps';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export type LanguageCode = 'en' | 'ar';

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
  newRequestStatus: string;

  officeManagerComment: string;
  hodComment: string;
  gmComment: string;
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

export default class GmApprovalDashboard extends React.Component<IGmApprovalDashboardProps, IGmApprovalDashboardState> {

  private gmSiteUrl: string = 'http://spse26h/GM';
  private requestsListName: string = 'GM Requests';
  private documentsLibraryName: string = 'GM Approval Documents';
  private documentsLibraryServerRelativeUrl: string = '/GM/GM Approval Documents';
  private sharedFolderRootPath: string = '\\\\SPSE26H\\GMApprovalShare';
  private pdfJsonFieldName: string = 'PDFFileJson';
  private odataJsonHeader: string = 'application/json;odata=verbose';
  private requestsListItemEntityTypeFullName: string = '';
  private documentsLibraryListItemEntityTypeFullName: string = '';
  private pdfFileInput: HTMLInputElement;
  private officeManagerEditor: HTMLDivElement;
  private officeManagerDocumentHtml: string = '';
  private besafeLogoUrl: string = require('../assets/besafe-logo.svg');
  private translations: { [key: string]: ITranslatedText } = {
    appTitle: { en: 'General Manager Approval Dashboard', ar: 'لوحة موافقات المدير العام' },
    english: { en: 'English', ar: 'English' },
    arabic: { en: 'Arabic', ar: 'العربية' },
    currentUser: { en: 'Current user:', ar: 'المستخدم الحالي:' },
    groups: { en: 'Groups:', ar: 'المجموعات:' },
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
    signedPdf: { en: 'Signed PDF:', ar: 'ملف PDF الموقع:' },
    openSignedPdf: { en: 'Open Signed PDF', ar: 'فتح ملف PDF الموقع' },
    documents: { en: 'Request documents', ar: 'مستندات الطلب' },
    requestPdf: { en: 'Supporting PDF', ar: 'ملف PDF الداعم' },
    officeManagerPdf: { en: 'Office Manager PDF', ar: 'ملف PDF الخاص بمدير المكتب' },
    signedPdfDocument: { en: 'Signed PDF', ar: 'ملف PDF الموقع' },
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
    generatingPdf: { en: 'Creating and uploading PDF...', ar: 'جاري إنشاء ملف PDF ورفعه...' },
    closeEditor: { en: 'Close editor', ar: 'إغلاق المحرر' },
    hodComment: { en: 'HOD Comment', ar: 'تعليق رئيس القسم' },
    gmCommentNotes: { en: 'GM Comment / Notes', ar: 'تعليق / ملاحظات المدير العام' },
    putOnHold: { en: 'Put On Hold', ar: 'تعليق الطلب' },
    sendToHod: { en: 'Send to HOD', ar: 'إرسال إلى رئيس القسم' },
    sendToGmSignature: { en: 'Send to GM Signature', ar: 'إرسال لتوقيع المدير العام' },
    reject: { en: 'Reject', ar: 'رفض' },
    returnToPendingOfficeManager: { en: 'Return to Pending Office Manager', ar: 'إرجاع إلى مدير المكتب' },
    confirmGmApproval: { en: 'Confirm GM Approval', ar: 'تأكيد موافقة المدير العام' },
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
    messageUploadingPdf: { en: 'Request created. Uploading PDF to the document library and list attachment...', ar: 'تم إنشاء الطلب. جاري رفع ملف PDF إلى مكتبة المستندات ومرفقات الطلب...' },
    messageCreated: { en: 'Request created successfully: {0}', ar: 'تم إنشاء الطلب بنجاح: {0}' },
    messagePdfStepFailed: { en: 'Request {0} was created, but the PDF step failed: ', ar: 'تم إنشاء الطلب {0}، ولكن فشلت خطوة PDF: ' },
    messageUpdatedTo: { en: 'Request updated to: {0}', ar: 'تم تحديث الطلب إلى: {0}' },
    messageReturnedToOfficeManager: { en: 'Request returned to Office Manager.', ar: 'تم إرجاع الطلب إلى مدير المكتب.' },
    messageGmApprovalConfirmed: { en: 'GM approval confirmed.', ar: 'تم تأكيد موافقة المدير العام.' },
    messageRequestClosed: { en: 'Request closed.', ar: 'تم إغلاق الطلب.' },
    messageOfficeManagerPdfCreated: { en: 'Office Manager PDF created and attached successfully: {0}', ar: 'تم إنشاء ملف PDF الخاص بمدير المكتب وإرفاقه بنجاح: {0}' },

    statusPendingOfficeManager: { en: 'Pending Office Manager', ar: 'بانتظار مدير المكتب' },
    statusOnHold: { en: 'On Hold', ar: 'معلق' },
    statusReturnedFromHod: { en: 'Returned from HOD', ar: 'مرجع من رئيس القسم' },
    statusPendingGmSignature: { en: 'Pending GM Signature', ar: 'بانتظار توقيع المدير العام' },
    statusGmSignedPendingOfficeManagerConfirmation: { en: 'GM Signed Pending Office Manager Confirmation', ar: 'تم توقيع المدير العام وبانتظار تأكيد مدير المكتب' },
    statusSentToHod: { en: 'Sent to HOD', ar: 'مرسل إلى رئيس القسم' },
    statusRejected: { en: 'Rejected', ar: 'مرفوض' },
    statusApprovedByGm: { en: 'Approved by GM', ar: 'معتمد من المدير العام' },
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
      newRequestStatus: '',

      officeManagerComment: '',
      hodComment: '',
      gmComment: '',
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

  public componentDidMount(): void {
    this.loadSecurity()
      .then(() => this.loadRequests())
      .catch((error: any) => {
        this.setState({
          loading: false,
          error: this.getErrorMessage(error)
        });
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
        'odata-version': ''
      }
    };
  }

  private getJsonPostOptions(values: any): ISPHttpClientOptions {
    return {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': this.odataJsonHeader,
        'odata-version': ''
      },
      body: JSON.stringify(values)
    };
  }

  private getJsonMergeOptions(values: any): ISPHttpClientOptions {
    return {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': this.odataJsonHeader,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
        'odata-version': ''
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
      "?$select=Id,Title,RequestNo,RequestDate,Department,RequestDetails,Status," +
      "OfficeManagerComment,HODComment,GMComment,PDFFileUrl,PDFFileJson,SignedPDFUrl,SharedFolderPath," +
      "GMApprovalDetected,GMApprovalConfirmed,Attachments,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl,Author/Id,Author/Title" +
      "&$expand=Author,AttachmentFiles" +
      "&$orderby=Id desc" +
      "&$top=200";
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

        return this.uploadPdfToLibraryAndListItem(pdfFile, requestNo, createdItemId)
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

  private updateItem(itemId: number, values: any): Promise<void> {
    const url: string = this.webUrl + "/_api/web/lists/GetByTitle('" + this.requestsListName + "')/items(" + itemId + ")";

    return this.getRequestsListItemEntityTypeFullName()
      .then((entityTypeName: string) => {
        const updateBody: any = this.addODataMetadata(values, entityTypeName);
        const options: ISPHttpClientOptions = this.getJsonMergeOptions(updateBody);

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
    this.setState({
      selectedRequest: item,
      showNewRequestForm: false,
      showOfficeManagerEditor: false,
      officeManagerComment: item.OfficeManagerComment || '',
      hodComment: item.HODComment || '',
      gmComment: item.GMComment || '',
      error: '',
      message: ''
    });
  }

  private officeManagerAction(status: string): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    const updateValues: any = {
      Status: status,
      OfficeManagerComment: this.state.officeManagerComment
    };

    this.applySharedFolderPathForStatus(updateValues, item.RequestNo, status);

    this.runUpdate(item.Id, updateValues, this.formatText('messageUpdatedTo', this.getStatusText(status)));
  }

  private hodReturnToOfficeManager(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    this.runUpdate(item.Id, {
      Status: 'Returned from HOD',
      HODComment: this.state.hodComment
    }, this.t('messageReturnedToOfficeManager'));
  }

  private confirmGmApproval(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    this.runUpdate(item.Id, {
      Status: 'Approved by GM',
      GMApprovalConfirmed: true,
      GMComment: this.state.gmComment,
      SharedFolderPath: this.getSharedFolderPathForStatus(item.RequestNo, 'Approved by GM')
    }, this.t('messageGmApprovalConfirmed'));
  }

  private closeRequest(): void {
    const item: IRequestItem = this.state.selectedRequest;
    if (!item) {
      return;
    }

    this.runUpdate(item.Id, {
      Status: 'Closed',
      SharedFolderPath: this.getSharedFolderPathForStatus(item.RequestNo, 'Closed')
    }, this.t('messageRequestClosed'));
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
      'GM Signed Pending Office Manager Confirmation'
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
      if (item.Status === 'Sent to HOD') {
        items.push(item);
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

  private isPdfFile(file: any): boolean {
    if (!file) {
      return false;
    }

    const fileName: string = (file.name || '').toLowerCase();
    return file.type === 'application/pdf' || fileName.lastIndexOf('.pdf') === fileName.length - 4;
  }

  private uploadPdfToLibraryAndListItem(file: any, requestNo: string, itemId: number): Promise<IPdfUploadResult> {
    const fileName: string = this.buildPdfFileName(requestNo, file.name || 'request.pdf');
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
                libraryResult.originalFileName = file.name || fileName;
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
    return this.getRequestDocumentFolderServerRelativeUrl(requestNo)
      .then((folderServerRelativeUrl: string) => {
        const uploadUrl: string = this.webUrl +
          "/_api/web/GetFolderByServerRelativeUrl('" + this.escapeODataString(folderServerRelativeUrl) + "')" +
          "/Files/add(url='" + this.escapeODataString(fileName) + "',overwrite=true)";

        const options: ISPHttpClientOptions = {
          headers: {
            'Accept': this.odataJsonHeader,
            'Content-Type': 'application/pdf',
            'odata-version': ''
          },
          body: fileContent
        };

        return this.props.context.spHttpClient
          .post(uploadUrl, SPHttpClient.configurations.v1, options)
          .then((response: SPHttpClientResponse) => {
            if (!response.ok) {
              return response.text().then((text: string) => {
                throw new Error(this.formatText('errorLibraryUpload', text));
              });
            }

            return response.json();
          })
          .then((data: any) => {
            const uploadedFile: any = data.d ? data.d : data;
            const serverRelativeUrl: string = uploadedFile.ServerRelativeUrl ||
              (folderServerRelativeUrl + '/' + fileName);

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
    return this.deleteRequestAttachmentIfExists(itemId, fileName)
      .then(() => this.attachPdfFileToRequestItem(itemId, fileName, fileContent));
  }

  private deleteRequestAttachmentIfExists(itemId: number, fileName: string): Promise<void> {
    const url: string = this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.escapeODataString(this.requestsListName) + "')" +
      "/items(" + itemId + ")/AttachmentFiles/getByFileName('" + this.escapeODataString(fileName) + "')";
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': this.odataJsonHeader,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE',
        'odata-version': ''
      }
    };

    return this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (response.ok || response.status === 404) {
          return;
        }

        return response.text().then((text: string) => {
          throw new Error(this.formatText('errorListAttachment', text));
        });
      });
  }

  private attachPdfFileToRequestItem(itemId: number, fileName: string, fileContent: ArrayBuffer): Promise<string> {
    const attachmentUrl: string = this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.escapeODataString(this.requestsListName) + "')" +
      "/items(" + itemId + ")/AttachmentFiles/add(FileName='" + this.escapeODataString(fileName) + "')";

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': 'application/pdf',
        'odata-version': ''
      },
      body: fileContent
    };

    return this.props.context.spHttpClient
      .post(attachmentUrl, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(this.formatText('errorListAttachment', text));
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

  private updatePdfMetadata(
    itemId: number,
    uploadResult: IPdfUploadResult,
    requestNo: string,
    existingPdfJson: string,
    updatePrimaryUrl: boolean
  ): Promise<string> {
    const pdfJson: string = this.buildPdfFileJson(uploadResult, requestNo, itemId, existingPdfJson);

    return this.updatePdfJson(itemId, pdfJson)
      .then(() => {
        if (!updatePrimaryUrl) {
          return;
        }

        return this.updatePdfFileUrl(itemId, uploadResult.url, uploadResult.fileName)
          .catch(() => {
            return;
          });
      })
      .then(() => pdfJson);
  }

  private updatePdfJson(itemId: number, pdfJson: string): Promise<void> {
    const values: any = {};
    values[this.pdfJsonFieldName] = pdfJson;

    return this.updateItem(itemId, values)
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
      schemaVersion: 2,
      type: 'GM_REQUEST_PDF_COLLECTION',
      requestItemId: itemId,
      requestNo: requestNo,
      documents: documents
    });
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
      role = documentType.toLowerCase().indexOf('office manager') >= 0 ? 'officeManager' : 'request';
    }

    return {
      role: role,
      documentType: documentType || (role === 'officeManager' ? 'Office Manager PDF' : 'Request PDF'),
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

  private getDocumentsLibraryServerRelativeUrl(): Promise<string> {
    const url: string = this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.escapeODataString(this.documentsLibraryName) + "')" +
      "?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder";

    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(this.formatText('errorDocumentsLibraryMissing', text));
          });
        }

        return response.json();
      })
      .then((data: any) => {
        const item: any = data.d ? data.d : data;
        const rootFolder: any = item.RootFolder;

        if (rootFolder && rootFolder.ServerRelativeUrl) {
          return rootFolder.ServerRelativeUrl;
        }

        return this.documentsLibraryServerRelativeUrl;
      })
      .catch(() => {
        return this.documentsLibraryServerRelativeUrl;
      });
  }

  private getRequestDocumentFolderServerRelativeUrl(requestNo: string): Promise<string> {
    return this.getDocumentsLibraryServerRelativeUrl()
      .then((libraryRootUrl: string) => {
        const folderName: string = this.sanitizeSharePointFolderName(requestNo);
        const folderUrl: string = libraryRootUrl.replace(/\/+$/, '') + '/' + folderName;

        return this.ensureSharePointFolder(folderUrl)
          .then(() => folderUrl);
      })
      .catch((error: any) => {
        throw new Error(this.formatText('errorRequestFolder', this.getErrorMessage(error)));
      });
  }

  private ensureSharePointFolder(folderServerRelativeUrl: string): Promise<void> {
    const getFolderUrl: string = this.webUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" + this.escapeODataString(folderServerRelativeUrl) + "')" +
      "?$select=ServerRelativeUrl";

    return this.props.context.spHttpClient
      .get(getFolderUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return;
        }

        if (response.status !== 404) {
          return response.text().then((text: string) => {
            throw new Error(text);
          });
        }

        const createFolderUrl: string = this.webUrl +
          "/_api/web/folders/add('" + this.escapeODataString(folderServerRelativeUrl) + "')";
        const options: ISPHttpClientOptions = {
          headers: {
            'Accept': this.odataJsonHeader,
            'Content-Type': this.odataJsonHeader,
            'odata-version': ''
          }
        };

        return this.props.context.spHttpClient
          .post(createFolderUrl, SPHttpClient.configurations.v1, options)
          .then((createResponse: SPHttpClientResponse) => {
            if (createResponse.ok) {
              return;
            }

            return this.props.context.spHttpClient
              .get(getFolderUrl, SPHttpClient.configurations.v1, this.getJsonRequestOptions())
              .then((retryResponse: SPHttpClientResponse) => {
                if (retryResponse.ok) {
                  return;
                }

                return createResponse.text().then((text: string) => {
                  throw new Error(text);
                });
              });
          });
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

    const legacyUrl: string = this.getUrlValue(item.PDFFileUrl) || this.getFirstAttachmentUrl(item);
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

  private getPdfDocumentByRole(item: IRequestItem, role: string): IPdfDocumentRecord {
    const documents: IPdfDocumentRecord[] = this.getPdfDocuments(item);

    for (let i = 0; i < documents.length; i++) {
      if (documents[i].role === role) {
        return documents[i];
      }
    }

    return null;
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

      if (serverRelativeUrl && attachmentFileName.indexOf('officemanager(') !== 0) {
        return this.getAbsoluteUrl(serverRelativeUrl);
      }
    }

    return '';
  }

  private openOfficeManagerDocumentEditor(item: IRequestItem): void {
    this.officeManagerDocumentHtml = this.buildOfficeManagerDocumentTemplate(item);
    this.setState({
      showOfficeManagerEditor: true,
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

  private buildOfficeManagerDocumentTemplate(item: IRequestItem): string {
    const title: string = this.escapeHtml(item.Title || '');
    const requestNo: string = this.escapeHtml(item.RequestNo || '');
    const department: string = this.escapeHtml(this.getDepartmentText(item.Department || ''));
    const createdBy: string = this.escapeHtml(item.Author ? item.Author.Title : '');
    const requestDetails: string = this.escapeHtml(item.RequestDetails || '').replace(/\r?\n/g, '<br />');
    const managerComment: string = this.escapeHtml(this.state.officeManagerComment || '').replace(/\r?\n/g, '<br />');

    return '' +
      '<div>' +
        '<p style="margin:0 0 8px;color:#107c41;font-size:12px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;">' +
          this.escapeHtml(this.t('officeManagerDocument')) +
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
        '<h2 style="margin:0 0 8px;color:#102a43;font-size:18px;">' + this.escapeHtml(this.t('officeManagerComment')) + '</h2>' +
        '<p style="min-height:92px;margin:0 0 22px;line-height:1.65;">' + (managerComment || '<br />') + '</p>' +
      '</div>';
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

  private createAndAttachOfficeManagerPdf(): void {
    const item: IRequestItem = this.state.selectedRequest;

    if (!item || !this.officeManagerEditor) {
      this.setState({ error: this.t('errorEditorUnavailable') });
      return;
    }

    const fileName: string = this.buildOfficeManagerPdfFileName(item.RequestNo);
    const editorElement: HTMLDivElement = this.officeManagerEditor;
    this.officeManagerDocumentHtml = editorElement.innerHTML;

    this.setState({
      saving: true,
      error: '',
      message: this.t('generatingPdf')
    });

    let generatedBlob: Blob = null;

    this.createPdfBlob(editorElement, fileName)
      .then((pdfBlob: Blob) => {
        generatedBlob = pdfBlob;
        return this.readFileAsArrayBuffer(pdfBlob);
      })
      .then((fileContent: ArrayBuffer) => {
        return this.uploadPdfFileToLibrary(fileName, fileContent, item.RequestNo, 'officeManager', 'Office Manager PDF')
          .then((uploadResult: IPdfUploadResult) => {
            return this.updateLibraryFileMetadata(
              uploadResult.serverRelativeUrl,
              fileName,
              item.RequestNo,
              item.Id,
              'Office Manager PDF'
            )
              .then(() => this.replacePdfFileOnRequestItem(item.Id, fileName, fileContent))
              .then((attachmentUrl: string) => {
                uploadResult.originalFileName = fileName;
                uploadResult.contentType = 'application/pdf';
                uploadResult.size = generatedBlob ? generatedBlob.size : fileContent.byteLength;
                uploadResult.attachmentUrl = attachmentUrl;
                return uploadResult;
              });
          });
      })
      .then((uploadResult: IPdfUploadResult) => {
        return this.updatePdfMetadata(item.Id, uploadResult, item.RequestNo, this.buildPdfJsonSnapshot(item), false);
      })
      .then((pdfJson: string) => {
        const updatedItem: IRequestItem = this.copyRequestItem(item);
        updatedItem.PDFFileJson = pdfJson;
        const updatedRequests: IRequestItem[] = this.replaceRequestInCollection(item.Id, updatedItem);

        this.setState({
          saving: false,
          message: this.formatText('messageOfficeManagerPdfCreated', fileName),
          selectedRequest: updatedItem,
          requests: updatedRequests,
          showOfficeManagerEditor: false
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
    return 'OfficeManager(' + this.sanitizeSharePointFolderName(requestNo) + ').pdf';
  }

  private buildPdfJsonSnapshot(item: IRequestItem): string {
    return JSON.stringify({
      schemaVersion: 2,
      type: 'GM_REQUEST_PDF_COLLECTION',
      requestItemId: item.Id,
      requestNo: item.RequestNo,
      documents: this.getPdfDocuments(item)
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
                  <div>{this.t('currentUser')} <strong>{this.state.currentUserTitle}</strong></div>
                  <div>{this.t('groups')} {this.state.groups.join(', ')}</div>
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

          {!this.state.loading && this.renderDetailsPanel()}

          <div className={styles.footer}>
            {this.t('footer')}
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
      <div className={styles.warning}>
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
              onChange={(e: any) => this.setState({ newTitle: e.target.value })}
            />
          </div>

          <div className={styles.formRow}>
            <label htmlFor="gm-new-request-department">{this.t('department')}</label>
            <select
              id="gm-new-request-department"
              value={this.state.newDepartment}
              onChange={(e: any) => this.setState({ newDepartment: e.target.value })}
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
              onChange={(e: any) => this.setState({ newDetails: e.target.value })}
            />
          </div>

          {this.renderNewRequestRouting()}

          <div className={styles.formRowWide}>
            <label>{this.t('pdfAttachment')}</label>
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
            {this.state.newPdfFile &&
              <div className={styles.selectedFile}>
                <span className={styles.fileBadge}>PDF</span>
                <span>
                  {this.state.newPdfFile.name}
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
          <div className={styles.tableViewport}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>{this.t('requestNo')}</th>
                  <th>{this.t('title')}</th>
                  <th>{this.t('department')}</th>
                  <th>{this.t('status')}</th>
                  <th>{this.t('createdBy')}</th>
                  <th>{this.t('pdf')}</th>
                  <th>{this.t('open')}</th>
                </tr>
              </thead>
              <tbody>
                {items.map((item: IRequestItem) => {
                  return (
                    <tr key={item.Id}>
                      <td>{item.RequestNo}</td>
                      <td>{item.Title}</td>
                      <td>{this.getDepartmentText(item.Department)}</td>
                      <td>{this.renderStatusBadge(item.Status)}</td>
                      <td>{item.Author ? item.Author.Title : ''}</td>
                      <td>
                        {this.getRequestPdfUrl(item) &&
                          <a
                            href={this.getRequestPdfUrl(item)}
                            target="_blank"
                            rel="noopener noreferrer"
                          >
                            {this.t('requestPdf')}
                          </a>
                        }
                        {this.getOfficeManagerPdfUrl(item) &&
                          <span>
                            {' · '}
                            <a
                              href={this.getOfficeManagerPdfUrl(item)}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {this.t('officeManagerPdf')}
                            </a>
                          </span>
                        }
                      </td>
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
            <h3 className={styles.detailsTitle}>{item.Title || this.t('requestDetailsTitle')}</h3>
            <div className={styles.sectionDescription}>{this.t('requestOverviewDescription')}</div>
          </div>
          {this.renderStatusBadge(item.Status)}
        </div>

        <div className={styles.metadataGrid}>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('id')}</div>
            <div className={styles.metadataValue}>{item.Id}</div>
          </div>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('requestNoLabel')}</div>
            <div className={styles.metadataValue}>{item.RequestNo}</div>
          </div>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('departmentLabel')}</div>
            <div className={styles.metadataValue}>{this.getDepartmentText(item.Department)}</div>
          </div>
          <div className={styles.metadataCard}>
            <div className={styles.metadataLabel}>{this.t('createdByLabel')}</div>
            <div className={styles.metadataValue}>{item.Author ? item.Author.Title : ''}</div>
          </div>
        </div>

        <div className={styles.formRowWide}>
          <label htmlFor="gm-request-details-readonly">{this.t('requestDetails')}</label>
          <textarea id="gm-request-details-readonly" value={item.RequestDetails || ''} readOnly={true} />
        </div>

        {item.SharedFolderPath &&
          <div className={styles.infoBox}>
            {this.t('gmSharedFolderPath')} {item.SharedFolderPath}
          </div>
        }

        <div className={styles.sectionHeader}>
          <div className={styles.sectionTitleBlock}>
            <div className={styles.sectionEyebrow}>{this.t('documents')}</div>
            <h3>{this.t('documents')}</h3>
          </div>
        </div>

        {this.renderRequestDocuments(item)}

        {this.renderOfficeManagerActions(item)}
        {this.renderHodActions(item)}

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
    const documents: IPdfDocumentRecord[] = this.getPdfDocuments(item);
    const signedPdfUrl: string = this.getUrlValue(item.SignedPDFUrl);

    if (documents.length === 0 && !signedPdfUrl) {
      return <div className={styles.emptyState}>{this.t('noDocuments')}</div>;
    }

    return (
      <div className={styles.documentGrid}>
        {documents.map((documentRecord: IPdfDocumentRecord, index: number) => {
          const documentUrl: string = this.getPdfDocumentUrl(documentRecord);
          const documentTitle: string = documentRecord.role === 'officeManager' ? this.t('officeManagerPdf') : this.t('requestPdf');
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
      </div>
    );
  }

  private renderOfficeManagerActions(item: IRequestItem): JSX.Element {
    if (!this.state.isOfficeManager) {
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
    if (!this.state.showOfficeManagerEditor) {
      return (
        <button
          type="button"
          className={styles.wordLaunchButton}
          disabled={this.state.saving}
          onClick={() => this.openOfficeManagerDocumentEditor(item)}
        >
          <span className={styles.wordIcon}>W</span>
          <span>
            {this.getOfficeManagerPdfUrl(item) ? this.t('editOfficeManagerDocument') : this.t('createOfficeManagerDocument')}
          </span>
        </button>
      );
    }

    return (
      <div className={styles.editorShell}>
        <div className={styles.editorHeader}>
          <span className={styles.wordIcon}>W</span>
          <span className={styles.editorHeaderText}>
            {this.t('editorTitle')} · {this.buildOfficeManagerPdfFileName(item.RequestNo)}
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
          <button type="button" className={styles.toolButton} title={this.t('bullets')} aria-label={this.t('bullets')} disabled={this.state.saving} onMouseDown={(e: any) => this.applyEditorCommand(e, 'insertUnorderedList')}>• List</button>
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
            onClick={() => this.createAndAttachOfficeManagerPdf()}
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
    if (!this.state.isHod || item.Status !== 'Sent to HOD') {
      return null;
    }

    return (
      <div className={styles.actionPanel}>
        <div className={styles.actionPanelHeader}>
          <h4>{this.t('hodActions')}</h4>
          {this.renderStatusBadge(item.Status)}
        </div>

        <div className={styles.actionPanelBody}>
          <div className={styles.formRow}>
            <label htmlFor="gm-hod-office-comment">{this.t('officeManagerComment')}</label>
            <textarea id="gm-hod-office-comment" value={item.OfficeManagerComment || ''} readOnly={true} />
          </div>

          <div className={styles.formRow}>
            <label htmlFor="gm-hod-comment">{this.t('hodComment')}</label>
            <textarea
              id="gm-hod-comment"
              value={this.state.hodComment}
              onChange={(e: any) => this.setState({ hodComment: e.target.value })}
            />
          </div>
        </div>

        <div className={styles.actionFooter}>
          <button
            type="button"
            className={styles.buttonPrimary}
            disabled={this.state.saving}
            onClick={() => this.hodReturnToOfficeManager()}
          >
            {this.t('returnToOfficeManager')}
          </button>
        </div>
      </div>
    );
  }
}
