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
}

interface IPdfUploadResult {
  fileName: string;
  originalFileName: string;
  contentType: string;
  size: number;
  serverRelativeUrl: string;
  url: string;
  attachmentUrl: string;
  json: string;
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
  private pdfFileInput: HTMLInputElement;
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

    newRequest: { en: 'New Request', ar: 'طلب جديد' },
    title: { en: 'Title', ar: 'العنوان' },
    department: { en: 'Department', ar: 'القسم' },
    selectDepartment: { en: 'Select department', ar: 'اختر القسم' },
    requestDetails: { en: 'Request Details', ar: 'تفاصيل الطلب' },
    pdfAttachment: { en: 'PDF Attachment', ar: 'مرفق PDF' },
    pdfHelp: { en: 'Stored in GM Approval Documents and the request Attachments', ar: 'سيتم حفظه في مكتبة مستندات موافقات المدير العام ومرفقات الطلب' },
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
    closeDetails: { en: 'Close Details', ar: 'إغلاق التفاصيل' },

    officeManagerActions: { en: 'Office Manager Actions', ar: 'إجراءات مدير المكتب' },
    officeManagerComment: { en: 'Office Manager Comment', ar: 'تعليق مدير المكتب' },
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
    messageUploadingPdf: { en: 'Request created. Uploading PDF to the document library and list attachment...', ar: 'تم إنشاء الطلب. جاري رفع ملف PDF إلى مكتبة المستندات ومرفقات الطلب...' },
    messageCreated: { en: 'Request created successfully: {0}', ar: 'تم إنشاء الطلب بنجاح: {0}' },
    messagePdfStepFailed: { en: 'Request {0} was created, but the PDF step failed: ', ar: 'تم إنشاء الطلب {0}، ولكن فشلت خطوة PDF: ' },
    messageUpdatedTo: { en: 'Request updated to: {0}', ar: 'تم تحديث الطلب إلى: {0}' },
    messageReturnedToOfficeManager: { en: 'Request returned to Office Manager.', ar: 'تم إرجاع الطلب إلى مدير المكتب.' },
    messageGmApprovalConfirmed: { en: 'GM approval confirmed.', ar: 'تم تأكيد موافقة المدير العام.' },
    messageRequestClosed: { en: 'Request closed.', ar: 'تم إغلاق الطلب.' },

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
      gmComment: ''
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

  private getListItemsUrl(): string {
    return this.webUrl +
      "/_api/web/lists/GetByTitle('" + this.requestsListName + "')/items" +
      "?$select=Id,Title,RequestNo,RequestDate,Department,RequestDetails,Status," +
      "OfficeManagerComment,HODComment,GMComment,PDFFileUrl,SignedPDFUrl,SharedFolderPath," +
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
            return this.updatePdfMetadata(createdItemId, result);
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
    this.setState({
      selectedRequest: item,
      showNewRequestForm: false,
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

    return this.readFileAsArrayBuffer(file)
      .then((fileContent: ArrayBuffer) => {
        return this.uploadPdfFileToLibrary(fileName, fileContent)
          .then((libraryResult: IPdfUploadResult) => {
            return this.updateLibraryFileMetadata(libraryResult.serverRelativeUrl, fileName, requestNo, itemId)
              .then(() => {
                return this.attachPdfFileToRequestItem(itemId, fileName, fileContent);
              })
              .then((attachmentUrl: string) => {
                libraryResult.originalFileName = file.name || fileName;
                libraryResult.contentType = file.type || 'application/pdf';
                libraryResult.size = file.size || 0;
                libraryResult.attachmentUrl = attachmentUrl;
                libraryResult.json = this.buildPdfFileJson(libraryResult, requestNo, itemId);

                return libraryResult;
              });
          });
      });
  }

  private uploadPdfFileToLibrary(fileName: string, fileContent: ArrayBuffer): Promise<IPdfUploadResult> {
    return this.getDocumentsLibraryServerRelativeUrl()
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
              fileName: fileName,
              originalFileName: '',
              contentType: 'application/pdf',
              size: 0,
              serverRelativeUrl: serverRelativeUrl,
              url: this.getAbsoluteUrl(serverRelativeUrl),
              attachmentUrl: '',
              json: ''
            };
          });
      });
  }

  private updateLibraryFileMetadata(serverRelativeUrl: string, fileName: string, requestNo: string, itemId: number): Promise<void> {
    const values: any = {
      Title: fileName,
      RequestNo: requestNo,
      RequestItemId: itemId,
      DocumentType: 'Request PDF',
      ApprovalProcessed: false
    };

    return this.mergeLibraryFileMetadata(serverRelativeUrl, values)
      .catch(() => {
        return this.mergeLibraryFileMetadata(serverRelativeUrl, {
          Title: fileName
        });
      })
      .catch(() => {
        return;
      });
  }

  private mergeLibraryFileMetadata(serverRelativeUrl: string, values: any): Promise<void> {
    const url: string = this.webUrl +
      "/_api/web/GetFileByServerRelativeUrl('" + this.escapeODataString(serverRelativeUrl) + "')" +
      "/ListItemAllFields";

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': this.odataJsonHeader,
        'Content-Type': this.odataJsonHeader,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
        'odata-version': ''
      },
      body: JSON.stringify(values)
    };

    return this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text: string) => {
            throw new Error(this.formatText('errorLibraryMetadata', text));
          });
        }
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

  private updatePdfMetadata(itemId: number, uploadResult: IPdfUploadResult): Promise<void> {
    return this.updatePdfJson(itemId, uploadResult.json)
      .then(() => {
        return this.updatePdfFileUrl(itemId, uploadResult.url, uploadResult.fileName)
          .catch(() => {
            return;
          });
      });
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

  private buildPdfFileJson(uploadResult: IPdfUploadResult, requestNo: string, itemId: number): string {
    return JSON.stringify({
      schemaVersion: 1,
      type: 'GM_REQUEST_PDF',
      requestItemId: itemId,
      requestNo: requestNo,
      fileName: uploadResult.fileName,
      originalFileName: uploadResult.originalFileName,
      contentType: uploadResult.contentType,
      size: uploadResult.size,
      uploadedAt: new Date().toISOString(),
      libraryUrl: uploadResult.url,
      serverRelativeUrl: uploadResult.serverRelativeUrl,
      attachmentUrl: uploadResult.attachmentUrl
    });
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
    const fieldUrl: string = this.getUrlValue(item.PDFFileUrl);

    if (fieldUrl) {
      return fieldUrl;
    }

    return this.getFirstAttachmentUrl(item);
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

    const firstAttachment: any = results[0];
    const serverRelativeUrl: string = firstAttachment.ServerRelativeUrl || '';

    if (!serverRelativeUrl) {
      return '';
    }

    return this.getAbsoluteUrl(serverRelativeUrl);
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
      <div className={this.getRootClassName()} dir={this.state.language === 'ar' ? 'rtl' : 'ltr'} lang={this.state.language}>
        <div className={styles.header}>
          <div className={styles.headerTop}>
            <div className={styles.brandBlock}>
              <img className={styles.logoImage} src={this.besafeLogoUrl} alt="Besafe" />
              <div className={styles.headerText}>
                <h2>{this.t('appTitle')}</h2>
                <div className={styles.userMeta}>
                  <div>{this.t('currentUser')} <strong>{this.state.currentUserTitle}</strong></div>
                  <div>{this.t('groups')} {this.state.groups.join(', ')}</div>
                </div>
              </div>
            </div>
            <div className={styles.languageSwitcher}>
              <button
                className={this.state.language === 'en' ? styles.languageButtonActive : styles.languageButton}
                onClick={() => this.setLanguage('en')}
              >
                {this.t('english')}
              </button>
              <button
                className={this.state.language === 'ar' ? styles.languageButtonActive : styles.languageButton}
                onClick={() => this.setLanguage('ar')}
              >
                {this.t('arabic')}
              </button>
            </div>
          </div>
        </div>

        {this.state.error &&
          <div className={styles.error}>{this.state.error}</div>
        }

        {this.state.message &&
          <div className={styles.success}>{this.state.message}</div>
        }

        {this.state.loading &&
          <div>{this.t('loadingRequests')}</div>
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
      <div className={styles.section}>
        <button disabled={this.state.saving} onClick={() => this.showNewRequestForm()}>
          {this.t('newRequest')}
        </button>
      </div>
    );
  }

  private renderNewRequestRouting(): JSX.Element {
    if (this.state.isOfficeManager) {
      return (
        <div className={styles.formRow}>
          <label>{this.t('routeTo')}</label>
          <select
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
        <h3>{this.t('newRequest')}</h3>

        <div className={styles.formRow}>
          <label>{this.t('title')}</label>
          <input
            type="text"
            value={this.state.newTitle}
            onChange={(e: any) => this.setState({ newTitle: e.target.value })}
          />
        </div>

        <div className={styles.formRow}>
          <label>{this.t('department')}</label>
          <select
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

        <div className={styles.formRow}>
          <label>{this.t('requestDetails')}</label>
          <textarea
            value={this.state.newDetails}
            onChange={(e: any) => this.setState({ newDetails: e.target.value })}
          />
        </div>

        {this.renderNewRequestRouting()}

        <div className={styles.formRow}>
          <label>{this.t('pdfAttachment')}</label>
          <input
            type="file"
            accept="application/pdf,.pdf"
            ref={(input: HTMLInputElement) => { this.pdfFileInput = input; }}
            onChange={(e: any) => this.onPdfFileChange(e)}
          />
          <div className={styles.helpText}>
            {this.state.newPdfFile ? this.state.newPdfFile.name : this.t('pdfHelp')}
          </div>
        </div>

        <button disabled={this.state.saving} onClick={() => this.createRequest()}>
          {this.state.saving ? this.t('submitting') : this.getNewRequestSubmitText()}
        </button>

        <button className={styles.secondaryButton} disabled={this.state.saving} onClick={() => this.cancelNewRequest()}>
          {this.t('cancel')}
        </button>
      </div>
    );
  }

  private renderRequestTable(title: string, items: IRequestItem[]): JSX.Element {
    return (
      <div className={styles.section}>
        <h3>{title}</h3>

        {items.length === 0 &&
          <div>{this.t('noRequestsFound')}</div>
        }

        {items.length > 0 &&
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
                    <td>{this.getStatusText(item.Status)}</td>
                    <td>{item.Author ? item.Author.Title : ''}</td>
                    <td>
                      {this.getRequestPdfUrl(item) &&
                        <a href={this.getRequestPdfUrl(item)} target="_blank">{this.t('open')}</a>
                      }
                    </td>
                    <td>
                      <button onClick={() => this.selectRequest(item)}>{this.t('open')}</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
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
        <h3>{this.t('requestDetailsTitle')}</h3>

        <div className={styles.detailsGrid}>
          <div><strong>{this.t('id')}</strong> {item.Id}</div>
          <div><strong>{this.t('requestNoLabel')}</strong> {item.RequestNo}</div>
          <div><strong>{this.t('titleLabel')}</strong> {item.Title}</div>
          <div><strong>{this.t('departmentLabel')}</strong> {this.getDepartmentText(item.Department)}</div>
          <div><strong>{this.t('statusLabel')}</strong> {this.getStatusText(item.Status)}</div>
          <div><strong>{this.t('createdByLabel')}</strong> {item.Author ? item.Author.Title : ''}</div>
        </div>

        <div className={styles.formRow}>
          <label>{this.t('requestDetails')}</label>
          <textarea value={item.RequestDetails || ''} readOnly={true} />
        </div>

        {item.SharedFolderPath &&
          <div className={styles.infoBox}>
            {this.t('gmSharedFolderPath')} {item.SharedFolderPath}
          </div>
        }

        {this.getRequestPdfUrl(item) &&
          <div className={styles.infoBox}>
            {this.t('attachedPdf')} <a href={this.getRequestPdfUrl(item)} target="_blank">{this.t('openPdf')}</a>
          </div>
        }

        {this.getUrlValue(item.SignedPDFUrl) &&
          <div className={styles.infoBox}>
            {this.t('signedPdf')} <a href={this.getUrlValue(item.SignedPDFUrl)} target="_blank">{this.t('openSignedPdf')}</a>
          </div>
        }

        {this.renderOfficeManagerActions(item)}
        {this.renderHodActions(item)}

        <button className={styles.secondaryButton} onClick={() => this.setState({ selectedRequest: null })}>
          {this.t('closeDetails')}
        </button>
      </div>
    );
  }

  private renderOfficeManagerActions(item: IRequestItem): JSX.Element {
    if (!this.state.isOfficeManager) {
      return null;
    }

    return (
      <div className={styles.actionBox}>
        <h4>{this.t('officeManagerActions')}</h4>

        <div className={styles.formRow}>
          <label>{this.t('officeManagerComment')}</label>
          <textarea
            value={this.state.officeManagerComment}
            onChange={(e: any) => this.setState({ officeManagerComment: e.target.value })}
          />
        </div>

        {item.Status === 'Returned from HOD' &&
          <div className={styles.formRow}>
            <label>{this.t('hodComment')}</label>
            <textarea value={item.HODComment || ''} readOnly={true} />
          </div>
        }

        {item.Status === 'GM Signed Pending Office Manager Confirmation' &&
          <div className={styles.formRow}>
            <label>{this.t('gmCommentNotes')}</label>
            <textarea
              value={this.state.gmComment}
              onChange={(e: any) => this.setState({ gmComment: e.target.value })}
            />
          </div>
        }

        {item.Status === 'Pending Office Manager' &&
          <div>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('On Hold')}>{this.t('putOnHold')}</button>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Sent to HOD')}>{this.t('sendToHod')}</button>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Pending GM Signature')}>{this.t('sendToGmSignature')}</button>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Rejected')}>{this.t('reject')}</button>
          </div>
        }

        {item.Status === 'On Hold' &&
          <div>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Pending Office Manager')}>{this.t('returnToPendingOfficeManager')}</button>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Rejected')}>{this.t('reject')}</button>
          </div>
        }

        {item.Status === 'Returned from HOD' &&
          <div>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Pending GM Signature')}>{this.t('sendToGmSignature')}</button>
            <button disabled={this.state.saving} onClick={() => this.officeManagerAction('Rejected')}>{this.t('reject')}</button>
          </div>
        }

        {item.Status === 'GM Signed Pending Office Manager Confirmation' &&
          <div>
            <button disabled={this.state.saving} onClick={() => this.confirmGmApproval()}>{this.t('confirmGmApproval')}</button>
          </div>
        }

        {item.Status === 'Approved by GM' &&
          <div>
            <button disabled={this.state.saving} onClick={() => this.closeRequest()}>{this.t('closeRequest')}</button>
          </div>
        }
      </div>
    );
  }

  private renderHodActions(item: IRequestItem): JSX.Element {
    if (!this.state.isHod || item.Status !== 'Sent to HOD') {
      return null;
    }

    return (
      <div className={styles.actionBox}>
        <h4>{this.t('hodActions')}</h4>

        <div className={styles.formRow}>
          <label>{this.t('officeManagerComment')}</label>
          <textarea value={item.OfficeManagerComment || ''} readOnly={true} />
        </div>

        <div className={styles.formRow}>
          <label>{this.t('hodComment')}</label>
          <textarea
            value={this.state.hodComment}
            onChange={(e: any) => this.setState({ hodComment: e.target.value })}
          />
        </div>

        <button disabled={this.state.saving} onClick={() => this.hodReturnToOfficeManager()}>
          {this.t('returnToOfficeManager')}
        </button>
      </div>
    );
  }
}
