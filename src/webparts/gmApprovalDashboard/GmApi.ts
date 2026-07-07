import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { IGMRequest, IUserContext } from './models';

export default class GmApi {
  private context: WebPartContext;
  private webUrl: string;
  private listTitle: string = 'GM Requests';
  private libraryTitle: string = 'GM Approval Documents';

  constructor(context: WebPartContext) {
    this.context = context;

    // Because your real site is http://spse26h/GM
    // If the web part is added inside the GM site, this will be correct.
    this.webUrl = this.context.pageContext.web.absoluteUrl;
  }

  public getCurrentUser(): Promise<IUserContext> {
    var userUrl = this.webUrl + "/_api/web/currentuser";
    var groupsUrl = this.webUrl + "/_api/web/currentuser/groups?$select=Title";

    return this.context.spHttpClient
      .get(userUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((user: any) => {
        return this.context.spHttpClient
          .get(groupsUrl, SPHttpClient.configurations.v1)
          .then((groupResponse: SPHttpClientResponse) => groupResponse.json())
          .then((groupData: any) => {
            var groups: string[] = [];

            if (groupData && groupData.value) {
              groups = groupData.value.map((g: any) => g.Title);
            }

            return {
              id: user.Id,
              title: user.Title,
              loginName: user.LoginName,
              groups: groups
            };
          });
      });
  }

  public getMyRequests(userId: number): Promise<IGMRequest[]> {
    var url =
      this.webUrl +
      "/_api/web/lists/getbytitle('" +
      this.listTitle +
      "')/items" +
      "?$select=Id,Title,RequestNo,Status,RequestDate,RequestDetails,OfficeManagerComment,HODComment,GMApprovalDetected,SignedPDFUrl,Author/Id,Author/Title" +
      "&$expand=Author" +
      "&$filter=AuthorId eq " +
      userId +
      "&$orderby=Id desc";

    return this.getItems(url);
  }

  public getOfficeManagerQueue(): Promise<IGMRequest[]> {
    var filter =
      "Status eq 'Pending Office Manager' or " +
      "Status eq 'On Hold' or " +
      "Status eq 'GM Signed Pending Office Manager Confirmation' or " +
      "Status eq 'Approved by GM'";

    var url =
      this.webUrl +
      "/_api/web/lists/getbytitle('" +
      this.listTitle +
      "')/items" +
      "?$select=Id,Title,RequestNo,Status,RequestDate,RequestDetails,OfficeManagerComment,HODComment,GMApprovalDetected,SignedPDFUrl,Author/Id,Author/Title" +
      "&$expand=Author" +
      "&$filter=" +
      encodeURIComponent(filter) +
      "&$orderby=Id desc";

    return this.getItems(url);
  }

  public getHodQueue(): Promise<IGMRequest[]> {
    var url =
      this.webUrl +
      "/_api/web/lists/getbytitle('" +
      this.listTitle +
      "')/items" +
      "?$select=Id,Title,RequestNo,Status,RequestDate,RequestDetails,OfficeManagerComment,HODComment,GMApprovalDetected,SignedPDFUrl,Author/Id,Author/Title" +
      "&$expand=Author" +
      "&$filter=Status eq 'Sent to HOD'" +
      "&$orderby=Id desc";

    return this.getItems(url);
  }

  public getRequestById(id: number): Promise<IGMRequest> {
    var url =
      this.webUrl +
      "/_api/web/lists/getbytitle('" +
      this.listTitle +
      "')/items(" +
      id +
      ")" +
      "?$select=Id,Title,RequestNo,Status,RequestDate,RequestDetails,OfficeManagerComment,HODComment,GMApprovalDetected,SignedPDFUrl,Author/Id,Author/Title" +
      "&$expand=Author";

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json());
  }

  public createRequest(title: string, details: string): Promise<IGMRequest> {
    var createUrl =
      this.webUrl +
      "/_api/web/lists/getbytitle('" +
      this.listTitle +
      "')/items";

    var body: any = {
      Title: title,
      RequestDetails: details,
      Status: 'Pending Office Manager',
      RequestDate: new Date().toISOString()
    };

    var options: ISPHttpClientOptions = {
      body: JSON.stringify(body)
    };

    return this.context.spHttpClient
      .post(createUrl, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => response.json())
      .then((createdItem: any) => {
        var requestNo = this.buildRequestNo(createdItem.Id);
        return this.updateRequest(createdItem.Id, {
          RequestNo: requestNo
        }).then(() => {
          return this.getRequestById(createdItem.Id);
        });
      });
  }

  public updateRequest(id: number, values: any): Promise<void> {
    var url =
      this.webUrl +
      "/_api/web/lists/getbytitle('" +
      this.listTitle +
      "')/items(" +
      id +
      ")";

    var options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(values)
    };

    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, options)
      .then(() => {
        return;
      });
  }

  public changeStatus(id: number, status: string, extraValues?: any): Promise<void> {
    var values: any = {
      Status: status
    };

    if (extraValues) {
      for (var key in extraValues) {
        if (extraValues.hasOwnProperty(key)) {
          values[key] = extraValues[key];
        }
      }
    }

    return this.updateRequest(id, values);
  }

  private getItems(url: string): Promise<IGMRequest[]> {
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: any) => {
        if (data && data.value) {
          return data.value as IGMRequest[];
        }

        return [];
      });
  }

  private buildRequestNo(id: number): string {
    var now = new Date();
    var year = now.getFullYear().toString();
    var month = this.pad(now.getMonth() + 1);
    var day = this.pad(now.getDate());

    return 'REQ-' + year + month + day + '-' + id;
  }

  private pad(value: number): string {
    return value < 10 ? '0' + value : value.toString();
  }
}