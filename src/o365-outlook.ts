// Imports here
import CNShell from "cn-shell";
import qs from "qs";

import * as MSGraph from "@microsoft/microsoft-graph-types";

// Misc config consts here
const CFG_MS_GRAPH_APP_ID = "MS_GRAPH_APP_ID";
const CFG_MS_GRAPH_CLIENT_SECRET = "MS_GRAPH_CLIENT_SECRET";
const CFG_MS_GRAPH_RESOURCE_ = "MS_GRAPH_RESOURCE";
const CFG_MS_GRAPH_TENANT_ID = "MS_GRAPH_TENANT_ID";
const CFG_MS_GRAPH_GRANT_TYPE = "MS_GRAPH_GRANT_TYPE";

// Misc consts here
const GRAPH_API_VERSION = "v1.0";
const EMAIL_REF_GUID = `String {350e62a0-efe4-431e-bfed-c38cf45740eb} Name cn-msg-api-ref-code`;

const CFG_TOKEN_GRACE_PERIOD = "TOKEN_GRACE_PERIOD";

const DEFAULT_TOKEN_GRACE_PERIOD = "5"; // In mins

process.on("unhandledRejection", error => {
  // Will print "unhandledRejection err is not defined"
  console.log("unhandledRejection", error);
});

// MsGraphApi class here
class CNMsGraphApi extends CNShell {
  // Properties here
  private _appId: string;
  private _clientSecret: string;
  private _resource: string;
  private _tenantId: string;
  private _grantType: string;

  private _token: string | undefined;
  private _tokenGracePeriod: number;
  private _tokenTimeout: NodeJS.Timeout;

  // Constructor here
  constructor(name: string, master?: CNShell) {
    super(name, master);

    this._appId = this.getRequiredCfg(CFG_MS_GRAPH_APP_ID, false, true);
    this._clientSecret = this.getRequiredCfg(
      CFG_MS_GRAPH_CLIENT_SECRET,
      false,
      true,
    );
    this._resource = this.getRequiredCfg(CFG_MS_GRAPH_RESOURCE_);
    this._tenantId = this.getRequiredCfg(CFG_MS_GRAPH_TENANT_ID, false, true);
    this._grantType = this.getRequiredCfg(CFG_MS_GRAPH_GRANT_TYPE);

    let gracePeriod = this.getCfg(
      CFG_TOKEN_GRACE_PERIOD,
      DEFAULT_TOKEN_GRACE_PERIOD,
    );

    this._tokenGracePeriod = parseInt(gracePeriod, 10) * 60 * 1000; // Convert to ms
  }

  // Abstract method implementations here
  async start(): Promise<boolean> {
    await this.renewToken();

    return true;
  }

  async stop(): Promise<void> {
    if (this._tokenTimeout !== undefined) {
      clearTimeout(this._tokenTimeout);
    }

    return;
  }

  async healthCheck(): Promise<boolean> {
    if (this._token === undefined) {
      return false;
    }

    return true;
  }

  // Private methods here
  private async renewToken(): Promise<void> {
    this.info("Renewing token now!");

    let data = {
      client_id: this._appId,
      client_secret: this._clientSecret,
      scope: `${this._resource}/.default`,
      grant_type: this._grantType,
    };

    let res = await this.httpReq({
      method: "post",
      url: `https://login.microsoftonline.com/${this._tenantId}/oauth2/v2.0/token`,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      data: qs.stringify(data),
    }).catch(e => {
      this.error("Error while renewing token - (%s)", e);
    });

    if (res !== undefined) {
      this._token = res.data.access_token;

      let renewIn = res.data.expires_in * 1000 - this._tokenGracePeriod;

      this._tokenTimeout = setTimeout(() => this.renewToken(), renewIn);

      this.info(
        "Will renew token again in (%s) mins",
        Math.round(renewIn / 1000 / 60),
      );
    } else {
      // Try again in 1 minute
      this._token = undefined;
      this.info("Will try and renew token again in 1 min");
      this._tokenTimeout = setTimeout(() => this.renewToken(), 60 * 1000);
    }
  }

  private async getConversationId(
    user: string,
    refCode: string,
  ): Promise<string> {
    let query = qs.stringify({
      $filter: `singleValueExtendedProperties/Any(ep: ep/id eq '${EMAIL_REF_GUID}' and ep/value eq '${refCode}')`,
      $expand: `singleValueExtendedProperties($filter=id eq '${EMAIL_REF_GUID}')`,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "outlook.body-content-type": "text",
      },
    }).catch(e => {
      this.error(
        "Error while getting conversation ID for ref code (%s) - (%s)",
        refCode,
        e,
      );
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return "";
    }

    return res.data.value[0].conversationId;
  }

  private async getLastMsgInConversation(
    user: string,
    refCode: string,
  ): Promise<MSGraph.Message | undefined> {
    let conversationId = await this.getConversationId(user, refCode);

    if (conversationId === "") {
      return undefined;
    }

    // Note: To use receivedDateTime in $orderby you must also have
    // the receivedDateTime in the $filter so just choose a valid date in the past
    // We only need the last msg so set $top to 1
    let query = qs.stringify({
      $filter: `receivedDateTime ge 2000-01-01T00:01:00Z and conversationId eq '${conversationId}'`,
      $orderby: "receivedDateTime desc",
      $top: 1,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "outlook.body-content-type": "text",
      },
    }).catch(e => {
      this.error(
        "Error while getting conversation for ref code (%s) - (%s)",
        refCode,
        e,
      );
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return undefined;
    }

    return res.data.value[0];
  }

  // Public methods here
  public async sendMessage(
    toRecipients: string[],
    ccRecipients: string[],
    user: string,
    subject: string,
    content: string,
    contentType: MSGraph.BodyType,
    refCode: string,
  ): Promise<boolean> {
    this.info("Attempting to send an email for refCode %s", refCode);

    let message: MSGraph.Message = {
      subject,
      body: {
        contentType,
        content,
      },
      singleValueExtendedProperties: [
        {
          id: EMAIL_REF_GUID,
          value: refCode,
        },
      ],
    };

    message.toRecipients = [];

    for (let recipient of toRecipients) {
      message.toRecipients.push({
        emailAddress: {
          address: recipient,
        },
      });
    }

    message.ccRecipients = [];

    for (let recipient of ccRecipients) {
      message.ccRecipients.push({
        emailAddress: {
          address: recipient,
        },
      });
    }

    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/sendMail`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: { message },
    }).catch(e => {
      this.error("Error while sending email - (%s)", e);
    });

    if (res === undefined || res.status !== 202) {
      return false;
    }

    return true;
  }

  public async replyToAllInConversation(
    user: string,
    comment: string,
    refCode: string,
  ): Promise<boolean> {
    let message = await this.getLastMsgInConversation(user, refCode);

    if (message === undefined) {
      return false;
    }

    let id = message.id;

    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${id}/replyAll`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: { comment },
    }).catch(e => {
      this.error("Error while replying to all - (%s)", e);
    });

    if (res === undefined || res.status !== 202) {
      return false;
    }

    return true;
  }

  public async getMessage(
    user: string,
    refCode: string,
  ): Promise<MSGraph.Message | undefined> {
    let query = qs.stringify({
      $filter: `singleValueExtendedProperties/Any(ep: ep/id eq '${EMAIL_REF_GUID}' and ep/value eq '${refCode}')`,
      $expand: `singleValueExtendedProperties($filter=id eq '${EMAIL_REF_GUID}')`,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "outlook.body-content-type": "text",
      },
    }).catch(e => {
      this.error(
        "Error while getting email with ref code (%s) - (%s)",
        refCode,
        e,
      );
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return undefined;
    }

    return res.data;
  }

  public async getRefCodeForConversation(
    user: string,
    conversationId: string,
  ): Promise<string | undefined> {
    let query = qs.stringify({
      $filter: `singleValueExtendedProperties/Any(ep: ep/id eq '${EMAIL_REF_GUID}' and ep/value ne null) and conversationId eq '${conversationId}'`,
      $expand: `singleValueExtendedProperties($filter=id eq '${EMAIL_REF_GUID}')`,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "outlook.body-content-type": "text",
      },
    }).catch(e => {
      this.error(
        "Error while getting ref code for conversation ID (%s) - (%s)",
        conversationId,
        e,
      );
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return undefined;
    }

    // Get the property value
    let messages: MSGraph.Message[] = res.data.value;

    if (
      messages.length === 0 ||
      messages[0].singleValueExtendedProperties === undefined
    ) {
      return undefined;
    }

    let props = <MSGraph.SingleValueLegacyExtendedProperty[]>(
      messages[0].singleValueExtendedProperties
    );

    for (let prop of props) {
      if (prop.id === EMAIL_REF_GUID) {
        return <string>prop.value;
      }
    }

    return undefined;
  }

  public async getConversation(
    user: string,
    refCode: string,
  ): Promise<MSGraph.Message[] | undefined> {
    // NOTE: This is not working fully - only returns the 1st 10 messages
    // in the conversation. Add a $top query string and use @odata.nextLink
    let conversationId = await this.getConversationId(user, refCode);

    if (conversationId === "") {
      return undefined;
    }

    let query = qs.stringify({
      $filter: `conversationId eq '${conversationId}'`,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "outlook.body-content-type": "text",
      },
    }).catch(e => {
      this.error(
        "Error while getting conversation for ref code (%s) - (%s)",
        refCode,
        e,
      );
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return undefined;
    }

    return res.data.value;
  }

  public async getUnreadMessages(
    user: string,
  ): Promise<MSGraph.Message[] | undefined> {
    let query = qs.stringify({
      $filter: `receivedDateTime ge 2021-06-30T00:01:00Z and isRead eq false`,
      $top: 10,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "outlook.body-content-type": "text",
      },
    }).catch(e => {
      this.error("Error while getting unread emails - (%s)", e);
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return undefined;
    }

    return res.data.value;
  }

  public async getUnreadMessagesInConversation(
    user: string,
    refCode: string,
  ): Promise<MSGraph.Message[] | undefined> {
    let conversationId = await this.getConversationId(user, refCode);

    if (conversationId === "") {
      return undefined;
    }

    let query = qs.stringify({
      $filter: `isRead eq false and conversationId eq '${conversationId}'`,
    });

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages?${query}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        Prefer: "outlook.body-content-type='text'",
      },
    }).catch(e => {
      this.error("Error while getting unread emails - (%s)", e);
    });

    if (
      res === undefined ||
      res.status !== 200 ||
      res.data.value.length === 0
    ) {
      return undefined;
    }

    return res.data.value;
  }

  public async markMessageRead(user: string, id: string): Promise<boolean> {
    let message: MSGraph.Message = {
      isRead: true,
    };

    let res = await this.httpReq({
      method: "patch",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${id}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: message,
    }).catch(e => {
      this.error("Error while marking email read - (%s)", e);
    });

    if (res === undefined || res.status !== 200) {
      return false;
    }

    return true;
  }
}

export { CNMsGraphApi };
