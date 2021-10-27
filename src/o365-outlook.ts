// Imports here
import CNShell from "cn-shell";
import { CNO365 } from "@cn-shell/o365";

import qs from "qs";

import * as MSGraph from "@microsoft/microsoft-graph-types";

// Misc config consts here

// Misc consts here
const GRAPH_API_VERSION = "v1.0";
const EMAIL_REF_GUID = `String {350e62a0-efe4-431e-bfed-c38cf45740eb} Name cn-msg-api-ref-code`;

// process.on("unhandledRejection", error => {
//   // Will print "unhandledRejection err is not defined"
//   console.log("unhandledRejection", error);
// });

// interfaces here
interface MessageAttachement {
  name: string;
  contentType: string;
  contentB64: string;
}

// CNO365Outlook class here
class CNO365Outlook extends CNO365 {
  // Properties here

  // Constructor here
  constructor(name: string, master?: CNShell) {
    super(name, master);
  }

  // Abstract method implementations here
  async start(): Promise<boolean> {
    await super.start();
    return true;
  }

  async stop(): Promise<void> {
    await super.stop();
    return;
  }

  async healthCheck(): Promise<boolean> {
    return await super.healthCheck();
  }

  // Private methods here
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

  private async getFirstMsgInConversation(
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
      $orderby: "receivedDateTime asc",
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
        "Error while getting first msg in conversation for ref code (%s) - (%s)",
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
        "Error while getting last msg in conversation for ref code (%s) - (%s)",
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
  public async createDraft(user: string, subject: string): Promise<string> {
    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: { message: { subject } },
    }).catch(e => {
      this.error("Error while creating draft - (%s)", e);
    });

    if (res === undefined || res.status !== 201) {
      return "";
    }

    return res.data.id;
  }

  public async findDraft(
    user: string,
    subject: string,
  ): Promise<MSGraph.Message[] | undefined> {
    let query = qs.stringify({
      $filter: `isDraft eq true and subject eq ${subject}`,
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

  public async copyDraft(id: string, user: string): Promise<string> {
    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${id}/copy`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: { destinationId: "Drafts" },
    }).catch(e => {
      this.error("Error while copying draft - (%s)", e);
    });

    if (res === undefined || res.status !== 201) {
      return "";
    }

    return res.data.id;
  }

  public async updateDraft(
    id: string,
    toRecipients: string[],
    ccRecipients: string[],
    bccRecipients: string[],
    user: string,
    subject: string,
    content: string,
    contentType: MSGraph.BodyType,
    refCode?: string,
  ): Promise<string> {
    let message: MSGraph.Message = {
      subject,
      body: {
        contentType,
        content,
      },
    };

    if (refCode !== undefined) {
      message.singleValueExtendedProperties = [
        {
          id: EMAIL_REF_GUID,
          value: refCode,
        },
      ];
    }

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

    message.bccRecipients = [];

    for (let recipient of bccRecipients) {
      message.bccRecipients.push({
        emailAddress: {
          address: recipient,
        },
      });
    }

    let res = await this.httpReq({
      method: "patch",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${id}`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: { message },
    }).catch(e => {
      this.error("Error while updating draft - (%s)", e);
    });

    if (res === undefined || res.status !== 201) {
      return "";
    }

    return res.data.id;
  }

  public async addAttachementToDraft(
    id: string,
    user: string,
    attachment: MessageAttachement,
  ): Promise<boolean> {
    let data = {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: attachment.name,
      contentType: attachment.contentType,
      contentBytes: attachment.contentB64,
      contentId: "",
      isInline: false,
    };

    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${id}/attachments`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data,
    }).catch(e => {
      this.error("Error while creating message - (%s)", e);
    });

    if (res === undefined || res.status !== 201) {
      return false;
    }

    return true;
  }

  public async sendDraft(id: string, user: string): Promise<boolean> {
    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${id}/send`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
    }).catch(e => {
      this.error("Error while creating message - (%s)", e);
    });

    if (res === undefined || res.status !== 202) {
      return false;
    }

    return true;
  }

  public async sendMessage(
    toRecipients: string[],
    ccRecipients: string[],
    bccRecipients: string[],
    user: string,
    subject: string,
    content: string,
    contentType: MSGraph.BodyType,
    refCode?: string,
    attachments?: MessageAttachement[],
  ): Promise<boolean> {
    let message: MSGraph.Message = {
      subject,
      body: {
        contentType,
        content,
      },
    };

    if (refCode !== undefined) {
      message.singleValueExtendedProperties = [
        {
          id: EMAIL_REF_GUID,
          value: refCode,
        },
      ];
    }

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

    message.bccRecipients = [];

    for (let recipient of bccRecipients) {
      message.bccRecipients.push({
        emailAddress: {
          address: recipient,
        },
      });
    }

    message.attachments = [];

    if (attachments !== undefined) {
      for (let attachment of attachments) {
        let aObj = {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: attachment.name,
          contentType: attachment.contentType,
          contentBytes: attachment.contentB64,
          contentId: "",
          isInline: false,
        };

        message.attachments.push(aObj);
      }
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

  public async replyToConversation(
    user: string,
    comment: string,
    refCode: string,
  ): Promise<boolean> {
    let firstMsg = await this.getFirstMsgInConversation(user, refCode);

    if (firstMsg === undefined) {
      return false;
    }

    let message: MSGraph.Message = {
      toRecipients: firstMsg.toRecipients,
      ccRecipients: firstMsg.ccRecipients,
    };

    let lastMsg = await this.getFirstMsgInConversation(user, refCode);

    if (lastMsg === undefined) {
      return false;
    }

    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/users/${user}/messages/${lastMsg.id}/reply`,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
      data: { comment, message },
    }).catch(e => {
      this.error("Error while replying to conversation - (%s)", e);
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
    numOfMsgs: number = 10,
  ): Promise<MSGraph.Message[] | undefined> {
    let query = qs.stringify({
      $filter: "isRead eq false",
      $top: numOfMsgs,
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

export { CNO365Outlook, MessageAttachement };
