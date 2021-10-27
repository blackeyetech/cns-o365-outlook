import { CNO365Outlook } from "./o365-outlook";

import inquirer from "inquirer";
// import EasyTable from "easy-table";

import fs from "fs";
import path from "path";

// enums here
enum Prompts {
  USER = "User",
  TEST = "Test",
  REF = "Ref",
  TO = "To",
  SUBJECT = "Subject",
  BODY = "Body",
  ATTACHMENT = "Attach",
  TYPE = "Type",
}

enum TestChoices {
  SEND = "Send Email",
  REPLY_ALL = "Reply to all email",
  REPLY_TO = "Reply to email",
  GET = "Get Conversation",
  GET_UNREAD = "Get for unread emails",
  GET_UNREAD_CONVERSATION = "Get for unread emails from conversation",
  MARK_READ = "Mark email as read",
  CREATE_DRAFT = "Create a Draft",
  COPY_DRAFT = "Copy Draft",
  SEND_DRAFT = "Send Draft",
  QUIT = "Quit",
}

// Tests here
async function sendMessage(userEmail: string, msGraphApi: CNO365Outlook) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code:",
    },
  ]);

  let refCode = answer[Prompts.REF];

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.TO,
      message: "Input address to email:",
    },
  ]);

  let toAddress = answer[Prompts.TO];

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.SUBJECT,
      message: "Input email subject:",
    },
  ]);

  let subject = answer[Prompts.SUBJECT];

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.BODY,
      message: "Input email body:",
    },
  ]);

  let body = answer[Prompts.BODY];

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.ATTACHMENT,
      message: "File to attach (leave empty to ignore):",
    },
  ]);

  let fileName: string = answer[Prompts.ATTACHMENT];
  let contentType = "";
  let content: Buffer = Buffer.from("");

  if (fileName.length) {
    content = fs.readFileSync(fileName);

    answer = await inquirer.prompt([
      {
        type: "input",
        name: Prompts.TYPE,
        message: "File type:",
      },
    ]);

    contentType = answer[Prompts.ATTACHMENT];
  }

  await msGraphApi.sendMessage(
    [toAddress],
    [],
    [],
    userEmail,
    subject,
    body,
    "text",
    refCode,
    fileName.length
      ? [
          {
            name: path.basename(fileName),
            contentType,
            contentB64: content.toString("base64"),
          },
        ]
      : undefined,
  );
}

async function replyToAll(userEmail: string, msGraphApi: CNO365Outlook) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code:",
    },
  ]);

  let refCode = answer[Prompts.REF];

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.BODY,
      message: "Input reply email body:",
    },
  ]);

  let body = answer[Prompts.BODY];

  await msGraphApi.replyToAllInConversation(userEmail, body, refCode);
}

async function replyTo(userEmail: string, msGraphApi: CNO365Outlook) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code:",
    },
  ]);

  let refCode = answer[Prompts.REF];

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.BODY,
      message: "Input reply email body:",
    },
  ]);

  let body = answer[Prompts.BODY];

  await msGraphApi.replyToConversation(userEmail, body, refCode);
}

async function getConversation(userEmail: string, msGraphApi: CNO365Outlook) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code:",
    },
  ]);

  let refCode = answer[Prompts.REF];

  let conversation = await msGraphApi.getConversation(userEmail, refCode);

  console.log(JSON.stringify(conversation, undefined, 2));
}

async function getUnreadMessages(userEmail: string, msGraphApi: CNO365Outlook) {
  let messages = await msGraphApi.getUnreadMessages(userEmail);

  console.log(JSON.stringify(messages, undefined, 2));
}

async function getUnreadMessagesFromConversation(
  userEmail: string,
  msGraphApi: CNO365Outlook,
) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code:",
    },
  ]);

  let refCode = answer[Prompts.REF];

  let messages = await msGraphApi.getUnreadMessagesInConversation(
    userEmail,
    refCode,
  );

  console.log(JSON.stringify(messages, undefined, 2));
}

async function copyDraft(userEmail: string, msGraphApi: CNO365Outlook) {
  let messages = await msGraphApi.getUnreadMessages(userEmail, 1);

  console.log(JSON.stringify(messages, undefined, 2));
}

async function markEmailRead(userEmail: string, msGraphApi: CNO365Outlook) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code:",
    },
  ]);

  let refCode = answer[Prompts.REF];

  let messages = await msGraphApi.getUnreadMessagesInConversation(
    userEmail,
    refCode,
  );

  if (messages === undefined || messages.length === 0) {
    return;
  }

  // Mark the first unread email and read
  msGraphApi.markMessageRead(userEmail, <string>messages[0].id);
}

// Main here
(async () => {
  let msGraphApi = new CNO365Outlook("Test-MS-Graph-API");
  await msGraphApi.init();

  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.USER,
      message: "Input user email to use:",
    },
  ]);

  let userEmail = answer[Prompts.USER];

  while (1) {
    answer = await inquirer.prompt([
      {
        type: "list",
        name: Prompts.TEST,
        choices: [
          TestChoices.GET,
          TestChoices.REPLY_ALL,
          TestChoices.REPLY_TO,
          TestChoices.SEND,
          TestChoices.GET_UNREAD,
          TestChoices.GET_UNREAD_CONVERSATION,
          TestChoices.MARK_READ,
          TestChoices.CREATE_DRAFT,
          TestChoices.COPY_DRAFT,
          TestChoices.SEND_DRAFT,
          TestChoices.QUIT,
        ],
        message: "What test do you want to run?",
      },
    ]);

    if (answer[Prompts.TEST] === TestChoices.QUIT) {
      break;
    }

    switch (answer[Prompts.TEST]) {
      case TestChoices.SEND:
        await sendMessage(userEmail, msGraphApi);
        break;
      case TestChoices.REPLY_ALL:
        await replyToAll(userEmail, msGraphApi);
        break;
      case TestChoices.REPLY_TO:
        await replyTo(userEmail, msGraphApi);
        break;
      case TestChoices.GET:
        await getConversation(userEmail, msGraphApi);
        break;
      case TestChoices.GET_UNREAD:
        await getUnreadMessages(userEmail, msGraphApi);
        break;
      case TestChoices.MARK_READ:
        await markEmailRead(userEmail, msGraphApi);
        break;
      case TestChoices.GET_UNREAD_CONVERSATION:
        await getUnreadMessagesFromConversation(userEmail, msGraphApi);
        break;
      case TestChoices.CREATE_DRAFT:
        await copyDraft(userEmail, msGraphApi);
        break;
      case TestChoices.COPY_DRAFT:
        await copyDraft(userEmail, msGraphApi);
        break;
      case TestChoices.SEND_DRAFT:
        await copyDraft(userEmail, msGraphApi);
        break;
    }
  }

  msGraphApi.exit();
})();
