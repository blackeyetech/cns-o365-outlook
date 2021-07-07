import { CNMsGraphApi } from "./o365-outlook";

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
  CHECK_SENT = "Check Email",
  REPLY = "Reply to email",
  GET = "Get Conversation",
  GET_UNREAD = "Get for unread emails",
  GET_UNREAD_CONVERSATION = "Get for unread emails from conversation",
  MARK_READ = "Mark email as read",
  QUIT = "Quit",
}

// Tests here
async function sendMessage(userEmail: string, msGraphApi: CNMsGraphApi) {
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

  let attach = answer[Prompts.ATTACHMENT];
  let contentType = "";
  let content: Buffer = Buffer.from("");

  if (attach.lenth) {
    content = fs.readFileSync(attach);

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
    attach.lenth
      ? [
          {
            name: path.basename(attach),
            contentType,
            contentB64: content.toString("base64"),
          },
        ]
      : undefined,
  );
}

async function replyTo(userEmail: string, msGraphApi: CNMsGraphApi) {
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

async function checkMessage(userEmail: string, msGraphApi: CNMsGraphApi) {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.REF,
      message: "Input ref code to check for:",
    },
  ]);

  let check = await msGraphApi.getMessage(userEmail, answer[Prompts.REF]);

  if (check !== undefined) {
    console.log("Email for ref code %s was sent", answer[Prompts.REF]);
  } else {
    console.log("Email for ref code %s was not sent", answer[Prompts.REF]);
  }
}

async function getConversation(userEmail: string, msGraphApi: CNMsGraphApi) {
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

async function getUnreadMessages(userEmail: string, msGraphApi: CNMsGraphApi) {
  let messages = await msGraphApi.getUnreadMessages(userEmail);

  console.log(JSON.stringify(messages, undefined, 2));
}

async function getUnreadMessagesFromConversation(
  userEmail: string,
  msGraphApi: CNMsGraphApi,
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

async function markEmailRead(userEmail: string, msGraphApi: CNMsGraphApi) {
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
  let msGraphApi = new CNMsGraphApi("Test-MS-Graph-API");
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
          TestChoices.CHECK_SENT,
          TestChoices.REPLY,
          TestChoices.SEND,
          TestChoices.GET_UNREAD,
          TestChoices.GET_UNREAD_CONVERSATION,
          TestChoices.MARK_READ,
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
      case TestChoices.CHECK_SENT:
        await checkMessage(userEmail, msGraphApi);
        break;
      case TestChoices.REPLY:
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
    }
  }

  msGraphApi.exit();
})();
