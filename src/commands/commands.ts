/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import { Configuration, OpenAIApi } from "openai";

const configuration = new Configuration({
    organization: "brunner40s",
    apiKey: "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",
});
const openai = new OpenAIApi(configuration);

type gptResponse = {
  eventName: string,
  eventDate: string,
  startTime: string,
  endTime: string,
  links?: string[],
}

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const open_message = Office.context.mailbox.item;

  Office.context.mailbox.item.body.getAsync("text", {}, function callback(result) {
    Office.context.mailbox.displayNewAppointmentForm({
      subject: "Event: " + open_message.subject,
      body: result.value,
    } as Office.AppointmentForm);
  });

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;


async function queryGPT(input: string): Promise<gptResponse>{

  const query = `The following message contains one or more events. \
  Figure out the date(eventDate), name(eventName), start time(startTime) and end time(endTime) of the events \
  and return it in a JSON format (put the dates in dd-mm-yyyy format and time in 24 hour format). If there are any links put them in an array of strings. "${input}"`;

  const completion = await openai.createChatCompletion({
    model: "gpt-3.5-turbo",
    messages: [{"role": "assistant", "content": query}],
  });

  const response  = JSON.parse(completion.data.choices[0].message.content);
  return response;
};