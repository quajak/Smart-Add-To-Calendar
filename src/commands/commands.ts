/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import { Configuration, OpenAIApi } from "openai";

const configuration = new Configuration({
  apiKey: "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",
});
const openai = new OpenAIApi(configuration);

type gptResponse = {
  eventName: string;
  eventDate: string;
  startTime: string;
  endTime: string;
  location?: string;
  meetingLinks?: string[];
};

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const open_message = Office.context.mailbox.item;

  open_message.body.getAsync("text", {}, async function callback(result) {
    const events = await queryGPT("Subject: " + open_message.subject + " Body: " + result.value);
    events.forEach((event) => {
      const appointment = {
        subject: event.eventName,
        body: `${event.meetingLinks || ""} \n\n ${result.value}`,
        start: new Date(event.eventDate + " " + event.startTime),
      } as Office.AppointmentForm;
      if (event.endTime) {
        appointment.end = new Date(event.eventDate + " " + event.endTime);
      }
      if (event.location) {
        appointment.location = event.location;
      }
      Office.context.mailbox.displayNewAppointmentForm(appointment);
    });
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

async function queryGPT(input: string): Promise<gptResponse[]> {
  const query = `The following message contains one or more events. \
  Figure out the date(eventDate), name(eventName), start time(startTime), end time(endTime),optional location(location) and optional meeting links(meetingLinks) of the events \
  and return it in a JSON format (put the dates in yyyy-mm-dd format and time in 24 hour format). "${input}"`;

  const completion = await openai.createChatCompletion({
    model: "gpt-3.5-turbo",
    messages: [{ role: "assistant", content: query }],
  });

  const response = JSON.parse(completion.data.choices[0].message.content);
  return response.events;
}
