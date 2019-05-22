/* eslint-disable  func-names */
/* eslint-disable  no-console */

const Alexa = require('ask-sdk-core')
const moment = require('moment')
const Graph = require('@microsoft/microsoft-graph-client')

const LaunchRequestHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'LaunchRequest'
  },
  handle (handlerInput) {
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    let responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    if (accessToken) {
      const speechText = 'Including yourself how many people are in this meeting?'
      sessionAttributes.init = false

      return responseBuilder
        .speak(speechText)
        .reprompt(speechText)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles number at start
const SetUpIntentHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()

    return request.type === 'IntentRequest' &&
      request.intent.name === 'SetUpIntent' &&
      !sessionAttributes.init
  },
  handle (handlerInput) {
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const responseBuilder = handlerInput.responseBuilder
    const personNumber = sessionAttributes.listOfAttendees.length + 1
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      sessionAttributes.init = true
      // SSML tags so Alexa can say things like "first", "second"
      const speechtext = `<speak>What is the first name of the <say-as interpret-as='ordinal'>${personNumber}</say-as> person you would like to add?</speak>`
      return responseBuilder
        .speak(speechtext)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles first names at if attendees are not set
const AddPersonByFirstNameIntentHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AddPersonIntent' &&
      sessionAttributes.init
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const { listOfAttendees } = sessionAttributes
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      const client = Graph.client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      // TODO: Remove in production
      console.log(client)
      // -------------------------
      let speechText
      // Looks for employee in bussiness outlook account. This returns an array with said first name
      const attendee = await findEmployee(client, slots.firstName.value).catch((error) => {
        console.log(error)
        responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
      })
      // Checks what findEmployee returns
      // Add attendee to the list and checks if the user is completed
      if (attendee.length === 1) {
        listOfAttendees.append(attendee[0])
        slots.firstName.value = ''
        speechText = `${attendee.value.displayName} has been added to the meeting.`
        if (attendee.length < numberOfAttendees) {
          speechText += ` Please say the first name of your next attendee`
        } else {
          speechText += ` Would you like to find a meeting time?`
        }
      // Ask user for last name if multiple are found
      } else if (attendee.length < 1) {
        speechText = `There was multiple ${slots.firstName.value}s found. Please say the full name of the attendee you would like to add.`
      // No employees were found
      } else {
        speechText = `I'm sorry but I could not find the employee. Please try again or try another first name.`
      }
      slots.firstName.value = ''
      return responseBuilder
        .speak(speechText)
        .reprompt(speechText)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles full names at if attendees are not set
const AddPersonByFullNameIntentHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AddPersonIntent' &&
      request.slots.firstName.value &&
      request.slots.lastName.value &&
      sessionAttributes.listOfAttendees < numberOfAttendees &&
      sessionAttributes.init
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const { listOfAttendees } = sessionAttributes
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      const client = Graph.client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })

      // Looks for employee in bussiness outlook account. This returns an array with said first name
      const attendee = findEmployee(client, `${slots.firstName.value} ${slots.lastName.value}`).catch((error) => {
        console.log(error)
        responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
      })
      let speechText
      // Checks what findEmployee returns
      // Add attendee to the list and checks if the user is completed
      if (attendee.length === 1) {
        listOfAttendees.append(attendee[0])
        slots.firstName.value = ''
        slots.lastName.value = ''
        speechText = `${attendee.value.displayName} has been added to the meeting.`
        if (attendee.length < numberOfAttendees) {
          speechText += ` Please say the first name of your next attendee`
        } else {
          speechText += ` Would you like to find the best available meeting time?`
        }
      // No employee was found
      } else if (attendee.length > 1) {
        speechText = `I can not find the employee with that last or first name. Please say the first name of your next attendee`
      } else {
        speechText = `There are multiple attendees with that first and last name. Please say the first name of your next attendee`
      }
      return responseBuilder
        .speak(speechText)
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles yes if attendees are set
const YesStartMeetingHandler = {
  canhandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.YesIntent' &&
      sessionAttributes.listOfAttendees >= numberOfAttendees &&
      !sessionAttributes.availableTimes
  },
  async handle (handlerInput) {
    const responseBuilder = handlerInput.responseBuilder
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
    // Collects the sltos needed for the meeting
      return responseBuilder
        .speak(`When is the earliest date for the meeting?`)
        .reprompt(`Say the earliest date for the meeting.`)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles no if attendees are set
const NoStartMeetingHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.NoIntent' &&
      sessionAttributes.listOfAttendees >= numberOfAttendees
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
    // Resets the skill or closes it base on user input
      sessionAttributes.listOfAttendees = []
      slots.number.value = ''
      return responseBuilder
        .speak(`If you would like to cancel the meeting say stop. Otherwise if you would like to start over, say the number of attendees attending the meeting.`)
        .reprompt(`Say stop or a first name`)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles meeting intent
const AvailableTimeIntent = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AvailableTimeIntent'
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    sessionAttributes.timeSlot = 0
    const timeSlot = sessionAttributes.timeSlot
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Graph.client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      sessionAttributes.listOfAttendees = sessionAttributes.listOfAttendees.map(attendee => {
        return {
          emailAddress: {
            name: `${attendee.givenName} ${attendee.surname}`,
            address: attendee.scoredEmailAddresses[0].address
          }
        }
      })
      // This is a list of all the times
      let { availableTimes } = await findAvailableTimes(client, sessionAttributes.listOfAttendees, slots).catch((error) => {
        console.log(error)
        responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
      })
      sessionAttributes.availableTimes = availableTimes
      // Checks if there is a available time
      if (availableTimes.length === 0) {
      // Ends the skill with a message
        return responseBuilder
          .speak(`There are no available times within the time frame.`)
          .getResponse()
      } else {
      // Asks the person to confirm the time
        return responseBuilder
          .speak(`<speak> Your <say-as interpret-as='ordinal'>${timeSlot + 1}</say-as> available time frame is ${availableTimes[timeSlot].start.value}.
            Would you like to set up a meeting then? Or find the next available time?</speak>`)
          .reprompt(`Would you like to set up a meeting at ${availableTimes[timeSlot].start.value}.`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handler for Meeting Intent
const MeetingIntent = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1

    return request.type === 'IntentRequest' &&
      request.intent.name === 'MeetingIntent' &&
      sessionAttributes.listOfAttendees >= numberOfAttendees &&
      sessionAttributes.availableTimes
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots

    const context = slots.context.value
    const subject = slots.subject.value
    const { availableTimes } = sessionAttributes
    const { timeSlot } = sessionAttributes
    const { listOfAttendees } = sessionAttributes
    const meetingTime = availableTimes[timeSlot]
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Graph.client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      let response = `Your meeting has been created!`

      let result = createMeeting(client, subject, context, meetingTime, listOfAttendees)
        .catch(error => {
          console.log(error)
          response = `There was an error speaking to outlook`
        })
      if (result) {
        return responseBuilder
          .speak(response)
          .getResponse()
      } else {
        return responseBuilder
          .speak(`Result came back empty`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles yes if available times are set
const YesAvailableTimeHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.YesIntent' &&
      sessionAttributes.listOfAttendees >= numberOfAttendees &&
      sessionAttributes.availableTimes
  },
  async handle (handlerInput) {
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      const responseBuilder = handlerInput.responseBuilder
      return responseBuilder
        .speak(`What is the subject of the meeting?`)
        .reprompt(`What would you like the subject of the meeting to be?`)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles no if available times are set
const NoAvailableTimeHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots
    const numberOfAttendees = slots.number.value - 1

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.NoIntent' &&
      sessionAttributes.listOfAttendees >= numberOfAttendees &&
      sessionAttributes.availableTimes
  },
  async handle (handlerInput) {
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      const responseBuilder = handlerInput.responseBuilder
      return responseBuilder
        .speak(`Goodbye!`)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

const HelpHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.HelpIntent'
  },
  handle (handlerInput) {
    const speechText = 'You can name someone to add to the meeting.'
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      return handlerInput.responseBuilder
        .speak(speechText)
        .reprompt(speechText)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

const ExitHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      (request.intent.name === 'AMAZON.CancelIntent' ||
        request.intent.name === 'AMAZON.StopIntent')
  },
  handle (handlerInput) {
    return handlerInput.responseBuilder
      .speak(`Good bye!`)
      .getResponse()
  }
}

const SessionEndedRequestHandler = {
  canHandle (handlerInput) {
    return handlerInput.requestEnvelope.request.type === 'SessionEndedRequest'
  },
  handle (handlerInput) {
    if (handlerInput.requestEnvelope.request.error != null) {
      console.log(JSON.stringify(handlerInput.requestEnvelope.request.error))
    }

    return handlerInput.responseBuilder.getResponse()
  }
}

const ErrorHandler = {
  canHandle () {
    return true
  },
  handle (handlerInput, error) {
    console.log(`Error handled: ${error.message}`)
    console.log(`Error stack: ${error.stack}`)
    const speechText = `Session ended in error`

    return handlerInput.responseBuilder
      .speak(speechText)
      .reprompt(speechText)
      .getResponse()
  }
}

// Handles non-linked accounts
function askToLink (handlerInput) {
  const speechText = 'Please link your account to use this skill.'
  return handlerInput.responseBuilder.speak(speechText).getResponse()
}

// Returns a array of all employees with a given first name.
const findEmployee = (client, givenName) => client.api('me/people').search(givenName).get()

// Returns an array of available times for the meetings
async function findAvailableTimes (client, attendees, slots) {
  const meetingDetail = () => {
    let duration = slots.duration.value
    let startDate = slots.startDate.value
    let endDate = slots.endDate.value
    let startTime = slots.startTime.value || '00:00:00'
    let endTime = slots.endTime.value || '23:59:59'

    // Moment time format Year - Month - Day - T - Hour - Minutes - Seconds - Milliseconds
    const timeFormat = 'YYYY-MM-DDTHH:mm:ss.SSS'
    let dateArray = []
    let currentDate = startDate
    // Creating a list of time periods from the start date to the end date
    while (currentDate.isSameOrBefore(endDate)) {
      // Object reprsenting the earilest someone could meet at a certain day to the latest
      let meetingPeriod = {
        start: {
          dateTime: moment(`${currentDate}T${startTime}`).format(timeFormat) + 'Z',
          timeZone: 'Eastern Standard Time'
        },
        end: {
          dateTime: moment(`${currentDate}T${endTime}`).format(timeFormat) + 'Z',
          timeZone: 'Eastern Standard Time'
        }
      }
      dateArray.push(meetingPeriod)
      // Add one day and skips weekends 0 = Sundays, 6 = Saturdays
      do {
        currentDate = currentDate.add(1, 'days')
      } while (currentDate.day() === 0 || currentDate.day() === 6)
    }

    // Format the email address information of every attendee
    attendees = attendees.map(attendee => {
      return Object.assign({}, attendee, { type: `required` })
    })

    // Formated for post request to graph api
    return {
      attendees: attendees,
      timeConstraint: {
        // Makes sure to always use work hours of the atteende
        activityDomain: `personal`,
        timeslots: dateArray
      },
      meetingDuration: duration
    }
  }

  return client.api('/me/findMeetingTimes').post(meetingDetail())
    .then(result => {
      const { meetingTimeSuggestions } = result
      const meetingTimes = []

      // Convert all time to moment objects
      for (let time of meetingTimeSuggestions) {
        const timeSlot = time.meetingTimeSlot
        timeSlot.start.value = moment(timeSlot.start.dateTime).format('MMMM Do, h:mma')
        timeSlot.end.value = moment(timeSlot.end.dateTime).format('MMMM Do, h:mma')
        meetingTimes.push(timeSlot)
      }

      // Sorting the times from soonest to latest
      return meetingTimes.sort((a, b) => a.start.value - b.start.value)
    })
}

async function createMeeting (client, subject, content, meetingTime, attendees) {
  const eventDetails = {
    subject: subject,
    body: {
      contentType: 'Text',
      content: content
    },
    start: {
      dateTime: meetingTime.start.dateTime,
      timeZone: meetingTime.start.timeZone
    },
    end: {
      dateTime: meetingTime.end.dateTime,
      timeZone: meetingTime.end.timeZone
    },
    attendees: attendees,
    type: 'singleInstance'
  }
  let response = await client.api('/me/calendar/events').post({ event: eventDetails })
  return response
}
const skillBuilder = Alexa.SkillBuilders.custom()

exports.handler = skillBuilder
  .addRequestHandlers(
    LaunchRequestHandler,
    SetUpIntentHandler,
    YesStartMeetingHandler,
    NoStartMeetingHandler,
    MeetingIntent,
    YesAvailableTimeHandler,
    NoAvailableTimeHandler,
    AvailableTimeIntent,
    AddPersonByFullNameIntentHandler,
    AddPersonByFirstNameIntentHandler,
    HelpHandler,
    ExitHandler,
    SessionEndedRequestHandler
  )
  .addErrorHandlers(ErrorHandler)
  .lambda()
