/* eslint-disable  func-names */
/* eslint-disable  no-console */

const Alexa = require('ask-sdk-core')
const moment = require('moment')
const Client = require('@microsoft/microsoft-graph-client').Client
const Fuse = require('fuse.js')

const LaunchRequestHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'LaunchRequest'
  },
  async handle (handlerInput) {
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    let responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    sessionAttributes.init = false
    sessionAttributes.listOfAttendees = []
    if (accessToken) {
      const speechText = 'Including yourself how many people are in this meeting?'
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
    sessionAttributes.size = handlerInput.requestEnvelope.request.intent.slots.number.value - 1
    sessionAttributes.timeSlot = 0
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      sessionAttributes.init = true
      // SSML tags so Alexa can say things like "first", "second"
      const speechtext = `<speak>What is the first name of the <say-as interpret-as='ordinal'>1</say-as> person you would like to add?</speak>`
      return responseBuilder
        .speak(speechtext)
        .reprompt(speechtext)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles first names at if attendees are not set
const AddPersonIntentHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AddPersonIntent' &&
      sessionAttributes.init
  },
  async handle (handlerInput) {
    const { request } = handlerInput.requestEnvelope
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const { slots } = request.intent

    const sessionAttributes = attributesManager.getSessionAttributes()
    const { size } = sessionAttributes
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      if (!slots.lastName.value) {
        const client = Client.init({
          authProvider: (done) => {
            done(null, accessToken)
          }
        })
        let speechText
        // Looks for employee in bussiness outlook account. This returns an array with said first name
        let attendee = await findEmployee(client, slots.firstName.value).catch((error) => {
          console.log(error)
          responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
        })
        attendee = attendee.value.filter(employee => {
          return employee.givenName.toLowerCase() === slots.firstName.value.toLowerCase()
        })
        // Checks what findEmployee returns
        // Add attendee to the list and checks if the user is completed
        if (attendee.length === 1) {
          attendee = attendee[0]
          sessionAttributes.listOfAttendees.push(attendee)
          speechText = `${attendee.displayName} has been added to the meeting.`
          if (sessionAttributes.listOfAttendees.length < size) {
            speechText += ` Please say the first name of your next attendee`
          } else {
            speechText += ` Would you like to find a meeting time?`
          }
          // Ask user for last name if multiple are found
        } else if (attendee.length > 1) {
          speechText = `There was multiple ${slots.firstName.value}s found. Please say the full name of the attendee you would like to add.`
          // No employees were found
        } else {
          speechText = `I'm sorry but I could not find the employee. Please try again or try another first name.`
        }
        return responseBuilder
          .speak(speechText)
          .reprompt(speechText)
          .getResponse()
      } else {
        const firstName = slots.firstName.value
        const lastName = slots.lastName.value
        const client = Client.init({
          authProvider: (done) => {
            done(null, accessToken)
          }
        })
        // Looks for employee in bussiness outlook account. This returns an array with said first name
        let attendee = await findEmployee(client, `${lastName}`).catch((error) => {
          console.log(error)
          responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
        })
        let attendeeFilter = attendee.value.filter(employee => {
          return employee.displayName.toLowerCase().includes(firstName.toLowerCase())
        })
        if (attendeeFilter.length === 0) {
          const options = {
            keys: 'givenName',
            id: 'displayName'
          }
          let fuse = new Fuse(attendee.value, options)
          attendee = fuse.search(firstName)
        } else {
          attendee = attendeeFilter
        }
        let speechText
        // Add attendee to the list and checks if the user is completed
        if (attendee.length === 1) {
          attendee = attendee[0]
          sessionAttributes.listOfAttendees.push(attendee)
          speechText = `${attendee.displayName} has been added to the meeting.`
          if (attendee.length < size) {
            speechText += ` Please say the first name of your next attendee`
          } else {
            speechText += ` Would you like to find the best available meeting time?`
          }
        // No employee was found
        } else if (attendee.length < 1) {
          speechText = `I can not find the employee with that last and first name. Please say the first name of your next attendee`
        } else {
          speechText = `There are multiple attendees with that first and last name. Please say the first name of your next attendee`
        }
        return responseBuilder
          .speak(speechText)
          .reprompt(speechText)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles yes based on if a time is set
const YesStartMeetingHandler = {
  canHandle (handlerInput) {
    const { request } = handlerInput.requestEnvelope
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.YesIntent'
  },
  handle (handlerInput) {
    const responseBuilder = handlerInput.responseBuilder
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    if (accessToken) {
      if (!sessionAttributes.availableTimes) {
        // Collects the sltos needed for the meeting
        return responseBuilder
          .speak(`When is the earliest date for the meeting?`)
          .reprompt(`Say the earliest date for the meeting.`)
          .getResponse()
      } else {
        const responseBuilder = handlerInput.responseBuilder
        return responseBuilder
          .speak(`What is the subject of the meeting?`)
          .reprompt(`What would you like the subject of the meeting to be?`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles no if attendees are set
const NoStartMeetingHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.NoIntent'
  },
  async handle (handlerInput) {
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      if (!sessionAttributes.availableTimes) {
        // Resets the skill or closes it base on user input
        sessionAttributes.listOfAttendees = []
        return responseBuilder
          .speak(`If you would like to cancel the meeting say stop. Otherwise if you would like to start over, say the number of attendees attending the meeting.`)
          .reprompt(`Say stop or a first name`)
          .getResponse()
      } else {
        const responseBuilder = handlerInput.responseBuilder
        return responseBuilder
          .speak(`Goodbye!`)
          .getResponse()
      }
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
    const { request } = handlerInput.requestEnvelope
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    const { slots } = request.intent
    const { timeSlot } = sessionAttributes
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      // This is a list of all the times
      let availableTimes = await findAvailableTimes(client, sessionAttributes.listOfAttendees, slots).catch((error) => {
        console.log(error)
        responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
      })
      sessionAttributes.availableTimes = availableTimes
      const currentMeetingTime = availableTimes[timeSlot].start.value
      // Checks if there is a available time
      if (availableTimes.length === 0) {
      // Ends the skill with a message
        return responseBuilder
          .speak(`There are no available times within the time frame.`)
          .getResponse()
      } else {
      // Asks the person to confirm the time
        return responseBuilder
          .speak(`Your first available time frame is ${currentMeetingTime}.
            Would you like to set up a meeting then? Or find the next available time?`)
          .reprompt(`Would you like to set up a meeting at ${currentMeetingTime}.`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

const TimeSlotHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'TimeSlotIntent'
  },
  handle (handlerInput) {
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    const { timeSlot, availableTimes } = sessionAttributes
    const currentMeetingTime = availableTimes[timeSlot].start.value
    return responseBuilder.speak(`<speak> Your next available time frame is ${currentMeetingTime}.
    Would you like to set up a meeting then? Or find the next available time?</speak>`)
      .reprompt(`Would you like to set up a meeting at ${currentMeetingTime}.`).getResponse()
  }
}

// Handler for Meeting Intent
const MeetingIntent = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'MeetingIntent'
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
      const client = Client.init({
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
        console.log('results: ' + JSON.stringify(result))
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
    let startTime = slots.startTime.value || '07:00:00'
    let endTime = slots.endTime.value || '18:59:59'
    // Moment time format Year - Month - Day - T - Hour - Minutes - Seconds - Milliseconds
    const timeFormat = 'YYYY-MM-DDTHH:mm:ss.SSS'
    let dateArray = []
    let currentMoment = moment(startDate)
    // Creating a list of time periods from the start date to the end date
    while (currentMoment.isSameOrBefore(endDate)) {
      // Object reprsenting the earilest someone could meet at a certain day to the latest
      let dateTimeStart = moment(`${currentMoment.format('YYYY-MM-DD')}T${startTime}`).format(timeFormat) + 'Z'
      let dateTimeEnd = moment(`${currentMoment.format('YYYY-MM-DD')}T${endTime}`).format(timeFormat) + 'Z'
      let meetingPeriod = {
        start: {
          dateTime: dateTimeStart,
          timeZone: 'Eastern Standard Time'
        },
        end: {
          dateTime: dateTimeEnd,
          timeZone: 'Eastern Standard Time'
        }
      }
      dateArray.push(meetingPeriod)
      // Add one day and skips weekends 0 = Sundays, 6 = Saturdays
      do {
        currentMoment = currentMoment.add(1, 'days')
      } while (currentMoment.day() === 0 || currentMoment.day() === 6)
    }

    // Format the email address information of every attendee
    attendees = attendees.map(attendee => {
      return {
        type: `required`,
        emailAddress: {
          address: attendee.scoredEmailAddresses[0].address,
          name: attendee.displayName
        }
      }
    })
    // Formated for post request to graph api
    const returnValue = {
      attendees: attendees,
      timeConstraint: {
        // Makes sure to always use work hours of the atteende
        activityDomain: `personal`,
        timeslots: dateArray
      },
      meetingDuration: duration
    }
    return returnValue
  }

  return client.api('/me/findMeetingTimes').post(meetingDetail())
    .then(result => {
      const { meetingTimeSuggestions } = result
      const meetingTimes = []

      // Convert all time to moment objects that are in a certain Format
      for (let time of meetingTimeSuggestions) {
        const timeFormat = 'MMMM Do, h:mma'
        const timeSlot = time.meetingTimeSlot

        timeSlot.start.value = moment(timeSlot.start.dateTime).format(timeFormat)
        timeSlot.end.value = moment(timeSlot.end.dateTime).format(timeFormat)
        meetingTimes.push(timeSlot)
      }

      // Sorting the times from soonest to latest
      return meetingTimes.sort((a, b) => a.start.value - b.start.value)
    })
}

async function createMeeting (client, subject, content, meetingTime, attendees) {
  attendees = attendees.map(attendee => {
    return {
      type: `required`,
      emailAddress: {
        address: attendee.scoredEmailAddresses[0].address,
        name: attendee.displayName
      }
    }
  })
  const event = {
    subject: subject,
    body: {
      contentType: 'HTML',
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
    attendees: attendees
  }
  return client.api('/me/events').post(event)
    .then(result => console.log(result))
}
const skillBuilder = Alexa.SkillBuilders.custom()

exports.handler = skillBuilder
  .addRequestHandlers(
    LaunchRequestHandler,
    SetUpIntentHandler,
    AddPersonIntentHandler,
    AvailableTimeIntent,
    MeetingIntent,
    TimeSlotHandler,
    YesStartMeetingHandler,
    NoStartMeetingHandler,
    HelpHandler,
    ExitHandler,
    SessionEndedRequestHandler
  )
  .addErrorHandlers(ErrorHandler)
  .lambda()
