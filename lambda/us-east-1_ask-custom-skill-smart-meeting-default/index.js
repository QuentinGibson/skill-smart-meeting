/* eslint-disable  func-names */
/* eslint-disable  no-console */

const Alexa = require('ask-sdk-core')
const moment = require('moment')
const Client = require('@microsoft/microsoft-graph-client').Client
const Fuse = require('fuse.js')

const LaunchRequestHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return handlerInput.requestEnvelope.session.new || request.type === 'LaunchRequest'
  },
  async handle (handlerInput) {
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    const { responseBuilder, attributesManager } = handlerInput
    const attributes = attributesManager.getSessionAttributes()
    attributes.listOfAttendees = []
    attributes.timeSlot = 0
    if (accessToken) {
      const speechText = 'Welcome the the smart meeting finder. Including yourself, how many people are in this meeting?'
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
    const attributes = attributesManager.getSessionAttributes()
    return request.type === 'IntentRequest' &&
      request.intent.name === 'SetUpIntent' &&
      !attributes.size
  },
  handle (handlerInput) {
    const { responseBuilder, attributesManager, requestEnvelope } = handlerInput
    const attributes = attributesManager.getSessionAttributes()
    attributes.size = requestEnvelope.request.intent.slots.number.value - 1
    attributes.duration = requestEnvelope.request.intent.slots.duration.value
    let { accessToken } = requestEnvelope.context.System.user

    if (accessToken) {
      console.log(`Session Attributes: ${JSON.stringify(attributes)}`)
      // SSML tags so Alexa can say things like "first", "second"
      const speechtext = `<speak>What is the name of the person you would like to add?</speak>`
      return responseBuilder
        .speak(speechtext)
        .reprompt(speechtext)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles names at if attendees are not set
const AddPersonIntentHandler = {
  canHandle (handlerInput) {
    let finishAdding = false

    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const attributes = attributesManager.getSessionAttributes()
    const { listOfAttendees, size } = attributes

    if (listOfAttendees.length >= size) {
      finishAdding = true
    }

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AddPersonIntent' &&
      finishAdding === false
  },
  async handle (handlerInput) {
    const { request } = handlerInput.requestEnvelope
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const { slots } = request.intent

    const attributes = attributesManager.getSessionAttributes()
    const { size, timeSlot, listOfAttendees, duration } = attributes

    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      if ((!slots.lastName.value && slots.firstName.value) || (!slots.firstName.value && slots.lastName.value)) {
        const client = Client.init({
          authProvider: (done) => {
            done(null, accessToken)
          }
        })
        let speechText
        let name
        if (slots.firstName.value) {
          name = slots.firstName.value
        } else {
          name = slots.lastName.value
        }
        // Looks for employee in bussiness outlook account. This returns an array with said first name
        let attendee = await findEmployee(client, name).catch((error) => {
          console.log(error)
          responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
        })
        // TODO: Check if attendee exists
        attendee = attendee.value
        const options = {
          shouldSort: true,
          threshold: 0.5,
          keys: ['givenName', 'surname']
        }
        let fuse = new Fuse(attendee, options)
        let result = fuse.search(name)
        // Checks what findEmployee returns
        // Add attendee to the list and checks if the user is completed
        if (result.length === 1) {
          result = result[0]
          attributes.listOfAttendees.push(result)
          speechText = `${result.displayName} has been added to the meeting.`
          if (listOfAttendees.length < size) {
            speechText += ` Please say the name of your next attendee`
          } else {
            console.log(`Session Attributes: ${JSON.stringify(attributes)}`)
            attributes.availableTimes = findFirstTime(attributes)
            let firstTime = attributes.availableTimes[timeSlot].start.value
            speechText = `Your first available time is ${firstTime}. Schedule meeting or find the next available time?`
          }
          // Ask user for last name if multiple are found
        } else if (result.length > 1) {
          speechText = `There was multiple ${name}'s found. Please say the full name of the attendee you would like to add.`
          // No employees were found
        } else {
          speechText = `I'm sorry but I could not find the employee. Please try again or try another name.`
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
        let fuseFilter
        if (attendeeFilter.length === 0) {
          const options = {
            keys: 'givenName'
          }
          let fuse = new Fuse(attendee.value, options)
          fuseFilter = fuse.search(firstName)
        } else {
          attendee = attendeeFilter
        }

        let speechText = ''
        // Add attendee to the list and checks if the user is completed
        if (attendee.length === 1 || fuseFilter) {
          attendee = attendee[0] || fuseFilter[0]
          attributes.listOfAttendees.push(attendee)
          speechText = `${attendee.displayName} has been added to the meeting.`
          if (attributes.listOfAttendees.length < size) {
            speechText += ` Please say the first name of your next attendee`
          } else {
            console.log(`Session Attributes: ${JSON.stringify(attributes)}`)
            let { accessToken } = handlerInput.requestEnvelope.context.System.user
            attributes.availableTimes = findFirstTime(attributes, accessToken)
            let firstTime = attributes.availableTimes[timeSlot].start.value
            speechText = `Your first available time is ${firstTime}. Schedule meeting or find the next available time?`
          }
        // No employee was found
        } else if (attendee.length < 1) {
          speechText = `I can not find the attendee with that last and first name. Please say the first name of your next attendee`
        } else {
          attendee = attendee[0] || fuseFilter[0]
          attributes.listOfAttendees.push(attendee)
          speechText = `${attendee.displayName} has been added to the meeting.`
          if (attributes.listOfAttendees.length < size) {
            speechText += ` Please say the first name of your next attendee`
          } else {
            const times = await findFirstTime(attributes, accessToken).then((result) => {
              console.log(result)
              attributes.availableTimes = result
            })
            let firstTime = times[timeSlot].start.value
            speechText = `Your first available time is ${firstTime}. Schedule meeting or find the next available time?`
          }
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
    let finishAdding = false

    const { request } = handlerInput.requestEnvelope
    const attributesManager = handlerInput.attributesManager
    const attributes = attributesManager.getSessionAttributes()

    if (attributes.listOfAttendees.length >= attributes.size) {
      finishAdding = true
    }

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.YesIntent' &&
      finishAdding
  },
  handle (handlerInput) {
    const responseBuilder = handlerInput.responseBuilder
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    const { attributesManager } = handlerInput
    const attributes = attributesManager.getSessionAttributes()
    if (accessToken) {
      if (!attributes.availableTimes) {
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
    const attributes = attributesManager.getSessionAttributes()
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      if (!attributes.availableTimes) {
        // Resets the skill or closes it base on user input
        attributes.listOfAttendees = []
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
    let finishAdding = false

    const attributesManager = handlerInput.attributesManager
    const attributes = attributesManager.getSessionAttributes()

    if (attributes.listOfAttendees.length >= attributes.size) {
      finishAdding = true
    }

    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AvailableTimeIntent' &&
      finishAdding
  },
  async handle (handlerInput) {
    const { request } = handlerInput.requestEnvelope
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const attributes = attributesManager.getSessionAttributes()
    const { slots } = request.intent
    const { timeSlot } = attributes
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      // This is a list of all the times
      let availableTimes = await findAvailableTimes(client, attributes.listOfAttendees, slots).catch((error) => {
        console.log(error)
        responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
      })
      attributes.availableTimes = availableTimes
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
    const { attributesManager } = handlerInput
    const attributes = attributesManager.getSessionAttributes()
    const request = handlerInput.requestEnvelope.request

    let finishAdding = false
    let finishTime = false

    if (attributes.listOfAttendees.length >= attributes.size) {
      finishAdding = true
    }
    if (attributes.availableTimes) {
      finishTime = true
    }

    return request.type === 'IntentRequest' &&
      request.intent.name === 'TimeSlotIntent' &&
      finishAdding &&
      finishTime
  },
  handle (handlerInput) {
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const attributes = attributesManager.getSessionAttributes()
    attributes.timeSlot += 1
    console.log(`Session Attributes: ${JSON.stringify(attributes)}`)
    const { availableTimes, timeSlot } = attributes
    const currentMeetingTime = availableTimes[timeSlot].start.value
    return responseBuilder.speak(`<speak> Your next available time frame is ${currentMeetingTime}.
    Say yes to set up a meeting  Or say find the next available time?</speak>`)
      .reprompt(`Would you like to set up a meeting at ${currentMeetingTime}.`).getResponse()
  }
}

// Handler for Meeting Intent
const MeetingIntent = {
  canHandle (handlerInput) {
    let finishAdding = false

    const attributesManager = handlerInput.attributesManager
    const attributes = attributesManager.getSessionAttributes()

    if (attributes.listOfAttendees.length >= attributes.size && attributes.availableTimes) {
      finishAdding = true
    }

    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'MeetingIntent' &&
      finishAdding
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const attributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots

    const subject = slots.subject.value
    const { availableTimes } = attributes
    const { timeSlot } = attributes
    const { listOfAttendees } = attributes
    const meetingTime = availableTimes[timeSlot]
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      let response = `Your meeting has been created!`

      let result = createMeeting(client, subject, meetingTime, listOfAttendees)
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
    const speechText = 'Find the best time for a group meeting and schedule it.'

    return handlerInput.responseBuilder
      .speak(speechText)
      .reprompt(speechText)
      .getResponse()
  }
}

// TODO: Add conditions for handlers that are strict in nature.
const FallbackHandler = {
  canHandle (handlerInput) {
    // handle fallback intent, yes and no when playing a game
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.FallbackIntent'
  },
  handle (handlerInput) {
    // currently playing
    return handlerInput.responseBuilder
      .speak('Im sorry I didnt understand that.')
      .reprompt('Im sorry I didnt understand that.')
      .getResponse()
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
const UnhandledIntent = {
  canHandle () {
    return true
  },
  handle (handlerInput) {
    const outputSpeech = 'Im sorry, I do not understand that phrase? Could you rephrase that?'
    return handlerInput.responseBuilder
      .speak(outputSpeech)
      .reprompt(outputSpeech)
      .getResponse()
  }
}

// Handles non-linked accounts
function askToLink (handlerInput) {
  const speechText = 'Please link your account to use this skill.'
  return handlerInput.responseBuilder.speak(speechText).getResponse()
}

async function findFirstTime (attributes, token) {
  let { listOfAttendees, duration } = attributes
  console.log(`Session Attributes: ${JSON.stringify(attributes)}`)
  const client = Client.init({
    authProvider: (done) => {
      done(null, token)
    }
  })
  let result = await findAvailableTimes(client, listOfAttendees, duration)
  console.log(JSON.stringify(result))
  return result
}

// Returns a array of all employees with a given first name.
const findEmployee = (client, givenName) => client.api('me/people').search(givenName).get()

// Returns an array of available times for the meetings
async function findAvailableTimes (client, attendees, duration) {
  const meetingDetail = () => {
    let startDate = moment()
    let endDate = moment(startDate).add(90, 'days')
    let startTime = '07:00:00'
    let endTime = '19:00:00'
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

      // Convert all time to moment objects that are in a certain Format
      for (let time of meetingTimeSuggestions) {
        const timeFormat = 'MMMM Do, h:mma'
        const timeSlot = time.meetingTimeSlot

        timeSlot.start.value = moment(timeSlot.start.dateTime).format(timeFormat)
        timeSlot.end.value = moment(timeSlot.end.dateTime).format(timeFormat)
        meetingTimes.push(timeSlot)
      }
      meetingTimes.sort((a, b) => new Date(a.start.dateTime) - new Date(b.end.dateTime))
      // Sorting the times from soonest to latest
      console.log(`Meeting Times => ${JSON.stringify(meetingTimes)}`)
      return meetingTimes
    })
}

async function createMeeting (client, subject, meetingTime, attendees) {
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
      content: `Created by Atlanticus's Smart Finder`
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
    FallbackHandler,
    UnhandledIntent,
    SessionEndedRequestHandler
  )
  .addErrorHandlers(ErrorHandler)
  .lambda()
