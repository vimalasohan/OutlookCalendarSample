//
//  GraphManager.swift
//  GraphTutorial
//
//  Copyright © 2019 Microsoft. All rights reserved.
//  Licensed under the MIT license. See LICENSE.txt in the project root for license information.
//

import Foundation
import MSGraphClientSDK
import MSGraphClientModels

class GraphManager {

    // Implement singleton pattern
    static let instance = GraphManager()

    private let client: MSHTTPClient?

    private init() {
        client = MSClientFactory.createHTTPClient(with: AuthenticationManager.instance)
    }

    public func getMe(completion: @escaping(MSGraphUser?, Error?) -> Void) {
        // GET /me
        let meRequest = NSMutableURLRequest(url: URL(string: "\(MSGraphBaseURL)/me")!)
        let meDataTask = MSURLSessionDataTask(request: meRequest, client: self.client, completion: {
            (data: Data?, response: URLResponse?, graphError: Error?) in
            guard let meData = data, graphError == nil else {
                completion(nil, graphError)
                return
            }

            do {
                // Deserialize response as a user
                let user = try MSGraphUser(data: meData)
                completion(user, nil)
            } catch {
                completion(nil, error)
            }
        })

        // Execute the request
        meDataTask?.execute()
    }
    
//    public func createEvent(completion: @escaping(Error?) -> Void) {
//        let availableRequest = NSMutableURLRequest(url: URL(string: "\(MSGraphBaseURL)/me/calendar/events")!)
//        availableRequest.httpMethod = "POST"
//        availableRequest.setValue("application/json", forHTTPHeaderField: "Content-Type")
//        let event = MSGraphEvent()
//        event.subject = "Meeting with FA"
//        let eventBody = MSGraphItemBody()
//        eventBody.contentType = MSGraphBodyType.html()
//        eventBody.content = "Mid Monthly Investment Review"
//        event.body = eventBody
//        let startTime = MSGraphDateTimeTimeZone()
//        startTime.dateTime = "2020-04-07T09:00:00"
//        startTime.timeZone = "Eastern Standard Time"
//        event.start = startTime
//        let endTime = MSGraphDateTimeTimeZone()
//        endTime.dateTime = "2020-04-07T10:00:00"
//        endTime.timeZone = "Eastern Standard Time"
//        event.end = endTime
//        let location = MSGraphLocation()
//        location.displayName = "Skype"
//        event.location = location
//        let attendeeList = NSMutableArray()
//        let attendees = MSGraphAttendee()
//        let emailAddress = MSGraphEmailAddress()
//        emailAddress.address = "vimalasohan@gmail.com"
//        emailAddress.name = "Vimal Asohan"
//        attendees.emailAddress = emailAddress
//        attendees.type = MSGraphAttendeeType.required()
//        attendeeList.add(attendees)
//        event.attendees = attendeeList as? [Any]
//        do {
//            let eventData = try event.getSerializedData()
//            availableRequest.httpBody = eventData
//            print(eventData)
//        }
//        catch let error
//        {
//            print(error)
//        }
//        let eventsDataTask = MSURLSessionDataTask(request: availableRequest, client: self.client, completion: {
//            (data: Data?, response: URLResponse?, graphError: Error?) in
//            guard let _ = data, graphError == nil else {
//                completion(graphError)
//                return
//            }
//        })
//
//        // Execute the request
//        eventsDataTask?.execute()
//
//    }
    
    public func getAvailabilty(completion: @escaping([MSGraphScheduleItem]?, Error?) -> Void) {
        let availableRequest = NSMutableURLRequest(url: URL(string: "\(MSGraphBaseURL)/me/calendar/getSchedule")!)
        availableRequest.httpMethod = "POST"
        availableRequest.setValue("outlook.timezone=\"Eastern Standard Time\"", forHTTPHeaderField: "Prefer")
        availableRequest.setValue("application/json", forHTTPHeaderField: "Content-Type")
        let payloadDictionary = NSMutableDictionary()
        let scheduleList = NSMutableArray()
        scheduleList.add("as_vimal@yahoo.co.in")
        payloadDictionary["schedules"] = scheduleList
        let startTime = [
            "dateTime" : "2020-04-07T09:00:00",
            "timeZone" : "Eastern Standard Time"
    ]
        payloadDictionary["startTime"] = startTime
        let endTime = [
            "dateTime" : "2020-04-07T18:00:00",
            "timeZone" : "Eastern Standard Time"
            ]
        payloadDictionary["endTime"] = endTime
        let availabilityInterval = 60
        payloadDictionary["availabilityViewInterval"] = availabilityInterval
        do {
            let data = try JSONSerialization.data(withJSONObject: payloadDictionary, options: .sortedKeys)
            availableRequest.httpBody = data
            print(data)
        }
        catch let error
        {
            print(error)
        }

        let availableDataTask = MSURLSessionDataTask(request: availableRequest, client: self.client, completion: {
            (data: Data?, response: URLResponse?, graphError: Error?) in
            guard let eventsData = data, graphError == nil else {
                completion(nil, graphError)
                return
            }

            do {
                // Deserialize response as a user
//                let user = try MSGraphAttendeeAvailability(data: meData).availability
//                completion(user, nil)
                 let eventsCollection = try MSCollection(data: eventsData)
                var eventArray: [MSGraphScheduleItem] = []
//
                eventsCollection.value.forEach({
                    (rawEvent: Any) in
                    // Convert JSON to a dictionary
                    guard let eventDict = rawEvent as? [String: Any] else {
                        return
                    }

                    // Deserialize event from the dictionary
                    let event = MSGraphScheduleItem(dictionary: eventDict)!
                    eventArray.append(event)
                })

                // Return the array
                completion(eventArray, nil)

            } catch {
                completion(nil, error)
            }
        })

        // Execute the request
        availableDataTask?.execute()

    }
    public func getEvents(completion: @escaping([MSGraphEvent]?, Error?) -> Void) {
        // GET /me/events?$select='subject,organizer,start,end'$orderby=createdDateTime DESC
        // Only return these fields in results
        let select = "$select=subject,organizer,start,end"
        // Sort results by when they were created, newest first
        let orderBy = "$orderby=createdDateTime+DESC"
        let eventsRequest = NSMutableURLRequest(url: URL(string: "\(MSGraphBaseURL)/me/events?\(select)&\(orderBy)")!)
        let eventsDataTask = MSURLSessionDataTask(request: eventsRequest, client: self.client, completion: {
            (data: Data?, response: URLResponse?, graphError: Error?) in
            guard let eventsData = data, graphError == nil else {
                completion(nil, graphError)
                return
            }
            
            do {
                // Deserialize response as events collection
                let eventsCollection = try MSCollection(data: eventsData)
                var eventArray: [MSGraphEvent] = []
                
                eventsCollection.value.forEach({
                    (rawEvent: Any) in
                    // Convert JSON to a dictionary
                    guard let eventDict = rawEvent as? [String: Any] else {
                        return
                    }
                    
                    // Deserialize event from the dictionary
                    let event = MSGraphEvent(dictionary: eventDict)!
                    eventArray.append(event)
                })
                
                // Return the array
                completion(eventArray, nil)
            } catch {
                completion(nil, error)
            }
        })
        
        // Execute the request
        eventsDataTask?.execute()
    }
}
