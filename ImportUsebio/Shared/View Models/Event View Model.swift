//
//  Event View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/08/2025.
//

import Combine
import SwiftUI
import CoreData

public class EventViewModel : ViewModel, ObservableObject {
    
    // Properties in core data model
    @Published private(set) var eventId: UUID = UUID() ; public override var id: UUID { eventId }
    @Published public var eventCode: String = ""
    @Published public var eventName: String = ""
    @Published public var active: Bool = true
    @Published public var startDate: Date?
    @Published public var endDate: Date?
    @Published public var validMinRank: Int = 0
    @Published public var validMaxRank: Int = 0
    @Published public var originatingClubCode: String = ""
    @Published public var clubMandatory: Bool = false
    @Published public var nationalAllowed: Bool = true
    @Published public var localAllowed: Bool = true
    
    @Published public var eventCodeMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.event")!)!
    
    // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.entity = eventEntity
        self.masterData = MasterData.shared.events
        self.setupMappings()
    }
    
    public convenience init(eventMO: EventMO) {
        self.init()
        self.managedObject = eventMO
        self.revert()
    }
    
    public convenience init(eventCode: String, eventName: String, active: Bool = true, startDate: Date? = nil, endDate: Date? = nil, validMinRank: Int = 0, validMaxRank: Int = 999, originatingClubCode: String = "", clubMandatory: Bool = false, nationalAllowed: Bool = true, localAllowed: Bool = true) {
        self.init()
        self.eventName = eventName
        self.eventCode = eventCode
        self.active = active
        self.startDate = startDate
        self.endDate = endDate
        self.validMinRank = Int(validMinRank)
        self.validMaxRank = Int(validMaxRank)
        self.originatingClubCode = originatingClubCode
        self.clubMandatory = clubMandatory
        self.nationalAllowed = nationalAllowed
        self.localAllowed = localAllowed
    }
    
    public override var newManagedObject: NSManagedObject { EventMO() }

    public static func == (lhs: EventViewModel, rhs: EventViewModel) -> Bool {
        return lhs.id == rhs.id
    }
    
    private func setupMappings() {
        $eventCode
            .receive(on: RunLoop.main)
            .map { (eventCode) in
                return (eventCode == "" ? "Event code must not be left blank. Either enter a valid code or delete this event" : (self.eventCodeExists(eventCode) ? "This event code already exists on another event. The event code must be unique" : ""))
            }
        .assign(to: \.saveMessage, on: self)
        .store(in: &cancellableSet)
        
        $eventCode
            .receive(on: RunLoop.main)
            .map { (eventCode) in
                return (eventCode == "" ? "Must be non-blank" : (self.eventCodeExists(eventCode) ? "Must be unique" : ""))
            }
        .assign(to: \.eventCodeMessage, on: self)
        .store(in: &cancellableSet)
              
        $saveMessage
            .receive(on: RunLoop.main)
            .map { (saveMessage) in
                return (saveMessage == "")
            }
        .assign(to: \.canSave, on: self)
        .store(in: &cancellableSet)
    }
    
    public override func beforeInsert() {
        assert(eventCode != "", "Event code must have a non-blank name")
    }
    
    public override var exists: Bool {
        return EventViewModel.event(eventCode: eventCode) != nil
    }
    
    public static func event(eventCode: String?) -> EventViewModel? {
        return EventViewModel.getLookup(eventCode: eventCode)
    }
    
    static public func getLookup(eventCode: String?) -> EventViewModel? {
        return (eventCode == nil ? nil : (MasterData.shared.events.array as! [EventViewModel]).first(where: {$0.eventCode == eventCode}))
    }
    
    private func eventCodeExists(_ eventCode: String) -> Bool {
    return !(MasterData.shared.events.array as! [EventViewModel]).filter({$0.eventCode == eventCode && $0.eventId != self.eventId}).isEmpty
    }
    
    public override var description: String {
        "Event: \(self.eventCode)"
    }
    
    public override var debugDescription: String { self.description }
    
    override public func value(forKey key: String) -> Any {
        switch key {
            case "eventId": return self.eventId as Any
            case "eventCode": return self.eventCode as Any
            case "eventName": return self.eventName as Any
            case "active": return self.active as Any
            case "startDate": return self.startDate as Any
            case "endDate": return self.endDate as Any
            case "validMinRank": return self.validMinRank as Any
            case "validMaxRank": return self.validMaxRank as Any
            case "originatingClubCode": return self.originatingClubCode as Any
            case "clubMandatory": return self.clubMandatory as Any
            case "nationalAllowed": return self.nationalAllowed as Any
            case "localAllowed": return self.localAllowed as Any
            default: fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "eventId": self.eventId = value as! UUID
            case "eventCode": self.eventCode = value as! String
            case "eventName": self.eventName = value as! String
            case "active": self.active = value as! Bool
            case "startDate": self.startDate = value as! Date?
            case "endDate": self.endDate = value as! Date?
            case "validMinRank": self.validMinRank = value as! Int
            case "validMaxRank": self.validMaxRank = value as! Int
            case "originatingClubCode": self.originatingClubCode = value as! String
            case "clubMandatory": self.clubMandatory = value as! Bool
            case "nationalAllowed": self.nationalAllowed = value as! Bool
            case "localAllowed": self.localAllowed = value as! Bool
            default: fatalError("Unknown property '\(key)'")
        }
    }
}
