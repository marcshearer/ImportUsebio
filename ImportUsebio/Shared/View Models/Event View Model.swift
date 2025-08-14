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
    
    override public var description: String {
        "Event: \(self.eventCode)"
    }
    
    override public var debugDescription: String { self.description }
}
