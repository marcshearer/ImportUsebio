//
//  Event Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/08/2025.
//

import CoreData

let eventEntity = Entity( "Event",
                        EventMO.self,
                        Attribute("eventId",               .UUIDAttributeType),
                        Attribute("eventCode",             .stringAttributeType),
                        Attribute("eventName",             .stringAttributeType),
                        Attribute("active",                .booleanAttributeType),
                        Attribute("startDate",             .dateAttributeType,
                                  isOptional: true),
                        Attribute("endDate",               .dateAttributeType,
                                  isOptional: true),
                        Attribute("validMinRank16",        .integer16AttributeType,
                                  equivalent: "validMinRank", equivalentType: .int),
                        Attribute("validMaxRank16",          .integer16AttributeType,
                                  equivalent: "validMaxRank", equivalentType: .int),
                        Attribute("originatingClubCode",   .stringAttributeType),
                        Attribute("clubMandatory",         .booleanAttributeType),
                        Attribute("nationalAllowed",       .booleanAttributeType),
                        Attribute("localAllowed",          .booleanAttributeType))

@objc public class EventMO: NSManagedObject, ManagedObject {
    public static let entity = eventEntity
    
    @NSManaged public var eventId: UUID         ; public var id: UUID { eventId }
    @NSManaged public var eventCode: String
    @NSManaged public var eventName: String
    @NSManaged public var active: Bool
    @NSManaged public var startDate: Date
    @NSManaged public var endDate: Date
    @NSManaged public var validMinRank16: Int16 ; @IntProperty(\EventMO.validMinRank16) @objc public var validMinRank: Int
    @NSManaged public var validMaxRank16: Int16 ; @IntProperty(\EventMO.validMaxRank16) @objc public var validMaxRank: Int
    @NSManaged public var originatingClubCode: String
    @NSManaged public var clubMandatory: Bool
    @NSManaged public var nationalAllowed: Bool
    @NSManaged public var localAllowed: Bool
    
    
    convenience init() {
        self.init(context: CoreData.context)
    }
}
