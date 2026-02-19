//
//  Blocked Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/02/2026.
//

import CoreData

let blockedEntity = Entity( "Blocked",
                        BlockedMO.self,
                        Attribute("blockedId",                  .UUIDAttributeType),
                        Attribute("nationalId",                 .stringAttributeType),
                        Attribute("reason",                     .stringAttributeType),
                        Attribute("warnOnly",                   .booleanAttributeType))

@objc public class BlockedMO: NSManagedObject, ManagedObject {
    public static let entity = blockedEntity
    
    @NSManaged public var blockedId: UUID     ; public var id: UUID { blockedId }
    @NSManaged public var nationalId$: String
    @NSManaged public var reason: String
    @NSManaged public var warnOnly: Bool
    
    convenience init() {
        self.init(context: CoreData.context)
    }
}
