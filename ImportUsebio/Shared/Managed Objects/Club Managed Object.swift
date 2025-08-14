//
//  Club Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/08/2025.
//

import CoreData

let clubEntity = Entity( "Club",
                        ClubMO.self,
                        Attribute("clubId",                .UUIDAttributeType),
                        Attribute("clubCode",              .stringAttributeType),
                        Attribute("clubName",              .stringAttributeType))

@objc public class ClubMO: NSManagedObject, ManagedObject {
    public static let entity = clubEntity
    
    @NSManaged public var clubId: UUID         ; public var id: UUID { clubId }
    @NSManaged public var clubCode: String
    @NSManaged public var clubName: String
    
    convenience init() {
        self.init(context: CoreData.context)
    }
}
