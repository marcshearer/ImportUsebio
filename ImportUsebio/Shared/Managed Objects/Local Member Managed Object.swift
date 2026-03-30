//
//  Local Member Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 10/02/2026.
//
import CoreData

let localMemberEntity = Entity( "LocalMember",
                           LocalMemberMO.self,
                           Attribute("memberId",              .UUIDAttributeType),
                           Attribute("nationalId",            .stringAttributeType),
                           Attribute("otherNames",            .stringAttributeType),
                           Attribute("lastName",              .stringAttributeType),
                           Attribute("status16",              .integer16AttributeType),
                           Attribute("created",               .dateAttributeType))

@objc public class LocalMemberMO: NSManagedObject, ManagedObject {
    public static let entity = localMemberEntity
    
    @NSManaged public var nationalId: String
    @NSManaged public var otherNames: String
    @NSManaged public var lastName: String
    @NSManaged public var missing: Bool
    @NSManaged public var created: Date
    @NSManaged public var status16: Int16
    
    convenience init() {
        self.init(context: CoreData.context)
    }
}
