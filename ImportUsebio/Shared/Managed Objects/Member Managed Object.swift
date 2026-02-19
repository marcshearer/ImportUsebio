//
//  Member Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 10/02/2026.
//
import CoreData

let memberEntity = Entity( "Member",
                           MemberMO.self,
                           Attribute("memberId",              .UUIDAttributeType),
                           Attribute("nationalId",            .stringAttributeType),
                           Attribute("firstName",             .stringAttributeType),
                           Attribute("lastName",              .stringAttributeType),
                           Attribute("homeClub",              .stringAttributeType),
                           Attribute("postCode",              .stringAttributeType),
                           Attribute("rankCode16",            .integer16AttributeType,
                                            equivalent: "rankCode", equivalentType: .int),
                           Attribute("downloaded",            .dateAttributeType))

@objc public class MemberMO: NSManagedObject, ManagedObject {
    public static let entity = memberEntity
    
    @NSManaged public var nationalId: String
    @NSManaged public var firstName: String
    @NSManaged public var lastName: String
    @NSManaged public var homeClub: String
    @NSManaged public var postCode: String
    @NSManaged public var rankCode16: Int16 ; @IntProperty(\MemberMO.rankCode16) @objc public var rankCode: Int
    @NSManaged public var downloaded: Date

    convenience init() {
        self.init(context: CoreData.context)
    }
}
