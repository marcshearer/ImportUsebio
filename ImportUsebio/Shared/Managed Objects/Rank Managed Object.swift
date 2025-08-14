//
//  Rank Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/08/2025.
//

import CoreData

let rankEntity = Entity( "Rank",
                        RankMO.self,
                        Attribute("rankId",                .UUIDAttributeType),
                        Attribute("rankCode16",            .integer16AttributeType,
                                  equivalent: "rankCode", equivalentType: .int),
                        Attribute("rankName",              .stringAttributeType))

@objc public class RankMO: NSManagedObject, ManagedObject {
    public static let entity = rankEntity
    
    @NSManaged public var rankId: UUID         ; public var id: UUID { rankId }
    @NSManaged public var rankCode16: Int16 ; @IntProperty(\RankMO.rankCode16) @objc public var rankCode: Int
    @NSManaged public var rankName: String
    
    convenience init() {
        self.init(context: CoreData.context)
    }
}
