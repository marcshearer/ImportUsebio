//
//  Strata Definition Managed Object.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 16/02/2026.
//

import CoreData

let strataDefEntity = Entity( "StrataType",
                           StrataDefMO.self,
                           Attribute("strataDefId",             .UUIDAttributeType),
                           Attribute("name",                    .stringAttributeType),
                           Attribute("customFooter",            .stringAttributeType),
                           Attribute("code1",                   .stringAttributeType,       custom: true),
                           Attribute("code2",                   .stringAttributeType,       custom: true),
                           Attribute("code3",                   .stringAttributeType,       custom: true),
                           Attribute("rank1",                   .integer16AttributeType,    custom: true),
                           Attribute("rank2",                   .integer16AttributeType,    custom: true),
                           Attribute("rank3",                   .integer16AttributeType,    custom: true),
                           Attribute("percent1",                .floatAttributeType,        custom: true),
                           Attribute("percent2",                .floatAttributeType,        custom: true),
                           Attribute("percent3",                .floatAttributeType,        custom: true))

@objc public class StrataDefMO: NSManagedObject, ManagedObject {
    public static let entity = strataDefEntity
    
    @NSManaged public var strataDefId: UUID     ; public var id: UUID { strataDefId }
    @NSManaged public var name: String
    @NSManaged public var customFooter: String
    @NSManaged public var code1: String
    @NSManaged public var code2: String
    @NSManaged public var code3: String
    @NSManaged public var rank1: Int16
    @NSManaged public var rank2: Int16
    @NSManaged public var rank3: Int16
    @NSManaged public var percent1: Float
    @NSManaged public var percent2: Float
    @NSManaged public var percent3: Float
    
    convenience init() {
        self.init(context: CoreData.context)
    }
}
