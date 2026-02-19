//
//  Strata Definition View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 16/02/2026.
//

import Combine
import SwiftUI
import CoreData

public class StrataElement {
    public var code: String
    public var rank: Int
    public var percent: Float
    
    init(code: String = "", rank: Int = 0, percent: Float = 0) {
        self.code = code
        self.rank = rank
        self.percent = percent
    }
    
    public func copy(from: StrataElement) {
        self.code = from.code
        self.rank = from.rank
        self.percent = from.percent
    }
}

public class StrataDefViewModel : ViewModel, ObservableObject {
    
    // Properties in core data model
    @Published private(set) var strataDefId: UUID = UUID() ; public override var id: UUID { strataDefId }
    @Published public var name: String = ""
    @Published public var strata: [StrataElement] = []
    
    @Published public var nameMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    let strataElements = 3
    
    static var data = MasterData.shared.strataDefs.array as! [StrataDefViewModel]
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.strata")!)!
    
    // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.strata = []
        self.strata.append(StrataElement(code: "", rank: 999, percent: 100))
        for _ in 1..<strataElements {
            self.strata.append(StrataElement())
        }
        self.entity = strataDefEntity
        self.masterData = MasterData.shared.strataDefs
        self.setupMappings()
    }
    
    static func defaultSort(_ first: StrataDefViewModel, _ second: StrataDefViewModel) -> Bool {
        return ViewModel.sort(first, second, sortKeys: [("name", .ascending)])
    }
    
    public convenience init(strataDefMO: StrataDefMO) {
        self.init()
        self.managedObject = strataDefMO
        self.revert()
    }
    
    public convenience init(name: String) {
        self.init()
        self.name = name
    }
    
    public override var newManagedObject: NSManagedObject { StrataDefMO() }

    public static func == (lhs: StrataDefViewModel, rhs: StrataDefViewModel) -> Bool {
        return lhs.id == rhs.id
    }
    
    private func setupMappings() {
        $name
            .receive(on: RunLoop.main)
            .map { (name) in
                return (name == "" ? "Name must be non-blank" : (self.strataDefExists(name) ? "Already used. The name must be unique" : ""))
            }
        .assign(to: \.nameMessage, on: self)
        .store(in: &cancellableSet)
              
        $nameMessage
            .receive(on: RunLoop.main)
            .map { (nameMessage) in
                return nameMessage
            }
        .assign(to: \.saveMessage, on: self)
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
        assert(name != "", "Name must be non-blank")
    }
    
    public override var exists: Bool {
        return StrataDefViewModel.strataDef(strataDefId: strataDefId) != nil
    }
    
    public static func strataDef(strataDefId: UUID?) -> StrataDefViewModel? {
        return (strataDefId == nil ? nil : (MasterData.shared.strataDefs.array as! [StrataDefViewModel]).first(where: {$0.strataDefId == strataDefId}))
    }
    
    public static func strata(name: String?) -> StrataDefViewModel? {
        return (name == nil ? nil : (MasterData.shared.strataDefs.array as! [StrataDefViewModel]).first(where: {$0.name == name}))
    }
    
    static public func getLookup(name: String?) -> StrataDefViewModel? {
        return (name == nil ? nil : (MasterData.shared.strataDefs.array as! [StrataDefViewModel]).first(where: {$0.name == name}))
    }
    
    private func strataDefExists(_ name: String) -> Bool {
    return !(MasterData.shared.strataDefs.array as! [StrataDefViewModel]).filter({$0.name == name && $0.strataDefId != self.strataDefId}).isEmpty
    }
    
    public override var description: String {
        "Strata Definition: \(self.name)"
    }
    
    public override var debugDescription: String { self.description }
    
    override public func value(forKey key: String) -> Any? {
        switch key {
            case "strataDefId": return self.strataDefId as Any
            case "name": return self.name as Any
            case "code1": return (strata[0].code) as Any?
            case "code2": return (strata[1].code) as Any?
            case "code3": return (strata[2].code) as Any?
            case "rank1": return (strata[0].rank) as Any?
            case "rank2": return (strata[1].rank) as Any?
            case "rank3": return (strata[2].rank) as Any?
            case "percent1": return (strata[0].percent) as Any?
            case "percent2": return (strata[1].percent) as Any?
            case "percent3": return (strata[2].percent) as Any?
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "strataDefId": self.strataDefId = value as! UUID
            case "name": self.name = value as! String
            case "code1": self.strata[0].code = value as! String
            case "code2": self.strata[1].code = value as! String
            case "code3": self.strata[2].code = value as! String
            case "rank1": self.strata[0].rank = value as! Int
            case "rank2": self.strata[1].rank = value as! Int
            case "rank3": self.strata[2].rank = value as! Int
            case "percent1": self.strata[0].percent = value as! Float
            case "percent2": self.strata[1].percent = value as! Float
            case "percent3": self.strata[2].percent = value as! Float
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
    public func removeStratum(at index: Int) {
        for copyIndex in index..<(strata.count - 1) {
            strata[copyIndex].copy(from: strata[copyIndex + 1])
        }
        strata.last!.copy(from: StrataElement())
    }
}
