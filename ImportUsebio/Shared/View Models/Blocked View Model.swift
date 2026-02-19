//
//  Blocked View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/02/2026.
//

import Combine
import SwiftUI
import CoreData

public class BlockedViewModel : ViewModel, ObservableObject {
    
    
    // Properties in core data model
    @Published private(set) var blockedId: UUID = UUID() ; public override var id: UUID { blockedId }
    @Published public var nationalId: String = ""
    @Published public var reason: String = ""
    @Published public var warnOnly: Bool = false
    
    @Published public var nationalIdMessage: String = ""
    @Published public var reasonMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    static var data = MasterData.shared.blocked.array as! [BlockedViewModel]
    public var intNationalId: Int { Int(nationalId) ?? -1 }
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.blocked")!)!
    
    // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.entity = blockedEntity
        self.masterData = MasterData.shared.blocked
        self.setupMappings()
    }
    
    static func defaultSort(_ first: BlockedViewModel, _ second: BlockedViewModel) -> Bool {
        return ViewModel.sort(first, second, sortKeys: [("intNationalId", .ascending), ("nationalId", .ascending)])
    }
    
    public convenience init(blockedMO: BlockedMO) {
        self.init()
        self.managedObject = blockedMO
        self.revert()
    }
    
    public convenience init(nationalId: String, reason: String, warnOnly: Bool = false) {
        self.init()
        self.nationalId = nationalId
        self.reason = reason
        self.warnOnly = warnOnly
    }
    
    public override var newManagedObject: NSManagedObject { BlockedMO() }

    public static func == (lhs: BlockedViewModel, rhs: BlockedViewModel) -> Bool {
        return lhs.id == rhs.id
    }
    
    private func setupMappings() {
        $nationalId
            .receive(on: RunLoop.main)
            .map { (nationalId) in
                return (nationalId == "" ? "National ID must be non-blank" : (self.blockedExists(nationalId) ? "Already used. The national ID must be unique" : ""))
            }
        .assign(to: \.nationalIdMessage, on: self)
        .store(in: &cancellableSet)
        
        $reason
            .receive(on: RunLoop.main)
            .map { (reason) in
                return (reason == "" ? "Reason must be non-blank" : "")
            }
        .assign(to: \.reasonMessage, on: self)
        .store(in: &cancellableSet)
              
        Publishers.CombineLatest($nationalIdMessage, $reasonMessage)
            .receive(on: RunLoop.main)
            .map { (nationalIdMessage, reasonMessage) in
                return (nationalIdMessage != "" ? nationalIdMessage : (reasonMessage != "" ? reasonMessage : ""))
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
        assert(nationalId != "", "National Id must be non-blank")
    }
    
    public override var exists: Bool {
        return BlockedViewModel.blocked(blockedId: blockedId) != nil
    }
    
    public static func blocked(blockedId: UUID?) -> BlockedViewModel? {
        return (blockedId == nil ? nil : (MasterData.shared.blocked.array as! [BlockedViewModel]).first(where: {$0.blockedId == blockedId}))
    }
    
    public static func blocked(nationalId: String?) -> BlockedViewModel? {
        return (nationalId == nil ? nil : (MasterData.shared.blocked.array as! [BlockedViewModel]).first(where: {$0.nationalId == nationalId}))
    }
    
    static public func getLookup(nationalId: String?) -> BlockedViewModel? {
        return (nationalId == nil ? nil : (MasterData.shared.blocked.array as! [BlockedViewModel]).first(where: {$0.nationalId == nationalId}))
    }
    
    private func blockedExists(_ nationalId: String) -> Bool {
    return !(MasterData.shared.blocked.array as! [BlockedViewModel]).filter({$0.nationalId == nationalId && $0.blockedId != self.blockedId}).isEmpty
    }
    
    public override var description: String {
        "Blocked: \(self.nationalId) \(self.reason)"
    }
    
    public override var debugDescription: String { self.description }
    
    override public func value(forKey key: String) -> Any? {
        switch key {
            case "blockedId": return self.blockedId as Any
            case "nationalId": return self.nationalId as Any
            case "reason": return self.reason as Any
            case "warnOnly": return self.warnOnly as Any
            case "intNationalId": return self.intNationalId as Any
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "blockedId": self.blockedId = value as! UUID
            case "nationalId": self.nationalId = value as! String
            case "reason": self.reason = value as! String
            case "warnOnly": self.warnOnly = value as! Bool
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
}
