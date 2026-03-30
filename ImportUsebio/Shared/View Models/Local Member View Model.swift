//
//  Local Member View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 10/02/2026.
//

import Combine
import SwiftUI
import CoreData

public class LocalMemberViewModel : ViewModel, ObservableObject {
    
    // Properties in core data model
    @Published private(set) var memberId: UUID = UUID() ; public override var id: UUID { memberId }
    @Published private(set) var nationalId = ""
    @Published public var otherNames: String = ""
    @Published public var lastName: String = ""
    @Published public var status16: Int16 = PlayerStatus.unknown.rawValue
    @Published public var created: Date = Date()
    
    @Published public var nationalIdMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    public var status: PlayerStatus {
        get { PlayerStatus(rawValue: status16)!}
        set { status16 = Int16(newValue.rawValue) }
    }
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.localmember")!)!
    
    public var names: String {
        get { "\(otherNames) \(lastName)" }
        set {
            var names = newValue.components(separatedBy: " ")
            if let lastName = names.last {
                self.lastName = lastName
                names.removeLast()
                if names.isEmpty {
                    self.otherNames = ""
                } else {
                    self.otherNames = names.joined(separator: " ")
                }
            } else {
                lastName = ""
                otherNames = ""
            }
        }
    }
    
    // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.entity = localMemberEntity
        self.masterData = MasterData.shared.localMembers
        self.setupMappings()
    }
    
    public convenience init(localMemberMO: LocalMemberMO) {
        self.init()
        self.managedObject = localMemberMO
        self.revert()
    }
    
    public convenience init(nationalId: String, otherNames: String, lastName: String, status: PlayerStatus) {
        self.init()
        self.nationalId = nationalId
        self.otherNames = otherNames
        self.lastName = lastName
        self.status = status
        self.created = Date()
    }
    
    public override var newManagedObject: NSManagedObject { LocalMemberMO() }
    
    public static func == (lhs: LocalMemberViewModel, rhs: LocalMemberViewModel) -> Bool {
        return lhs.nationalId == rhs.nationalId
    }
    
    public var memberViewModel: MemberViewModel {
        MemberViewModel(nationalId: nationalId, otherNames: otherNames, lastName: lastName, status: status)
    }
    
    private func setupMappings() {
        $nationalId
            .receive(on: RunLoop.main)
            .map { (nationalId) in
                return (nationalId == "" ? "National ID must not be blank. Either enter a non-blank value or delete this member" : (self.nationalIdExists(nationalId) ? "This national ID already exists on another member. The national ID must be unique" : ""))
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
        assert(nationalId != "", "National ID must have a non-blank value")
    }
    
    public override var exists: Bool {
        return LocalMemberViewModel.member(nationalId: nationalId) != nil
    }
    
    public static func member(nationalId: String) -> LocalMemberViewModel? {
        return LocalMemberViewModel.getLookup(nationalId: nationalId)
    }
    
    public static func member(names: String) -> LocalMemberViewModel? {
        return (MasterData.shared.localMembers.array as! [LocalMemberViewModel]).filter({"\($0.otherNames) \($0.lastName)" == names }).first
    }
    
    static public func getLookup(nationalId: String) -> LocalMemberViewModel? {
        return (MasterData.shared.localMembers.array as! [LocalMemberViewModel]).first(where: {$0.nationalId == nationalId})
    }
    
    private func nationalIdExists(_ nationalId: String) -> Bool {
        return !(masterData.array as! [LocalMemberViewModel]).filter({$0.nationalId == nationalId && $0.memberId != self.memberId}).isEmpty
    }
    
    override public var description: String {
        "Member: \(self.nationalId) - \(self.otherNames) \(self.lastName)"
    }
    
    override public var debugDescription: String { self.description }
    
    override public func value(forKey key: String) -> Any? {
        switch key {
            case "memberId": return memberId as Any
            case "nationalId": return nationalId as Any
            case "otherNames": return otherNames as Any
            case "lastName": return lastName as Any
            case "status16": return status16 as Any
            case "created": return created as Any
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "memberId": self.memberId = value as! UUID
            case "nationalId": self.nationalId = value as! String
            case "otherNames": self.otherNames = value as! String
            case "lastName": self.lastName = value as! String
            case "status16": self.status16 = value as! Int16
            case "created": self.created = value as! Date
            default : fatalError("Unknown property '\(key)'")
        }
    }
}
