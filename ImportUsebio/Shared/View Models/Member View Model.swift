//
//  Member View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 10/02/2026.
//

import Combine
import SwiftUI
import CoreData

public class MemberViewModel : ViewModel, ObservableObject {
    
        // Properties in core data model
    @Published private(set) var memberId: UUID = UUID() ; public override var id: UUID { memberId }
    @Published private(set) var nationalId = ""
    @Published public var firstName: String = ""
    @Published public var lastName: String = ""
    @Published public var homeClub: String = ""
    @Published public var postCode: String = ""
    @Published public var rankCode: Int = 0
    @Published public var downloaded: Date = Date()
    
    @Published public var nationalIdMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.member")!)!
    
        // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.entity = memberEntity
        self.masterData = MasterData.shared.members
        self.setupMappings()
    }
    
    public convenience init(memberMO: MemberMO) {
        self.init()
        self.managedObject = memberMO
        self.revert()
    }
    
    public convenience init(nationalId: String, firstName: String, lastName: String, homeClub: String, postCode: String, rankCode: Int, downloaded: Date) {
        self.init()
        self.nationalId = nationalId
        self.firstName = firstName
        self.lastName = lastName
        self.homeClub = homeClub
        self.postCode = postCode
        self.rankCode = rankCode
        self.downloaded = downloaded
    }
    
    public override var newManagedObject: NSManagedObject { MemberMO() }
    
    public static func == (lhs: MemberViewModel, rhs: MemberViewModel) -> Bool {
        return lhs.nationalId == rhs.nationalId
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
        return MemberViewModel.member(nationalId: nationalId) != nil
    }
    
    public static func member(nationalId: String) -> MemberViewModel? {
        return MemberViewModel.getLookup(nationalId: nationalId)
    }
    
    static public func getLookup(nationalId: String) -> MemberViewModel? {
        return (MasterData.shared.members.array as! [MemberViewModel]).first(where: {$0.nationalId == nationalId})
    }
    
    private func nationalIdExists(_ nationalId: String) -> Bool {
        return !(masterData.array as! [MemberViewModel]).filter({$0.nationalId == nationalId && $0.memberId != self.memberId}).isEmpty
    }
    
    override public var description: String {
        "Member: \(self.nationalId) - \(self.firstName) \(self.lastName)"
    }
    
    override public var debugDescription: String { self.description }
    
    override public func value(forKey key: String) -> Any? {
        switch key {
            case "memberId": return memberId as Any
            case "nationalId": return nationalId as Any
            case "firstName": return firstName as Any
            case "lastName": return lastName as Any
            case "homeClub": return homeClub as Any
            case "postCode": return postCode as Any
            case "rankCode": return rankCode as Any
            case "downloaded": return downloaded as Any
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "memberId": self.memberId = value as! UUID
            case "nationalId": self.nationalId = value as! String
            case "firstName": self.firstName = value as! String
            case "lastName": self.lastName = value as! String
            case "homeClub": self.homeClub = value as! String
            case "postCode": self.postCode = value as! String
            case "rankCode": self.rankCode = value as! Int
            case "downloaded": self.downloaded = value as! Date
            default : fatalError("Unknown property '\(key)'")
        }
    }
}
