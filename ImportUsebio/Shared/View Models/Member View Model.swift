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
    @Published public var nationalId = ""
    @Published public var otherNames: String = ""
    @Published public var lastName: String = ""
    @Published public var status16: Int16 = PlayerStatus.missing.rawValue
    @Published public var homeClub: String = ""
    @Published public var postCode: String = ""
    @Published public var rankCode: Int = 0
    @Published public var downloaded: Date = Date()
    
    @Published public var nationalIdMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    public var status: PlayerStatus {
        get { PlayerStatus(rawValue: status16)!}
        set { status16 = Int16(newValue.rawValue) }
    }
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.member")!)!
    
    public var names: String {
        get { "\(otherNames) \(lastName)".trim() }
        set {
            var names = newValue.trim().components(separatedBy: " ")
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
        self.entity = memberEntity
        self.masterData = MasterData.shared.members
    }
    
    public convenience init(memberMO: MemberMO) {
        self.init()
        self.managedObject = memberMO
        self.revert()
    }
    
    public convenience init(nationalId: String, otherNames: String, lastName: String, status: PlayerStatus, homeClub: String = "", postCode: String = "", rankCode: Int = 0, downloaded: Date = Date()) {
        self.init()
        self.nationalId = nationalId
        self.otherNames = otherNames
        self.lastName = lastName
        self.homeClub = homeClub
        self.postCode = postCode
        self.rankCode = rankCode
        self.status = status
        self.downloaded = downloaded
    }
    
    public override var newManagedObject: NSManagedObject { MemberMO() }
    
    public static func == (lhs: MemberViewModel, rhs: MemberViewModel) -> Bool {
        return lhs.nationalId == rhs.nationalId
    }
    
    public override func beforeInsert() {
        assert(nationalId != "", "National ID must have a non-blank value")
    }
    
    public override var exists: Bool {
        return MemberViewModel.member(nationalId: nationalId) != nil
    }
    
    public static func member(nationalId: String) -> MemberViewModel? {
        return MasterData.shared.member(nationalId: nationalId)
    }
    
    public static func member(names: String) -> [MemberViewModel] {
        return MasterData.shared.member(names: names)
    }
    
    
    private func nationalIdExists(_ nationalId: String) -> Bool {
        return !(masterData.array as! [MemberViewModel]).filter({$0.nationalId == nationalId && $0.memberId != self.memberId}).isEmpty
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
            case "homeClub": return homeClub as Any
            case "postCode": return postCode as Any
            case "rankCode": return rankCode as Any
            case "status16": return status16 as Any
            case "downloaded": return downloaded as Any
            default : fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "memberId": self.memberId = value as! UUID
            case "nationalId": self.nationalId = value as! String
            case "otherNames": self.otherNames = value as! String
            case "lastName": self.lastName = value as! String
            case "homeClub": self.homeClub = value as! String
            case "postCode": self.postCode = value as! String
            case "rankCode": self.rankCode = value as! Int
            case "status16": self.status16 = value as! Int16
            case "downloaded": self.downloaded = value as! Date
            default : fatalError("Unknown property '\(key)'")
        }
    }
}
