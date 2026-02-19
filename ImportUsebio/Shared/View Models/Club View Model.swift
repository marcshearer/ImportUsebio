//
//  Club View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/08/2025.
//

import Combine
import SwiftUI
import CoreData

public class ClubViewModel : ViewModel, ObservableObject {
    
    // Properties in core data model
    @Published private(set) var clubId: UUID = UUID() ; public override var id: UUID { clubId }
    @Published public var clubCode: String = ""
    @Published public var clubName: String = ""
    
    @Published public var clubCodeMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.club")!)!
    
    // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.entity = clubEntity
        self.masterData = MasterData.shared.clubs
        self.setupMappings()
    }
    
    public convenience init(clubMO: ClubMO) {
        self.init()
        self.managedObject = clubMO
        self.revert()
    }
    
    public convenience init(clubCode: String, clubName: String) {
        self.init()
        self.clubName = clubName
        self.clubCode = clubCode
    }
    
    public override var newManagedObject: NSManagedObject { ClubMO() }

    public static func == (lhs: ClubViewModel, rhs: ClubViewModel) -> Bool {
        return lhs.id == rhs.id
    }
    
    private func setupMappings() {
        $clubCode
            .receive(on: RunLoop.main)
            .map { (clubCode) in
                return (clubCode == "" ? "Club code must not be left blank. Either enter a valid code or delete this club" : (self.clubCodeExists(clubCode) ? "This club code already exists on another club. The club code must be unique" : ""))
            }
        .assign(to: \.saveMessage, on: self)
        .store(in: &cancellableSet)
        
        $clubCode
            .receive(on: RunLoop.main)
            .map { (clubCode) in
                return (clubCode == "" ? "Must be non-blank" : (self.clubCodeExists(clubCode) ? "Must be unique" : ""))
            }
        .assign(to: \.clubCodeMessage, on: self)
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
        assert(clubCode != "", "Club code must have a non-blank name")
    }
    
    public override var exists: Bool {
        return ClubViewModel.club(clubCode: clubCode) != nil
    }
    
    public static func club(clubCode: String?) -> ClubViewModel? {
        return ClubViewModel.getLookup(clubCode: clubCode)
    }
    
    public static func club(clubName: String) -> ClubViewModel? {
        return (clubName == "" ? nil : (MasterData.shared.clubs.array as! [ClubViewModel]).first(where: {$0.clubName.uppercased() == clubName.uppercased()}))
    }
    
    static public func getLookup(clubCode: String?) -> ClubViewModel? {
        return (clubCode == nil ? nil : (MasterData.shared.clubs.array as! [ClubViewModel]).first(where: {$0.clubCode == clubCode}))
    }
    
    private func clubCodeExists(_ clubCode: String) -> Bool {
    return !(MasterData.shared.clubs.array as! [ClubViewModel]).filter({$0.clubCode == clubCode && $0.clubId != self.clubId}).isEmpty
    }
    
    public override var description: String {
        "Club: \(self.clubCode)"
    }
    
    public override var debugDescription: String { self.description }
    
    override public func value(forKey key: String) -> Any {
        switch key {
            case "clubId": return clubId as Any
            case "clubCode": return clubCode as Any
            case "clubName": return clubName as Any
        default: fatalError("Unknown property '\(key)'")
        }
    }
    
    override public func setValue(_ value: Any?, forKey key: String) {
        switch key {
            case "clubId": self.clubId = value as! UUID
            case "clubCode": self.clubCode = value as! String
            case "clubName": self.clubName = value as! String
            default: fatalError("Unknown property '\(key)'")
        }
    }
}
