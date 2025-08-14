//
//  Rank View Model.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/08/2025.
//

import Combine
import SwiftUI
import CoreData

public class RankViewModel : ViewModel, ObservableObject {
    
    // Properties in core data model
    @Published private(set) var rankId: UUID = UUID() ; public override var id: UUID { rankId }
    @Published public var rankCode: Int = 0
    @Published public var rankName: String = ""
    
    @Published public var rankCodeMessage: String = ""
    @Published private(set) var saveMessage: String = ""
    @Published private(set) var canSave: Bool = false
    
    public let itemProvider = NSItemProvider(contentsOf: URL(string: "com.sheareronline.importusebio.rank")!)!
    
    // Auto-cleanup
    private var cancellableSet: Set<AnyCancellable> = []
    
    override public init() {
        super.init()
        self.entity = rankEntity
        self.masterData = MasterData.shared.ranks
        self.setupMappings()
    }
    
    public convenience init(rankMO: RankMO) {
        self.init()
        self.managedObject = rankMO
        self.revert()
    }
    
    public convenience init(rankCode: Int, rankName: String) {
        self.init()
        self.rankCode = rankCode
        self.rankName = rankName
    }
    
    public override var newManagedObject: NSManagedObject { RankMO() }

    public static func == (lhs: RankViewModel, rhs: RankViewModel) -> Bool {
        return lhs.id == rhs.id
    }
    
    private func setupMappings() {
        $rankCode
            .receive(on: RunLoop.main)
            .map { (rankCode) in
                return (rankCode == 0 ? "Rank code must not be left zero. Either enter a valid code or delete this rank" : (self.rankCodeExists(rankCode) ? "This rank code already exists on another rank. The rank code must be unique" : ""))
            }
        .assign(to: \.saveMessage, on: self)
        .store(in: &cancellableSet)
        
        $rankCode
            .receive(on: RunLoop.main)
            .map { (rankCode) in
                return (rankCode == 0 ? "Must be non-zero" : (self.rankCodeExists(rankCode) ? "Must be unique" : ""))
            }
        .assign(to: \.rankCodeMessage, on: self)
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
        assert(rankCode != 0, "Rank code must have a non-blank name")
    }
    
    public override var exists: Bool {
        return RankViewModel.rank(rankCode: rankCode) != nil
    }
    
    public static func rank(rankCode: Int?) -> RankViewModel? {
        return RankViewModel.getLookup(rankCode: rankCode)
    }
    
    public static func rank(rankName: String) -> RankViewModel? {
        return (rankName == "" ? nil : (MasterData.shared.ranks.array as! [RankViewModel]).first(where: {$0.rankName.uppercased() == rankName.uppercased()}))
    }
    
    static public func getLookup(rankCode: Int?) -> RankViewModel? {
        return (rankCode == nil ? nil : (MasterData.shared.ranks.array as! [RankViewModel]).first(where: {$0.rankCode == rankCode}))
    }
    
    private func rankCodeExists(_ rankCode: Int) -> Bool {
    return !(MasterData.shared.ranks.array as! [RankViewModel]).filter({$0.rankCode == rankCode && $0.rankId != self.rankId}).isEmpty
    }
    
    override public var description: String {
        "Rank: \(self.rankCode)"
    }
    
    override public var debugDescription: String { self.description }
}
