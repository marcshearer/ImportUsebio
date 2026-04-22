//
//  Master Data.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/08/2025.
//

import Foundation
import CoreData

class MasterData: ObservableObject {
    
    public static let shared = MasterData()
    
    @Published private(set) var events = WrappedArray()
    @Published private(set) var clubs = WrappedArray()
    @Published private(set) var ranks = WrappedArray()
    @Published private(set) var members = WrappedArray()
    @Published private(set) var localMembers = WrappedArray()
    @Published private(set) var blocked = WrappedArray()
    @Published private(set) var strataDefs = WrappedArray()
    
    public func load() {
        
        /// **Builds in-memory mirror of event codes, club codes and ranks **
        /// with pointers to managed objects
        /// Note that this infers that there will only ever be 1 instance of the app accessing the database
        
        let createDefaultData = true
        
            // Setup events
        let eventMOs = CoreData.fetch(from: EventMO.entity.name) as! [EventMO]
        
        self.events.array = []
        for eventMO in eventMOs {
            events.array.append(EventViewModel(eventMO: eventMO))
        }
        if events.array.count == 0 && createDefaultData {
            // No events - create defaults
        }
        
        // Setup clubs
        let clubMOs = CoreData.fetch(from: ClubMO.entity.name) as! [ClubMO]
        
        self.clubs.array = []
        for clubMO in clubMOs {
            clubs.array.append(ClubViewModel(clubMO: clubMO))
        }
        if clubs.array.count == 0 && createDefaultData {
            // No clubs - create defaults
        }
        
        // Setup ranks
        let rankMOs = CoreData.fetch(from: RankMO.entity.name) as! [RankMO]
        
        self.ranks.array = []
        for rankMO in rankMOs {
            ranks.array.append(RankViewModel(rankMO: rankMO))
        }
        if ranks.array.count == 0 && createDefaultData {
            // No ranks - create defaults
        }
        
        // Setup blocked
        let blockedMOs = CoreData.fetch(from: BlockedMO.entity.name) as! [BlockedMO]

        self.blocked.array = []
        for blockedMO in blockedMOs {
            blocked.array.append(BlockedViewModel(blockedMO: blockedMO))
        }
        
        if blocked.array.count == 0 && createDefaultData {
        
            BlockedViewModel(nationalId: "superuser", reason: "Special login for Rohallion").insert()
            BlockedViewModel(nationalId: "18000", reason: "Special login for Bill Whyte").insert()
            BlockedViewModel(nationalId: "19000", reason: "Special login for Gordon Milne").insert()
            BlockedViewModel(nationalId: "20000", reason: "Special login for Marc Shearer").insert()
            BlockedViewModel(nationalId: "21000", reason: "Special login for Ann Binder").insert()
            BlockedViewModel(nationalId: "22000", reason: "Special login for Steven Henderson").insert()
            BlockedViewModel(nationalId: "22222", reason: "Special login for Margaret Thompson").insert()
            BlockedViewModel(nationalId: "11111", reason: "Special login").insert()
            BlockedViewModel(nationalId: "17777", reason: "Special login").insert()
        
            BlockedViewModel(nationalId: "13037", reason: "This is almost certainly the wrong Ian Jones and should be 22199").insert()
            
        }
            
        // Setup strata definitions
        let strataDefMOs = CoreData.fetch(from: StrataDefMO.entity.name) as! [StrataDefMO]

        self.strataDefs.array = []
        for strataDefMO in strataDefMOs {
             strataDefs.array.append(StrataDefViewModel(strataDefMO: strataDefMO))
        }
        
        if strataDefs.array.count == 0 && createDefaultData {
        
            let bronze50 = StrataDefViewModel(name: "Bronze only at 50%")
            bronze50.strata[1].code = "Bronze"
            bronze50.strata[1].rank = 160
            bronze50.strata[1].percent = 50
            bronze50.insert()

            let silver50Bronze25 = StrataDefViewModel(name: "Silver at 50% Bronze at 25%")
            silver50Bronze25.strata[1].code = "Silver"
            silver50Bronze25.strata[1].rank = 180
            silver50Bronze25.strata[1].percent = 50
            silver50Bronze25.strata[2].code = "Bronze"
            silver50Bronze25.strata[2].rank = 160
            silver50Bronze25.strata[2].percent = 25
            silver50Bronze25.insert()
        }
        
        loadMembers()
        
        MyApp.masterDataLoaded = true
    }
    
    func loadMembers() {
        
        let createDefaultData = true
        
        // Setup members
        let memberMOs = CoreData.fetch(from: MemberMO.entity.name) as! [MemberMO]

        self.members.array = []
        for memberMO in memberMOs {
            self.members.array.append(MemberViewModel(memberMO: memberMO))
        }
        
        // Setup local members
        let localMemberMOs = CoreData.fetch(from: LocalMemberMO.entity.name) as! [LocalMemberMO]

        self.localMembers.array = []
        for localMemberMO in localMemberMOs {
            if member(nationalId: localMemberMO.nationalId) != nil {
                // Already in members file - remove from local additions
                CoreData.update {
                    CoreData.context.delete(localMemberMO)
                }
            } else {
                self.localMembers.array.append(LocalMemberViewModel(localMemberMO: localMemberMO))
            }
        }

        // If no data then create defaults
        if  members.array.count == 0 && createDefaultData {
            // No members - create defaults
        }
    }
    
    public func member(nationalId: String) -> MemberViewModel? {
        return (members.array as! [MemberViewModel]).first(where: {$0.nationalId == nationalId})
    }
    
    public func member(names: String) -> [MemberViewModel] {
        return (members.array as! [MemberViewModel]).filter({"\($0.otherNames) \($0.lastName)".trim() == names })
    }
}
