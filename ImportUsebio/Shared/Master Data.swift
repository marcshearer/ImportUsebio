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
    }
}
