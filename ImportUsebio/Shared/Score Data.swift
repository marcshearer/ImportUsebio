//
//  Score Data.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 09/02/2023.
//

import Foundation

public enum ScoringMethod: String {
    case vps
    case imps
    case match_points
    case percentage
    case cross_imps
    case aggregate
    
    init?(_ string: String) {
        self.init(rawValue: string.lowercased())
    }
}

public class Club {
    var name: String?
    var id: String?
}

public enum EventType: String {
    case ko
    case ladder
    case mp_pairs
    case butler_pairs
    case individual
    case swiss_pairs
    case swiss_teams
    case teams_of_four
    case cross_imps
    case aggregate
    case swiss_pairs_cross_imps
    case swiss_pairs_butler_imps
    case pairs
    case teams
    case invalid
    
    init?(_ string: String) {
        self.init(rawValue: string.lowercased())
    }
    
    var string: String { "\(self)".replacingOccurrences(of: "_", with: " ").capitalized }
    
    var supported: Bool {
        return self == .swiss_pairs || self == .swiss_teams || self == .mp_pairs || self == .pairs || self == .cross_imps || self == .teams || self == .individual
    }
    
    var participantType: ParticipantType? {
        switch self {
        case .individual:
            return .player
        case .mp_pairs, .butler_pairs, .swiss_pairs, .cross_imps, .aggregate, .swiss_pairs_cross_imps, .swiss_pairs_butler_imps, .pairs:
            return .pair
        case .swiss_teams, .teams, .teams_of_four:
            return .team
        default:
            return nil
        }
    }
    
    var requiresWinDraw: Bool {
        switch self {
        case .swiss_pairs, .swiss_pairs_cross_imps, .swiss_pairs_butler_imps, .swiss_teams:
            return true
        default:
            return false
        }
    }
}

public class Event {
    var type: EventType?
    var programName: String?
    var programVersion: String?
    var description: String?
    var date: Date?
    var eventCode: String?
    var boardScoring: ScoringMethod?
    var matchScoring: ScoringMethod?
    var boards: Int?
    var boardsPerRound: Int?
    var contact: String?
    var participants: [Participant] = []
    var matches: [Match] = []
    var winnerType: Int = 1
    var sectionCount: Int = 1
    var sessionCount: Int = 1
}

public enum ParticipantType: Int {
    case player = 1
    case pair = 2
    case team = 4
    
    public var players: Int { self.rawValue }
    
    public var string: String { "\(self)".capitalized }
}

public enum Direction: String {
    case ns
    case ew
    
    var sort: Int {
        switch self {
        case .ns:
            return 1
        default:
            return 2
        }
    }
    
    init?(_ string: String) {
        self.init(rawValue: string.lowercased())
    }
    
    var string: String { "\(self)".uppercased()}
}

public enum Seat {
    case north
    case south
    case east
    case west
    case invalid
    
    init(_ string: String) {
        switch string.uppercased() {
        case "N": self = .north
        case "S": self = .south
        case "E": self = .east
        case "W": self = .west
        default:
            self = .invalid
        }
    }
    
    var string: String { "\(self)".uppercased()}
}

public class Member {
    var name: String?
    var number: String?
    
    var type: ParticipantType { fatalError("Must be overridden") }
    var playerList: [Player] { fatalError("Must be overridden") }
    var additional: String { fatalError("Must be overridden") }
    var description: String {
        if let name = name {
            return "\(type.string) \(name)"
        } else if let id = number {
            return "\(type.string) \(id) \(additional)"
        } else {
            return "Unknown \(type.string)"
        }
    }
    
    weak var participant: Participant?
}

public class Player : Member {
    var nationalId: String?
    var seat: Seat?
    weak var pair: Pair? = nil

    override var type: ParticipantType { .player }
    override var additional: String { "" }
    override var playerList: [Player] { [self] }
    
    var accumulatedBoardsPlayed: Int?
    var accumulatedWinDraw: Float?
    
    var boardsPlayed: Int { accumulatedBoardsPlayed ?? pair?.boardsPlayed ?? participant?.event?.boards ?? 0}
    var winDraw: Float { accumulatedWinDraw ?? pair?.winDraw ?? participant?.winDraw ?? 0}
    
    init(pair: Pair? = nil) {
        super.init()
        self.participant = nil
        self.pair = pair
    }
    
    public func copy() -> Player {
        let player = Player()
        player.name = name
        player.number = number
        player.nationalId = nationalId
        player.seat = seat
        player.participant = participant
        return player
    }
}

public class Pair : Member {
    var imps: Float?
    var players: [Player] = []
    var boardsPlayed: Int?
    var direction: Direction?
    var winDraw: Float?
    
    override var type: ParticipantType { .pair }
    override var additional: String { "" }  // Could be set to Direction but need to
                                            // work out what happens in matches with pair IDs
                                            // being unique
    override var playerList: [Player] { return players }
}

public class Team : Member {
    var pairs: [Pair] = []
    
    override var type: ParticipantType { .team }
    override var additional: String { "" }
    
    override var playerList: [Player] {
        let fullList = pairs.flatMap{$0.playerList}
        var list: [Player] = []
        for player in fullList {
            if let existing = list.first(where: {$0.name?.lowercased() == player.name?.lowercased() && $0.nationalId == player.nationalId}) {
                existing.accumulatedBoardsPlayed = existing.accumulatedBoardsPlayed! + (player.pair?.boardsPlayed ?? 0)
                existing.accumulatedWinDraw = existing.accumulatedWinDraw! + (player.pair?.winDraw ?? 0)
            } else {
                let new = player.copy()
                new.accumulatedBoardsPlayed = player.pair?.boardsPlayed ?? 0
                new.accumulatedWinDraw = player.pair?.winDraw ?? 0
                list.append(new)
            }
        }
        return list
    }
}

public class Participant {
    var place: Int?
    var score: Float?
    var winDraw: Float?
    var member: Member
    weak var event: Event?
    
    var type: ParticipantType { member.type }
    var description: String { member.description }
    var number: String { member.number ?? "" }
    
    init(_ type: ParticipantType, from event: Event) {
        self.event = event
        
        var member: Member?
        switch type {
        case .player:
            member = Player()
        case .pair:
            member = Pair()
        case .team:
            member = Team()
        }
        self.member = member!
        self.member.participant = self
    }
}

public class Match {
    var round: Int?
    var number: String?
    var opposingNumber: String?
    var score: Float?
    var opposingScore: Float?
    var pairNumbers: Set<String> = []
    var opposingPairNumbers: Set<String> = []
}

public class ScoreData {
    public var fileUrl: URL?
    public var roundName: String?
    public var national: Bool = false
    public var maxAward: Float = 0.0
    public var minEntry: Int = 0
    public var awardTo: Float = 0.0
    public var perWin: Float = 0.0
    public var version: String?
    public var clubs: [Club] = []
    public var events: [Event] = []
    internal var errors: [String] = []
    internal var warnings: [String] = []
    internal var validateMissingNationalIds = false
}
