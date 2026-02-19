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
        self.init(rawValue: string.lowercased().replacing(" ", with: "_"))
    }
    
    func combine(scores: Float?...) -> Float? {
        self.combine(scores: scores)
    }
    
    func combine(scores: [Float?]) -> Float? {
        if scores.isEmpty || scores.contains(nil) {
            return nil
        } else {
            switch self {
            case .vps, .imps, .match_points, .cross_imps, .aggregate:
                return scores.reduce(0, {$0 + $1!})
            case .percentage:
                return scores.reduce(0, {$0 + $1!}) / Float(scores.count)
            }
        }
    }
    
    func invert(score: Float?) -> Float? {
        if let score = score {
            switch self {
            case .vps:
                return 20 - score
            case .percentage:
                return 100 - score
            case .imps, .cross_imps, .aggregate:
                return -score
            case .match_points:
                return nil
            }
        } else {
            return nil
        }
    }
    
}

public enum WinDrawLevel: Int, Comparable {
    case participant = 1
    case match = 2
    case board = 3
    
    init(_ string: String) {
        self = switch string.uppercased() {
        case WinDrawLevel.match.string.uppercased():
            .match
        case WinDrawLevel.board.string.uppercased():
            .board
        default:
            .participant
        }
    }
    
    var string: String {
        "\(self)"
    }
    
    var plural: String {
        switch self {
        case .match:
            "matches"
        default:
            "\(self)s"
        }
    }
    
    public static func < (lhs: WinDrawLevel, rhs: WinDrawLevel) -> Bool {
        lhs.rawValue < rhs.rawValue
    }
}

public enum VpType {
    case continuous
    case discrete
    
    init(_ string: String) {
        self = switch string.uppercased() {
        case VpType.continuous.string.uppercased():
            .continuous
        default:
            .discrete
        }
    }
    
    var string: String {
        "\(self)"
    }
}

public class Club {
    var name: String?
    var id: String?
}

public enum Source {
    case usebio
    case bridgewebs
    case manual
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
    case head_to_head
    case invalid
    
    init?(_ string: String) {
        self.init(rawValue: string.lowercased())
    }
    
    var string: String { "\(self)".replacingOccurrences(of: "_", with: " ").capitalized }
    
    var supported: Bool {
        return self == .swiss_pairs || self == .swiss_teams || self == .mp_pairs || self == .pairs || self == .cross_imps || self == .teams || self == .individual || self == .teams_of_four || self == .head_to_head
    }
    
    var participantType: ParticipantType? {
        switch self {
        case .individual:
            return .player
        case .mp_pairs, .butler_pairs, .swiss_pairs, .cross_imps, .aggregate, .swiss_pairs_cross_imps, .swiss_pairs_butler_imps, .pairs:
            return .pair
        case .swiss_teams, .teams, .teams_of_four, .head_to_head:
            return .team
        default:
            return nil
        }
    }
    
    var requiresWinDraw: Bool {
        switch self {
        case .swiss_pairs, .swiss_pairs_cross_imps, .swiss_pairs_butler_imps, .swiss_teams, .head_to_head:
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
    
    var boardsPlayed: Int? { accumulatedBoardsPlayed ?? pair?.boardsPlayed ?? participant?.event?.boards}
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
    var players: [Player] = [] // Data may contain either pairs or players - cope with both
    
    override var type: ParticipantType { .team }
    override var additional: String { "" }
    
    override var playerList: [Player] {
        var list: [Player] = []
        if pairs.isEmpty {
            list = players
        } else {
            let fullList = pairs.flatMap{$0.playerList}
            for player in fullList {
                if let existing = list.first(where: {$0.name?.lowercased() == player.name?.lowercased() && $0.nationalId == player.nationalId}) {
                    if let boardsPlayed = player.pair?.boardsPlayed {
                        existing.accumulatedBoardsPlayed = (existing.accumulatedBoardsPlayed ?? 0) + boardsPlayed
                    }
                    if let winDraw = player.pair?.winDraw {
                        existing.accumulatedWinDraw = (existing.accumulatedWinDraw ?? 0) + winDraw
                    }
                } else {
                    let new = player.copy()
                    new.accumulatedBoardsPlayed = player.pair?.boardsPlayed
                    new.accumulatedWinDraw = player.pair?.winDraw
                    list.append(new)
                }
            }
        }
        return list
    }
}

public class Participant {
    var place: Int?
    var score: Float?
    var winDraw: Float?
    var manualMps: Float?
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
    
    var names: String {
        Utility.toString(member.playerList.map{$0.name ?? "Unknown"})
    }
}

public class Match {
    var round: Int?
    var number: String?
    var opposingNumber: String?
    var score: Float?
    var opposingScore: Float?
    var vp: Float?
    var opposingVP: Float?
    var pairNumbers: Set<String> = []
    var opposingPairNumbers: Set<String> = []
    var boards: [Board] = []
}

public class Board {
    var nsScore: Float?
    var ewScore: Float?
}

public class ScoreData {
    public var source: Source?
    public var fileUrl: URL?
    public var roundName: String?
    public var national: Bool = false
    public var manualMPs: Bool = false
    public var maxAward: Float = 0.0
    public var ewMaxAward: Float?
    public var reducedTo: Float = 0.0
    public var minEntry: Int = 0
    public var awardTo: Float = 0.0
    public var perWin: Float = 0.0
    public var filterSessionId: String = ""
    public var otherSessionData: ScoreData? = nil
    public var aggreateAs: String?
    public var overrideTeamMembers: Int?
    public var manualPointsColumn: String?
    public var version: String?
    public var clubs: [Club] = []
    public var events: [Event] = []
    public var roundContinuousVPDraw = true
    public var winDrawLevel: WinDrawLevel = .match
    public var mergeMatches: Bool = false
    public var vpType: VpType = .discrete
    internal var errors: [String] = []
    internal var warnings: [String] = []
    internal var validateMissingNationalIds = false
}

public class Stratum {
    public var maxRank: Int = 0
    public var maxAward: Float = 0.0
    
    init(maxRank: Int = 0, maxAward: Float = 0) {
        self.maxRank = maxRank
        self.maxAward = maxAward
    }
}
