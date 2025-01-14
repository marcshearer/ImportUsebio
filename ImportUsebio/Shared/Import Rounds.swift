//
//  Import.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 28/02/2023.
//

import SwiftUI

fileprivate enum ImportPhase {
    case eventHeader
    case eventLine
    case roundHeader
    case roundLine
    case finish
}

fileprivate protocol EnumProtocol {
    var string: String {get}
}

fileprivate enum EventColumn: String, EnumProtocol {
    case description = "EVENT DESCRIPTION"
    case code = "EVENT CODE"
    case minRank = "MIN RANK"
    case maxRank = "MAX RANK"
    case localNational = "LOCAL / NATIONAL"
    
    var string: String { self.rawValue.capitalized }
}

fileprivate enum RoundColumn: String, EnumProtocol {
    case name = "ROUND NAME"
    case shortName = "SHORT NAME"
    case toe = "TOE"
    case exclude = "EXCLUDE"
    case basis = "BASIS"
    case localNational = "LOCAL / NATIONAL"
    case maxAward = "MAX AWARD"
    case nsMaxAward = "NS MAX AWARD"
    case ewMaxAward = "EW MAX AWARD"
    case reducedTo = "REDUCED TO %"
    case minEntry = "MIN ENTRY"
    case awardTo = "AWARD TO"
    case perWin = "PER WIN"
    case filterSessionId = "FILTER SESSION ID"
                    // Used to restrict the import of a Usebio raw data file to a specific Session Id
    case aggregateAs = "AGGREGATE AS"
                    // Used to combine columns in the Formatted tab using this column title
    case filterParticipantNumberMin = "FILTER PARTICIPANT NUMBER MIN"
    case filterParticipantNumberMax = "FILTER PARTICIPANT NUMBER MAX"
                    // Used when a single file/session contains multiple separate teams quali events as at Peebles 2024
    case manualPointsColumn = "POINTS COLUMN"
                    // Column with this title in Manual CSV data will get
                    // number of master points from column with this name
                    // rather than the default MANUAL MPS
    case filename = "FILENAME"
                    // Source file name
    case maxTeamMembers = "MAX TEAM MEMBERS"
                    // Will ignore data from team members beyond this number - useful to get rid of subs where scores not materially affected
    case winDrawLevel = "WIN/DRAW LEVEL"
                    // Use Win/Draw data from PARTICIPANT or from MATCH or from MATCH recalculated from BOARD scores - Falls back if data not available at level requested
    case mergeMatches = "MERGE MATCHES"
                    // Only relevant if above is MATCH/BOARD
    case vpType = "VP TYPE"
                    // Only relevant if above is MATCH/BOARD
    case roundContinuousVPDraw = "ROUND VP DRAWS"
                    // If using continuous VPs then setting this to "Yes" treats a score in range 9.50-10.49 as a draw
    
    var string: String { self.rawValue.capitalized }
}

enum LocalNational: String {
    case local = "LOCAL"
    case national = "NATIONAL"
    case byRound = "BY ROUND"
}

class ImportEvent {
    var description: String?
    var code: String?
    var minRank: Int?
    var maxRank: Int?
    var localNational: LocalNational?
}

class ImportRound {
    var exclude = false
    var name: String?
    var shortName: String?
    var toe: Int?
    var manualMPs = false
    var headToHead = false
    var localNational: LocalNational?
    var maxAward: Float?
    var ewMaxAward: Float?
    var reducedTo: Float?
    var minEntry: Int?
    var awardTo: Float?
    var perWin: Float?
    var filename: String?
    var filterSessionId: String?
    var aggregateAs: String?
    var roundContinuousVPDraw: Bool = true
    var winDrawLevel: WinDrawLevel = .board // Start at board and then fail back to level of data available
    var mergeMatches: Bool = false
    var vpType: VpType = .discrete
    var filterParticipantNumberMin: String?
    var filterParticipantNumberMax: String?
    var manualPointsColumn: String?
    var maxTeamMembers: Int?
    var scoreData: ScoreData?
}

class ImportRounds {
    private var completion: (ImportRounds?, String?, String?)->()
    private var data: [[String]]
    public var event: ImportEvent?
    public var rounds: [ImportRound] = []
    private var eventColumnNumbers: [Int:EventColumn] = [:]
    private var eventColumns: [EventColumn:Int] = [:]
    private var roundColumnNumbers: [Int:RoundColumn] = [:]
    private var roundColumns: [RoundColumn:Int] = [:]
    private var failed = false
    
    private init(_ data: [[String]], completion: @escaping (ImportRounds?, String?, String?)->()) {
        self.data = data
        self.completion = completion
    }
    
    public static func process(_ data: [[String]], completion: @escaping (ImportRounds?, String?, String?)->()) {
        let imported = ImportRounds(data, completion: completion)
        imported.parse()
        if !imported.failed {
            completion(imported, nil, nil)
        }
    }
    
    private func parse() {
        var phase = ImportPhase.eventHeader
        var current = 0
        
    loop: while true {
            if current >= data.count {
                if phase != .finish {
                    report(error: "Data does not contain all the required lines")
                } else {
                    finalChecks()
                }
                break
            }
            if data[current].allSatisfy({$0.isEmpty || $0 == "0"}) {
                current += 1
                continue
            }
            switch phase {
            case .eventHeader:
                if data[current].contains(where: { $0.uppercased() == EventColumn.description.rawValue }) {
                    // Found header line
                    if let error = parseEventHeader(columns: data[current]) {
                        report(error: error)
                        break loop
                    } else {
                        phase = .eventLine
                        current += 1
                    }
                } else {
                    report(error: "Unexpected data encountered when seeking event header")
                    break loop
                }
            case .eventLine:
                if let error = parseEventLine(columns: data[current]) {
                    report(error: error)
                    break loop
                } else {
                    phase = .roundHeader
                    current += 1
                }
            case .roundHeader:
                if data[current].contains(where: { $0.uppercased() == RoundColumn.name.rawValue }) {
                    // Found header line
                    if let error = parseRoundHeader(columns: data[current]) {
                        report(error: error)
                        break loop
                    } else {
                        phase = .roundLine
                        current += 1
                    }
                } else {
                    report(error: "Unexpected data encountered when seeking round header")
                    break loop
                }
            case .roundLine, .finish:
                if let error = parseRoundLine(columns: data[current]) {
                    report(error: error)
                    break loop
                } else {
                    phase = .finish
                    current += 1
                }
            }
        }
    }
    
    private func finalChecks() {
        let shortNames = rounds.map{$0.shortName ?? $0.name!}
        let names = rounds.map{$0.name!}
        if Set(names).count != names.count || Set(shortNames).count != shortNames.count {
            report(error: "Round names / Short names must be unique")
        }
    }
    
    private func report(error: String) {
        failed = true
        completion(nil, error, nil)
    }
    
    private func parseEventHeader(columns:[String])->String? {
        var error: String? = nil
        
        for (columnNumber, column) in columns.enumerated() {
            if let eventColumn = EventColumn(rawValue: column.uppercased()) {
                eventColumnNumbers[columnNumber] = eventColumn
                eventColumns[eventColumn] = columnNumber
            }
        }
        
        error = checkExists(in: eventColumns, columns: .description, .code, .localNational)
        
        return error
    }
    
    private func parseEventLine(columns:[String])->String? {
        var error: String? = nil
        
        event = ImportEvent()
        for (columnNumber, columnValue) in columns.enumerated() {
            if !columnValue.isEmpty {
                if let column = eventColumnNumbers[columnNumber] {
                    switch column {
                    case .description:
                        event?.description = columnValue
                    case .code:
                        event?.code = columnValue
                    case .localNational:
                        event?.localNational = LocalNational(rawValue: columnValue.uppercased())
                    case .minRank:
                        event?.minRank = Int(columnValue)
                    case .maxRank:
                        event?.maxRank = Int(columnValue)
                    }
                }
            }
        }
        
        if event?.description == nil {
            error = "\(EventColumn.description.string) must not be blank"
        } else if event?.code == nil {
            error = "\(EventColumn.code.string) must not be blank"
        } else if event?.localNational == nil {
            error = "\(EventColumn.localNational.string) must not be blank"
        }
        
        return error
    }
    
    private func parseRoundHeader(columns:[String])->String? {
        var error: String? = nil
        
        for (columnNumber, column) in columns.enumerated() {
            if let roundColumn = RoundColumn(rawValue: column.uppercased()) {
                roundColumnNumbers[columnNumber] = roundColumn
                roundColumns[roundColumn] = columnNumber
            }
        }
        
        error = checkExists(in: roundColumns, columns: .name, .awardTo, .filename)
        if error == nil {
            error = checkExists(in: roundColumns, columns: .nsMaxAward, .ewMaxAward)
            if error != nil {
                error = checkExists(in: roundColumns, columns: .maxAward)
            }
        }
        
        return error
    }
    
    private func parseRoundLine(columns:[String])->String? {
        var error: String? = nil
        
        let round = ImportRound()
        for (columnNumber, columnValue) in columns.enumerated() {
            if !columnValue.isEmpty {
                if let column = roundColumnNumbers[columnNumber] {
                    switch column {
                    case .exclude:
                        round.exclude = (columnValue.uppercased().left(1) == "Y")
                    case .name:
                        round.name = columnValue
                    case .shortName:
                        round.shortName = columnValue
                    case .toe:
                        round.toe = Int(columnValue)
                    case .basis:
                        switch columnValue.uppercased() {
                        case "MANUAL":
                            round.manualMPs = true
                        case "HEAD-TO-HEAD":
                            round.headToHead = true
                        default:
                            break
                        }
                    case .localNational:
                        round.localNational = LocalNational(rawValue: columnValue.uppercased())
                    case .maxAward, .nsMaxAward:
                        round.maxAward = Float(columnValue)
                    case .ewMaxAward:
                        round.ewMaxAward = Float(columnValue)
                    case .reducedTo:
                        round.reducedTo = (Float(columnValue.replacingOccurrences(of: "%", with: "")) ?? 0) / 100
                    case .minEntry:
                        round.minEntry = Int(columnValue)
                    case .awardTo:
                        if columnValue.contains(where: {$0 == "%"}) {
                            round.awardTo = (Float(columnValue.replacingOccurrences(of: "%", with: "")) ?? 0) / 100
                        } else {
                            round.awardTo = Float(columnValue)
                        }
                    case .perWin:
                        round.perWin = Float(columnValue)
                    case .filterSessionId:
                        round.filterSessionId = columnValue
                    case .aggregateAs:
                        round.aggregateAs = columnValue
                    case .winDrawLevel:
                        round.winDrawLevel = WinDrawLevel(columnValue)
                    case .mergeMatches:
                        round.mergeMatches = (columnValue.uppercased().left(1) == "Y")
                    case .vpType:
                        round.vpType = VpType(columnValue)
                    case .roundContinuousVPDraw:
                        round.roundContinuousVPDraw = (columnValue.uppercased() != "EXACT")
                    case .filterParticipantNumberMin:
                        round.filterParticipantNumberMin = columnValue
                    case .filterParticipantNumberMax:
                        round.filterParticipantNumberMax = columnValue
                    case .manualPointsColumn:
                        round.manualPointsColumn = columnValue
                    case .maxTeamMembers:
                        round.maxTeamMembers = Int(columnValue)
                    case .filename:
                        round.filename = columnValue
                    }
                }
            }
        }
        
        if !round.exclude {
            if round.name == nil {
                error = "\(RoundColumn.name.string) must not be blank"
            } else if round.localNational == nil && event?.localNational == .byRound {
                error = "\(RoundColumn.localNational.string) must not be blank on rounds"
            } else if (!round.headToHead && !round.manualMPs) && round.maxAward ?? 0 <= 0 {
                error = "\(RoundColumn.maxAward.string) must be a positive number"
            } else if !round.manualMPs && !round.headToHead && (round.reducedTo ?? 1 <= 0 || round.reducedTo ?? 1 > 1) {
                error = "\(RoundColumn.awardTo.string) must be greater than 0% less than or equal to 100%"
            } else if !round.manualMPs && !round.headToHead && (round.awardTo ?? 0 <= 0 || round.awardTo ?? 0 > 1) {
                error = "\(RoundColumn.awardTo.string) must be greater than 0% less than or equal to 100%"
            } else if round.filename == nil {
                error = "\(RoundColumn.filename.string) must not be blank"
            } else {
                rounds.append(round)
            }
        }
        
        return error
    }
    
    private func checkExists<T:EnumProtocol>(in columnList: [T:Int], columns: T...) -> String? {
        var result: String? = nil
        for column in columns {
            if columnList[column] == nil {
                result = "Mandatory column '\(column.string)' not found"
                break
            }
        }
        return result
    }
}

