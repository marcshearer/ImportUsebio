//
//  BridgeWebs Csv Parser.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 06/03/2023.
//
import Foundation

fileprivate enum Phase : String {
    case starting = ""
    case header = "VERSION"
    case headerLine = "LINE"
    case scores = "SCORES"
    case scoresHeader = "SCORESHEADER"
    case score = "SCORE"
    case boardFormat = "BOARDFORMAT"
    case travellersHeader = "TRAVELLERSHEADER"
    case travellers = "TRAVELLERS"
    case board = "BOARD"
    case traveller = "TRAVELLER"
    
    var nextPhases : [Phase] {
        switch self {
        case .starting:
            return [.header]
        case .header:
            return []
        case .headerLine:
            return [.scores]
        case .scores, .scoresHeader:
            return []
        case .score:
            return [.scores, .boardFormat]
        case .boardFormat:
            return [.travellers]
        case .travellers, .travellersHeader:
            return [.board]
        case .board:
            return []
        case .traveller:
            return [.board]
        }
    }
}

fileprivate enum Header : String {
    case date = "DATE"
    case title = "TITLE"
    case director = "DIRECTOR"
    case scorer = "SCORER"
    case pairs = "PAIRS"
    case boards = "BOARDS"
    case winners = "WINNERS"
    case rounds = "ROUNDS"
}

fileprivate enum Score : String {
    case position = "POSITION"
    case pair = "PAIR"
    case name1 = "NAME1"
    case name2 = "NAME2"
    case score = "PERCENT"
}

public class BridgeWebsCsvParser {
    
    private var data: [[String]]
    private var completion: (ScoreData?, [String])->()
    private let scoreData = ScoreData()
    private var errors: [String] = []
    private var warnings: [String] = []
    private var rounds: Int?
    private var scoreColumns: [String] = []
    private var boardColumns: [String] = []
    private var travellerColumns: [String] = []
    private var direction: Direction?
    
    init(fileUrl: URL, data: [[String]], completion: @escaping (ScoreData?, [String])->()) {
        self.scoreData.fileUrl = fileUrl
        self.scoreData.source = .bridgewebs
        self.completion = completion
        self.data = data
        self.parse()
    }
    
    private func parse() {
        var phase = Phase.starting
        var current = -1
        var foundPhase: Phase?
        
        loop: while true {
            current += 1
            
            if current >= data.count {
                if phase != .traveller {
                    report(error: "Data does not contain all the required lines")
                }
                finalUpdates()
                completion(scoreData, [])
                break
            }
            let columns = data[current]
            if columns.isEmpty {
                continue
            }
            
            if columns.first?.left(1) == "#" {
                foundPhase = Phase(rawValue: columns[0].uppercased().right(columns[0].count - 1))
            } else {
                foundPhase = nil
            }
            
            if let foundPhase = foundPhase {
                if !phase.nextPhases.contains(foundPhase) {
                    report(error: "Unexpected tag \(foundPhase.rawValue) found")
                } else {
                    phase = foundPhase
                }
            }
            
            switch phase {
            case .starting:
                break
            case .header:
                scoreData.events.append(Event())
                if columns.count >= 2 {
                    scoreData.events.first!.programName = columns[1]
                }
                phase = .headerLine
            case .headerLine:
                processHeaderElement(columns)
            case .scores:
                phase = .scoresHeader
            case .scoresHeader:
                scoreColumns = columns.map{$0.uppercased()}
                phase = .score
                direction = scoreData.events.first!.participants.isEmpty ? .ns : .ew
            case .score:
                processScoreLine(columns)
            case .boardFormat:
                boardColumns = columns.map{$0.uppercased()}
                boardColumns.removeFirst()
            case .board:
                phase = .traveller
            case .travellers:
                phase = .travellersHeader
            case .travellersHeader:
                travellerColumns = columns.map{$0.uppercased()}
            case .traveller:
                break
            }
        }
    }
    
    private func processHeaderElement(_ columns: [String]) {
        if columns.count >= 2 {
            let first = columns.first!.uppercased()
            if first.left(1) == "#" {
                if let entry = Header(rawValue: first.right(first.count - 1)) {
                    let string = columns[1].ltrim().rtrim()
                    let value = Int(string)
                    switch entry {
                    case .date:
                        scoreData.events.first!.date = Date(from: string, format: "dd/MM/yyyy")
                    case .title:
                        scoreData.events.first!.description = string
                    case .scorer, .director:
                        if string != "" {
                            scoreData.events.first!.contact = string
                        }
                    case .winners:
                        scoreData.events.first!.winnerType = value ?? 1
                    case .boards:
                        scoreData.events.first!.boards = value ?? 0
                    case .rounds:
                        rounds = Int(string)
                    case .pairs:
                        scoreData.events.first!.type = .pairs
                    }
                }
            }
        }
    }
    
    private func processScoreLine(_ columns: [String]) {
        if columns.count != scoreColumns.count {
            report(error: "Column headers do not match score data")
        } else {
            let event = scoreData.events.first!
            let participant = Participant(.pair, from: event)
            let pair = Pair()
            pair.direction = direction
            participant.member = pair
            event.participants.append(participant)
            for (columnNumber, name) in scoreColumns.enumerated() {
                if let key = Score(rawValue: name.uppercased()) {
                    let string = columns[columnNumber].ltrim().rtrim()
                    let value = Int(string)
                    switch key {
                    case .position:
                        participant.place = value
                    case .pair:
                        pair.number = string
                    case .name1, .name2:
                        let player = Player(pair: pair)
                        player.name = string
                        pair.players.append(player)
                    case .score:
                        participant.score = Float(string)
                    }
                }
            }
        }
    }
    
    private func finalUpdates() {
        if let event = scoreData.events.first {
            if let boards = event.boards, let rounds = self.rounds {
                if boards != 0 && rounds != 0 {
                    event.boardsPerRound = boards / rounds
                }
            }
            event.boardScoring = .imps
            event.matchScoring = .imps
        }
    }
    
    private func report(error: String) {
        completion(nil, [error])
    }
}
