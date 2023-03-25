//
//  Manual Csv Parser.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 06/03/2023.
//

import Foundation

fileprivate enum Phase : String {
    case starting = ""
    case parameters = "PARAMETERS"
    case parameterValues = "PARAMETERVALUES"
    case round = "ROUND"
    case roundValues = "ROUNDVALUES"
    case participants = "PARTICIPANTS"
    case participantValues = "PARTICIPANTVALUES"
    
    var nextPhases : [Phase] {
        switch self {
        case .starting:
            return [.parameters]
        case .parameters:
            return [.parameterValues]
        case .parameterValues:
            return [.round]
        case .round:
            return [.roundValues]
        case .roundValues:
            return [.participants]
        case .participants:
            return [.participantValues]
        case .participantValues:
            return []
        }
    }
}

fileprivate enum Parameters: String {
    case roundType = "ROUND TYPE"
    case maxPlayers = "MAX PLAYERS"
    case winners = "WINNERS"
    case winBonus = "WIN BONUS"
    case version = "VERSION"
}

fileprivate enum RoundHeader : String {
    case title = "ROUND NAME"
    case date = "ROUND DATE"
    case contact = "CONTACT"
    case boards = "BOARDS"
    case rounds = "ROUNDS"
    case boardScoring = "BOARD SCORING"
    case matchScoring = "MATCH SCORING"
}

fileprivate enum ParticipantHeader: String {
    case place = "PLACE"
    case direction = "DIRECTION"
    case number = "NUMBER"
    case score = "SCORE"
    case boardsPlayed = "BOARDS PLAYED"
    case winsDraws = "WINS/DRAWS"
    case name = "NAME"
    case nationalId = "NATIONAL ID"
}

public class ManualCsvParser {
    
    private var data: [[String]]
    private var completion: (ScoreData?, String?)->()
    private let scoreData = ScoreData()
    private var errors: [String] = []
    private var warnings: [String] = []
    private var rounds: Int?
    private var requiresWinDraw = false
    private var parameterColumns: [String] = []
    private var roundColumns: [String] = []
    private var participantColumns: [String] = []
    
    init(fileUrl: URL, data: [[String]], completion: @escaping (ScoreData?, String?)->()) {
        self.scoreData.fileUrl = fileUrl
        self.scoreData.source = .manual
        self.completion = completion
        self.data = data
        self.parse()
    }
    
    func parse() {
        var phase = Phase.starting
        var current = -1
        var foundPhase: Phase?
        
        loop: while true {
            current += 1
            
            if current >= data.count {
                if phase != .participantValues {
                    report(error: "Data does not contain all the required lines")
                }
                finalUpdates()
                completion(scoreData, nil)
                break
            }
            var columns = data[current]
            if columns.isEmpty || columns.first(where: {!$0.isEmpty}) == nil {
                continue
            }
            
            if columns.first != "" {
                foundPhase = Phase(rawValue: columns.first!.uppercased())
            } else {
                foundPhase = nil
            }
            
            if let foundPhase = foundPhase {
                if !phase.nextPhases.contains(foundPhase) {
                    report(error: "Unexpected section \(foundPhase.rawValue) found")
                } else {
                    phase = foundPhase
                }
            }
            
            switch phase {
            case .starting:
                break
            case .parameters:
                scoreData.events.append(Event())
                scoreData.events.first!.programName = "Manual"
                parameterColumns = columns
                parameterColumns.removeFirst()
                phase = .parameterValues
            case .parameterValues:
                columns.removeFirst()
                processParameterElement(columns)
            case .round:
                roundColumns = columns
                roundColumns.removeFirst()
                phase = .roundValues
            case .roundValues:
                columns.removeFirst()
                processRoundElement(columns)
            case .participants:
                participantColumns = columns
                participantColumns.removeFirst()
                phase = .participantValues
            case .participantValues:
                columns.removeFirst()
                processParticipantElement(columns)
            }
        }
    }
    
    private func processParameterElement(_ columns: [String]) {
        if columns.count != parameterColumns.count {
            report(error: "Parameter column headers do not match data")
        } else {
            for (columnNumber, column) in columns.enumerated() {
                let header = parameterColumns[columnNumber].uppercased()
                if let entry = Parameters(rawValue: header) {
                    let string = column.ltrim().rtrim()
                    let value = Int(string)
                    switch entry {
                    case .version:
                        scoreData.version = string
                    case .roundType:
                        scoreData.events.first!.type = EventType(string.uppercased())
                    case .winners:
                        scoreData.events.first!.winnerType = value ?? 1
                    case .winBonus:
                        requiresWinDraw = true
                    case .maxPlayers:
                        break
                    }
                }
            }
        }
    }
    
    private func processRoundElement(_ columns: [String]) {
        if columns.count != roundColumns.count {
            report(error: "Round column headers do not match data")
        } else {
            for (columnNumber, column) in columns.enumerated() {
                let header = roundColumns[columnNumber].uppercased()
                if let entry = RoundHeader(rawValue: header) {
                    let string = column.ltrim().rtrim()
                    let value = Int(string)
                    switch entry {
                    case .title:
                        scoreData.events.first!.description = string
                    case .date:
                        scoreData.events.first!.date = Utility.dateFromString(string, format: "dd/MM/yyyy")
                    case .contact:
                        scoreData.events.first!.contact = string
                    case .boards:
                        scoreData.events.first!.boards = value
                    case .rounds:
                        rounds = value
                    case .boardScoring:
                        scoreData.events.first!.boardScoring = ScoringMethod(string)
                    case .matchScoring:
                        scoreData.events.first!.boardScoring = ScoringMethod(string)
                    }
                }
            }
        }
    }
    
    private func processParticipantElement(_ columns: [String]) {
        if columns.count != participantColumns.count {
            report(error: "Round column headers do not match data")
        } else {
            let event = scoreData.events.first!
            let type = event.type!.participantType!
            let participant = Participant(type, from: event)
            switch type {
            case .pair:
                participant.member = Pair()
            case .team:
                participant.member = Team()
            case .player:
                participant.member = Player()
            }
            event.participants.append(participant)
            for (columnNumber, column) in columns.enumerated() {
                let header = participantColumns[columnNumber].uppercased()
                let string = column.ltrim().rtrim()
                let value = Int(string)
                if let entry = ParticipantHeader(rawValue: header) {
                    switch entry {
                    case .place:
                        participant.place = value
                    case .direction:
                        if let pair = participant.member as? Pair {
                            pair.direction = Direction(string)
                        }
                    case .number:
                        participant.member.number = string
                    case .score:
                        participant.score = Float(value ?? 0)
                    case .boardsPlayed:
                        if let pair = participant.member as? Pair {
                            pair.boardsPlayed = value
                        }
                    case .winsDraws:
                        participant.winDraw = Float(value ?? 0)
                    case .name, .nationalId:
                        // Only with indexes
                        break
                    }
                } else {
                    // Check for index
                    let components = header.split(at: "(")
                    if components.count == 2 {
                        let header = components[0]
                        let components = components[1].split(at: ")")
                        if components.count == 2 {
                            let indexString = components[0]
                            if let index = Int(indexString) {
                                if let entry = ParticipantHeader(rawValue: header) {
                                    switch entry {
                                    case .name:
                                        participant.member.playerList[index].name = string
                                    case .nationalId:
                                        participant.member.playerList[index].nationalId = string
                                    case .boardsPlayed:
                                        participant.member.playerList[index].accumulatedBoardsPlayed = value ?? 0
                                    case .winsDraws:
                                        participant.member.playerList[index].accumulatedWinDraw = Float(value ?? 0)
                                    default:
                                        break
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    private func finalUpdates() {
        if let boards = scoreData.events.first!.boards, let rounds = self.rounds {
            if boards != 0 && rounds != 0 {
                scoreData.events.first!.boardsPerRound = boards / rounds
            }
        }
        if requiresWinDraw {
            scoreData.events.first!.type = EventType("SWISS_" + scoreData.events.first!.type!.rawValue)
        }
    }
    
    private func report(error: String) {
        completion(nil, error)
    }
}
