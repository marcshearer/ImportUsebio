//
//  Parser.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import Foundation

fileprivate class Node {
    var parent: Node?
    var child: Node? = nil
    let name: String
    let attributes: [String : String]
    var value: String
    var process: ((String, [String:String])->())?
    var completion: ((String)->())?
    
    init(name: String = "", attributes: [String : String] = [:], value: String = "", process: ((String,[String:String])->())? = nil, completion: ((String)->())? = nil) {
        self.name = name.uppercased()
        self.attributes = attributes
        self.value = value
        self.process = process
        self.completion = completion
    }
    
    public func addValue(string: String) {
        self.value += string
    }
    
    public func add(child: Node) -> Node {
        child.parent = self
        self.child = child
        return child
    }
}

public class UsebioParser: NSObject, XMLParserDelegate {

    private var data: Data!
    private var completion: (ScoreData?, String?)->()
    private var parser: XMLParser!
    private let replacingSingleQuote = "@@replacingSingleQuote@@"
    private var root: Node?
    private var current: Node?
    private let scoreData = ScoreData()
    private var errors: [String] = []
    private var warnings: [String] = []
    private var travellerDirection: Direction?
    
    init(fileUrl: URL, data: Data, completion: @escaping (ScoreData?, String?)->()) {
        self.scoreData.fileUrl = fileUrl
        self.scoreData.source = .usebio
        self.completion = completion
        let string = String(decoding: data, as: UTF8.self)
        let replacedQuote = string.replacingOccurrences(of: "&#39;", with: replacingSingleQuote)
        self.data = replacedQuote.data(using: .utf8)
        super.init()
        root = Node(name: "MAIN", process: processMain)
        current = root
        parser = XMLParser(data: self.data)
        parser.delegate = self
        parser.parse()
    }
    
    func parseComplete() {
        finalUpdates()
        checkWinDraw()
        completion(scoreData, nil)
    }
            
    // MARK: - Parser Delegate ========================================================================== -

    public func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String] = [:]) {
        if let process = current?.process {
            process(elementName, attributeDict)
        } else {
            self.current = self.current?.add(child: Node(name: elementName))
        }
    }
    
    public func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        current?.completion?(current?.value ?? "")
        current = current?.parent
    }

    public func parser(_ parser: XMLParser, foundCharacters string: String) {
        let string = string.replacingOccurrences(of: replacingSingleQuote, with: "'")
        current?.addValue(string: string)
    }
    
    public func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        print("error")
    }
    
    public func parser(_ parser: XMLParser, validationErrorOccurred validationError: Error) {
        print("error")
    }
    
    public func parserDidEndDocument(_ parser: XMLParser) {
        parseComplete()
    }
    
    // MARK: Processors
    
    private func processMain(name: String, attributes: [String : String]) {
        switch name {
        case "USEBIO":
            current = current?.add(child: Node(name: name, attributes: attributes, process: processUsebio))
            scoreData.version = current?.attributes["Version"]
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processUsebio(name: String, attributes: [String : String]) {
        switch name {
        case "CLUB":
            scoreData.clubs.append(Club())
            current = current?.add(child: Node(name: name, attributes: attributes, process: processClub))
        case "EVENT":
            let event = Event()
            event.type = EventType(attributes["EVENT_TYPE"] ?? "INVALID")
            scoreData.events.append(event)
            current = current?.add(child: Node(name: name, attributes: attributes, process: processEvent))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processClub(name: String, attributes: [String : String]) {
        switch name {
        case "CLUB_NAME":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.clubs.last?.name = value
            }))
        case "CLUB_ID_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.clubs.last?.id = value
            }))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processEvent(name: String, attributes: [String : String]) {
        switch name {
        case "PROGRAM_NAME":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.programName = value
            }))
        case "PROGRAM_VERSION":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.programVersion = value
            }))
        case "EVENT_DESCRIPTION":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.description = value
            }))
        case "DATE":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.date = Utility.dateFromString(value, format: "dd/MM/yyyy")
            }))
        case "EVENT_IDENTIFIER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.eventCode = value
            }))
        case "BOARD_SCORING_METHOD":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.boardScoring = ScoringMethod(value)
            }))
        case "MATCH_SCORING_METHOD":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.matchScoring = ScoringMethod(value)
            }))
        case "BOARDS_PLAYED":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.boards = Int(value)
            }))
        case "BOARDS_PER_ROUND":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.boardsPerRound = Int(value)
            }))
        case "WINNER_TYPE":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.winnerType = Int(value) ?? 0
            }))
        case "SECTION_COUNT":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.sectionCount = Int(value) ?? 0
            }))
        case "SESSION_COUNT":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.sessionCount = Int(value) ?? 0
            }))
        case "CONTACT":
            current = current?.add(child: Node(name: name, process: processName))
        case "PARTICIPANTS":
            current = current?.add(child: Node(name: name, process: processParticipants))
        case "MATCH":
            let match = Match()
            scoreData.events.last?.matches.append(match)
            current = current?.add(child: Node(name: name, process: processMatch))
        case "SESSION":
            current = current?.add(child: Node(name: name, process: processEvent))
        case "SECTION":
            current = current?.add(child: Node(name: name, process: processEvent))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processName(name: String, attributes: [String : String]) {
        switch name {
        case "FULL_NAME":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.scoreData.events.last?.contact = value
            }))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processParticipants(name: String, attributes: [String : String]) {
        let event = scoreData.events.first!
        switch name {
        case "TEAM":
            current = current?.add(child: Node(name: name, process: processTeam))
            let participant = Participant(.team, from: event)
            let team = participant.member as! Team
            team.number = attributes["TEAM_ID"]
            team.name = attributes["TEAM_NAME"]
            self.scoreData.events.last?.participants.append(participant)
        case "PAIR":
            current = current?.add(child: Node(name: name, process: processPair))
            self.scoreData.events.last?.participants.append(Participant(.pair, from: event))
        case "PLAYER":
            current = current?.add(child: Node(name: name, process: processPlayer))
            self.scoreData.events.last?.participants.append(Participant(.player, from: event))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processTeam(name: String, attributes: [String : String]) {
        let participant = self.scoreData.events.last!.participants.last!
        let team = participant.member as! Team
        switch name {
        case "PAIR":
            current = current?.add(child: Node(name: name, process: processPair))
            team.pairs.append(Pair())
        case "PLAYER":
            current = current?.add(child: Node(name: name, process: processPlayer))
            team.players.append(Player())
        default:
            processParticipant(participant: participant, name: name, attributes: attributes)
        }
    }
    
    func processPair(name: String, attributes: [String : String]) {
        let participant = self.scoreData.events.last!.participants.last!
        var pair: Pair
        if participant.type == .team {
            pair = (participant.member as! Team).pairs.last!
        } else {
            pair = participant.member as! Pair
        }
        switch name {
        case "PAIR_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                pair.number = value
            }))
        case "PLAYER":
            current = current?.add(child: Node(name: name, process: processPlayer))
            pair.players.append(Player(pair: pair))
        case "BOARDS_PLAYED":
            current = current?.add(child: Node(name: name, completion: { (value) in
                pair.boardsPlayed = Int(value)
            }))
        case "DIRECTION":
            current = current?.add(child: Node(name: name, completion: { (value) in
                pair.direction = Direction(value)
            }))
        default:
            processParticipant(participant: participant, name: name, attributes: attributes)
        }
    }
    
    func processPlayer(name: String, attributes: [String : String]) {
        let participant = self.scoreData.events.last!.participants.last!
        var player: Player!
        if participant.type == .team {
            if let team = participant.member as? Team {
                if team.pairs.isEmpty {
                    player = team.players.last!
                } else {
                    player = team.pairs.last!.players.last!
                }
            }
        } else if participant.type == .pair {
            if let pair = participant.member as? Pair {
                player = pair.players.last!
            }
        } else {
            player = participant.member as? Player
        }
        switch name {
        case "PLAYER_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                player.number = value
            }))
        case "PLAYER_NAME":
            current = current?.add(child: Node(name: name, completion: { (value) in
                player.name = value
            }))
        case "NATIONAL_ID_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                player.nationalId = value
            }))
        default:
            processParticipant(participant: participant, name: name, attributes: attributes)
        }
    }
    
    func processParticipant(participant: Participant, name: String, attributes: [String : String]) {
        switch name {
        case "PLACE":
            current = current?.add(child: Node(name: name, completion: { (value) in
                participant.place = Int(value.replacingOccurrences(of: "=", with: ""))
            }))
        case "TOTAL_SCORE", "PERCENTAGE":
            current = current?.add(child: Node(name: name, completion: { (value) in
                participant.score = Float(value)
            }))
        case "WINS_OR_DRAWS":
            current = current?.add(child: Node(name: name, completion: { (value) in
                participant.winDraw = Float(value)
            }))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    func processMatch(name: String, attributes: [String : String]) {
        let match = scoreData.events.last?.matches.last
        switch name {
        case "ROUND_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.round = Int(value)
            }))
        case "TEAM", "NS_PAIR_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.number = value
            }))
        case "OPPOSING_TEAM", "EW_PAIR_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.opposingNumber = value
            }))
        case "TEAM_SCORE", "NS_SCORE":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.score = Float(value)
            }))
        case "OPPOSING_TEAM_SCORE", "EW_SCORE":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.opposingScore = Float(value)
            }))
        case "BOARD":
            current = current?.add(child: Node(name: name, process: processBoard))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    private func processBoard(name: String, attributes: [String : String]) {
        switch name {
        case "TRAVELLER_LINE":
            travellerDirection = nil
            current = current?.add(child: Node(name: name, process: processTravellerLine, completion: { (value) in
                self.travellerDirection = nil
            }))
        default:
            current = current?.add(child: Node(name: name))
        }
    }

    private func processTravellerLine(name: String, attributes: [String : String]) {
        let match = scoreData.events.last?.matches.last
        switch name {
        case "DIRECTION":
            current = current?.add(child: Node(name: name, completion: { (value) in
                self.travellerDirection = Direction(value)
            }))
        case "NS_PAIR_NUMBER", "EW_PAIR_NUMBER":
            current = current?.add(child: Node(name: name, completion: { (value) in
                if name.left(2) == self.travellerDirection?.string.uppercased() {
                    match?.pairNumbers.insert(value)
                } else {
                    match?.opposingPairNumbers.insert(value)
                }
            }))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    private func finalUpdates() {
        if let event = scoreData.events.first {
            if event.boardScoring == nil {
                if event.type?.participantType?.players == 4 {
                    event.boardScoring = .imps
                } else {
                    event.boardScoring = .percentage
                }
            }
        }
    }

    private func checkWinDraw() {
        if let event = scoreData.events.first {
            if event.type?.requiresWinDraw ?? false {
                for participant in event.participants.filter({$0.winDraw == nil}) {
                    
                    participant.winDraw = 0
                    if let team = participant.member as? Team {
                        for pair in team.pairs {
                            pair.winDraw = 0
                        }
                    }
                    for match in event.matches.filter({$0.number == participant.member.number || $0.opposingNumber == participant.member.number}) {
                        
                        // Add to participant wins/draws
                        var increment:Float = 0.0
                        if let score = match.score, let opposingScore = match.opposingScore {
                            if score == opposingScore {
                                increment = 0.5
                            } else if score > opposingScore && match.number == participant.member.number {
                                increment = 1.0
                            } else if score < opposingScore && match.number != participant.member.number {
                                increment = 1.0
                            }
                        }
                        participant.winDraw! += increment

                        // Add to wins/draws in teams pairs if relevant
                        if increment != 0 {
                            if let team = participant.member as? Team {
                                for pair in team.pairs {
                                    if (match.number == participant.member.number &&            match.pairNumbers.contains(pair.number!)) ||
                                        (match.opposingNumber == participant.member.number && match.opposingPairNumbers.contains(pair.number!)) {
                                        pair.winDraw = pair.winDraw! + increment
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
 }
