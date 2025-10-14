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
    private var completion: (ScoreData?, [String])->()
    private var parser: XMLParser!
    private let replacingSingleQuote = "@@replacingSingleQuote@@"
    private var root: Node?
    private var current: Node?
    private var scoreData = ScoreData()
    private var errors: [String] = []
    private var warnings: [String] = []
    private var travellerDirection: Direction?
    private var filterSessionId: String?
    private var roundContinuousVPDraw = true
    private var filterParticipantNumberMin: String?
    private var filterParticipantNumberMax: String?
    private var overrideEventType: EventType? // Used to switch to a specific event type (currently only for head-to-head teams league)
    
    init(fileUrl: URL, data: Data, filterSessionId: String? = nil, filterParticipantNumberMin: String? = nil, filterParticipantNumberMax: String? = nil, overrideEventType: EventType? = nil, roundContinuousVPDraw: Bool = false, winDrawLevel: WinDrawLevel? = nil, mergeMatches: Bool = false, vpType: VpType? = nil, completion: @escaping (ScoreData?, [String])->()) {
        self.scoreData.fileUrl = fileUrl
        self.scoreData.source = .usebio
        self.scoreData.winDrawLevel = winDrawLevel ?? .board
        self.scoreData.roundContinuousVPDraw = roundContinuousVPDraw
        self.scoreData.mergeMatches = mergeMatches
        self.scoreData.vpType = vpType ?? .discrete
        self.completion = completion
        let string = String(decoding: data, as: UTF8.self)
        let replacedQuote = string.replacingOccurrences(of: "&#39;", with: replacingSingleQuote)
        self.data = replacedQuote.data(using: .utf8)
        self.filterSessionId = filterSessionId
        self.filterParticipantNumberMin = filterParticipantNumberMin
        self.filterParticipantNumberMax = filterParticipantNumberMax
        self.overrideEventType = overrideEventType
        super.init()
        root = Node(name: "MAIN", process: processMain)
        current = root
        parser = XMLParser(data: self.data)
        parser.delegate = self
        parser.parse()
    }
    
    func parseComplete() {
        filterParticipantNumbers()
        finalUpdates()
        UsebioParser.calculatePlace(scoreData: scoreData)
        let message = UsebioParser.calculateWinDraw(scoreData: scoreData)
        completion(scoreData, message)
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
            if let overrideEventType = overrideEventType {
                event.type = overrideEventType
            } else {
                event.type = EventType(attributes["EVENT_TYPE"] ?? "INVALID")
            }
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
                self.scoreData.events.last?.date = Date(timeInterval: TimeInterval(12*60*60), since: Utility.dateFromString(value, format: "dd/MM/yyyy")!)
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
            var matched = true
            if let filterSessionId = filterSessionId {
                if let id = attributes["SESSION_ID"] {
                    if id.uppercased() != filterSessionId.uppercased() {
                        matched = false
                    }
                }
            }
            if matched {
                // Matched session data
                current = current?.add(child: Node(name: name, process: processEvent))
            } else {
                // Unmatched session data - gather to one side
                let mainSessionData = scoreData
                if mainSessionData.otherSessionData == nil {
                    // Create scoreData object for unmatched session with a dummy event
                    mainSessionData.otherSessionData = ScoreData()
                    mainSessionData.otherSessionData!.events.append(Event())
                }
                scoreData = mainSessionData.otherSessionData!
                current = current?.add(child: Node(name: name, process: processEvent, completion: { (_) in
                    self.scoreData = mainSessionData
                }))
            }
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
        case "TEAM_VICTORY_POINTS", "NS_VICTORY_POINTS":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.vp = Float(value)
            }))
        case "OPPOSING_TEAM_VICTORY_POINTS", "EW_VICTORY_POINTS":
            current = current?.add(child: Node(name: name, completion: { (value) in
                match?.opposingVP = Float(value)
            }))
        case "BOARD":
            current = current?.add(child: Node(name: name, process: processBoard))
        default:
            current = current?.add(child: Node(name: name))
        }
    }
    
    private func processBoard(name: String, attributes: [String : String]) {
        switch name {
        case "IMPS":
            // Used if rescoring wins/losses at board level
            let match = scoreData.events.last?.matches.last
            current = current?.add(child: Node(name: name, completion: { (value) in
                let board = Board()
                if let value = Float(value) {
                    board.nsScore = value
                    board.ewScore = -value
                }
                match?.boards.append(board)
            }))
        case "TRAVELLER_LINE":
            // Not sure this is used
            travellerDirection = nil
            current = current?.add(child: Node(name: name, process: processTravellerLine, completion: { (value) in
                self.travellerDirection = nil
            }))
        default:
            current = current?.add(child: Node(name: name))
        }
    }

    private func processTravellerLine(name: String, attributes: [String : String]) {
        // Used if match doesn't have pair numbers in it so need to look at travellers
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
    
    private func filterParticipantNumbers() {
        var nextPlace = 1
        if let filterParticipantNumberMin = self.filterParticipantNumberMin, let filterParticipantNumberMax = self.filterParticipantNumberMax {
            var newParticipants: [Participant] = []
            for participant in scoreData.events.first!.participants.sorted(by: {$0.place ?? 0 < $1.place ?? 0}) {
                if let numericNumber = Float(participant.member.number ?? "?"), let min = Float(filterParticipantNumberMin), let max = Float(filterParticipantNumberMax) {
                        // Integer values
                    if numericNumber < min || numericNumber > max {
                        continue
                    }
                } else if let stringNumber = participant.member.number {
                        // String values
                    if stringNumber < filterParticipantNumberMin || stringNumber > filterParticipantNumberMax {
                        continue
                    }
                } else {
                    continue
                }
                newParticipants.append(participant)
                newParticipants.last!.place = nextPlace
                nextPlace += 1
            }
            scoreData.events.first!.participants = newParticipants
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

    public static func calculatePlace(scoreData: ScoreData) {
        if let event = scoreData.events.first {
            if event.participants.contains(where: {($0.place ?? 0) == 0}) {
                var lastScore: Float?
                var lastPlace = 1
                var count = 0
                for participant in event.participants.sorted(by: { $0.score ?? 0 > $1.score ?? 0 }) {
                    count += 1
                    if participant.score ?? 0 == lastScore {
                        participant.place = lastPlace
                    } else {
                        participant.place = count
                        lastPlace = count
                    }
                    lastScore = participant.score ?? 0
                }
            }
        }
    }
    
    public static func calculateWinDraw(scoreData: ScoreData) -> [String] {
        var useData: ScoreData?
        var messages: [String] = []
        var suffix: String?
        // Use other sessions if no matches in session being filtered
        if let event = scoreData.events.first {
            if (event.matches).isEmpty {
                useData = scoreData.otherSessionData
                suffix = " (using scores in other sessions)"
            } else {
                useData = scoreData
            }
            if let useEvent = useData?.events.first {
                if event.type?.requiresWinDraw ?? false && scoreData.winDrawLevel > .participant {
                    if !useEvent.matches.isEmpty {
                        if let matchScoring = event.matchScoring, let boardScoring = event.boardScoring {
                            if let message = recalculateMatches(event: useEvent, winDrawLevel: scoreData.winDrawLevel, merge: scoreData.mergeMatches, matchScoring: matchScoring, vpType: scoreData.vpType, boardScoring: boardScoring) {
                                messages.append(message + (suffix ?? ""))
                            }
                        } else {
                            messages.append("Wins/draws rescored from match scores" + (suffix ?? ""))
                        }
                        for participant in event.participants.sorted(by: {$0.place ?? 0 < $1.place ?? 0 }) {
                            if let team = participant.member as? Team {
                                for pair in team.pairs {
                                    pair.winDraw = 0
                                }
                            }
                            var winDraw:Float = 0
                            for match in useEvent.matches.filter({$0.number == participant.member.number || $0.opposingNumber == participant.member.number}) {
                                
                                // Add to participant wins/draws
                                var increment:Float = 0.0
                                if let score = match.vp ?? match.score, let opposingScore = match.opposingVP ?? match.opposingScore {
                                    if score == opposingScore {
                                        increment = 0.5
                                    } else if useEvent.matchScoring == .vps && scoreData.roundContinuousVPDraw && score.rounded() == opposingScore.rounded() {
                                        increment = 0.5
                                    } else if score > opposingScore && match.number == participant.member.number {
                                        increment = 1.0
                                    } else if score < opposingScore && match.number != participant.member.number {
                                        increment = 1.0
                                    }
                                }
                                winDraw += increment
                                // Add to wins/draws in teams pairs if relevant (more than 2 pairs/team)
                                if increment != 0 {
                                    if let team = participant.member as? Team {
                                        for pair in team.pairs {
                                            if (match.number == participant.member.number && match.pairNumbers.contains(pair.number!)) ||
                                                (match.opposingNumber == participant.member.number && match.opposingPairNumbers.contains(pair.number!)) {
                                                pair.winDraw = pair.winDraw! + increment
                                            }
                                        }
                                    }
                                }
                            }
                            // Add to participant
                            if participant.winDraw != winDraw {
                                var name = participant.member.name
                                if name == nil || name == "" {
                                    name = participant.names
                                }
                                messages.append("Win/draws \(participant.winDraw == nil ? "set" : "updated from \(participant.winDraw!)") to \(winDraw) - \(participant.type.string) \(name!)")
                                participant.winDraw = winDraw
                            }
                        }
                    } else {
                        messages.append("No \(scoreData.winDrawLevel.string) level data available for win/draws")
                    }
                }
            }
        }
        return messages
    }
    
    private static func recalculateMatches(event: Event, winDrawLevel: WinDrawLevel, merge: Bool, matchScoring: ScoringMethod, vpType: VpType, boardScoring: ScoringMethod) -> String? {
        var rescored = 0
        
        if merge {
            UsebioParser.mergeMatches(event: event, scoring: matchScoring)
        }
        if winDrawLevel == .board {
            // Re-score matches from boards
            for match in event.matches {
                if !match.boards.isEmpty {
                    let boardTotal = boardScoring.combine(scores: match.boards.map{$0.nsScore})
                    if let boardTotal = boardTotal {
                        var score: Float? = nil
                        switch matchScoring {
                        case .aggregate, .cross_imps, .imps, .match_points, .percentage:
                            if matchScoring == boardScoring {
                                score = boardTotal
                            }
                        case .vps:
                            if vpType == .discrete {
                                score = Float(BridgeImps(Int(boardTotal)).discreteVp(boards: match.boards.count, maxVp: 20))
                            } else {
                                score = Float(BridgeImps(Int(boardTotal)).vp(boards: match.boards.count, maxVp: 20, places: 2))
                            }
                        }
                        if let score = score, let opposingScore = matchScoring.invert(score: score) {
                            match.score = score
                            match.opposingScore = opposingScore
                            if matchScoring == .vps {
                                match.vp = score
                                match.opposingVP = opposingScore
                            } else {
                                match.vp = nil
                                match.opposingVP = nil
                            }
                            rescored += 1
                        }
                    }
                }
            }
        } else if winDrawLevel == .match {
            rescored = event.matches.count
        }
        return "Wins/draws \(rescored == 0 ? "not " : (rescored == event.matches.count ? "" : "partially "))rescored from \(winDrawLevel.plural)"
    }
    
    private static func mergeMatches(event: Event, scoring: ScoringMethod) {
        var remove: [Int] = []
        for (index, match) in event.matches.enumerated() {
            if index < event.matches.count - 1 {
                for subIndex in index+1..<event.matches.count {
                    let subMatch = event.matches[subIndex]
                    if (match.number == subMatch.number && match.opposingNumber == subMatch.opposingNumber) || (match.number == subMatch.opposingNumber && match.opposingNumber == subMatch.number) {
                            // Match found - merge in
                        var subScore: Float?
                        var subOpposingScore: Float?
                        var subVp: Float?
                        var subOpposingVp: Float?
                        var subPairNumbers: Set<String>
                        var subOpposingPairNumbers: Set<String>
                        if match.number == subMatch.number {
                            subScore = subMatch.score
                            subOpposingScore = subMatch.opposingScore
                            subVp = subMatch.vp
                            subOpposingVp = subMatch.opposingVP
                            subPairNumbers = subMatch.pairNumbers
                            subOpposingPairNumbers = subMatch.opposingPairNumbers
                        } else {
                            subScore = subMatch.opposingScore
                            subOpposingScore = subMatch.score
                            subVp = subMatch.opposingVP
                            subOpposingVp = subMatch.vp
                            subPairNumbers = subMatch.opposingPairNumbers
                            subOpposingPairNumbers = subMatch.pairNumbers
                        }
                        match.score = scoring.combine(scores: match.score, subScore)
                        match.opposingScore = scoring.combine(scores: match.opposingScore, subOpposingScore)
                        if let vp = match.vp, let subVp = subVp {
                            match.vp = vp + subVp
                        }
                        if let opposingVp = match.opposingVP, let subOpposingVp = subOpposingVp {
                            match.opposingVP = opposingVp + subOpposingVp
                        }
                        match.pairNumbers = match.pairNumbers.union(subPairNumbers)
                        match.opposingPairNumbers = match.opposingPairNumbers.union(subOpposingPairNumbers)
                        for board in subMatch.boards {
                            if match.number != subMatch.number {
                                    // Swap scores if pair numbers other way round
                                let ewScore = board.ewScore
                                board.ewScore = board.nsScore
                                board.nsScore = ewScore
                            }
                            match.boards.append(board)
                        }
                        remove.append(subIndex)
                    }
                }
            }
        }
        // Now remove the duplicates
        for removeIndex in remove.reversed() {
            event.matches.remove(at: removeIndex)
        }
    }
 }
