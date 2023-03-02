//
//  Validation.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 09/02/2023.
//

import Foundation

extension ScoreData {
    
    public func validate() -> (Bool, [String]?, [String]?) {
        errors = []
        warnings = []
        validateMissingNationalIds = false
        
        validateFile()
        if errors.isEmpty {
            validateEvent()
            if errors.isEmpty {
                validateParticipants()
            }
        }        
        return (validateMissingNationalIds, errors.count == 0 ? nil : errors, warnings.count == 0 ? nil: warnings)
    }
    
    // Mark: - Main file validation
    
    private func validateFile() {
        
        if version == nil {
            error("Invalid USEBIO file")
        } else if version != "1.2" {
            error("Only Usebio v1.2 is supported")
        }
        if events.count > 1 {
            error("File contains more than 1 event")
        } else if events.count == 0 {
            error("No events found")
        }
        if clubs.count > 1 {
            error("File contains more than 1 club")
        } else if clubs.count == 0 {
            error("No clubs found")
        }
    }

    // MARK: - Event validation
        
    private func validateEvent() {
        let event = events.first!
        
        if event.type == nil { 
            error("No event type specified")
        } else if !event.type!.supported {
            error("\(event.type!.string) event type not currently supported")
        }
        
        if event.boardScoring == nil {
            error("Invalid board scoring method")
        }
        
        if (event.boards ?? 0) <= 0 {
            error("Number of boards played unspecified / zero")
        }
        
        if (event.boardsPerRound ?? 0) == 0 && (event.type?.requiresWinDraw ?? false) {
            error("Number of boards per round unspecified / invalid")
        }
        
        if event.sectionCount != 1 {
            warning("Section count is \(event.sectionCount)")
        }
        
        if event.sessionCount != 1 {
            warning("Session count is \(event.sessionCount)")
        }
        
        if event.winnerType < 1 || event.winnerType > 2 {
            warning("Winner type is \(event.winnerType)")
        } else if event.winnerType == 2 && event.type?.participantType != .pair {
            error("2 winners in \(event.type?.string ?? "incompatible") event")
        }
    }
    
    // MARK: - Participant validation
        
    private func validateParticipants() {
        let event = events.first!
        var participantNumbers: [String : Bool] = [:]
        var playerNationalIds: [ String : [String] ] = [:]
        
        for participant in event.participants {
            
            if participant.member.number == nil {
                error("Participant found with no number")
            } else {
                // Check no duplicates
                if participantNumbers[participant.number] == nil {
                    participantNumbers[participant.number] = false
                } else {
                    if participantNumbers[participant.number] == false {
                        error("Duplicate found for \(participant.type.string) '\(participant.number)'")
                    }
                    participantNumbers[participant.description] = true
                }
            }
            
            if (participant.place ?? 0) <= 0 || (participant.place ?? 0) > (event.participants.count / event.winnerType) {
                error("Invalid place (\(participant.place ?? 0) for \(participant.description)")
            }
            
            if participant.score == nil {
                error("No score for \(participant.description)")
            }
            
            if participant.type != events.first!.type!.participantType {
                error("Participant is \(participant.type.string) but event requires \(event.type!.participantType!.string) participants")
            }
            
            switch participant.type {
            case .player:
                validate(player: participant.member as! Player)
            case .pair:
                validate(pair: participant.member as! Pair)
            case .team:
                validate(team: participant.member as! Team)
            }
            
            for player in participant.member.playerList {
                // Build national Id list
                if let nationalId = player.nationalId {
                    if nationalId != "0" && nationalId != "" {
                        if playerNationalIds[nationalId] == nil {
                            playerNationalIds[nationalId] = []
                        }
                        playerNationalIds[nationalId]!.append("\(player.description) in \(participant.description)")
                    }
                }
                validate(player: player, errorSuffix: " in \(participant.description)")
            }
            
            if (event.type?.requiresWinDraw ?? false) && participant.winDraw == nil {
                error("No wins/draws for \(participant.description)")
            }
        }
        // Report duplicate National IDs
        for (nationalId, players) in playerNationalIds {
            if players.count > 1 {
                var message = "Duplicate National Id \(nationalId) for "
                for (playerNumber, player) in players.enumerated() {
                    if playerNumber == players.count - 1 {
                        message += " & "
                    } else if playerNumber != 0 {
                        message += ", "
                    }
                    message += player
                }
                warning(message)
            }
        }
    }
    
    private func validate(player: Player, errorSuffix: String = "") {
        
        if (player.nationalId ?? "") == "" || player.nationalId == "0" {
            validateMissingNationalIds = true
        }
    }
    
    private func validate(pair: Pair, errorSuffix: String = "") {
        
        if pair.players.count != 2 {
            error("Pair \(pair.description) has \(pair.players.count) players\(errorSuffix)")
        }        
    }
    
    private func validate(team: Team, errorSuffix: String = "") {
        
        if team.pairs.count < 2 {
            error("Team \(team.description) has \(team.pairs.count) pairs\(errorSuffix)")
        }
    }
       
    // MARK: - Utility routines

    func error(_ text: String) {
        errors.append(text)
    }

    func warning(_ text: String) {
        warnings.append(text)
    }
}
