//
//  Writer.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 08/02/2023.
//

import xlsxwriter

enum CellType {
    case string
    case integer
    case float
    case date
    case stringFormula
    case integerFormula
    case floatFormula
    case dataFormula
}

class Column {
    var title: String
    var content: ((Participant, Int)->String)?
    var playerContent: ((Participant, Player, Int, Int)->String)?
    var referenceContent: ((Int)->Int)?
    var referenceDivisor: Int?
    var referenceCalculated: ((Int)->String)?
    var referenceDynamic: (()->String)?
    var playerNumber: Int?
    var cellType: CellType
    var width: Float
    
    init(title: String, content: ((Participant, Int)->String)? = nil, playerContent: ((Participant, Player, Int, Int)->String)? = nil, playerNumber: Int? = nil, referenceContent: ((Int)->Int)? = nil, referenceDivisor: Int? = nil, referenceCalculated: ((Int)->String)? = nil, referenceDynamic: (()->String)? = nil, cellType: CellType = .string, width: Float = 10.0) {
        self.title = title
        self.content = content
        self.playerContent = playerContent
        self.playerNumber = playerNumber
        self.referenceContent = referenceContent
        self.referenceCalculated = referenceCalculated
        self.referenceDivisor = referenceDivisor
        self.referenceDynamic = referenceDynamic
        self.cellType = cellType
        self.width = width
    }
}

class Writer {
    let scoreData: ScoreData
    
    var summaryWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let summaryName = "Summary"
    let summaryDescriptionColumn = 0
    let summaryNationalLocalColumn = 3
    
    var consolidatedWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let consolidatedName = "Consolidated"
    
    var ranksPlusMpsWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let ranksPlusMPsName = "RanksPlusMPs"
    var ranksPlusMpsColumns: [Column] = []
    var ranksPlusMpsPositionColumn: Int?
    var ranksPlusMPsDirectionColumn: Int?
    var ranksPlusMPsParticipantNoColumn: Int?
    var ranksPlusMPsfFirstNameColumn: [Int] = []
    var ranksPlusMPsOtherNameColumn: [Int] = []
    var ranksPlusMPsNationalIdColumn: [Int] = []
    var ranksPlusMPsBoardsPlayedColumn: [Int] = []
    var ranksPlusMPsWinDrawColumn: [Int] = []
    var ranksPlusMPsBonusMPColumn: [Int] = []
    var ranksPlusMPsWinDrawMPColumn: [Int] = []
    var ranksPlusMPsTotalMPColumn: [Int] = []
    var ranksPlusMPsEventDescriptionCell: String?
    var ranksPlusMPsToeCell: String?
    var ranksPlusMPsTablesCell: String?
    var ranksPlusMPsMaxAwardCell: String?
    var ranksPlusMPsMaxEwAwardCell: String?
    var ranksPlusMPsAwardToCell: String?
    var ranksPlusMPsEwAwardToCell: String?
    var ranksPlusMPsPerWinCell: String?
    var ranksPlusMPsEventDateCell: String?
    var ranksPlusMPsEventIdCell: String?
    var ranksPlusMPsNationalLocalCell: String?
    let ranksPlusMPsHeaderRows = 3
    
    var individualMpsWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let individualMPsName = "IndividualMPs"
    var individualMpsColumns: [Column] = []
    var individualMPsPositionColumn: Int?
    var individualMPsUniqueColumn: Int?
    var individualMPsDecimalColumn: Int?
    var individualMPsNationalIdColumn: Int?
    var individualMPsLocalMPsColumn: Int?
    var individualMPsNationalMPsColumn: Int?
    var individualMPsLocalTotalColumn: Int?
    var individualMPsNationalTotalColumn: Int?
    var individualMPsChecksumColumn: Int?
    
    var workbook: UnsafeMutablePointer<lxw_workbook>?
    var formatInt: UnsafeMutablePointer<lxw_format>?
    var formatFloat: UnsafeMutablePointer<lxw_format>?
    var formatZeroInt: UnsafeMutablePointer<lxw_format>?
    var formatZeroFloat: UnsafeMutablePointer<lxw_format>?
    var formatBold: UnsafeMutablePointer<lxw_format>?
    var formatRightBold: UnsafeMutablePointer<lxw_format>?
    var formatBoldUnderline: UnsafeMutablePointer<lxw_format>?
    var formatRightBoldUnderline: UnsafeMutablePointer<lxw_format>?
    var formatFloatBoldUnderline: UnsafeMutablePointer<lxw_format>?
    var formatPercent: UnsafeMutablePointer<lxw_format>?
    var formatDate: UnsafeMutablePointer<lxw_format>?
    var fieldSize: Int? = nil
    var nsPairs: Int? = nil

   init(scoreData: ScoreData) {
       self.scoreData = scoreData
       setupRanksPlusMpsColumns()
       
       let name = "\(scoreData.fileUrl!.deletingPathExtension().lastPathComponent.removingPercentEncoding!).xlsm"
       workbook = workbook_new(name)
       summaryWorksheet = workbook_add_worksheet(workbook, summaryName)
       consolidatedWorksheet = workbook_add_worksheet(workbook, consolidatedName)
       individualMpsWorksheet = workbook_add_worksheet(workbook, individualMPsName)
       ranksPlusMpsWorksheet = workbook_add_worksheet(workbook, ranksPlusMPsName)
       workbook_add_vba_project(workbook, "./Award.bin")
       
       setupFormats()
    }
    
    func write() {
        writeRanksPlusMPs()
        writeIndividualMPs()
        writeSummary()
        writeConsolidated()
        workbook_close(workbook)
    }
    
    func writeSummary() {
        
        let toeColumn = 1
        let tablesColumn = 2
        let localMPsColumn = 4
        let nationalMPsColumn = 5
       
        setColumn(worksheet: summaryWorksheet, column: summaryDescriptionColumn, width: 30)
        setColumn(worksheet: summaryWorksheet, column: localMPsColumn, width: 12)
        setColumn(worksheet: summaryWorksheet, column: nationalMPsColumn, width: 12)

        write(worksheet: summaryWorksheet, row: 0, column: summaryDescriptionColumn, string: "Round", format: formatBold)
        write(worksheet: summaryWorksheet, row: 0, column: toeColumn, string: "TOE", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: 0, column: tablesColumn, string: "Tables", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: 0, column: summaryNationalLocalColumn, string: "Nat/Local", format: formatBold)
        write(worksheet: summaryWorksheet, row: 0, column: localMPsColumn, string: "Local MPs", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: 0, column: nationalMPsColumn, string: "National MPs", format: formatRightBold)
        
        write(worksheet: summaryWorksheet, row: 1, column: summaryDescriptionColumn, formula: "=\(ranksPlusMPsName)!\(ranksPlusMPsEventDescriptionCell!)")
        write(worksheet: summaryWorksheet, row: 1, column: toeColumn, integerFormula: "=\(ranksPlusMPsName)!\(ranksPlusMPsToeCell!)")
        write(worksheet: summaryWorksheet, row: 1, column: tablesColumn, integerFormula: "=\(ranksPlusMPsName)!\(ranksPlusMPsTablesCell!)")
        write(worksheet: summaryWorksheet, row: 1, column: summaryNationalLocalColumn, formula: "=\(ranksPlusMPsName)!\(ranksPlusMPsNationalLocalCell!)")
        write(worksheet: summaryWorksheet, row: 1, column: localMPsColumn, floatFormula: "=\(individualMPsName)!\(cell(1, rowFixed: true, individualMPsLocalTotalColumn!, columnFixed: true))", format: formatZeroFloat)
        write(worksheet: summaryWorksheet, row: 1, column: nationalMPsColumn, floatFormula: "=\(individualMPsName)!\(cell(1, rowFixed: true, individualMPsNationalTotalColumn!, columnFixed: true))", format: formatZeroFloat)
        
        let tablesColumnRef = columnRef(column: tablesColumn, fixed: true)
        write(worksheet: summaryWorksheet, row: 13, column: 0, string: "Total Tables", format: formatBold)
        write(worksheet: summaryWorksheet, row: 13, column: 1, integerFormula: "=SUM(\(tablesColumnRef)2:\(tablesColumnRef)11)", format: formatZeroInt)
        
        let localMPsColumnRef = columnRef(column: localMPsColumn, fixed: true)
        write(worksheet: summaryWorksheet, row: 14, column: 0, string: "Total Local MPs", format: formatBold)
        write(worksheet: summaryWorksheet, row: 14, column: 1, floatFormula: "=SUM(\(localMPsColumnRef)2:\(localMPsColumnRef)11)", format: formatZeroFloat)

        let nationalMPsColumnRef = columnRef(column: nationalMPsColumn, fixed: true)
        write(worksheet: summaryWorksheet, row: 15, column: 0, string: "Total National MPs", format: formatBold)
        write(worksheet: summaryWorksheet, row: 15, column: 1, floatFormula: "=SUM(\(nationalMPsColumnRef)2:\(nationalMPsColumnRef)11)", format: formatZeroFloat)
    }
    
    func writeConsolidated() {
        let maxRounds = 10
        var nameCell: [String] = []
        
        let uniqueColumn = 0
        let firstNameColumn = 1
        let otherNamesColumn = 2
        let nationalIdColumn = 3
        let localMPsColumn = 4
        let nationalMPsColumn = 5
        let dataColumn = 5
        
        let titleRow = 3
        let nationalLocalRow = 0
        let totalRow = 1
        let checksumRow = 2
        let dataRow = 4
        
        let nationalLocalRange = "\(cell(nationalLocalRow, rowFixed: true, dataColumn, columnFixed: true)):\(cell(nationalLocalRow, rowFixed: true, dataColumn + maxRounds - 1, columnFixed: true))"
        
        setRow(worksheet: consolidatedWorksheet, row: titleRow, format: formatBoldUnderline)
        setRow(worksheet: consolidatedWorksheet, row: nationalLocalRow, format: formatBoldUnderline)
        setRow(worksheet: consolidatedWorksheet, row: totalRow, format: formatBoldUnderline)
        setRow(worksheet: consolidatedWorksheet, row: checksumRow, format: formatBoldUnderline)

        for column in 0..<(dataColumn + maxRounds) {
            setColumn(worksheet: consolidatedWorksheet, column: column, width: 16.5)
        }
        
        // Title row
        write(worksheet: consolidatedWorksheet, row: titleRow, column: firstNameColumn, string: "First Name", format: formatBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: otherNamesColumn, string: "Other Names", format: formatBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: nationalIdColumn, string: "SBU", format: formatRightBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: localMPsColumn, string: "Local MPs", format: formatRightBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: nationalMPsColumn, string: "National MPs", format: formatRightBoldUnderline)
        
        for column in 0..<maxRounds {
            let summaryCell = "\(summaryName)!\(cell(column+1, rowFixed: true, summaryDescriptionColumn, columnFixed: true))"
            write(worksheet: consolidatedWorksheet, row: titleRow, column: dataColumn + column, formula: "=IF(\(summaryCell)=0,\"\",\(summaryCell))", format: formatRightBoldUnderline)
            nameCell.append(cell(titleRow, rowFixed: true, dataColumn + column))
        }
        
        // National/Local row
        for column in 0..<maxRounds {
            let cell = "\(summaryName)!\(cell(column+1, rowFixed: true, summaryNationalLocalColumn, columnFixed: true))"
            write(worksheet: consolidatedWorksheet, row: nationalLocalRow, column: dataColumn + column, formula: "=IF(\(cell)=0,\"\",\(cell))", format: formatRightBoldUnderline)
        }
        
        for element in 0...1 {
            // Total row and checksum row
            let row = (element == 0 ? totalRow : checksumRow)
            let totalRange = "\(cell(row, rowFixed: true, dataColumn, columnFixed: true)):\(cell(row, rowFixed: true, dataColumn + maxRounds - 1, columnFixed: true))"
            
            write(worksheet: consolidatedWorksheet, row: row, column: nationalIdColumn, string: (row == totalRow ? "Total" : "Checksum"), format: formatRightBoldUnderline)

            write(worksheet: consolidatedWorksheet, row: row, column: localMPsColumn, floatFormula: "=SUMIF(\(nationalLocalRange), \"<>National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            write(worksheet: consolidatedWorksheet, row: row, column: nationalMPsColumn, floatFormula: "=SUMIF(\(nationalLocalRange), \"=National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            for column in 0..<maxRounds {
                let valueRange = "\(cell(dataRow, rowFixed: true, dataColumn + column)):\(cell(dataRow + (maxPlayers * fieldSize!) - 1, rowFixed: true, dataColumn + column))"
                let nationalIdRange = "\(cell(dataRow, rowFixed: true, nationalIdColumn, columnFixed: true)):\(cell(dataRow + (maxPlayers * fieldSize!) - 1, rowFixed: true, nationalIdColumn, columnFixed: true))"
                
                if row == totalRow {
                    write(worksheet: consolidatedWorksheet, row: row, column: dataColumn + column, floatFormula: "=IF(\(nameCell[column])=\"\",\"\",SUM(\(valueRange)))", format: formatFloatBoldUnderline)
                } else {
                    write(worksheet: consolidatedWorksheet, row: row, column: dataColumn + column, floatFormula: "=IF(\(nameCell[column])=\"\",\"\",SUMPRODUCT(\(valueRange),\(nationalIdRange)))", format: formatFloatBoldUnderline)
                }
            }
        }
        
        // Data rows
        let formula = "=_xlfn._xlws.UNIQUE(_xlfn._xlws.VSTACK(\(individualMPsName)!\(cell(1,rowFixed: true, individualMPsUniqueColumn!, columnFixed: true)):\(cell((maxPlayers * fieldSize!),rowFixed: true, individualMPsUniqueColumn!, columnFixed: true))))"
        let wsDataRow = lxw_row_t(UInt32(dataRow))
        let wsUniqueColumn = lxw_col_t(UInt32(uniqueColumn))
        worksheet_write_dynamic_array_formula(consolidatedWorksheet, wsDataRow, wsUniqueColumn, 999, wsUniqueColumn, formula, nil)
        
        for row in dataRow...(dataRow + (maxPlayers * fieldSize!) - 1) {
            let uniqueIdCell = cell(row, uniqueColumn, columnFixed: true)
            let dataRange = "\(cell(row, dataColumn, columnFixed: true)):\(cell(row, rowFixed: true, dataColumn + maxRounds - 1, columnFixed: true))"
            worksheet_write_formula(consolidatedWorksheet, lxw_row_t(row), lxw_col_t(nationalIdColumn), "=IF(\(uniqueIdCell)=\"\",\"\",IFERROR(NUMBERVALUE(TEXTBEFORE(\(uniqueIdCell), \"+\")),TEXTBEFORE(\(uniqueIdCell), \"+\")))", formatFloat)
            // write(worksheet: consolidatedWorksheet, row: row, column: nationalIdColumn, integerFormula: "=IF(\(uniqueIdCell)=\"\",\"\",IFERROR(NUMBERVALUE(TEXTBEFORE(\(uniqueIdCell), \"+\")),TEXTBEFORE(\(uniqueIdCell), \"+\")))")
            write(worksheet: consolidatedWorksheet, row: row, column: firstNameColumn, integerFormula: "=IF(\(uniqueIdCell)=\"\",\"\",TEXTBEFORE(TEXTAFTER(\(uniqueIdCell),\"+\"),\"+\"))")
            write(worksheet: consolidatedWorksheet, row: row, column: otherNamesColumn, integerFormula: "=IF(\(uniqueIdCell)=\"\",\"\",TEXTAFTER(\(uniqueIdCell), \"+\", 2))")
            
            write(worksheet: consolidatedWorksheet, row: row, column: localMPsColumn, floatFormula: "=IF(\(uniqueIdCell)=\"\",\"\",SUMIF(\(nationalLocalRange), \"<>National\", \(dataRange)))")
            write(worksheet: consolidatedWorksheet, row: row, column: nationalMPsColumn, floatFormula: "==IF(\(uniqueIdCell)=\"\",\"\",SUMIF(\(nationalLocalRange), \"=National\", \(dataRange)))")
            
            let sourceDataRange = "\(cell(1, rowFixed: true, individualMPsUniqueColumn!, columnFixed: true)):\(cell((maxPlayers * fieldSize!), rowFixed: true, individualMPsDecimalColumn!, columnFixed: true))"
            let sourceOffset = (individualMPsDecimalColumn! - individualMPsUniqueColumn! + 1)
            
            for column in 0..<maxRounds {
                write(worksheet: consolidatedWorksheet, row: row, column: dataColumn + column, floatFormula: "=IF(OR(\(uniqueIdCell)=\"\",\(nameCell[column])=\"\"),0,IFERROR(VLOOKUP(\(uniqueIdCell),INDIRECT(\(nameCell[column])&\"IndividualMPs!\(sourceDataRange)\"),\(sourceOffset),FALSE),0))")
            }
        }
    }
    
    func writeRanksPlusMPs() {
        let event = scoreData.events.first!
        let participants = scoreData.events.first!.participants.sorted(by: sortCriteria)
        fieldSize = participants.count
        if event.winnerType == 2 && event.type?.participantType == .pair {
            nsPairs = participants.filter{(($0.member as? Pair)?.direction ?? .ns) == .ns}.count
        }
        
        writeRanksPlusMPsHeader()
        
        for (columnNumber, column) in ranksPlusMpsColumns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: ranksPlusMpsWorksheet, column: columnNumber, width: column.width)
            write(worksheet: ranksPlusMpsWorksheet, row: ranksPlusMPsHeaderRows, column: columnNumber, string: replace(column.title), format: format)
        }
        
        for (rowSequence, participant) in participants.enumerated() {
            
            let rowNumber = rowSequence + ranksPlusMPsHeaderRows + 1

            for (columnNumber, column) in ranksPlusMpsColumns.enumerated() {
                
                if let content = column.content?(participant, rowNumber) {
                    write(cellType: (content == "" ? .string : column.cellType), worksheet: ranksPlusMpsWorksheet, row: rowNumber, column: columnNumber, content: content)
                }
                
                if let playerNumber = column.playerNumber {
                    let playerList = participant.member.playerList
                    if playerNumber < playerList.count {
                        if let playerContent = column.playerContent?(participant, playerList[playerNumber], playerNumber, rowNumber) {
                            write(cellType: (playerContent == "" ? .string : column.cellType), worksheet: ranksPlusMpsWorksheet, row: rowNumber, column: columnNumber, content: playerContent)
                        }
                    } else {
                        write(cellType: .string, worksheet: ranksPlusMpsWorksheet, row: rowNumber, column: columnNumber, content: "")
                    }
                }
            }
        }
    }
    
    private func setupRanksPlusMpsColumns() {
        let event = scoreData.events.first!
        let playerCount = maxPlayers
        let winDraw = event.type?.requiresWinDraw ?? false
        
        ranksPlusMpsColumns.append(Column(title: "Place", content: { (participant, _) in "\(participant.place!)" }, cellType: .integer))
        ranksPlusMpsPositionColumn = ranksPlusMpsColumns.count - 1

        if event.winnerType == 2 && event.type?.participantType == .pair {
            ranksPlusMpsColumns.append(Column(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
            ranksPlusMPsDirectionColumn = ranksPlusMpsColumns.count - 1
        }
        
        ranksPlusMpsColumns.append(Column(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer)) ; ranksPlusMPsParticipantNoColumn = ranksPlusMpsColumns.count - 1

        ranksPlusMpsColumns.append(Column(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float))
        
        if winDraw && maxPlayers <= event.type?.participantType?.players ?? maxPlayers {
            ranksPlusMpsColumns.append(Column(title: "Win/Draw", content: { (participant, _) in "\(participant.winDraw!)" }, cellType: .float))
            ranksPlusMPsWinDrawColumn.append(ranksPlusMpsColumns.count - 1)
        }

        for playerNumber in 0..<playerCount {
            ranksPlusMpsColumns.append(Column(title: "First Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 0) }, playerNumber: playerNumber, cellType: .string))
            ranksPlusMPsfFirstNameColumn.append(ranksPlusMpsColumns.count - 1)

            ranksPlusMpsColumns.append(Column(title: "Other Names (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 1) }, playerNumber: playerNumber, cellType: .string))
            ranksPlusMPsOtherNameColumn.append(ranksPlusMpsColumns.count - 1)

            ranksPlusMpsColumns.append(Column(title: "SBU No (\(playerNumber+1))", playerContent: { (_, player,_, _) in player.nationalId! }, playerNumber: playerNumber, cellType: .integer))
            ranksPlusMPsNationalIdColumn.append(ranksPlusMpsColumns.count - 1)
            
            if maxPlayers > event.type?.participantType?.players ?? maxPlayers {
                
                ranksPlusMpsColumns.append(Column(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.boardsPlayed)" }, playerNumber: playerNumber, cellType: .integer))
                ranksPlusMPsBoardsPlayedColumn.append(ranksPlusMpsColumns.count - 1)
                
                if winDraw {
                    ranksPlusMpsColumns.append(Column(title: "Win/Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                    ranksPlusMPsWinDrawColumn.append(ranksPlusMpsColumns.count - 1)
                }
                
                ranksPlusMpsColumns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP (\(playerNumber+1))", playerContent: playerBonusAward, playerNumber: playerNumber, cellType: .integerFormula))
                ranksPlusMPsBonusMPColumn.append(ranksPlusMpsColumns.count - 1)
                
                if winDraw {
                    ranksPlusMpsColumns.append(Column(title: "Win/Draw MP (\(playerNumber+1))", playerContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .integerFormula))
                    ranksPlusMPsWinDrawMPColumn.append(ranksPlusMpsColumns.count - 1)
                    
                    ranksPlusMpsColumns.append(Column(title: "Total MP (\(playerNumber+1))", playerContent: playerTotalAward, playerNumber: playerNumber, cellType: .integerFormula))
                    ranksPlusMPsTotalMPColumn.append(ranksPlusMpsColumns.count - 1)
                }
            }
        }
            
        if maxPlayers <= event.type?.participantType?.players ?? maxPlayers {
            
            ranksPlusMpsColumns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP", content: bonusAward, cellType: .integerFormula))
            ranksPlusMPsBonusMPColumn.append(ranksPlusMpsColumns.count - 1)
            
            if winDraw {
                
                ranksPlusMpsColumns.append(Column(title: "Win/Draw MP", content: winDrawAward, cellType: .integerFormula))
                ranksPlusMPsWinDrawMPColumn.append(ranksPlusMpsColumns.count - 1)
                
                ranksPlusMpsColumns.append(Column(title: "Total MP", content: totalAward, cellType: .integerFormula))
                for _ in 0..<maxPlayers {
                    ranksPlusMPsTotalMPColumn.append(ranksPlusMpsColumns.count - 1)
                }
            }
        }
    }
    
    // MARK: - Ranks plus MPs Content getters
    
    private func bonusAward(participant: Participant, rowNumber: Int) -> String {
        return playerBonusAward(participant: participant, rowNumber: rowNumber)
    }
    
    private func playerBonusAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "="
        let event = scoreData.events.first!
        var useAwardCell = ranksPlusMPsMaxAwardCell
        var useAwardToCell = ranksPlusMPsAwardToCell
        var firstRow = ranksPlusMPsHeaderRows + 1
        var lastRow = ranksPlusMPsHeaderRows + event.participants.count
        if let nsPairs = nsPairs {
            if event.winnerType == 2 && participant.type == .pair {
                if let pair = participant.member as? Pair {
                    if pair.direction == .ew {
                        useAwardCell = ranksPlusMPsMaxEwAwardCell
                        useAwardToCell = ranksPlusMPsEwAwardToCell
                        firstRow = ranksPlusMPsHeaderRows + nsPairs + 1
                    } else {
                        lastRow = ranksPlusMPsHeaderRows + nsPairs
                    }
                }
            }
        }
        
        if let playerNumber = playerNumber {
            // Need to calculate percentage played
            let totalBoardsPlayed = participant.member.playerList.map{$0.boardsPlayed}.reduce(0, +) / event.type!.participantType!.players
            result += "ROUNDUP((\(cell(rowNumber, ranksPlusMPsBoardsPlayedColumn[playerNumber], columnFixed: true))/\(totalBoardsPlayed))*"
        }
        
        let positionCell = cell(rowNumber, ranksPlusMpsPositionColumn!)
        let allPositionsRange = "\(cell(firstRow, rowFixed: true, ranksPlusMpsPositionColumn!, columnFixed: true)):\(cell(lastRow, rowFixed: true, ranksPlusMpsPositionColumn!, columnFixed: true))"
        
        result += "Award(\(useAwardCell!), \(positionCell), \(useAwardToCell!), \(allPositionsRange))"
        
        if playerNumber != nil {
            result += ", 0)"
        }

        return result
    }
    
    private func winDrawAward(participant: Participant, rowNumber: Int) -> String {
        return playerWinDrawAward(participant: participant, rowNumber: rowNumber)
    }

    private func playerWinDrawAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let winDrawsCell = cell(rowNumber, ranksPlusMPsWinDrawColumn[playerNumber ?? 0], columnFixed: true)
        result = "=ROUNDUP(\(winDrawsCell) * \(ranksPlusMPsPerWinCell!), 0)"
        
        return result
    }
    
    private func totalAward(participant: Participant, rowNumber: Int) -> String {
        return playerTotalAward(participant: participant, rowNumber: rowNumber)
    }

    private func playerTotalAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let bonusMPCell = cell(rowNumber, ranksPlusMPsBonusMPColumn[playerNumber ?? 0], columnFixed: true)
        let winDrawMPCell = cell(rowNumber, ranksPlusMPsWinDrawMPColumn[playerNumber ?? 0], columnFixed: true)
        result = "\(bonusMPCell)+\(winDrawMPCell)"
        
        return result
    }
    
    // MARK: - Ranks plus MPs header
    
    func writeRanksPlusMPsHeader() {
        let event = scoreData.events.first!
        var column = -1
        var row = 0
        var prefix = ""
        var toe = fieldSize!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        if twoWinners {
            prefix = "NS "
            toe = nsPairs!
        }
        
        func writeCell(string: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, string: string, format: format)
        }
        
        func writeCell(integer: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, integer: integer, format: format)
        }
        
        func writeCell(float: Float, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, float: float, format: format)
        }
        
        func writeCell(date: Date, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, date: date, format: format)
        }
        
        func writeCell(formula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, floatFormula: formula, format: format)
        }

        func writeCell(integerFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, integerFormula: integerFormula, format: format)
        }

        writeCell(string: "Event", format: formatBold)
        writeCell(string: "\(prefix)TOE", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW TOE", format: formatRightBold)
        }
        writeCell(string: "Tables", format: formatRightBold)
        writeCell(string: "Boards", format: formatRightBold)
        writeCell(string: "\(prefix)Max Award", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW Max Award", format: formatRightBold)
        }
        writeCell(string: "Factor%", format: formatRightBold)
        writeCell(string: "\(prefix)Award", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW Award", format: formatRightBold)
        }
        writeCell(string: "Award %", format: formatRightBold)
        writeCell(string: "\(prefix)Award to", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW Award to", format: formatRightBold)
        }
        writeCell(string: "Per Win", format: formatRightBold)
        
        writeCell(string: "Event Date", format: formatBold)
        
        writeCell(string: "Event ID", format: formatBold)
        
        writeCell(string: "Local/Nat", format: formatBold)
        
        column = -1
        row = 1
        
        writeCell(string: event.description ?? "") ; ranksPlusMPsEventDescriptionCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: toe) ; ranksPlusMPsToeCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        var toeCell = ranksPlusMPsToeCell!
        var ewToeRef = ""
        if twoWinners {
            writeCell(integer: fieldSize! - nsPairs!) ; ewToeRef = cell(row, rowFixed: true, column, columnFixed: true)
            toeCell += "+" + ewToeRef
        }
        writeCell(formula: "=ROUNDUP(\(toeCell)*(\(event.type?.participantType?.players ?? 4)/4),0)") ; ranksPlusMPsTablesCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: event.boards ?? 0)
        
        var baseMaxEwAwardCell = ""
        writeCell(integer: 1000) ; let baseMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(integerFormula: "ROUNDUP(\(baseMaxAwardCell)*\(ewToeRef)/\(ranksPlusMPsToeCell!),0)") ; baseMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(integer: 1, format: formatPercent) ; let factorCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(baseMaxAwardCell)*\(factorCell),0)") ; ranksPlusMPsMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(baseMaxEwAwardCell)*\(factorCell),0)") ; ranksPlusMPsMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 0.25, format: formatPercent) ; let awardPercentCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(ranksPlusMPsToeCell!)*\(awardPercentCell),0)") ; ranksPlusMPsAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(ewToeRef)*\(awardPercentCell),0)") ; ranksPlusMPsEwAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(integer: 25) ; ranksPlusMPsPerWinCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(date: event.date ?? Date.today) ; ranksPlusMPsEventDateCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.eventId ?? "") ; ranksPlusMPsEventIdCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: "Local") ; ranksPlusMPsNationalLocalCell = cell(row, rowFixed: true, column, columnFixed: true)
    }
    
    // MARK: - Indiviual MPs worksheet
    
    func writeIndividualMPs() {
        setupIndividualMpsColumns()
        
        for (columnNumber, column) in individualMpsColumns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: individualMpsWorksheet, column: columnNumber, width: column.width)
            write(worksheet: individualMpsWorksheet, row: 0, column: columnNumber, string: replace(column.title), format: format)
        }
        
        for (columnNumber, column) in individualMpsColumns.enumerated() {
            
            if let referenceContent = column.referenceContent {
                referenceColumn(columnNumber: columnNumber, referencedContent: referenceContent, divisor: column.referenceDivisor)
            }
            
            if let referenceCalculatedContent = column.referenceCalculated {
                referenceCalculated(columnNumber: columnNumber, calculated: referenceCalculatedContent, cellType: column.cellType)
            }
            
            if let referenceDynamicContent = column.referenceDynamic {
                print(referenceDynamicContent())
                referenceDynamic(columnNumber: columnNumber, content: referenceDynamicContent)
            }
        }
        
        let localColumnRef = columnRef(column: individualMPsLocalMPsColumn!)
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsLocalTotalColumn!, floatFormula: "=SUM(\(localColumnRef):\(localColumnRef))", format: formatZeroFloat)
        
        let nationalColumnRef = columnRef(column: individualMPsNationalMPsColumn!)
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsNationalTotalColumn!, floatFormula: "=SUM(\(nationalColumnRef):\(nationalColumnRef))", format: formatZeroFloat)
        
        let totalColumnRef = columnRef(column: individualMPsDecimalColumn!)
        let nationalIdColumnRef = columnRef(column: individualMPsNationalIdColumn!)
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsChecksumColumn!, floatFormula: "=SUMPRODUCT(\(totalColumnRef):\(totalColumnRef),\(nationalIdColumnRef):\(nationalIdColumnRef))", format: formatZeroFloat)
        
        setColumn(worksheet: individualMpsWorksheet, column: individualMPsUniqueColumn!, hidden: true)

    }
    
    private func setupIndividualMpsColumns() {
        let event = scoreData.events.first!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        
        individualMpsColumns.append(Column(title: "Place", referenceContent: { (_) in self.ranksPlusMpsPositionColumn! }, cellType: .integerFormula)) ; individualMPsPositionColumn = individualMpsColumns.count - 1
        
        if twoWinners {
            individualMpsColumns.append(Column(title: "Direction", referenceContent: { (_) in self.ranksPlusMPsDirectionColumn! }, cellType: .stringFormula))
        }
        
        individualMpsColumns.append(Column(title: "@P no", referenceContent: { (_) in self.ranksPlusMPsParticipantNoColumn! }, cellType: .integerFormula))
        
        individualMpsColumns.append(Column(title: "Unique", cellType: .floatFormula)) ; individualMPsUniqueColumn = individualMpsColumns.count - 1
        let unique = individualMpsColumns.last!
        
        individualMpsColumns.append(Column(title: "Names", referenceContent: { (playerNumber) in self.ranksPlusMPsfFirstNameColumn[playerNumber] }, cellType: .stringFormula)) ; let firstNameColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "", referenceContent: { (playerNumber) in self.ranksPlusMPsOtherNameColumn[playerNumber] }, cellType: .stringFormula)) ; let otherNamesColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "SBU No", referenceContent: { (playerNumber) in self.ranksPlusMPsNationalIdColumn[playerNumber] }, cellType: .integerFormula)) ; individualMPsNationalIdColumn = individualMpsColumns.count - 1
        
        unique.referenceCalculated = { (rowNumber) in "CONCATENATE(\(self.cell(rowNumber, self.individualMPsNationalIdColumn!, columnFixed: true)), \"+\", \(self.cell(rowNumber, firstNameColumn, columnFixed: true)), \"+\", \(self.cell(rowNumber, otherNamesColumn, columnFixed: true)))" }
        
        individualMpsColumns.append(Column(title: "Total MPs", referenceContent: { (playerNumber) in self.ranksPlusMPsTotalMPColumn[playerNumber] }, referenceDivisor: 100, cellType: .floatFormula)) ; individualMPsDecimalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Local MPs", referenceCalculated: { (rowNumber) in "IF(\(self.ranksPlusMPsName)!\(self.ranksPlusMPsNationalLocalCell!) = \"National\",0,\(self.cell(rowNumber, self.individualMPsDecimalColumn!, columnFixed: true)))" }, cellType: .floatFormula)) ; individualMPsLocalMPsColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "National MPs", referenceCalculated: { (rowNumber) in "IF(\(self.ranksPlusMPsName)!\(self.ranksPlusMPsNationalLocalCell!) <> \"National\",0,\(self.cell(rowNumber, self.individualMPsDecimalColumn!, columnFixed: true)))" }, cellType: .floatFormula)) ; individualMPsNationalMPsColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Total Local", cellType: .floatFormula)) ; individualMPsLocalTotalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Total National", cellType: .floatFormula)) ; individualMPsNationalTotalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Checksum", cellType: .floatFormula)) ; individualMPsChecksumColumn = individualMpsColumns.count - 1
    }
    
    private func referenceColumn(columnNumber: Int, referencedContent: (Int)->Int, divisor: Int? = nil) {
        
        var result = "=_xlfn._xlws.FILTER(_xlfn._xlws.VSTACK("
        
        for playerNumber in 0..<maxPlayers {
            let columnReference = referencedContent(playerNumber)
            if playerNumber != 0 {
                result += ","
            }
            result += "\(ranksPlusMPsName)!" + cell(ranksPlusMPsHeaderRows + 1, columnReference)
            result += ":"
            result += cell(ranksPlusMPsHeaderRows + fieldSize!, columnReference)
        }
        
        result += ")"
        
        if let divisor = divisor {
            result += "/\(divisor)"
        }
        result += ",_xlfn._xlws.VSTACK("
        
        for playerNumber in 0..<maxPlayers {
            if playerNumber != 0 {
                result += ","
            }
            result += "\(ranksPlusMPsName)!" + cell(ranksPlusMPsHeaderRows + 1, ranksPlusMPsTotalMPColumn[playerNumber])
            result += ":"
            result += cell(ranksPlusMPsHeaderRows + fieldSize!, ranksPlusMPsTotalMPColumn[playerNumber])
        }
        
        result += ")<>0)"
                
        let column = lxw_col_t(Int32(columnNumber))
        worksheet_write_dynamic_array_formula(individualMpsWorksheet, 1, column, 999, column, result, nil)
    }
    
    private func referenceCalculated(columnNumber: Int, calculated: (Int)->String, cellType: CellType? = nil) {
        
        var format = formatInt
        if let cellType = cellType {
            switch cellType {
            case .date, .dataFormula:
                format = formatDate
            case .float, .floatFormula:
                format = formatFloat
            case .string, .stringFormula:
                format = nil
            default:
                break
            }
        }
        let column = lxw_col_t(Int32(columnNumber))
        
        for rowNumber in 1...(maxPlayers * fieldSize!) {
            let placeRef = cell(rowNumber, individualMPsPositionColumn!, columnFixed: true)
            var result = "=IF(\(placeRef)<>\"\","
            
            result += calculated(rowNumber)
            
            result += ",\"\")"
            
            let row = lxw_row_t(Int32(rowNumber))
            worksheet_write_formula(individualMpsWorksheet, row, column, result, format)
        }
    }
    
    private func referenceDynamic(columnNumber: Int, content: ()->String) {
                
        let column = lxw_col_t(Int32(columnNumber))
            worksheet_write_dynamic_array_formula(individualMpsWorksheet, 1, column, 999, column, content(), nil)
    }
    
    // MARK: - Utility routines
    
    private func sortCriteria(_ a: Participant, _ b: Participant) -> Bool {
        let aPlace = a.place ?? 0
        let bPlace = b.place ?? 0
        var aDirection: Int = 0
        var bDirection: Int = 0
        if scoreData.events.first!.winnerType == 2 {
            if let aPair = a.member as? Pair, let bPair = b.member as? Pair {
                aDirection = aPair.direction?.sort ?? 0
                bDirection = bPair.direction?.sort ?? 0
            }
        }
        if aDirection < bDirection {
            return true
        } else if aDirection == bDirection {
            if aPlace < bPlace {
                return true
            } else if (a.place ?? 0) == (b.place ?? 0) {
                if a.number < b.number {
                    return true
                } else {
                    return false
                }
            } else {
                return false
            }
        } else {
            return false
        }
    }
    
    private func nameColumn(name: String, element: Int) -> String {
        let names = name.components(separatedBy: " ")
        if element == 0 {
            return names[0]
        } else {
            var otherNames = names
            otherNames.removeFirst()
            return otherNames.joined(separator: " ")
        }
    }
    
    func replace(_ text: String) -> String {
        var text = text
        
        text = text.replacingOccurrences(of: "@P", with: scoreData.events.first!.type!.participantType!.string)
        
        return text
    }
    
    private var maxPlayers: Int {
        var result = 0
        
        let event = scoreData.events.first!
        if event.type?.participantType == .team {
            result = event.participants.map{$0.member.playerList.count}.max() ?? 0
        } else {
            result = event.type?.participantType?.players ?? 2
        }
        
        return result
    }
    
    func cell(_ row: Int, rowFixed: Bool = false, _ column: Int, columnFixed: Bool = false) -> String {
        let rowRef = rowRef(row: row, fixed: rowFixed)
        let columnRef = columnRef(column: column, fixed: columnFixed)
        return "\(columnRef)\(rowRef)"
    }
    
    func rowRef(row: Int, fixed: Bool = false) -> String {
        let rowRef = (fixed ? "$" : "") + "\(row + 1)"
        return rowRef
    }
    
    func columnRef(column: Int, fixed: Bool = false) -> String {
        var columnRef = ""
        var remaining = column
        while remaining >= 0 {
            let letter = remaining % 26
            columnRef = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".mid(letter,1) + columnRef
            remaining = ((remaining - letter) / 26) - 1
        }
        return (fixed ? "$" : "") + columnRef
    }
    
    private func setupFormats() {
        formatInt = workbook_add_format(workbook)
        format_set_num_format(formatInt, "0;-0;")
        formatFloat = workbook_add_format(workbook)
        format_set_num_format(formatFloat, "0.00;-0.00;")
        formatZeroInt = workbook_add_format(workbook)
        format_set_num_format(formatInt, "0")
        formatZeroFloat = workbook_add_format(workbook)
        format_set_num_format(formatZeroFloat, "0.00")
        formatDate = workbook_add_format(workbook)
        format_set_num_format(formatDate, "dd/MM/yyyy")
        formatPercent = workbook_add_format(workbook)
        format_set_num_format(formatPercent, "0.00%")
        formatBold = workbook_add_format(workbook)
        format_set_bold(formatBold)
        formatBold = workbook_add_format(workbook)
        format_set_bold(formatBold)
        formatRightBold = workbook_add_format(workbook)
        format_set_bold(formatRightBold)
        format_set_align(formatRightBold, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_num_format(formatRightBold, "0;-0;")
        formatBoldUnderline = workbook_add_format(workbook)
        format_set_bold(formatBoldUnderline)
        format_set_bottom(formatBoldUnderline, UInt8(LXW_BORDER_THIN.rawValue))
        formatRightBoldUnderline = workbook_add_format(workbook)
        format_set_bold(formatRightBoldUnderline)
        format_set_align(formatRightBoldUnderline, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_bottom(formatRightBoldUnderline, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_num_format(formatRightBoldUnderline, "0;-0;")
        formatFloatBoldUnderline = workbook_add_format(workbook)
        format_set_bold(formatFloatBoldUnderline)
        format_set_align(formatFloatBoldUnderline, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_bottom(formatFloatBoldUnderline, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_num_format(formatFloatBoldUnderline, "0.00;-0.00;")
    }
    
    // MARK: - Helper routines
    
    func write(cellType: CellType, worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, content: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        switch cellType {
        case .string:
            write(worksheet: worksheet, row: row, column: column, string: content, format: format)
        case .integer:
            if let integer = Int(content) {
                write(worksheet: worksheet, row: row, column: column, integer: integer, format: format)
            } else {
                write(worksheet: worksheet, row: row, column: column, string: content, format: format)
            }
        case .float:
            if let float = Float(content) {
                write(worksheet: worksheet, row: row, column: column, float: float, format: format)
            } else {
                write(worksheet: worksheet, row: row, column: column, string: content, format: format)
            }
        case .stringFormula:
            write(worksheet: worksheet, row: row, column: column, formula: content, format: format)
        case .integerFormula:
            write(worksheet: worksheet, row: row, column: column, integerFormula: content, format: format)
        case .floatFormula:
            write(worksheet: worksheet, row: row, column: column, floatFormula: content, format: format)
        default:
            break
        }
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, string: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        worksheet_write_string(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), string, format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, integer: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatInt
        worksheet_write_number(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), Double(integer), format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, float: Float, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatFloat
        worksheet_write_number(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), Double(float), format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, formula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        worksheet_write_formula(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), formula, format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, date: Date, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatDate
        worksheet_write_unixtime(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), Int64(date.timeIntervalSince1970), format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, integerFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatInt
        worksheet_write_formula(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), integerFormula, format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, floatFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatFloat
        worksheet_write_formula(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), floatFormula, format)
    }
    
    func setColumn(worksheet: UnsafeMutablePointer<lxw_worksheet>?, column: Int, width: Float = 0, hidden: Bool = false, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let column = lxw_col_t(Int32(column))
        var options = lxw_row_col_options()
        options.hidden = hidden ? 1 : 0
        worksheet_set_column_opt(worksheet, column, column, Double(width), format, &options)
    }
    
    func setRow(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let row = lxw_row_t(Int32(row))
        worksheet_set_row(worksheet, row,  LXW_DEF_ROW_HEIGHT, format)
    }
}
