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
    var playerNumber: Int?
    var cellType: CellType
    var width: Float
    
    init(title: String, content: ((Participant, Int)->String)? = nil, playerContent: ((Participant, Player, Int, Int)->String)? = nil, playerNumber: Int? = nil, referenceContent: ((Int)->Int)? = nil, referenceDivisor: Int? = nil, referenceCalculated: ((Int)->String)? = nil, cellType: CellType = .string, width: Float = 10.0) {
        self.title = title
        self.content = content
        self.playerContent = playerContent
        self.playerNumber = playerNumber
        self.referenceContent = referenceContent
        self.referenceCalculated = referenceCalculated
        self.referenceDivisor = referenceDivisor
        self.cellType = cellType
        self.width = width
    }
}

class Writer {
    let scoreData: ScoreData
    var ranksPlusMpsColumns: [Column] = []
    let ranksPlusMPsName = "ranksPlusMPs"
    var individualMpsColumns: [Column] = []
    let individualMPsName = "individualMPs"
    var positionColumn: Int?
    var directionColumn: Int?
    var participantNoColumn: Int?
    var individualMPsPositionColumn: Int?
    var individualMPsDecimalColumn: Int?
    var individualMPsNationalIdColumn: Int?
    var individualMPsTotalColumn: Int?
    var individualMPsChecksumColumn: Int?
    var firstNameColumn: [Int] = []
    var otherNameColumn: [Int] = []
    var nationalIdColumn: [Int] = []
    var boardsPlayedColumn: [Int] = []
    var winDrawColumn: [Int] = []
    var bonusMPColumn: [Int] = []
    var winDrawMPColumn: [Int] = []
    var totalMPColumn: [Int] = []
    var maxAwardCell: String?
    var maxEwAwardCell: String?
    var awardToCell: String?
    var ewAwardToCell: String?
    var perWinCell: String?
    var eventDateCell: String?
    var eventIdCell: String?
    var nationalLocalCell: String?
    let ranksPlusMPsHeaderRows = 3
    var workbook: UnsafeMutablePointer<lxw_workbook>?
    var ranksPlusMpsWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    var individualMpsWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    var formatInt: UnsafeMutablePointer<lxw_format>?
    var formatFloat: UnsafeMutablePointer<lxw_format>?
    var formatBold: UnsafeMutablePointer<lxw_format>?
    var formatRightBold: UnsafeMutablePointer<lxw_format>?
    var formatPercent: UnsafeMutablePointer<lxw_format>?
    var formatTopLine: UnsafeMutablePointer<lxw_format>?
    var formatDate: UnsafeMutablePointer<lxw_format>?
    var fieldSize: Int? = nil
    var nsPairs: Int? = nil

   init(scoreData: ScoreData) {
       self.scoreData = scoreData
       setupRanksPlusMpsColumns()
       
       let name = "\(scoreData.fileUrl!.deletingPathExtension().lastPathComponent.removingPercentEncoding!).xlsm"
       workbook = workbook_new(name)
       individualMpsWorksheet = workbook_add_worksheet(workbook, individualMPsName)
       ranksPlusMpsWorksheet = workbook_add_worksheet(workbook, ranksPlusMPsName)
       workbook_add_vba_project(workbook, "./Award.bin")
       formatInt = workbook_add_format(workbook)
       format_set_num_format(formatInt, "0;-0;")
       formatFloat = workbook_add_format(workbook)
       format_set_num_format(formatFloat, "0.00;-0.00;")
       formatDate = workbook_add_format(workbook)
       format_set_num_format(formatDate, "dd/MM/yyyy")
       formatPercent = workbook_add_format(workbook)
       format_set_num_format(formatPercent, "#.00%")
       formatBold = workbook_add_format(workbook)
       format_set_bold(formatBold)
       formatBold = workbook_add_format(workbook)
       format_set_bold(formatBold)
       formatRightBold = workbook_add_format(workbook)
       format_set_bold(formatRightBold)
       format_set_align(formatRightBold, UInt8(LXW_ALIGN_RIGHT.rawValue))
       formatTopLine = workbook_add_format(workbook)
       format_set_top(formatTopLine, UInt8(LXW_BORDER_THIN.rawValue))
    }
    
    func write() {
        writeRanksPlusMPs()
        writeIndividualMPs()
        workbook_close(workbook)
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
        positionColumn = ranksPlusMpsColumns.count - 1

        if event.winnerType == 2 && event.type?.participantType == .pair {
            ranksPlusMpsColumns.append(Column(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
            directionColumn = ranksPlusMpsColumns.count - 1
        }
        
        ranksPlusMpsColumns.append(Column(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer)) ; participantNoColumn = ranksPlusMpsColumns.count - 1

        ranksPlusMpsColumns.append(Column(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float))
        
        if winDraw && maxPlayers <= event.type?.participantType?.players ?? maxPlayers {
            ranksPlusMpsColumns.append(Column(title: "Win/Draw", content: { (participant, _) in "\(participant.winDraw!)" }, cellType: .float))
            winDrawColumn.append(ranksPlusMpsColumns.count - 1)
        }

        for playerNumber in 0..<playerCount {
            ranksPlusMpsColumns.append(Column(title: "First Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 0) }, playerNumber: playerNumber, cellType: .string))
            firstNameColumn.append(ranksPlusMpsColumns.count - 1)

            ranksPlusMpsColumns.append(Column(title: "Other Names (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 1) }, playerNumber: playerNumber, cellType: .string))
            otherNameColumn.append(ranksPlusMpsColumns.count - 1)

            ranksPlusMpsColumns.append(Column(title: "SBU No (\(playerNumber+1))", playerContent: { (_, player,_, _) in player.nationalId! }, playerNumber: playerNumber, cellType: .integer))
            nationalIdColumn.append(ranksPlusMpsColumns.count - 1)
            
            if maxPlayers > event.type?.participantType?.players ?? maxPlayers {
                
                ranksPlusMpsColumns.append(Column(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.boardsPlayed)" }, playerNumber: playerNumber, cellType: .integer))
                boardsPlayedColumn.append(ranksPlusMpsColumns.count - 1)
                
                if winDraw {
                    ranksPlusMpsColumns.append(Column(title: "Win/Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                    winDrawColumn.append(ranksPlusMpsColumns.count - 1)
                }
                
                ranksPlusMpsColumns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP (\(playerNumber+1))", playerContent: playerBonusAward, playerNumber: playerNumber, cellType: .integerFormula))
                bonusMPColumn.append(ranksPlusMpsColumns.count - 1)
                
                if winDraw {
                    ranksPlusMpsColumns.append(Column(title: "Win/Draw MP (\(playerNumber+1))", playerContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .integerFormula))
                    winDrawMPColumn.append(ranksPlusMpsColumns.count - 1)
                    
                    ranksPlusMpsColumns.append(Column(title: "Total MP (\(playerNumber+1))", playerContent: playerTotalAward, playerNumber: playerNumber, cellType: .integerFormula))
                    totalMPColumn.append(ranksPlusMpsColumns.count - 1)
                }
            }
        }
            
        if maxPlayers <= event.type?.participantType?.players ?? maxPlayers {
            
            ranksPlusMpsColumns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP", content: bonusAward, cellType: .integerFormula))
            bonusMPColumn.append(ranksPlusMpsColumns.count - 1)
            
            if winDraw {
                
                ranksPlusMpsColumns.append(Column(title: "Win/Draw MP", content: winDrawAward, cellType: .integerFormula))
                winDrawMPColumn.append(ranksPlusMpsColumns.count - 1)
                
                ranksPlusMpsColumns.append(Column(title: "Total MP", content: totalAward, cellType: .integerFormula))
                for _ in 0..<maxPlayers {
                    totalMPColumn.append(ranksPlusMpsColumns.count - 1)
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
        var useAwardCell = maxAwardCell
        var useAwardToCell = awardToCell
        var firstRow = ranksPlusMPsHeaderRows + 1
        var lastRow = ranksPlusMPsHeaderRows + event.participants.count
        if let nsPairs = nsPairs {
            if event.winnerType == 2 && participant.type == .pair {
                if let pair = participant.member as? Pair {
                    if pair.direction == .ew {
                        useAwardCell = maxEwAwardCell
                        useAwardToCell = ewAwardToCell
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
            result += "ROUNDUP((\(cell(rowNumber, boardsPlayedColumn[playerNumber], columnFixed: true))/\(totalBoardsPlayed))*"
        }
        
        let positionCell = cell(rowNumber, positionColumn!)
        let allPositionsRange = "\(cell(firstRow, rowFixed: true, positionColumn!, columnFixed: true)):\(cell(lastRow, rowFixed: true, positionColumn!, columnFixed: true))"
        
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
        
        let winDrawsCell = cell(rowNumber, winDrawColumn[playerNumber ?? 0], columnFixed: true)
        result = "=ROUNDUP(\(winDrawsCell) * \(perWinCell!), 0)"
        
        return result
    }
    
    private func totalAward(participant: Participant, rowNumber: Int) -> String {
        return playerTotalAward(participant: participant, rowNumber: rowNumber)
    }

    private func playerTotalAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let bonusMPCell = cell(rowNumber, bonusMPColumn[playerNumber ?? 0], columnFixed: true)
        let winDrawMPCell = cell(rowNumber, winDrawMPColumn[playerNumber ?? 0], columnFixed: true)
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
        
        writeCell(string: "Event Code", format: formatBold)
        
        writeCell(string: "Local/Nat", format: formatBold)
        
        column = -1
        row = 1
        
        writeCell(string: event.description ?? "")
        
        writeCell(integer: toe) ; let toeRef = cell(row, rowFixed: true, column, columnFixed: true)
        
        var toeCell = toeRef
        var ewToeRef = ""
        if twoWinners {
            writeCell(integer: fieldSize! - nsPairs!) ; ewToeRef = cell(row, rowFixed: true, column, columnFixed: true)
            toeCell += "+" + ewToeRef
        }
        writeCell(formula: "=ROUNDUP(\(toeCell)*(\(event.type?.participantType?.players ?? 4)/4),0)")
        
        writeCell(integer: event.boards ?? 0)
        
        var baseMaxEwAwardCell = ""
        writeCell(integer: 1000) ; let baseMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(integerFormula: "ROUNDUP(\(baseMaxAwardCell)*\(ewToeRef)/\(toeRef),0)") ; baseMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(integer: 1, format: formatPercent) ; let factorCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(baseMaxAwardCell)*\(factorCell),0)") ; maxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(baseMaxEwAwardCell)*\(factorCell),0)") ; maxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 0.25, format: formatPercent) ; let awardPercentCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(toeRef)*\(awardPercentCell),0)") ; awardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(ewToeRef)*\(awardPercentCell),0)") ; ewAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(integer: 25) ; perWinCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(date: event.date ?? Date.today) ; eventDateCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.eventId ?? "") ; eventIdCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: "Local") ; nationalLocalCell = cell(row, rowFixed: true, column, columnFixed: true)
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
        }
        
        let totalColumnRef = columnRef(column: individualMPsDecimalColumn!)
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsTotalColumn!, floatFormula: "=SUM(\(totalColumnRef):\(totalColumnRef))")
        
        let nationalIdColumnRef = columnRef(column: individualMPsNationalIdColumn!)
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsChecksumColumn!, floatFormula: "=SUMPRODUCT(\(totalColumnRef):\(totalColumnRef),\(nationalIdColumnRef):\(nationalIdColumnRef))")

    }
    
    private func setupIndividualMpsColumns() {
        let event = scoreData.events.first!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        
        individualMpsColumns.append(Column(title: "Place", referenceContent: { (_) in self.positionColumn! }, cellType: .integerFormula)) ; individualMPsPositionColumn = individualMpsColumns.count - 1
        
        if twoWinners {
            individualMpsColumns.append(Column(title: "Direction", referenceContent: { (_) in self.directionColumn! }, cellType: .stringFormula))
        }
        
        individualMpsColumns.append(Column(title: "@P no", referenceContent: { (_) in self.participantNoColumn! }, cellType: .integerFormula))
        
        individualMpsColumns.append(Column(title: "Total MPs", referenceContent: { (playerNumber) in self.totalMPColumn[playerNumber] }, referenceDivisor: 100, cellType: .floatFormula)) ; individualMPsDecimalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Names", referenceContent: { (playerNumber) in self.firstNameColumn[playerNumber] }, cellType: .stringFormula))
        
        individualMpsColumns.append(Column(title: "", referenceContent: { (playerNumber) in self.otherNameColumn[playerNumber] }, cellType: .stringFormula))
        
        individualMpsColumns.append(Column(title: "Event Date", referenceCalculated: { (_) in "\(self.ranksPlusMPsName)!\(self.eventDateCell!)" }, cellType: .dataFormula))

        individualMpsColumns.append(Column(title: "MemNo", referenceContent: { (playerNumber) in self.nationalIdColumn[playerNumber] }, cellType: .integerFormula)) ; individualMPsNationalIdColumn = individualMpsColumns.count - 1
                                    
        individualMpsColumns.append(Column(title: "Event Code", referenceCalculated: { (_) in "\(self.ranksPlusMPsName)!\(self.eventIdCell!)" }, cellType: .stringFormula))
        
        individualMpsColumns.append(Column(title: "Club Code"))

        individualMpsColumns.append(Column(title: "Local MPs", referenceCalculated: { (rowNumber) in "IF(\(self.ranksPlusMPsName)!\(self.nationalLocalCell!) = \"National\",0,\(self.cell(rowNumber, self.individualMPsDecimalColumn!, columnFixed: true)))" }, cellType: .floatFormula))
        
        individualMpsColumns.append(Column(title: "National MPs", referenceCalculated: { (rowNumber) in "IF(\(self.ranksPlusMPsName)!\(self.nationalLocalCell!) <> \"National\",0,\(self.cell(rowNumber, self.individualMPsDecimalColumn!, columnFixed: true)))" }, cellType: .floatFormula))
        
        individualMpsColumns.append(Column(title: "Total")) ; individualMPsTotalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Checksum")) ; individualMPsChecksumColumn = individualMpsColumns.count - 1
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
            result += "\(ranksPlusMPsName)!" + cell(ranksPlusMPsHeaderRows + 1, totalMPColumn[playerNumber])
            result += ":"
            result += cell(ranksPlusMPsHeaderRows + fieldSize!, totalMPColumn[playerNumber])
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
    
    func setColumn(worksheet: UnsafeMutablePointer<lxw_worksheet>?, column: Int, width: Float = 0, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let column = lxw_col_t(Int32(column))
        worksheet_set_column(worksheet, column, column, Double(width), format)
    }
    
    func setRow(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let row = lxw_row_t(Int32(row))
        worksheet_set_row(worksheet, row,  LXW_DEF_ROW_HEIGHT, format)
    }
}
