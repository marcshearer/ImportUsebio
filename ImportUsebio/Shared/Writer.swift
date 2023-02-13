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
    var content: ((Participant, Int)->(String))?
    var playerContent: ((Participant, Player, Int, Int)->(String))?
    var playerNumber: Int?
    var cellType: CellType
    var width: Float
    
    init(title: String, content: ((Participant, Int)->(String))? = nil, playerContent: ((Participant, Player, Int, Int)->(String))? = nil, playerNumber: Int? = nil, cellType: CellType = .string, width: Float = 10.0) {
        self.title = title
        self.content = content
        self.playerContent = playerContent
        self.playerNumber = playerNumber
        self.cellType = cellType
        self.width = width
    }
}

class Writer {
    let scoreData: ScoreData
    var columns: [Column] = []
    var positionColumn: Int?
    var nameColumn: [Int] = []
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
    let headerRows = 3
    var workbook: UnsafeMutablePointer<lxw_workbook>?
    var formatInt: UnsafeMutablePointer<lxw_format>?
    var formatFloat: UnsafeMutablePointer<lxw_format>?
    var formatBold: UnsafeMutablePointer<lxw_format>?
    var formatRightBold: UnsafeMutablePointer<lxw_format>?
    var formatPercent: UnsafeMutablePointer<lxw_format>?
    var formatTopLine: UnsafeMutablePointer<lxw_format>?
    var fieldSize: Int? = nil
    var nsPairs: Int? = nil

   init(scoreData: ScoreData) {
       self.scoreData = scoreData
       setupColumns()
       
       let name = "\(scoreData.fileUrl!.deletingPathExtension().lastPathComponent.removingPercentEncoding!).xlsm"
       workbook = workbook_new(name)
       formatInt = workbook_add_format(workbook)
       format_set_num_format(formatInt, "0;-0;")
       formatFloat = workbook_add_format(workbook)
       format_set_num_format(formatFloat, "0.00;-0.00;")
       formatPercent = workbook_add_format(workbook)
       format_set_num_format(formatPercent, "#.00%")
       formatBold = workbook_add_format(workbook)
       format_set_bold(formatBold)
       formatBold = workbook_add_format(workbook)
       format_set_bold(formatBold)
       formatRightBold = workbook_add_format(workbook)
       format_set_bold(formatRightBold)
       format_set_align(formatRightBold, UInt8(LXW_ALIGN_RIGHT.rawValue))
       workbook_add_vba_project(workbook, "./Award.bin")
       formatTopLine = workbook_add_format(workbook)
       format_set_top(formatTopLine, UInt8(LXW_BORDER_THIN.rawValue))
    }
    
    func write() {
        let event = scoreData.events.first!
        let participants = scoreData.events.first!.participants.sorted(by: sortCriteria)
        fieldSize = participants.count
        if event.winnerType == 2 && event.type?.participantType == .pair {
            nsPairs = participants.filter{(($0.member as? Pair)?.direction ?? .ns) == .ns}.count
        }
        
        let worksheet1 = workbook_add_worksheet(workbook, "Ranks plus MPs")

        writeHeader(worksheet: worksheet1)
        
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: worksheet1, column: columnNumber, width: column.width)
            write(worksheet: worksheet1, row: headerRows, column: columnNumber, string: replace(column.title), format: format)
        }
        
        for (rowSequence, participant) in participants.enumerated() {
            
            let rowNumber = rowSequence + headerRows + 1

            for (columnNumber, column) in columns.enumerated() {
                
                if let content = column.content?(participant, rowNumber) {
                    write(cellType: (content == "" ? .string : column.cellType), worksheet: worksheet1, row: rowNumber, column: columnNumber, content: content)
                }
                
                if let playerNumber = column.playerNumber {
                    let playerList = participant.member.playerList
                    if playerNumber < playerList.count {
                        if let playerContent = column.playerContent?(participant, playerList[playerNumber], playerNumber, rowNumber) {
                            write(cellType: (playerContent == "" ? .string : column.cellType), worksheet: worksheet1, row: rowNumber, column: columnNumber, content: playerContent)
                        }
                    } else {
                        write(cellType: .string, worksheet: worksheet1, row: rowNumber, column: columnNumber, content: "")
                    }
                }
            }
        }
        
        workbook_close(workbook)
    }
    
    private func setupColumns() {
        let event = scoreData.events.first!
        let playerCount = maxPlayers
        let winDraw = event.type?.requiresWinDraw ?? false
        
        columns.append(Column(title: "Place", content: { (participant, _) in "\(participant.place!)" }, cellType: .integer))
        positionColumn = columns.count - 1

        if event.winnerType == 2 && event.type?.participantType == .pair {
            columns.append(Column(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
        }
        
        columns.append(Column(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer))

        columns.append(Column(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float))
        
        if winDraw && maxPlayers <= event.type?.participantType?.players ?? maxPlayers {
            columns.append(Column(title: "Win/Draw", content: { (participant, _) in "\(participant.winDraw!)" }, cellType: .float))
            winDrawColumn.append(columns.count - 1)
        }

        for playerNumber in 0..<playerCount {
            columns.append(Column(title: "Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in player.name! }, playerNumber: playerNumber, cellType: .string))
            nameColumn.append(columns.count - 1)
            
            columns.append(Column(title: "SBU No (\(playerNumber+1))", playerContent: { (_, player,_, _) in player.nationalId! }, playerNumber: playerNumber, cellType: .integer))
            nationalIdColumn.append(columns.count - 1)
            
            if maxPlayers > event.type?.participantType?.players ?? maxPlayers {
                
                columns.append(Column(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.boardsPlayed)" }, playerNumber: playerNumber, cellType: .integer))
                boardsPlayedColumn.append(columns.count - 1)
                
                if winDraw {
                    columns.append(Column(title: "Win/Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                    winDrawColumn.append(columns.count - 1)
                }
                
                columns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP (\(playerNumber+1))", playerContent: playerBonusAward, playerNumber: playerNumber, cellType: .integerFormula))
                bonusMPColumn.append(columns.count - 1)
                
                if winDraw {
                    columns.append(Column(title: "Win/Draw MP (\(playerNumber+1))", playerContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .integerFormula))
                    winDrawMPColumn.append(columns.count - 1)
                    
                    columns.append(Column(title: "Total MP (\(playerNumber+1))", playerContent: playerTotalAward, playerNumber: playerNumber, cellType: .integerFormula))
                    totalMPColumn.append(columns.count - 1)
                }
            }
        }
            
        if maxPlayers <= event.type?.participantType?.players ?? maxPlayers {
            
            columns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP", content: bonusAward, cellType: .integerFormula))
            bonusMPColumn.append(columns.count - 1)
            
            if winDraw {
                
                columns.append(Column(title: "Win/Draw MP", content: winDrawAward, cellType: .integerFormula))
                winDrawMPColumn.append(columns.count - 1)
                
                columns.append(Column(title: "Total MP", content: totalAward, cellType: .integerFormula))
                totalMPColumn.append(columns.count - 1)
            }
        }
    }
    
    // MARK: - Content getters
    
    private func bonusAward(participant: Participant, rowNumber: Int) -> String {
        return playerBonusAward(participant: participant, rowNumber: rowNumber)
    }
    
    private func playerBonusAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "="
        let event = scoreData.events.first!
        var useAwardCell = maxAwardCell
        var useAwardToCell = awardToCell
        var firstRow = headerRows + 1
        var lastRow = headerRows + event.participants.count
        if let nsPairs = nsPairs {
            if event.winnerType == 2 && participant.type == .pair {
                if let pair = participant.member as? Pair {
                    if pair.direction == .ew {
                        useAwardCell = maxEwAwardCell
                        useAwardToCell = ewAwardToCell
                        firstRow = headerRows + nsPairs + 1
                    } else {
                        lastRow = headerRows + nsPairs
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
    
    // MARK: - Heading
    
    func writeHeader(worksheet: UnsafeMutablePointer<lxw_worksheet>?) {
        let event = scoreData.events.first!
        var column = -1
        var row = 0
        var prefix = ""
        var toe = fieldSize!
        if event.winnerType == 2 && event.type?.participantType == .pair {
            prefix = "NS "
            toe = nsPairs!
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "Event", format: formatBold)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "\(prefix)TOE", format: formatRightBold)
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "EW TOE", format: formatRightBold)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "Tables", format: formatRightBold)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "Boards", format: formatRightBold)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "\(prefix)Max Award", format: formatRightBold)
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "EW Max Award", format: formatRightBold)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "Factor%", format: formatRightBold)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "\(prefix)Award", format: formatRightBold)
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "EW Award", format: formatRightBold)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "Award %", format: formatRightBold)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "\(prefix)Award to", format: formatRightBold)
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "EW Award to", format: formatRightBold)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: "Per Win", format: formatRightBold)
        
        column = -1
        row = 1
        column += 1 ; write(worksheet: worksheet, row: row, column: column, string: event.description ?? "")
        column += 1 ; write(worksheet: worksheet, row: row, column: column, integer: toe) ; let toeRef = cell(row, rowFixed: true, column, columnFixed: true)
        var toeCell = toeRef
        var ewToeRef = ""
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, integer: fieldSize! - nsPairs!) ; ewToeRef = cell(row, rowFixed: true, column, columnFixed: true)
            toeCell += "+" + ewToeRef
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, formula: "=ROUNDUP(\(toeCell)*(\(event.type?.participantType?.players ?? 4)/4),0)")
        column += 1 ; write(worksheet: worksheet, row: row, column: column, integer: event.boards ?? 0)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, integer: 1000) ; let baseMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        var baseMaxEwAwardCell = ""
        if event.winnerType == 2 && event.type?.participantType == .pair {
            let ewPairs = fieldSize! - nsPairs!
            column += 1 ; write(worksheet: worksheet, row: row, column: column, integerFormula: "ROUNDUP(\(baseMaxAwardCell)*\(ewToeRef)/\(toeRef),0)") ; baseMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, integer: 1, format: formatPercent) ; let factorCell = cell(row, rowFixed: true, column, columnFixed: true)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, formula: "=ROUNDUP(\(baseMaxAwardCell)*\(factorCell),0)") ; maxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, formula: "=ROUNDUP(\(baseMaxEwAwardCell)*\(factorCell),0)") ; maxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, float: 0.25, format: formatPercent) ; let awardPercentCell = cell(row, rowFixed: true, column, columnFixed: true)
        column += 1 ; write(worksheet: worksheet, row: row, column: column, formula: "=ROUNDUP(\(toeRef)*\(awardPercentCell),0)") ; awardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        if event.winnerType == 2 && event.type?.participantType == .pair {
            column += 1 ; write(worksheet: worksheet, row: row, column: column, formula: "=ROUNDUP(\(ewToeRef)*\(awardPercentCell),0)") ; ewAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        column += 1 ; write(worksheet: worksheet, row: row, column: column, integer: 25) ; perWinCell = cell(row, rowFixed: true, column, columnFixed: true)
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
        worksheet_set_row(worksheet, row, 0, format)
    }
}
