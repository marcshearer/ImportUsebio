//
//  Writer.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 08/02/2023.
//

import xlsxwriter
import SwiftUI

enum CellType {
    case string
    case integer
    case float
    case date
    case numeric
    case stringFormula
    case integerFormula
    case floatFormula
    case numericFormula
    
    var isFormula: Bool {
        return self == .stringFormula || self == .integerFormula || self == .floatFormula || self == .numericFormula
    }
}

class Column {
    var title: String
    var content: ((Participant, Int)->String)?
    var playerContent: ((Participant, Player, Int, Int)->String)?
    var calculatedContent: ((Int?, Int) -> String)?
    var referenceContent: ((Int)->Int)?
    var referenceDynamic: (()->String)?
    var playerNumber: Int?
    var cellType: CellType
    var format: UnsafeMutablePointer<lxw_format>?
    var width: Float?
    
    init(title: String, content: ((Participant, Int)->String)? = nil, playerContent: ((Participant, Player, Int, Int)->String)? = nil, calculatedContent: ((Int?, Int) -> String)? = nil, playerNumber: Int? = nil, referenceContent: ((Int)->Int)? = nil, referenceDynamic: (()->String)? = nil, cellType: CellType = .string, format: UnsafeMutablePointer<lxw_format>? = nil, width: Float? = nil) {
        self.title = title
        self.content = content
        self.playerContent = playerContent
        self.calculatedContent = calculatedContent
        self.playerNumber = playerNumber
        self.referenceContent = referenceContent
        self.referenceDynamic = referenceDynamic
        self.cellType = cellType
        self.format = format
        self.width = width
    }
}

class Round {
    var writer: Writer
    var scoreData: ScoreData
    var name: String
    var shortName: String
    var toe: Int?
    var individualMPs: IndividualMPsWriter!
    var ranksPlusMps: RanksPlusMPsWriter!
    var workbook: UnsafeMutablePointer<lxw_workbook>?
    
    init(writer: Writer, name: String, shortName: String? = nil, scoreData: ScoreData, workbook: UnsafeMutablePointer<lxw_workbook>?) {
        self.writer = writer
        self.name = name
        self.shortName = shortName ?? name.left(16)
        self.scoreData = scoreData
        self.individualMPs = IndividualMPsWriter(writer: writer, round: self, scoreData: scoreData)
        self.ranksPlusMps = RanksPlusMPsWriter(writer: writer, round: self, scoreData: scoreData)
    }
    
    func replace(_ text: String) -> String {
        var text = text
        text = text.replacingOccurrences(of: "@P", with: scoreData.events.first!.type!.participantType!.string)
        return text
    }
    
    var maxParticipantPlayers: Int {
        var result = 0
        
        let event = scoreData.events.first!
        if event.type?.participantType == .team {
            result = event.participants.map{$0.member.playerList.count}.max() ?? 0
        } else {
            result = event.type?.participantType?.players ?? 2
        }
        if let maxTeamMembers = scoreData.maxTeamMembers {
            result = max(result, maxTeamMembers)
        }
        
        return result
    }
    
    var fieldSize: Int { Settings.current.largestFieldSize }
    
    var maxPlayers: Int { Settings.current.largestPlayerCount }
    
    var nsPairs : Int? {
        let event = scoreData.events.first!
        if event.winnerType == 2 && event.type?.participantType == .pair {
            return event.participants.filter{(($0.member as? Pair)?.direction ?? .ns) == .ns}.count
        } else {
            return nil
        }
    }
    
}

class Writer: WriterBase {
    var parameters: ParametersWriter!
    var summary: SummaryWriter!
    var csvImport:CsvImportWriter!
    var consolidated: ConsolidatedWriter!
    var missing: MissingNumbersWriter!
    var formatted: FormattedWriter!
    var rounds: [Round] = []
    var minRank = 0
    var maxRank = 9999
    var eventCode: String = ""
    var eventDescription: String = ""
    var missingNumbers: [String: (NationalId: String, Nbo: String)] = [:]
    
    var maxPlayers: Int { min(1000, rounds.map{ $0.maxPlayers }.reduce(0, +)) }
    
    init() {
        super.init()
        parameters = ParametersWriter(writer: self)
        summary = SummaryWriter(writer: self)
        csvImport = CsvImportWriter(writer: self)
        consolidated = ConsolidatedWriter(writer: self)
        missing = MissingNumbersWriter(writer: self)
        formatted = FormattedWriter(writer: self)
    }
        
    @discardableResult func add(name: String, shortName: String? = nil, scoreData: ScoreData) -> Round {
        if scoreData.source == .usebio {
            // Need to recalculate the wins/draws in case rounding mode changed
            UsebioParser.calculateWinDraw(scoreData: scoreData)
        }
        let round = Round(writer: self, name: name, shortName: shortName, scoreData: scoreData, workbook: workbook)
        rounds.append(round)
        return round
    }
    
    func write(as filename: String) {
        // Create workbook
        self.workbook = workbook_new(filename)
        setupFormats()
        // Create worksheets
        summary.prepare(workbook: workbook)
        csvImport.prepare(workbook: workbook)
        formatted.prepare(workbook: workbook)
        consolidated.prepare(workbook: workbook)
        missing.prepare(workbook: workbook)
        for round in rounds {
            round.ranksPlusMps.prepare(workbook: workbook)
        }
        for round in rounds {
            round.individualMPs.prepare(workbook: workbook)
        }
        parameters.prepare(workbook: workbook)
        workbook_add_vba_project(workbook, "./Award.bin")
        
        // Process data
        for round in rounds {
            round.ranksPlusMps.write()
            round.individualMPs.write()
        }
        consolidated.write()
        parameters.writeSortBy()
        csvImport.write()
        summary.write()
        missing.write()
        if missingNumbers.count <= 0 {
            worksheet_hide(missing.worksheet)
        }
        formatted.write()
        parameters.write()
        worksheet_hide(parameters.worksheet)
        
        // Finish up
        workbook_close(workbook)
    }
    
    fileprivate var maxEventDate: String {
        var formula = "=MAX("
        for (roundNumber, round) in rounds.enumerated() {
            let source = round.ranksPlusMps!
            if roundNumber != 0 {
                formula += ", "
            }
            formula += cell(writer: source, source.eventDateCell!)
        }
        formula += ")"
        return formula
    }
}

// MARK: Parameters Writer

class ParametersWriter : WriterBase {
    override var name: String { "Parameters" }
    let sortNameColumn = 0
    let sortAddressColumn = 1
    let sortDirectionColumn = 2
   
    let headerRow = 0
    let dataRow = 1
    
    var sortByNameRange: String!
    var sortByAddressRange: String!
    var sortByDirectionRange: String!
    
    var sortData: [(name: String, column: Int, direction: Int)] = []
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        freezePanes(worksheet: worksheet, row: dataRow, column: 0)
    }
    
    func writeSortBy() {
        let consolidated = writer.consolidated!
        
        sortData =
        [("National Id",  consolidated.nationalIdColumn!, 1),
         ("Local MPs",    consolidated.localMPsColumn!, -1),
         ("National MPs", consolidated.nationalMPsColumn!, -1),
         ("First Name", consolidated.firstNameColumn!, 1),
         ("Other Names", consolidated.otherNamesColumn!, 1)]
        
        for column in 0...sortDirectionColumn {
                setColumn(worksheet: worksheet, column: column, width: 17)
        }
        
        for (column, header) in ["Sort Text", "Sort Address", "Sort Direction"].enumerated() {
            write(worksheet: worksheet, row: headerRow, column: column, string: header, format: formatBold)
        }
        
        for (row,element) in sortData.enumerated() {
            write(worksheet: worksheet, row: dataRow + row, column: sortNameColumn, string: element.name)
            let arrayAddress = cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, element.column, columnFixed: true)
            write(worksheet: worksheet, row: dataRow + row, column: sortAddressColumn, formula: "=ADDRESS(ROW(\(arrayAddress)),COLUMN(\(arrayAddress)),1,1,\"\(consolidated.name)\")")
            write(worksheet: worksheet, row: dataRow + row, column: sortDirectionColumn, integer: element.direction)
        }
        
        sortByNameRange = "\(cell(writer: self, dataRow, rowFixed: true, sortNameColumn, columnFixed: true)):\(cell(dataRow + sortData.count - 1, rowFixed: true, sortNameColumn, columnFixed: true))"
        sortByAddressRange = "\(cell(writer: self, dataRow, rowFixed: true, sortAddressColumn, columnFixed: true)):\(cell(dataRow + sortData.count - 1, rowFixed: true, sortAddressColumn, columnFixed: true))"
        sortByDirectionRange = "\(cell(writer: self, dataRow, rowFixed: true, sortDirectionColumn, columnFixed: true)):\(cell(dataRow + sortData.count - 1, rowFixed: true, sortDirectionColumn, columnFixed: true))"
    }
}


// MARK: - Summary
    
class SummaryWriter : WriterBase {
    override var name: String { "Summary" }
    var entryColumn: Int?
    var tablesColumn: Int?
    var descriptionColumn: Int?
    var localNationalColumn: Int?
    var localMPsColumn: Int?
    var nationalMPsColumn: Int?
    var checksumColumn: Int?
    var exportedRow: Int?
    var headerRow: Int?
    var detailRow: Int?
    var totalRow: Int?
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        descriptionColumn = 0
        entryColumn = 1
        tablesColumn = 2
        localNationalColumn = 3
        localMPsColumn = 4
        nationalMPsColumn = 5
        checksumColumn = 6
        headerRow = 0
        detailRow = 1
        totalRow = detailRow! + writer.rounds.count + 1
        exportedRow = totalRow! + 1
               
        freezePanes(worksheet: worksheet, row: detailRow!, column: 0)
        
        setColumn(worksheet: worksheet, column: descriptionColumn!, width: 30)
        setColumn(worksheet: worksheet, column: localMPsColumn!, width: 12)
        setColumn(worksheet: worksheet, column: nationalMPsColumn!, width: 12)
        setColumn(worksheet: worksheet, column: checksumColumn!, width: 16)
        
        write(worksheet: worksheet, row: headerRow!, column: descriptionColumn!, string: "Round", format: formatBold)
        write(worksheet: worksheet, row: headerRow!, column: entryColumn!, string: "Entry", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: tablesColumn!, string: "Tables", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: localNationalColumn!, string: "Nat/Local", format: formatBold)
        write(worksheet: worksheet, row: headerRow!, column: localMPsColumn!, string: "Local MPs", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: nationalMPsColumn!, string: "National MPs", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: checksumColumn!, string: "Checksum", format: formatRightBold)
        
        for (roundNumber, round) in writer.rounds.enumerated() {
            let source = round.ranksPlusMps!
            let row = detailRow! + roundNumber
            write(worksheet: worksheet, row: row, column: descriptionColumn!, string: "\(round.name)")
            write(worksheet: worksheet, row: row, column: entryColumn!, integerFormula: "=\(cell(writer: source, source.entryCell!))")
            write(worksheet: worksheet, row: row, column: tablesColumn!, integerFormula: "=\(cell(writer: source, source.tablesCell!))")
            write(worksheet: worksheet, row: row, column: localNationalColumn!, formula: "=\(cell(writer: source, source.localCell!))")
            write(worksheet: worksheet, row: row, column: localMPsColumn!, floatFormula: "=\(cell(writer: source, source.localMPsCell!))", format: formatZeroFloat)
            write(worksheet: worksheet, row: row, column: nationalMPsColumn!, floatFormula: "=\(cell(writer: source, source.nationalMPsCell!))", format: formatZeroFloat)
            write(worksheet: worksheet, row: row, column: checksumColumn!, floatFormula: "=\(cell(writer: source, source.checksumCell!))", format: formatZeroFloat)
        }
        
        for column in [entryColumn, tablesColumn, localNationalColumn, localMPsColumn, nationalMPsColumn, checksumColumn] {
            write(worksheet: worksheet, row: detailRow! + writer.rounds.count, column: column!, string: "", format: formatBoldUnderline)
        }
        
        write(worksheet: worksheet, row: totalRow!, column: descriptionColumn!, string: "Round totals", format: formatBold)
        
        writeTotal(column: tablesColumn, format: formatInt)
        writeTotal(column: localMPsColumn)
        writeTotal(column: nationalMPsColumn)
        writeTotal(column: checksumColumn)
        
        let csvImport = writer.csvImport!
        write(worksheet: worksheet, row: exportedRow!, column: descriptionColumn!, string: "Exported totals", format: formatBold)
        write(worksheet: worksheet, row: exportedRow!, column: localMPsColumn!, formula: "=\(cell(writer: csvImport, csvImport.localMpsCell!))", format: formatZeroFloat)
        write(worksheet: worksheet, row: exportedRow!, column: nationalMPsColumn!, formula: "=\(cell(writer: csvImport, csvImport.nationalMpsCell!))", format: formatZeroFloat)
        write(worksheet: worksheet, row: exportedRow!, column: checksumColumn!, formula: "=\(cell(writer: csvImport, csvImport.checksumCell!))", format: formatZeroFloat)
        
        for column in [localMPsColumn, nationalMPsColumn, checksumColumn] {
            highlightTotalDifferent(row: exportedRow!, compareRow: totalRow!, column: column!)
        }
        
        // Add macro buttons to summary
        writer.createMacroButton(worksheet: worksheet, title: "Create PDF", macro: "PrintFormatted", row: 1, column: checksumColumn! + 1)
        writer.createMacroButton(worksheet: worksheet, title: "Select Formatted", macro: "SelectFormatted", row: 4, column: checksumColumn! + 1)
        
    }
    
    private func writeTotal(column: Int?, format: UnsafeMutablePointer<lxw_format>? = nil) {
        write(worksheet: worksheet, row: totalRow!, column: column!, integerFormula: "=ROUND(SUM(\(cell(detailRow!, rowFixed: true, column!)):\(cell(detailRow! + writer.rounds.count - 1, rowFixed: true, column!))),2)", format: format ?? formatZeroFloat)
    }
    
    private func highlightTotalDifferent(row: Int, compareRow: Int, column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let field = "\(cell(row, rowFixed: true, column))"
        let matchField = "\(cell(compareRow, rowFixed: true, column))"
        let formula = "\(field)<>\(matchField)"
        setConditionalFormat(worksheet: worksheet, fromRow: row, fromColumn: column, toRow: row, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}

// MARK: - Formatted

class FormattedWriter: WriterBase {
    override var name: String { "Formatted" }
    let titleRow = 0
    let dataRow = 1
    let nameColumn = 0
    let roundsColumn = 1
    var columns: [Column] = []
    var leftRightColumns: [String] = []
    var directionColumn: Int?
    
    var formatBannerString: UnsafeMutablePointer<lxw_format>?
    var formatBannerFloat: UnsafeMutablePointer<lxw_format>?
    var formatBannerNumeric: UnsafeMutablePointer<lxw_format>?
    var formatBodyString: UnsafeMutablePointer<lxw_format>?
    var formatBodyFloat: UnsafeMutablePointer<lxw_format>?
    var formatBodyNumeric: UnsafeMutablePointer<lxw_format>?
    var formatBottom: UnsafeMutablePointer<lxw_format>?
    var formatLeftRight: UnsafeMutablePointer<lxw_format>?
    var formatLeftRightBottom: UnsafeMutablePointer<lxw_format>?
    
    var singleEvent: Bool {
        let rounds = writer.rounds
        return (rounds.count == 1)
    }
    
    var twoWinners: Bool {
        let rounds = writer.rounds
        let round = rounds.first!
        let event = round.scoreData.events.first!
        return (event.winnerType == 2 && event.type?.participantType == .pair)
    }
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        
        setupFormattedFormats()
        setupColumns()
        
        // Define ranges
        workbook_define_name(writer.workbook, "FormattedNameArray", "=\(arrayRef)(\(cell(writer: self, dataRow, rowFixed: true, nameColumn, columnFixed: true)))")
        workbook_define_name(writer.workbook, "FormattedTitleRow", "=\(cell(writer: self, titleRow, rowFixed: true, 0, columnFixed: true)):\(cell(titleRow, rowFixed: true, columns.count - 1, columnFixed: true))")
        workbook_define_name(writer.workbook, "Printing", "=false")
        
        worksheet_set_default_row(worksheet, 25, 0)
        setRow(worksheet: worksheet, row: titleRow, height: 50)
        worksheet_fit_to_pages(worksheet, 1, 0)
        freezePanes(worksheet: worksheet, row: dataRow, column: 0)
        var rows: [lxw_row_t] = []
        for row in 1...((writer.maxPlayers / Settings.current.linesPerFormattedPage)) {
            rows.append(lxw_row_t(row*Settings.current.linesPerFormattedPage + 1))
            worksheet_set_h_pagebreaks(worksheet, &rows)
        }
        worksheet_repeat_rows(worksheet, 0, 0)
        worksheet_set_header(worksheet, "&C&14\(writer.eventDescription) - Master Point Allocations")
        worksheet_set_footer(worksheet,"&RPage &P of &N")
        
        for (columnNumber, column) in columns.enumerated() {
            var bannerFormat = formatBannerString
            var bodyFormat = formatBodyString
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                bannerFormat = formatBannerFloat
                bodyFormat = formatBodyFloat
            } else if column.cellType == .numeric || column.cellType == .numericFormula {
                bannerFormat = formatBannerNumeric
                bodyFormat = formatBodyNumeric
            }
            if let width = column.width {
                setColumn(worksheet: worksheet, column: columnNumber, width: width)
            }
            write(worksheet: worksheet, row: titleRow, column: columnNumber, string: column.title, format: bannerFormat)
            
            if singleEvent {
                if let referenceDynamic = column.referenceDynamic {
                    writeDynamicReference(rowNumber: dataRow, columnNumber: columnNumber, content: { referenceDynamic() }, format: column.format ?? bodyFormat)
                }
            } else {
                let consolidated = writer.consolidated!
                if let referenceDynamic = column.referenceDynamic {
                    let output = "=\(sortBy)(\(referenceDynamic()),\(arrayRef)(\(cell(writer: writer.consolidated, consolidated.dataRow!, rowFixed: true, consolidated.localMPsColumn!, columnFixed: true)))+\(arrayRef)(\(cell(writer: writer.consolidated, consolidated.dataRow!, rowFixed: true, consolidated.nationalMPsColumn!, columnFixed: true))),-1)"
                    writeDynamicReference(rowNumber: dataRow, columnNumber: columnNumber, content: {output}, format: column.format ?? bodyFormat)
                }
            }
        }
        
        var leftRightFormula = "AND($A2<>\"\",OR("
        for (columnNumber, column) in leftRightColumns.enumerated() {
            if columnNumber != 0 {
                leftRightFormula += ","
            }
            leftRightFormula += "A$1=\"\(column)\""
        }
        leftRightFormula += "))"
        
        var bottomFormula = "OR(AND($A3=\"\",$A2<>\"\"),AND(Printing,MOD(ROW($A2), \(Settings.current.linesPerFormattedPage ?? 32))=1)"
        if singleEvent && twoWinners{
            let columnRef = columnRef(directionColumn!, fixed: true)
            bottomFormula += ",AND($A2<>\"\",\(columnRef)2<>\(columnRef)3))"
        } else {
            bottomFormula += ")"
        }
        
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: "AND(\(leftRightFormula),\(bottomFormula))", stopIfTrue: true, format: formatLeftRightBottom!)
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: leftRightFormula, stopIfTrue: true, format: formatLeftRight!)
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: bottomFormula, stopIfTrue: true, format: formatBottom!)
    }
    
    private func setupColumns() {
        let consolidated = writer.consolidated!
        let rounds = writer.rounds
        let round = rounds.first!
        let event = round.scoreData.events.first!
        let winDraw = event.type?.requiresWinDraw ?? false
        var local = false
        var national = false
            
        if (singleEvent) {
            let round = writer.rounds.first!
            let ranksPlusMPs = round.ranksPlusMps!
            
            columns.append(Column(title: "Position", referenceDynamic: { [self] in "=\(ranksArrayRef(arrayContent: sourceRef(column: ranksPlusMPs.positionColumn!)))" }, cellType: .numericFormula, width: 9))
            leftRightColumns.append(columns.last!.title)
            
            if twoWinners {
                columns.append(Column(title: "Direction", referenceDynamic: { [self] in "=\(ranksArrayRef(arrayContent: sourceRef(column: ranksPlusMPs.directionColumn!)))" }, cellType: .numericFormula, width: 9))
                directionColumn = columns.count - 1
            }
            leftRightColumns.append(columns.last!.title)
            
            columns.append(Column(title: (round.maxParticipantPlayers == event.type!.participantType!.players ? "Names" : "Names           *Awards for team members will vary by boards played"), referenceDynamic: { [self] in ranksNamesRef() }, cellType: .string, width: Float(round.maxParticipantPlayers) * (18.0 - Float(round.maxParticipantPlayers))))
            
            columns.append(Column(title: "Score", referenceDynamic: { [self] in "=\(ranksArrayRef(arrayContent: sourceRef(column: ranksPlusMPs.scoreColumn!)))" }, cellType: .numericFormula))
            
            if winDraw {
                columns.append(Column(title: "Wins / Draws", referenceDynamic: { [self] in "=\(ranksArrayRef(arrayContent: sourceRef(column: ranksPlusMPs.winDrawColumn[0])))" }, cellType: .numericFormula))
            }

            columns.append(Column(title: "\(round.scoreData.national ? "National" : "Local") MPs", referenceDynamic: { [self] in "=\(ranksArrayRef(arrayContent: sourceRef(column: ranksPlusMPs.totalMPColumn[0])))" }, cellType: .floatFormula, width: 10))
            leftRightColumns.append(columns.last!.title)
            
        } else {
            
            columns.append(Column(title: "Name", referenceDynamic: { [self] in "CONCATENATE(\(consolidatedArrayRef(column: consolidated.firstNameColumn!)),\" \",\(consolidatedArrayRef(column: consolidated.otherNamesColumn!)))" }, cellType: .stringFormula, width: 30))
            leftRightColumns.append(columns.last!.title)
            
            for (roundNumber, round) in rounds.enumerated() {
                columns.append(Column(title: round.shortName.replacingOccurrences(of: " ", with: "\n"), referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.dataColumn! + roundNumber))" }, cellType: .floatFormula, width: 10))
                if round.scoreData.national {
                    national = true
                } else {
                    local = true
                }
            }
            
            if local {
                columns.append(Column(title: "Local MPs", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.localMPsColumn!))" }, cellType: .floatFormula, width: 10))
                leftRightColumns.append(columns.last!.title)
            }
            if national {
                columns.append(Column(title: "National MPs", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.nationalMPsColumn!))" }, cellType: .floatFormula, width: 10))
                leftRightColumns.append(columns.last!.title)
                
            }
            if local && national {
                columns.append(Column(title: "Total MPs", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.localMPsColumn!))+\(consolidatedArrayRef(column: consolidated.nationalMPsColumn!))" }, cellType: .floatFormula, width: 10))
                leftRightColumns.append(columns.last!.title)
            }
        }
    }
    
    private func consolidatedArrayRef(column: Int) -> String {
        return "\(arrayRef)(\(cell(writer: writer.consolidated, writer.consolidated.dataRow!, rowFixed: true, column, columnFixed: true)))"
    }
    
    private func ranksArrayRef(arrayContent: String) -> String {
        let ranksPlusMps = writer.rounds.first!.ranksPlusMps!
        
        var result = "\(sortBy)(\(zeroFiltered(arrayContent: arrayContent)), "
        
        if twoWinners {
            let direction = zeroFiltered(arrayContent: sourceRef(column: ranksPlusMps.directionColumn!))
            result += "\(direction), -1, "
        }
        
        let position = zeroFiltered(arrayContent: sourceRef(column: ranksPlusMps.positionColumn!))
        result += "\(position), 1)"
        
        return result
    }
    
    private func ranksNamesRef() -> String {
        let round = writer.rounds.first!
        let ranksPlusMps = round.ranksPlusMps!
        
        var arrayContent = "FormatNames("
        for playerNumber in 0..<round.maxParticipantPlayers {
            if playerNumber != 0 {
                arrayContent += ","
            }
            arrayContent += sourceRef(column: ranksPlusMps.firstNameColumn[playerNumber])
            arrayContent += ","
            arrayContent += sourceRef(column: ranksPlusMps.otherNamesColumn[playerNumber])
        }
        arrayContent += ")"
        
        return ranksArrayRef(arrayContent: arrayContent)
    }
    
    private func zeroFiltered(arrayContent: String) -> String {
        let round = writer.rounds.first!
        
        var content = "\(filter)(\(vstack)("
        
        content += arrayContent
        
        content += "),\(vstack)("
        
        content += sourceRef(column: round.ranksPlusMps.totalMPColumn.first!)
        content += ")<>0)"
        
        return content
    }
    
    private func sourceRef(column: Int) -> String {
        let round = writer.rounds.first!
        var content = ""
        content += cell(writer: round.ranksPlusMps, round.ranksPlusMps.dataRow, rowFixed: true, column, columnFixed: true)
        content += ":"
        content += cell(round.ranksPlusMps.dataRow + round.fieldSize - 1, rowFixed: true, column, columnFixed: true)
        return content
    }
    
    private func names(_ participant: Participant) -> String {
        var result = ""
        let list = participant.member.playerList
        for (playerNumber, player) in list.enumerated() {
            result += player.name!
            if playerNumber != list.count - 1 {
                if playerNumber != list.count - 2 {
                    result += ", "
                } else {
                    result += " & "
                }
            }
        }
        return result
    }
    
    func setupFormattedFormats() {
        formatBannerString = workbook_add_format(workbook)
        format_set_align(formatBannerString, UInt8(LXW_ALIGN_LEFT.rawValue))
        format_set_align(formatBannerString, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_bold(formatBannerString)
        format_set_pattern (formatBannerString, UInt8(LXW_PATTERN_SOLID.rawValue))
        format_set_bg_color(formatBannerString, UInt32(0x1F036C))
        format_set_font_color(formatBannerString, LXW_COLOR_WHITE.rawValue)
        format_set_text_wrap(formatBannerString)
        format_set_indent(formatBannerString, 2)
        
        formatBannerFloat = workbook_add_format(workbook)
        format_set_align(formatBannerFloat, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBannerFloat, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_bold(formatBannerFloat)
        format_set_pattern (formatBannerFloat, UInt8(LXW_PATTERN_SOLID.rawValue))
        format_set_bg_color(formatBannerFloat, UInt32(0x1F036C))
        format_set_font_color(formatBannerFloat, LXW_COLOR_WHITE.rawValue)
        format_set_text_wrap(formatBannerFloat)
        
        formatBannerNumeric = formatBannerFloat
        
        formatBodyString = workbook_add_format(workbook)
        format_set_align(formatBodyString, UInt8(LXW_ALIGN_LEFT.rawValue))
        format_set_align(formatBodyString, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_indent(formatBodyString, 2)
        
        formatBodyFloat = workbook_add_format(workbook)
        format_set_align(formatBodyFloat, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBodyFloat, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_num_format(formatBodyFloat, "0.00;-0.00;")
        
        formatBodyNumeric = workbook_add_format(workbook)
        format_set_align(formatBodyNumeric, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBodyNumeric, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))

        formatLeftRight = workbook_add_format(workbook)
        format_set_left(formatLeftRight, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_right(formatLeftRight, UInt8(LXW_BORDER_THIN.rawValue))
        
        formatLeftRightBottom = workbook_add_format(workbook)
        format_set_left(formatLeftRightBottom, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_right(formatLeftRightBottom, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_bottom(formatLeftRightBottom, UInt8(LXW_BORDER_THIN.rawValue))
        
        formatBottom = workbook_add_format(workbook)
        format_set_bottom(formatBottom, UInt8(LXW_BORDER_THIN.rawValue))
    }
    
    func colorValue(_ color: Color) -> UInt32 {
        let color = NSColor(color)
        return UInt32((((color.redComponent * 256) + color.greenComponent) * 256) + color.blueComponent)
    }
}

// MARK: - Export CSV

class CsvImportWriter: WriterBase {
    override var name: String { "Import" }
    var localMpsCell: String?
    var nationalMpsCell: String?
    var checksumCell: String?
    
    let eventDescriptionRow = 0
    let eventCodeRow = 1
    let minRankRow = 2
    let maxRankRow = 3
    let eventDateRow = 4
    let clubCodeRow = 5
    let sortByRow = 7
    let awardsRow = 8
    let localMPsRow = 9
    let nationalMPsRow = 10
    let checksumRow = 11
    
    let titleRow = 13
    let dataRow = 14
    
    let titleColumn = 0
    let valuesColumn = 1
    let firstNameColumn = 0
    let otherNamesColumn = 1
    let eventDateColumn = 2
    let nationalIdColumn = 3
    let eventCodeColumn = 4
    let clubCodeColumn = 5
    let localMPsColumn = 6
    let nationalMPsColumn = 7
    let lookupFirstNameColumn = 10
    let lookupOtherNamesColumn = 11
    let lookupOtherUnionColumn = 12
    let lookupHomeClubColumn = 13
    let lookupRankColumn = 14
    let lookupEmailColumn = 15
    let lookupStatusColumn = 16
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        // Define ranges
        workbook_define_name(writer.workbook, "ImportDateArray", "=\(arrayRef)(\(cell(writer: self, dataRow, rowFixed: true, eventDateColumn, columnFixed: true)))")
        workbook_define_name(writer.workbook, "ImportTitleRow", "=\(cell(writer: self, titleRow, rowFixed: true, eventDateColumn, columnFixed: true)):\(cell(titleRow, rowFixed: true, nationalMPsColumn, columnFixed: true))")
        
        freezePanes(worksheet: worksheet, row: dataRow, column: 0)

        // Add macro button
        writer.createMacroButton(worksheet: worksheet, title: "Copy Import", macro: "CopyImport", row: 1, column: 4)
        
        // Format rows/columns
        setRow(worksheet: worksheet, row: titleRow, height: 30)
        
        setColumn(worksheet: worksheet, column: titleColumn, width: 12)
        setColumn(worksheet: worksheet, column: valuesColumn, width: 16)
        setColumn(worksheet: worksheet, column: eventDateColumn, width: 10, format: formatDate)
        setColumn(worksheet: worksheet, column: nationalIdColumn, width: 13, format: formatInt)
        setColumn(worksheet: worksheet, column: localMPsColumn, format: formatFloat)
        setColumn(worksheet: worksheet, column: nationalMPsColumn, format: formatFloat)
        
        setColumn(worksheet: worksheet, column: lookupFirstNameColumn, width: 12)
        setColumn(worksheet: worksheet, column: lookupOtherNamesColumn, width: 12)
        setColumn(worksheet: worksheet, column: lookupOtherUnionColumn, width: 12, format: formatZeroBlank)
        setColumn(worksheet: worksheet, column: lookupHomeClubColumn, width: 25)
        setColumn(worksheet: worksheet, column: lookupRankColumn, width: 8)
        setColumn(worksheet: worksheet, column: lookupEmailColumn, width: 25)
        setColumn(worksheet: worksheet, column: lookupStatusColumn, width: 25)
        
        // Parameters etc
        write(worksheet: worksheet, row: eventDescriptionRow, column: titleColumn, string: "Event:", format: formatBold)
        write(worksheet: worksheet, row: eventDescriptionRow, column: valuesColumn, string: writer.eventDescription)
        
        write(worksheet: worksheet, row: eventCodeRow, column: titleColumn, string: "Event Code:", format: formatBold)
        write(worksheet: worksheet, row: eventCodeRow, column: valuesColumn, string: writer.eventCode)
        
        write(worksheet: worksheet, row: minRankRow, column: titleColumn, string: "Minimum Rank:", format: formatBold)
        write(worksheet: worksheet, row: minRankRow, column: valuesColumn, formula: "=\(writer.minRank)", format: formatString)
        
        write(worksheet: worksheet, row: maxRankRow, column: titleColumn, string: "Maximum Rank:", format: formatBold)
        write(worksheet: worksheet, row: maxRankRow, column: valuesColumn, formula: "=\(writer.maxRank)", format: formatString)
        
        write(worksheet: worksheet, row: eventDateRow, column: titleColumn, string: "Event Date:", format: formatBold)
        write(worksheet: worksheet, row: eventDateRow, column: valuesColumn, floatFormula: "=\(writer.maxEventDate)", format: formatDate)
        
        write(worksheet: worksheet, row: clubCodeRow, column: titleColumn, string: "Club Code:", format: formatBold)
        write(worksheet: worksheet, row: clubCodeRow, column: valuesColumn, string: "", format: formatString)
        
        write(worksheet: worksheet, row: sortByRow, column: titleColumn, string: "Sort by:", format: formatBold)
        let parameters = writer.parameters!
        let sortData = parameters.sortData
        let validationRange = "=\(cell(writer: parameters, parameters.dataRow, rowFixed: true, parameters.sortNameColumn, columnFixed: true)):\(cell(parameters.dataRow + sortData.count - 1, rowFixed: true, parameters.sortNameColumn, columnFixed: true))"
        setDataValidation(row: sortByRow, column: valuesColumn, formula: validationRange)
        write(worksheet: worksheet, row: sortByRow, column: valuesColumn, string: sortData.first!.name, format: formatInt)

        write(worksheet: worksheet, row: awardsRow, column: titleColumn, string: "Award count:", format: formatBold)
        write(worksheet: worksheet, row: awardsRow, column: valuesColumn, dynamicFormula: "=ROWS(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn, columnFixed: true))))", format: formatZeroInt)
        
        write(worksheet: worksheet, row: localMPsRow, column: titleColumn, string: "Local MPs:", format: formatBold)
        write(worksheet: worksheet, row: localMPsRow, column: valuesColumn, dynamicFormula: "=ROUND(SUM(\(arrayRef)(\(cell(dataRow, rowFixed: true, localMPsColumn, columnFixed: true)))),2)", format: formatZeroFloat)
        localMpsCell = cell(localMPsRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: worksheet, row: nationalMPsRow, column: titleColumn, string: "National MPs:", format: formatBold)
        write(worksheet: worksheet, row: nationalMPsRow, column: valuesColumn, dynamicFormula: "=ROUND(SUM(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalMPsColumn, columnFixed: true)))),2)", format: formatZeroFloat)
        nationalMpsCell = cell(nationalMPsRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: worksheet, row: checksumRow, column: titleColumn, string: "Checksum:", format: formatBold)
        write(worksheet: worksheet, row: checksumRow, column: valuesColumn, dynamicFormula: "=CheckSum(\(vstack)(\(arrayRef)(\(cell(dataRow, rowFixed: true, localMPsColumn, columnFixed: true)))+\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalMPsColumn, columnFixed: true)))),\(vstack)(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn, columnFixed: true)))))", format: formatZeroFloat)
        checksumCell = cell(checksumRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        // Data
        let consolidated = writer.consolidated!
        let sortByCell = cell(sortByRow, rowFixed: true, valuesColumn, columnFixed: true)
        let sortByLogic = ",\(fnPrefix)INDIRECT(\(fnPrefix)XLOOKUP(\(sortByCell),\(parameters.sortByNameRange!),\(parameters.sortByAddressRange!),,0)&\"#\"),\(fnPrefix)XLOOKUP(\(sortByCell),\(parameters.sortByNameRange!),\(parameters.sortByDirectionRange!),,0)"
        
        write(worksheet: worksheet, row: titleRow, column: firstNameColumn, string: "Names", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow, column: firstNameColumn, dynamicFormula: "=\(sortBy)(\(sourceArray(consolidated.firstNameColumn!))\(sortByLogic))")
        
        write(worksheet: worksheet, row: titleRow, column: otherNamesColumn, string: "", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow, column: otherNamesColumn, dynamicFormula: "=\(sortBy)(\(sourceArray(consolidated.otherNamesColumn!))\(sortByLogic))")
        
        write(worksheet: worksheet, row: titleRow, column: eventDateColumn, string: "Claim Date", format: formatBoldUnderline)
        let firstNameCell = cell(dataRow, rowFixed: true, firstNameColumn, columnFixed: true)
        write(worksheet: worksheet, row: dataRow, column: eventDateColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(firstNameCell)), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\", \"\", \(cell(eventDateRow, rowFixed: true, valuesColumn, columnFixed: true)))))", format: formatDate)
        
        write(worksheet: worksheet, row: titleRow, column: nationalIdColumn, string: "Membership ID", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: dataRow, column: nationalIdColumn, dynamicFormula: "=\(sortBy)(\(sourceArray(consolidated.nationalIdColumn!))\(sortByLogic))", format: formatInt)
        
        write(worksheet: worksheet, row: titleRow, column: eventCodeColumn, string: "Event Code", format: formatBoldUnderline)
        let eventCell = cell(eventCodeRow, rowFixed: true, valuesColumn, columnFixed: true)
        write(worksheet: worksheet, row: dataRow, column: eventCodeColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(cell(dataRow, rowFixed: true, firstNameColumn, columnFixed: true))), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\", \"\", IF(\(eventCell)=0,\"\",\(eventCell)))))")
        
        write(worksheet: worksheet, row: titleRow, column: clubCodeColumn, string: "Club Code", format: formatBoldUnderline)
        let clubCell = cell(clubCodeRow, rowFixed: true, valuesColumn, columnFixed: true)
        write(worksheet: worksheet, row: dataRow, column: clubCodeColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(firstNameCell)), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\",\"\", IF(\(clubCell)=0,\"\",\(clubCell)))))")
        
        write(worksheet: worksheet, row: titleRow, column: localMPsColumn, string: "Local Points", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: dataRow, column: localMPsColumn, dynamicFormula: "=\(sortBy)(\(sourceArray(consolidated.localMPsColumn!))\(sortByLogic))", format: formatFloat)
        
        write(worksheet: worksheet, row: titleRow, column: nationalMPsColumn, string: "National Points", format: formatRightBoldUnderline)  
        write(worksheet: worksheet, row: dataRow, column: nationalMPsColumn, dynamicFormula: "=\(sortBy)(\(sourceArray(consolidated.nationalMPsColumn!))\(sortByLogic))", format: formatFloat)
        
        //Lookups
        writeLookup(title: "First Name", column: lookupFirstNameColumn, lookupColumn: Settings.current.userDownloadFirstNameColumn)
        writeLookup(title: "Other Names", column: lookupOtherNamesColumn, lookupColumn: Settings.current.userDownloadOtherNamesColumn)
        writeLookup(title: "Other Union", column: lookupOtherUnionColumn, lookupColumn: Settings.current.userDownloadOtherUnionColumn)
        writeLookup(title: "Home Club", column: lookupHomeClubColumn, lookupColumn: Settings.current.userDownloadHomeClubColumn)
        writeLookup(title: "Rank", column: lookupRankColumn, lookupColumn: Settings.current.userDownloadRankColumn, format: formatRightBoldUnderline, numeric: true)
        writeLookup(title: "Email", column: lookupEmailColumn, lookupColumn: Settings.current.userDownloadEmailColumn)
        writeLookup(title: "Status", column: lookupStatusColumn, lookupColumn: Settings.current.userDownloadStatusColumn)
        
        highlightLookupDifferent(column: firstNameColumn, lookupColumn: lookupFirstNameColumn, format: formatYellow)
        highlightLookupDifferent(column: otherNamesColumn, lookupColumn: lookupOtherNamesColumn)
        highlightLookupError(fromColumn: lookupFirstNameColumn, toColumn: lookupStatusColumn, format: formatGrey)
        highlightBadNationalId(column: nationalIdColumn, firstNameColumn: firstNameColumn)
        highlightBadDate(column: eventDateColumn, firstNameColumn: firstNameColumn)
        highlightBadMPs(column: localMPsColumn, firstNameColumn: firstNameColumn)
        highlightBadMPs(column: nationalMPsColumn, firstNameColumn: firstNameColumn)
        highlightBadRank(column: lookupRankColumn)
        highlightBadStatus(column: lookupStatusColumn)
    }
    
    private func sourceArray(_ column: Int) -> String {
        let consolidated = writer.consolidated!
        return "\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, column, columnFixed: true)))"
    }
    
    private func writeLookup(title: String, column: Int, lookupColumn: String, format: UnsafeMutablePointer<lxw_format>? = nil, numeric: Bool = false) {
        write(worksheet: worksheet, row: titleRow, column: column, string: title, format: format ?? formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow, column: column, dynamicFormula: "=\(numeric ? "\(fnPrefix)NUMBERVALUE(" : "")\(fnPrefix)XLOOKUP(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn, columnFixed: true))),\(vstack)(\(lookupRange(Settings.current.userDownloadNationalIdColumn))),\(vstack)(\(lookupRange(lookupColumn))),,0)\(numeric ? ")" : "")")
    }
    
    private func lookupRange(_ column: String) -> String {
        return "'\(Settings.current.userDownloadData!)'!\(column)\(Settings.current.userDownloadMinRow!):\(column)\(Settings.current.userDownloadMaxRow!)"
    }
    
    private func highlightLookupDifferent(column: Int, lookupColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let field = "\(cell(dataRow, column, columnFixed: true))"
        let lookupfield = "\(cell(dataRow, lookupColumn, columnFixed: true))"
        let formula = "\(field)<>\(lookupfield)"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightLookupError(fromColumn: Int, toColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let lookupCell = "\(cell(dataRow, fromColumn, columnFixed: true))"
        let formula = "=ISNA(\(lookupCell))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: fromColumn, toRow: dataRow + writer.maxPlayers - 1, toColumn: toColumn, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadStatus(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let statusCell = "\(cell(dataRow, column, columnFixed: true))"
        let formula = "=AND(\(statusCell)<>\"\(Settings.current.goodStatus!)\", \(statusCell)<>\"\")"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadNationalId(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let nationalIdCell = cell(dataRow, column, columnFixed: true)
        let firstNameCell = cell(dataRow, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(nationalIdCell)<=0, \(nationalIdCell)>\(Settings.current.maxNationalIdNumber!), NOT(ISNUMBER(\(nationalIdCell)))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, checkDuplicates: true, format: format ?? formatRed!, duplicateFormat: formatRedHatched!)
    }
    
    private func highlightBadMPs(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let pointsCell = cell(dataRow, column, columnFixed: true)
        let firstNameCell = cell(dataRow, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(pointsCell)>\(Settings.current.maxPoints!), AND(\(pointsCell)<>\"\", NOT(ISNUMBER(\(pointsCell))))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadDate(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let dateCell = cell(dataRow, column, columnFixed: true)
        let firstNameCell = cell(dataRow, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(dateCell)>DATEVALUE(\"\(Date().toString(format: "dd/MM/yyyy"))\")+1, \(dateCell)<DATEVALUE(\"\(Date().toString(format: "dd/MM/yyyy"))\")-30))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadRank(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let rankCell = cell(dataRow, column, columnFixed: true)
        let formula = "=AND(\(rankCell)<>\"\", OR(\(rankCell)<\(cell(minRankRow, rowFixed: true, valuesColumn, columnFixed: true)), \(rankCell)>\(cell(maxRankRow, rowFixed: true, valuesColumn, columnFixed: true))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}
    
// MARK: - Consolidated
 
class ConsolidatedWriter: WriterBase {
    override var name: String { "Consolidated" }
    var dataRow: Int?
    var firstNameColumn: Int?
    var otherNamesColumn: Int?
    var nationalIdColumn: Int?
    var dataColumn: Int?
    var localMPsColumn: Int?
    var nationalMPsColumn: Int?
    
    init(writer: Writer) {
        super.init(writer: writer)
    }
    
    func write() {
        let uniqueColumn = 0
        firstNameColumn = 1
        otherNamesColumn = 2
        nationalIdColumn = 3
        localMPsColumn = 4
        nationalMPsColumn = 5
        dataColumn = 6
        
        let titleRow = 3
        let localNationalRow = 0
        let totalRow = 1
        let checksumRow = 2
        dataRow = 4
        
        freezePanes(worksheet: worksheet, row: dataRow!, column: dataColumn!)
        
        let localNationalRange = "\(cell(localNationalRow, rowFixed: true, dataColumn!, columnFixed: true)):\(cell(localNationalRow, rowFixed: true, dataColumn! + writer.rounds.count - 1, columnFixed: true))"
        
        setRow(worksheet: worksheet, row: titleRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: localNationalRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: totalRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: checksumRow, format: formatBoldUnderline)
        
        for column in 0..<(dataColumn! + writer.rounds.count) {
            if column == uniqueColumn {
                setColumn(worksheet: worksheet, column: column, hidden: true)
            } else if column == nationalIdColumn {
                setColumn(worksheet: worksheet, column: column, format: formatInt)
            } else if column == firstNameColumn || column == otherNamesColumn {
                setColumn(worksheet: worksheet, column: column, width: 16)
            } else {
                setColumn(worksheet: worksheet, column: column, width: 12, format: formatFloat)
            }
        }
        
        // Title row
        write(worksheet: worksheet, row: titleRow, column: firstNameColumn!, string: "First Name", format: formatBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: otherNamesColumn!, string: "Other Names", format: formatBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: nationalIdColumn!, string: "SBU", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: localMPsColumn!, string: "Local MPs", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: nationalMPsColumn!, string: "National MPs", format: formatRightBoldUnderline)
        
        for (column, round) in writer.rounds.enumerated() {
            // Round titles
            write(worksheet: worksheet, row: titleRow, column: (dataColumn!) + column, string: round.shortName, format: formatRightBoldUnderline)
        
            // National/Local row
            let cell = cell(writer: round.ranksPlusMps, round.ranksPlusMps.localCell!)
            write(worksheet: worksheet, row: localNationalRow, column: dataColumn! + column, formula: "=IF(\(cell)=0,\"\",\(cell))", format: formatRightBoldUnderline)
        }
        
        for element in 0...1 {
            // Total row and checksum row
            let row = (element == 0 ? totalRow : checksumRow)
            let totalRange = "\(cell(row, rowFixed: true, dataColumn!, columnFixed: true)):\(cell(row, rowFixed: true, dataColumn! + writer.rounds.count - 1, columnFixed: true))"
            
            write(worksheet: worksheet, row: row, column: nationalIdColumn!, string: (row == totalRow ? "Total" : "Checksum"), format: formatRightBoldUnderline)
            
            write(worksheet: worksheet, row: row, column: localMPsColumn!, floatFormula: "=SUMIF(\(localNationalRange), \"<>National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            write(worksheet: worksheet, row: row, column: nationalMPsColumn!, floatFormula: "=SUMIF(\(localNationalRange), \"=National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            for (column, _) in writer.rounds.enumerated() {
                let valueRange = "\(arrayRef)(\(cell(dataRow!, rowFixed: true, dataColumn! + column)))"
                let nationalIdRange = "\(arrayRef)(\(cell(dataRow!, rowFixed: true, nationalIdColumn!, columnFixed: true)))"
                
                if row == totalRow {
                    write(worksheet: worksheet, row: row, column: dataColumn! + column, floatFormula: "=ROUND(SUM(\(valueRange)),2)", format: formatFloatBoldUnderline)
                } else {
                    write(worksheet: worksheet, row: row, column: dataColumn! + column, floatFormula: "=CheckSum(\(vstack)(\(valueRange)),\(vstack)(\(nationalIdRange)))", format: formatFloatBoldUnderline)
                }
            }
        }
        
        // Data rows
        let uniqueIdCell = "\(arrayRef)(\(cell(dataRow!, rowFixed: true, uniqueColumn, columnFixed: true)))"
        
        // Unique ID column
        var formula = "=\(unique)(\(vstack)("
        for (roundNumber, round) in writer.rounds.enumerated() {
            if roundNumber != 0 {
                formula += ","
            }
            formula += "\(arrayRef)(\(cell(writer: round.individualMPs, 1,rowFixed: true, round.individualMPs.uniqueColumn!, columnFixed: true)))"
        }
        formula += "))"
        write(worksheet: worksheet, row: dataRow!, column: uniqueColumn, dynamicFormula: formula)
        
        // Name columns
        write(worksheet: worksheet, row: dataRow!, column: firstNameColumn!, dynamicFormula: "=IF(\(uniqueIdCell)=\"\",\"\",\(fnPrefix)TEXTBEFORE(\(fnPrefix)TEXTAFTER(\(uniqueIdCell),\"+\"),\"+\"))")
        write(worksheet: worksheet, row: dataRow!, column: otherNamesColumn!, dynamicFormula: "=IF(\(uniqueIdCell)=\"\",\"\",\(fnPrefix)TEXTAFTER(\(uniqueIdCell), \"+\", 2))")
        
        // National ID column
        write(worksheet: worksheet, row: dataRow!, column: nationalIdColumn!, dynamicIntegerFormula: "=IF(\(uniqueIdCell)=\"\",\"\",IFERROR(\(fnPrefix)NUMBERVALUE(\(fnPrefix)TEXTBEFORE(\(uniqueIdCell), \"+\")),\(fnPrefix)TEXTBEFORE(\(uniqueIdCell), \"+\")))")
        
        // Total local/national columns
        let dataRange = "\(arrayRef)(\(cell(dataRow!, dataColumn!, columnFixed: true))):\(arrayRef)(\(cell(dataRow!, rowFixed: true, dataColumn! + writer.rounds.count - 1, columnFixed: true)))"
        
        write(worksheet: worksheet, row: dataRow!, column: localMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SUMIF(\(localNationalRange), \"<>National\", \(lambdaParam))))")
        write(worksheet: worksheet, row: dataRow!, column: nationalMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SUMIF(\(localNationalRange), \"=National\", \(lambdaParam))))")
        
        // Lookup data columns
        for (column, round) in writer.rounds.enumerated() {
            let indivUniqueArray = "\(arrayRef)(\(cell(writer: round.individualMPs,round.individualMPs.dataRow, rowFixed: true, round.individualMPs.uniqueColumn!, columnFixed: true)))"
            let indivValueArray = "\(arrayRef)(\(cell(writer: round.individualMPs,round.individualMPs.dataRow, rowFixed: true, round.individualMPs.decimalColumn!, columnFixed: true)))"
            
            write(worksheet: worksheet, row: dataRow!, column: dataColumn! + column, dynamicFloatFormula: "=\(byRow)(\(uniqueIdCell),\(lambda)(\(lambdaParam),SUMIF(\(indivUniqueArray),\(lambdaParam),\(indivValueArray))))")
        }
    }
}
    
// MARK: - Ranks plus MPs
 
class RanksPlusMPsWriter: WriterBase {
    private var scoreData: ScoreData!
    
    override var name: String { "Source" }
    private var columns: [Column] = []
    private var headerTitleRow = 0
    private var headerDataRow = 1
    private var dataTitleRow = 3
    var dataRow = 4
    private var headerColumns = 0
    
    var positionColumn: Int?
    var directionColumn: Int?
    var participantNoColumn: Int?
    var firstNameColumn: [Int] = []
    var otherNamesColumn: [Int] = []
    var nationalIdColumn: [Int] = []
    var boardsPlayedColumn: [Int] = []
    var winDrawColumn: [Int] = []
    var bonusMPColumn: [Int] = []
    var winDrawMPColumn: [Int] = []
    var scoreColumn: Int?
    var totalMPColumn: [Int] = []
    
    var eventDescriptionCell: String?
    var entryCell: String?
    var tablesCell: String?
    var maxAwardCell: String?
    var ewMaxAwardCell: String?
    var minEntryCell: String?
    var awardToCell: String?
    var ewAwardToCell: String?
    var perWinCell: String?
    var eventDateCell: String?
    var eventIdCell: String?
    var localCell: String?
    var localMPsCell :String?
    var nationalMPsCell :String?
    var checksumCell :String?
    
    init(writer: Writer, round: Round, scoreData: ScoreData) {
        super.init(writer: writer)
        self.round = round
        self.scoreData = scoreData
        self.workbook = workbook
    }
    
    func write() {
        let participants = scoreData.events.first!.participants.sorted(by: sortCriteria)
        
        setupColumns()
        writeheader()
        
        freezePanes(worksheet: worksheet, row: dataRow, column: firstNameColumn[0])
                
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .numeric || column.cellType == .integerFormula || column.cellType == .floatFormula || column.cellType == .numericFormula {
                format = formatRightBold
            }
            if let width = column.width {
                setColumn(worksheet: worksheet, column: columnNumber, width: width)
            }
            write(worksheet: worksheet, row: dataTitleRow, column: columnNumber, string: round.replace(column.title), format: format)
        }
        
        for playerNumber in 0..<round.maxParticipantPlayers {
            let nationalIdColumn = nationalIdColumn[playerNumber]
            let firstNameColumn = firstNameColumn[playerNumber]
            let otherNamesColumn = otherNamesColumn[playerNumber]
            let firstNameNonBlank = "\(cell(dataRow, firstNameColumn, columnFixed: true))<>\"\""
            let otherNamesNonBlank = "\(cell(dataRow, otherNamesColumn, columnFixed: true))<>\"\""
            let nationalIdZero = "\(cell(dataRow, nationalIdColumn, columnFixed: true))<=0"
            let nationalIdLarge = "\(cell(dataRow, nationalIdColumn, columnFixed: true))>\(Settings.current.maxNationalIdNumber!)"
            let formula = "=AND(OR(\(firstNameNonBlank),\(otherNamesNonBlank)), OR(\(nationalIdZero),\(nationalIdLarge)))"
            setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: nationalIdColumn, toRow: dataRow + round.fieldSize - 1, toColumn: nationalIdColumn, formula: formula, format: formatRed!)
        }
        
        for participantNumber in 0..<round.fieldSize {
            
            let rowNumber = participantNumber + dataRow
            
            for (columnNumber, column) in columns.enumerated() {
            
                if participantNumber < participants.count {
                    let participant = participants[participantNumber]
                    if let content = column.content?(participant, rowNumber) {
                        write(cellType: (content == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: content)
                    }
                    
                    if let playerNumber = column.playerNumber {
                        let playerList = participant.member.playerList
                        if playerNumber < playerList.count {
                            if let playerContent = column.playerContent?(participant, playerList[playerNumber], playerNumber, rowNumber) {
                                let cellType = (playerContent.left(1) == "=" ? .integerFormula : column.cellType)
                                write(cellType: (playerContent == "" ? .string : cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: playerContent)
                            }
                        } else {
                            write(cellType: .string, worksheet: worksheet, row: rowNumber, column: columnNumber, content: "")
                        }
                    }
                }
                
                if let calculatedContent = column.calculatedContent?(column.playerNumber, rowNumber) {
                    write(cellType: (calculatedContent == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: calculatedContent)
                }
            }
        }
        
        setColumn(worksheet: worksheet, column: 0, toColumn: headerColumns - 1, width: 12.0)
    }
    
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
    
    private func setupColumns() {
        let event = scoreData.events.first!
        let playerCount = round.maxParticipantPlayers
        let winDraw = event.type?.requiresWinDraw ?? false
        
        columns.append(Column(title: "Place", content: { (participant, _) in "\(participant.place ?? 0)" }, cellType: .integer))
        positionColumn = columns.count - 1
        
        if event.winnerType == 2 && event.type?.participantType == .pair {
            columns.append(Column(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
            directionColumn = columns.count - 1
        }
        
        columns.append(Column(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer)) ; participantNoColumn = columns.count - 1
        
        columns.append(Column(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float)); scoreColumn = columns.count - 1
        
        if winDraw && playerCount <= event.type?.participantType?.players ?? playerCount {
            columns.append(Column(title: "Win/Draw", content: { (participant, _) in "\(participant.winDraw!)" }, cellType: .float))
            winDrawColumn.append(columns.count - 1)
        }
        
        for playerNumber in 0..<playerCount {
            columns.append(Column(title: "First Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 0) }, playerNumber: playerNumber, cellType: .string))
            firstNameColumn.append(columns.count - 1)
            
            columns.append(Column(title: "Other Names (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 1) }, playerNumber: playerNumber, cellType: .string))
            otherNamesColumn.append(columns.count - 1)
            
            columns.append(Column(title: "SBU No (\(playerNumber+1))", playerContent: playerNationalId, playerNumber: playerNumber, cellType: .integer))
            nationalIdColumn.append(columns.count - 1)
            
            if !scoreData.manualMPs {
                
                if playerCount > event.type?.participantType?.players ?? playerCount {
                    
                    columns.append(Column(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in
                        "\(player.boardsPlayed ?? event.boards ?? 1)"
                    }, playerNumber: playerNumber, cellType: .integer))
                    boardsPlayedColumn.append(columns.count - 1)
                    
                    if winDraw {
                        columns.append(Column(title: "Win/Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                        winDrawColumn.append(columns.count - 1)
                    }
                    
                    columns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP (\(playerNumber+1))", calculatedContent: playerBonusAward, playerNumber: playerNumber, cellType: .floatFormula))
                    if !winDraw {
                        totalMPColumn.append(columns.count - 1)
                    } else {
                        bonusMPColumn.append(columns.count - 1)
                        columns.append(Column(title: "Win/Draw MP (\(playerNumber+1))", calculatedContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .floatFormula))
                        winDrawMPColumn.append(columns.count - 1)
                        
                        columns.append(Column(title: "Total MP (\(playerNumber+1))", calculatedContent: playerTotalAward, playerNumber: playerNumber, cellType: .floatFormula))
                        totalMPColumn.append(columns.count - 1)
                    }
                }
            }
        }
        
        if scoreData.manualMPs {
            
            columns.append(Column(title: "Total MPs", content: { (participant, _) in "\(participant.manualMps ?? 0)" }, cellType: .floatFormula))
            for _ in 0..<playerCount {
                totalMPColumn.append(columns.count - 1)
            }
            
        } else {
            
            if playerCount <= event.type?.participantType?.players ?? playerCount {
                
                columns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP", calculatedContent: bonusAward, cellType: .floatFormula))
                
                if !winDraw {
                    for _ in 0..<playerCount {
                        totalMPColumn.append(columns.count - 1)
                    }
                } else {
                    bonusMPColumn.append(columns.count - 1)
                    
                    columns.append(Column(title: "Win/Draw MP", calculatedContent: winDrawAward, cellType: .floatFormula))
                    winDrawMPColumn.append(columns.count - 1)
                    
                    columns.append(Column(title: "Total MP", calculatedContent: totalAward, cellType: .floatFormula))
                    for _ in 0..<playerCount {
                        totalMPColumn.append(columns.count - 1)
                    }
                }
            }
        }
    }
    
    // Ranks plus MPs Content getters
    
    private func playerNationalId(_: Participant, player: Player, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        let nationalId = Int(player.nationalId ?? "0") ?? 0
        result = "=IFERROR(VLOOKUP(\(fnPrefix)CONCATENATE(\(cell(rowNumber, firstNameColumn[playerNumber!])),\" \",\(cell(rowNumber, otherNamesColumn[playerNumber!]))), \(cell(writer: writer.missing, writer.missing.dataRow, rowFixed: true, writer.missing.nameColumn, columnFixed: true)):\(cell(writer.missing.dataRow + Settings.current.largestPlayerCount, rowFixed: true, writer.missing.nationalIdColumn, columnFixed: true)),2,FALSE),\(nationalId == 0 ? ("\"" + (player.nationalId ?? "0") + "\"") : player.nationalId!))"
        if (nationalId <= 0 || nationalId > Settings.current.maxNationalIdNumber!) {
            if writer.missingNumbers[player.name!] == nil {
                if player.nationalId == nil || player.nationalId == "0" {
                    writer.missingNumbers[player.name!] = ("\(-(writer.missingNumbers.count + 1))", "")
                } else {
                    writer.missingNumbers[player.name!] = (player.nationalId!, "EBU")
                }
            }
        }
        return result
    }
    
    private func bonusAward(_: Int?, rowNumber: Int) -> String {
        return playerBonusAward(rowNumber: rowNumber)
    }
    
    private func playerBonusAward(playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "="
        let event = scoreData.events.first!
        var useAwardCell: String
        var useAwardToCell: String
        let positionCell = cell(rowNumber, positionColumn!)
        var allPositionsRange = ""
        
        if event.winnerType == 2 && event.type?.participantType == .pair {
            let directionCell = cell(rowNumber, directionColumn!, columnFixed: true)
            useAwardCell = "IF(\(directionCell)=\"\(Direction.ns.string)\",\(maxAwardCell!),\(ewMaxAwardCell!))"
            useAwardToCell = "IF(\(directionCell)=\"\(Direction.ns.string)\",\(awardToCell!),\(ewAwardToCell!))"
            allPositionsRange = "\(filter)(\(vstack)(\(cell(dataRow, rowFixed: true, positionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, positionColumn!, columnFixed: true))),\(vstack)(\(cell(dataRow, rowFixed: true, directionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, directionColumn!, columnFixed: true)))=\(directionCell))"
        } else {
            useAwardCell = maxAwardCell!
            useAwardToCell = awardToCell!
            allPositionsRange = "\(filter)(\(vstack)(\(cell(dataRow, rowFixed: true, positionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, positionColumn!, columnFixed: true))),\(vstack)(\(cell(dataRow, rowFixed: true, positionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, positionColumn!, columnFixed: true)))<>0)"
        }
        
        if let playerNumber = playerNumber {
            // Need to calculate percentage played
            var totalBoardsPlayed = "SUM("
            for playerNumber in 0..<round.maxParticipantPlayers {
                if playerNumber != 0 {
                    totalBoardsPlayed += ", "
                }
                totalBoardsPlayed += cell(rowNumber, boardsPlayedColumn[playerNumber], columnFixed: true)
            }
            totalBoardsPlayed += ")"

            result += "ROUNDUP((\(cell(rowNumber, boardsPlayedColumn[playerNumber], columnFixed: true))/(MAX(1,\(totalBoardsPlayed))/\(event.type!.participantType!.players)))*"
        }
        
        result += "IF(\(positionCell)=0,0,Award(\(useAwardCell), \(positionCell), \(useAwardToCell), 2, \(allPositionsRange)))"
        
        if playerNumber != nil {
            result += ", 2)"
        }
        
        return result
    }
    
    private func winDrawAward(_: Int?, rowNumber: Int) -> String {
        return playerWinDrawAward(rowNumber: rowNumber)
    }
    
    private func playerWinDrawAward(playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let winDrawsCell = cell(rowNumber, winDrawColumn[playerNumber ?? 0], columnFixed: true)
        result = "=ROUNDUP(\(winDrawsCell) * \(perWinCell!), 2)"
        
        return result
    }
    
    private func totalAward(_: Int?, rowNumber: Int) -> String {
        return playerTotalAward(rowNumber: rowNumber)
    }
    
    private func playerTotalAward(playerNumber: Int? = nil, rowNumber: Int) -> String {
        let event = scoreData.events.first!
        let winDraw = event.type?.requiresWinDraw ?? false
        var result = ""
        
        
        let bonusMPCell = cell(rowNumber, bonusMPColumn[playerNumber ?? 0], columnFixed: true)
        result = bonusMPCell
        if winDraw {
            let winDrawMPCell = cell(rowNumber, winDrawMPColumn[playerNumber ?? 0], columnFixed: true)
            result += "+\(winDrawMPCell)"
        }
        return result
    }
    
    // Ranks plus MPs header
    
    func writeheader() {
        let event = scoreData.events.first!
        var column = -1
        var row = headerTitleRow
        var prefix = ""
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        if twoWinners {
            prefix = "NS "
        }
        
        func writeCell(string: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, string: string, format: format)
        }
        
        func writeCell(integer: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, integer: integer, format: format)
        }
        
        func writeCell(float: Float, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, float: float, format: format)
        }
        
        func writeCell(date: Date, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, date: date, format: format)
        }
        
        func writeCell(formula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, floatFormula: formula, format: format)
        }
        
        func writeCell(integerFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, integerFormula: integerFormula, format: format)
        }
        
        func writeCell(floatFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: row, column: column, floatFormula: floatFormula, format: format)
        }
        
        writeCell(string: "Description", format: formatBold)
        if round.toe != nil {
            writeCell(string: "TOE", format: formatRightBold)
        }
        writeCell(string: "\(prefix)Entry", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW Entry", format: formatRightBold)
        }
        writeCell(string: "Tables", format: formatRightBold)
        writeCell(string: "Boards", format: formatRightBold)
        writeCell(string: "\(prefix)Full Award", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW Full Award", format: formatRightBold)
        }
        writeCell(string: "Min Entry", format: formatRightBold)
        writeCell(string: "Entry %", format: formatRightBold)
        writeCell(string: "Factor %", format: formatRightBold)
        writeCell(string: "\(prefix)Max Award", format: formatRightBold)
        if twoWinners {
            writeCell(string: "EW Max Award", format: formatRightBold)
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
        
        writeCell(string: "Local MPs", format: formatRightBold)
        writeCell(string: "National MPs", format: formatRightBold)
        writeCell(string: "Checksum", format: formatRightBold)
        
        headerColumns = column + 1

        column = -1
        row = headerDataRow
        
        writeCell(string: event.description ?? "") ; eventDescriptionCell = cell(row, rowFixed: true, column, columnFixed: true)
        if let toe = round.toe {
            writeCell(integer: toe, format: formatInt)
        }
        
        var entryCells = ""
        var ewEntryRef = ""
        if twoWinners {
            let nsCell = "COUNTIF(\(cell(dataRow, rowFixed: true, directionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, directionColumn!, columnFixed: true)),\"\(Direction.ns.string)\")"
            writeCell(integerFormula: "=\(nsCell)")
            entryCell = cell(row, rowFixed: true, column, columnFixed: true)
                        
            let ewCell = "COUNTIF(\(cell(dataRow, rowFixed: true, directionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, directionColumn!, columnFixed: true)),\"\(Direction.ew.string)\")"
            writeCell(integerFormula: "=\(ewCell)")
            ewEntryRef = cell(row, rowFixed: true, column, columnFixed: true)

            entryCells = "(\(nsCell)+\(ewCell))"
        } else {
            entryCells = "COUNTIF(\(cell(dataRow, rowFixed: true, positionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, positionColumn!, columnFixed: true)),\">0\")"
            writeCell(integerFormula: "=\(entryCells)") ; entryCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        writeCell(integerFormula: "=ROUNDUP(\(entryCells)*(\(event.type?.participantType?.players ?? 4)/4),0)") ; tablesCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: event.boards ?? 0)
        
        var baseEwMaxAwardCell: String?
        writeCell(floatFormula: "=ROUND(\(scoreData.maxAward),2)") ; let baseMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(floatFormula: "=ROUND(\(scoreData.ewMaxAward ?? scoreData.maxAward),2)") ; baseEwMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(integer: round.scoreData.minEntry, format: formatInt) ; minEntryCell = cell(row, rowFixed: true, column, columnFixed: true) ; minEntryCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(minEntryCell!)=0,1,MIN(1, ROUNDUP((\(entryCells))/\(minEntryCell!), 4)))", format: formatPercent) ; let entryFactorCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(float: round.scoreData.reducedTo, format: formatPercent) ; let reduceToCell = cell(row, rowFixed: true, column, columnFixed: true)

        writeCell(floatFormula: "=ROUNDUP(ROUNDUP(\(baseMaxAwardCell)*\(entryFactorCell),2)*\(reduceToCell),2)") ; maxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(floatFormula: "=ROUNDUP(ROUNDUP(\(baseEwMaxAwardCell!)*\(entryFactorCell),2)*\(reduceToCell),2)") ; ewMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(floatFormula: "=ROUND(\(scoreData.awardTo)/100,4)", format: formatPercent) ; let awardPercentCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(entryCell!)*\(awardPercentCell),0)") ; awardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(ewEntryRef)*\(awardPercentCell),0)") ; ewAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(floatFormula: "=ROUND(\(scoreData.perWin),2)") ; perWinCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(date: event.date ?? Date.today) ; eventDateCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.eventCode ?? "") ; eventIdCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: scoreData.national ? "National" : "Local") ; localCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(localCell!)=\"National\",0,ROUND(SUM(\(range(column: totalMPColumn))),2))") ; localMPsCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(localCell!)<>\"National\",0,ROUND(SUM(\(range(column: totalMPColumn))),2))") ; nationalMPsCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=CheckSum(\(vstack)(\(range(column: totalMPColumn))),\(vstack)(\(range(column: nationalIdColumn))))") ; checksumCell = cell(row, rowFixed: true, column, columnFixed: true)
    }
    
    private func range(column: [Int])->String {
        let firstRow = dataRow
        let lastRow = dataRow + round.fieldSize - 1
        var result = ""
        for playerNumber in 0..<round.maxParticipantPlayers {
            if playerNumber != 0 {
                result += ","
            }
            let from = cell(firstRow, rowFixed: true, column[playerNumber], columnFixed: true)
            let to = cell(lastRow, rowFixed: true, column[playerNumber], columnFixed: true)
            result += from + ":" + to
        }
        return result
    }
}
    
// MARK: - Indiviual MPs worksheet

class IndividualMPsWriter: WriterBase {
    var scoreData: ScoreData!
    override var name: String { "Indiv" }
    var columns: [Column] = []
    var positionColumn: Int?
    var directionColumn: Int?
    var uniqueColumn: Int?
    var decimalColumn: Int?
    var nationalIdColumn: Int?
    var localMPsColumn: Int?
    var nationalMPsColumn: Int?
    var localTotalColumn: Int?
    var nationalTotalColumn: Int?
    var checksumColumn: Int?
    var firstNameColumn: Int?
    
    let titleRow = 0
    let dataRow = 1
    
    init(writer: Writer, round: Round, scoreData: ScoreData) {
        super.init(writer: writer)
        self.scoreData = scoreData
        self.round = round
     }
    
    func write() {
        setupColumns()
        
        freezePanes(worksheet: worksheet, row: dataRow, column: nationalIdColumn!)
        
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .numeric || column.cellType == .integerFormula || column.cellType == .floatFormula || column.cellType == .numericFormula {
                format = formatRightBold
            }
            setColumn(worksheet: worksheet, column: columnNumber, width: column.width)
            write(worksheet: worksheet, row: titleRow, column: columnNumber, string: round.replace(column.title), format: format)
        }
        
        for (columnNumber, column) in columns.enumerated() {
            
            if let referenceContent = column.referenceContent {
                referenceColumn(columnNumber: columnNumber, referencedContent: referenceContent, cellType: column.cellType)
            }
            
            if let referenceDynamicContent = column.referenceDynamic {
                writeDynamicReference(rowNumber: dataRow, columnNumber: columnNumber, content: referenceDynamicContent, cellType: column.cellType)
            }
        }
        
        let localArrayRef = "\(arrayRef)(\(cell(dataRow, rowFixed: true, localMPsColumn!, columnFixed: true)))"
        write(worksheet: worksheet, row: dataRow, column: localTotalColumn!, floatFormula: "=SUM(\(localArrayRef))", format: formatZeroFloat)
        
        let nationalArrayRef = "\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalMPsColumn!, columnFixed: true)))"
        write(worksheet: worksheet, row: dataRow, column: nationalTotalColumn!, floatFormula: "=SUM(\(nationalArrayRef))", format: formatZeroFloat)
        
        let totalArrayRef = "\(arrayRef)(\(cell(dataRow, rowFixed: true, decimalColumn!, columnFixed: true)))"
        let nationalIdArrayRef = "\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn!, columnFixed: true)))"
        write(worksheet: worksheet, row: dataRow, column: checksumColumn!, floatFormula: "=CheckSum(\(vstack)(\(totalArrayRef)),\(vstack)(\(nationalIdArrayRef)))", format: formatZeroFloat)
        
        setColumn(worksheet: worksheet, column: uniqueColumn!, hidden: true)
        
    }
    
    private func setupColumns() {
        let event = scoreData.events.first!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        
        columns.append(Column(title: "Place", referenceContent: { [self] (_) in round.ranksPlusMps.positionColumn! }, cellType: .integerFormula)) ; positionColumn = columns.count - 1
        
        if twoWinners {
            columns.append(Column(title: "Direction", referenceContent: { [self] (_) in round.ranksPlusMps.directionColumn! }, cellType: .stringFormula))
            directionColumn = columns.count - 1
        }
        
        columns.append(Column(title: "@P no", referenceContent: { [self] (_) in round.ranksPlusMps.participantNoColumn! }, cellType: .integerFormula))
        
        columns.append(Column(title: "Unique", cellType: .floatFormula)) ; uniqueColumn = columns.count - 1
        let unique = columns.last!
        
        columns.append(Column(title: "Names", referenceContent: { [self] (playerNumber) in round.ranksPlusMps.firstNameColumn[playerNumber] }, cellType: .stringFormula, width: 16)) ; firstNameColumn = columns.count - 1
        
        columns.append(Column(title: "", referenceContent: { [self] (playerNumber) in round.ranksPlusMps.otherNamesColumn[playerNumber] }, cellType: .stringFormula, width: 16)) ; let otherNamesColumn = columns.count - 1
        
        columns.append(Column(title: "SBU No", referenceContent: { [self] (playerNumber) in round.ranksPlusMps.nationalIdColumn[playerNumber] }, cellType: .integerFormula)) ; nationalIdColumn = columns.count - 1
        
        unique.referenceDynamic = { [self] in "CONCATENATE(\(arrayRef)(\(cell(dataRow, nationalIdColumn!, columnFixed: true))), \"+\", \(arrayRef)(\(cell(dataRow, firstNameColumn!, columnFixed: true))), \"+\", \(arrayRef)(\(cell(dataRow, otherNamesColumn, columnFixed: true))))" }
        
        columns.append(Column(title: "Total MPs", referenceContent: { [self] (playerNumber) in round.ranksPlusMps.totalMPColumn[playerNumber] }, cellType: .floatFormula, width: 12)) ; decimalColumn = columns.count - 1
        
        columns.append(Column(title: "Local MPs", referenceDynamic: { [self] in "=\(byRow)(\(arrayRef)(\(cell(dataRow, decimalColumn!, columnFixed: true))),\(lambda)(\(lambdaParam), IF(\(cell(writer: round.ranksPlusMps, round.ranksPlusMps.localCell!))<>\"National\",\(lambdaParam),0)))" }, cellType: .floatFormula, width: 12)) ; localMPsColumn = columns.count - 1
        
        columns.append(Column(title: "National MPs", referenceDynamic: { [self] in "=\(byRow)(\(arrayRef)(\(cell(dataRow, decimalColumn!, columnFixed: true))),\(lambda)(\(lambdaParam), IF(\(cell(writer: round.ranksPlusMps, round.ranksPlusMps.localCell!))=\"National\",\(lambdaParam),0)))" }, cellType: .floatFormula, width: 12)) ; nationalMPsColumn = columns.count - 1
        
        columns.append(Column(title: "Total Local", cellType: .floatFormula, width: 12)) ; localTotalColumn = columns.count - 1
        
        columns.append(Column(title: "Total National", cellType: .floatFormula, width: 12)) ; nationalTotalColumn = columns.count - 1
        
        columns.append(Column(title: "Checksum", cellType: .floatFormula, width: 16)) ; checksumColumn = columns.count - 1
    }
    
    private func referenceColumn(columnNumber: Int, referencedContent: (Int)->Int, cellType: CellType? = nil) {
        let event = scoreData.events.first!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        
        let content = zeroFiltered(referencedContent: referencedContent)
        let position = zeroFiltered(referencedContent: { (_) in positionColumn! })
        
        var result = "=\(sortBy)(\(content), "
        
        if twoWinners {
            let direction = zeroFiltered(referencedContent: { (_) in directionColumn! })
            result += "\(direction), -1, "
        }
        result += "\(position), 1)"
        
        let column = lxw_col_t(Int32(columnNumber))
        let maxRow = lxw_row_t(Int32(Settings.current.largestPlayerCount))
        worksheet_write_dynamic_array_formula(worksheet, 1, column, maxRow, column, result, formatFrom(cellType: cellType))
    }
    
    private func zeroFiltered(referencedContent: (Int)->Int) -> String {
        
        var result = "\(filter)(\(vstack)("
        
        for playerNumber in 0..<round.maxParticipantPlayers {
            let columnReference = referencedContent(playerNumber)
            if playerNumber != 0 {
                result += ","
            }
            result += cell(writer: round.ranksPlusMps, round.ranksPlusMps.dataRow, rowFixed: true, columnReference)
            result += ":"
            result += cell(round.ranksPlusMps.dataRow + round.fieldSize - 1, rowFixed: true, columnReference)
        }
        
        result += "),\(vstack)("
        
        for playerNumber in 0..<round.maxParticipantPlayers {
            if playerNumber != 0 {
                result += ","
            }
            result += cell(writer: round.ranksPlusMps, round.ranksPlusMps.dataRow, rowFixed: true, round.ranksPlusMps.totalMPColumn[playerNumber])
            result += ":"
            result += cell(round.ranksPlusMps.dataRow + round.fieldSize - 1, rowFixed: true, round.ranksPlusMps.totalMPColumn[playerNumber])
        }
        result += ")<>0)"
        
        return result
    }
    
}

// MARK: - Summary
    
class MissingNumbersWriter : WriterBase {
    
    override var name: String { "Missing Numbers" }
    let nameColumn = 0
    let nationalIdColumn = 1
    let nboColumn = 2
    let suggestColumn = 3
    let duplicateColumn = 4
    let headerRow = 0
    let dataRow = 1

    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        setColumn(worksheet: worksheet, column: nameColumn, width: 30)
        setColumn(worksheet: worksheet, column: nationalIdColumn, width: 12)
        setColumn(worksheet: worksheet, column: nboColumn, width: 8)
        
        write(worksheet: worksheet, row: headerRow, column: nameColumn, string: "Names", format: formatBold)
        write(worksheet: worksheet, row: headerRow, column: nationalIdColumn, string: "SBU No", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow, column: nboColumn, string: "NBO", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow, column: suggestColumn, string: "Suggest", format: formatRightBold)
        
        if writer.missingNumbers.count == 0 {
            writeLogic(row: 0)
        } else {
            var row = 0
            for (name, (number, nbo)) in writer.missingNumbers.sorted(by: {$0.value > $1.value}) {
                var nbo = nbo
                var number = number
                let split = number.split(at: "-")
                if split.count == 2 {
                    nbo = split[0]
                    number = split[1]
                }
                write(worksheet: worksheet, row: dataRow + row, column: nameColumn, string: name)
                let nboCell = cell(dataRow + row, nboColumn, columnFixed: true)
                write(worksheet: worksheet, row: dataRow + row, column: nationalIdColumn, formula: "=CONCATENATE(\(nboCell),IF(\(nboCell)=\"\",\"\",\"-\"),\"\(number)\")", format: formatInt)
                write(worksheet: worksheet, row: dataRow + row, column: nboColumn, string: nbo, format: formatInt)
                writeLogic(row: row)
                row += 1
            }
        }
    }
    
    func writeLogic(row: Int) {
        let xlookupStem = "IFERROR(\(fnPrefix)XLOOKUP(\(cell(dataRow + row, nameColumn, columnFixed: true)),CONCATENATE(\(vstack)(\(lookupRange(.firstName))),\" \",\(vstack)(\(lookupRange(.otherNames)))),\(vstack)(\(lookupRange(.nationalId))),\"\",FALSE"
        write(worksheet: worksheet, row: dataRow + row, column: suggestColumn, integerFormula: "=\(xlookupStem),1),\"\")")
        write(worksheet: worksheet, row: dataRow + row, column: duplicateColumn, integerFormula: "=IF(\(xlookupStem),-1),\"\")=\(cell(dataRow + row, suggestColumn, columnFixed: true)),\"\",\"DUPLICATES\")", format: formatBold)
    }
    
    enum LookupColumn {
        case nationalId
        case firstName
        case otherNames
    }
    
    
    func lookupRange(_ requiredCol: LookupColumn) -> String {
        return "\(lookupCell(true, requiredCol, true)):\(lookupCell(false, requiredCol, false))"
    }
    
    func lookupCell(_ firstRow: Bool, _ requiredCol: LookupColumn, _ filename: Bool) -> String {
        let row = (firstRow ? Settings.current.userDownloadMinRow! : Settings.current.userDownloadMaxRow!)
        var column : Int
        switch requiredCol {
        case .nationalId: column = columnNumber(Settings.current.userDownloadNationalIdColumn)
        case .firstName: column = columnNumber(Settings.current.userDownloadFirstNameColumn)
        case .otherNames: column = columnNumber(Settings.current.userDownloadOtherNamesColumn)
        }
        return "\(filename ? "'\(Settings.current.userDownloadData!)'!" : "")\(cell(row, rowFixed: true, column, columnFixed: true))"
    }
}
// MARK - Writer base class

class WriterBase {
    
    let fnPrefix = "_xlfn."
    let dynamicFnPrefix = "_xlfn._xlws."
    let paramPrefix = "_xlpm."
    
    var vstack: String { "\(fnPrefix)VSTACK" }
    var hstack: String { "\(fnPrefix)HSTACK" }
    var arrayRef: String { "\(fnPrefix)ANCHORARRAY" }
    var filter: String { "\(dynamicFnPrefix)FILTER" }
    var sortBy: String { "\(dynamicFnPrefix)SORTBY" }
    var unique: String { "\(dynamicFnPrefix)UNIQUE" }
    var byRow: String { "\(fnPrefix)BYROW" }
    var lambda: String { "\(fnPrefix)LAMBDA"}
    var lambdaParam: String { "\(paramPrefix)x"}
    
    var writer: Writer! = nil
    var workbook: UnsafeMutablePointer<lxw_workbook>?
    var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    var formatString: UnsafeMutablePointer<lxw_format>?
    var formatZeroBlank: UnsafeMutablePointer<lxw_format>?
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
    var formatRed: UnsafeMutablePointer<lxw_format>?
    var formatRedHatched: UnsafeMutablePointer<lxw_format>?
    var formatYellow: UnsafeMutablePointer<lxw_format>?
    var formatGrey: UnsafeMutablePointer<lxw_format>?
    
    var round: Round!
    var name: String { fatalError() }
    
    init(writer: Writer? = nil) {
        if let writer = writer {
            self.writer = writer
        }
    }
    
    func prepare(workbook: UnsafeMutablePointer<lxw_workbook>?) {
        self.workbook = workbook
        var fullName = ""
        if let round = round {
            fullName += round.shortName.replacingOccurrences(of: "[", with: "")
                                       .replacingOccurrences(of: "]", with: "")
                                       .replacingOccurrences(of: ":", with: "-")
                                       .replacingOccurrences(of: "*", with: "-")
                                       .replacingOccurrences(of: "?", with: "-")
                                       .replacingOccurrences(of: "/", with: "-")
                                       .replacingOccurrences(of: "\\", with: "-") + " "
        }
        fullName += name
        worksheet = workbook_add_worksheet(workbook, fullName)
        worksheet_set_zoom(worksheet, UInt16(Settings.current.defaultWorksheetZoom))
        if let writer = writer {
            // Copy from writer
            self.formatString = writer.formatString
            self.formatZeroBlank = writer.formatZeroBlank
            self.formatInt = writer.formatInt
            self.formatFloat = writer.formatFloat
            self.formatZeroInt = writer.formatZeroInt
            self.formatZeroFloat = writer.formatZeroFloat
            self.formatBold = writer.formatBold
            self.formatRightBold = writer.formatRightBold
            self.formatBoldUnderline = writer.formatBoldUnderline
            self.formatRightBoldUnderline = writer.formatRightBoldUnderline
            self.formatFloatBoldUnderline = writer.formatFloatBoldUnderline
            self.formatPercent = writer.formatPercent
            self.formatDate = writer.formatDate
            self.formatRed = writer.formatRed
            self.formatRedHatched = writer.formatRedHatched
            self.formatYellow = writer.formatYellow
            self.formatGrey = writer.formatGrey
        }
    }
    
    // MARK: - Utility routines
    
    fileprivate func createMacroButton(worksheet: UnsafeMutablePointer<lxw_worksheet>?, title: String, macro: String, row: Int, column: Int, height: Int = 30, width: Int = 80, xScale: Double = 1.5, yScale: Double = 1.5, xOffset: Int = 2, yOffset: Int = 2) {
        // Add macro buttons
        var options = lxw_button_options()
        options.macro = UnsafeMutablePointer<Int8>(mutating: (macro as NSString).utf8String)
        options.caption = UnsafeMutablePointer<Int8>(mutating: (title as NSString).utf8String)
        options.height = UInt16(height)
        options.width = UInt16(width)
        options.x_scale = xScale
        options.y_scale = yScale
        options.x_offset = Int32(xOffset)
        options.y_offset = Int32(yOffset)
        worksheet_insert_button(worksheet, lxw_row_t(row + 1), lxw_col_t(column + 1), &options)
    }
    
    fileprivate func formatFrom(cellType: CellType? = nil) -> UnsafeMutablePointer<lxw_format>? {
        var format = formatInt
        if let cellType = cellType {
            switch cellType {
            case .date:
                format = formatDate
            case .float, .floatFormula:
                format = formatFloat
            case .string, .stringFormula:
                format = nil
            default:
                break
            }
        }
        return format
    }
    
    fileprivate func nameColumn(name: String, element: Int) -> String {
        let names = name.components(separatedBy: " ")
        if element == 0 {
            return names[0]
        } else {
            var otherNames = names
            otherNames.removeFirst()
            return otherNames.joined(separator: " ")
        }
    }
    
    func cell(writer: WriterBase? = nil, _ row: Int, rowFixed: Bool = false, _ column: Int, columnFixed: Bool = false) -> String {
        let rowRef = rowRef(row, fixed: rowFixed)
        let columnRef = columnRef(column, fixed: columnFixed)
        return cell(writer: writer, "\(columnRef)\(rowRef)")
    }
    
    func cell(writer: WriterBase? = nil, _ cellRef: String) -> String {
        var roundRef = ""
        if let writer = writer {
            var prefix = ""
            if let round = writer.round {
                prefix = "\(round.shortName) "
            }
            roundRef = "'\(prefix)\(writer.name)'!"
        }
        return "\(roundRef)\(cellRef)"
    }
    
    func rowRef(_ row: Int, fixed: Bool = false) -> String {
        let rowRef = (fixed ? "$" : "") + "\(row + 1)"
        return rowRef
    }
    
    func columnRef(_ column: Int, fixed: Bool = false) -> String {
        var columnRef = ""
        var remaining = column
        while remaining >= 0 {
            let letter = remaining % 26
            columnRef = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".mid(letter,1) + columnRef
            remaining = ((remaining - letter) / 26) - 1
        }
        return (fixed ? "$" : "") + columnRef
    }
    
    func columnNumber(_ ref: String) -> Int {
        var result = 0
        for index in 0..<ref.count {
            let letter = ref.mid(index,1)
            result = (result * 26) + ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".position(letter) ?? 0)
        }
        return result
    }
    
    func setupFormats() {
        formatString = workbook_add_format(workbook)
        format_set_align(formatString, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatZeroBlank = workbook_add_format(workbook)
        format_set_num_format(formatZeroBlank, "0;-0;")
        format_set_align(formatZeroBlank, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatInt = workbook_add_format(workbook)
        format_set_num_format(formatInt, "0;-0;")
        format_set_align(formatInt, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatFloat = workbook_add_format(workbook)
        format_set_num_format(formatFloat, "0.00;-0.00;")
        format_set_align(formatFloat, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatZeroInt = workbook_add_format(workbook)
        format_set_num_format(formatInt, "0")
        format_set_align(formatZeroInt, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatZeroFloat = workbook_add_format(workbook)
        format_set_num_format(formatZeroFloat, "0.00")
        format_set_align(formatZeroFloat, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatDate = workbook_add_format(workbook)
        format_set_num_format(formatDate, "dd/MM/yyyy")
        format_set_align(formatDate, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatPercent = workbook_add_format(workbook)
        format_set_num_format(formatPercent, "0.00%")
        format_set_align(formatPercent, UInt8(LXW_ALIGN_RIGHT.rawValue))
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
        format_set_text_wrap(formatBoldUnderline)
        formatRightBoldUnderline = workbook_add_format(workbook)
        format_set_bold(formatRightBoldUnderline)
        format_set_align(formatRightBoldUnderline, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_bottom(formatRightBoldUnderline, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_text_wrap(formatRightBoldUnderline)
        format_set_num_format(formatRightBoldUnderline, "0;-0;")
        formatFloatBoldUnderline = workbook_add_format(workbook)
        format_set_bold(formatFloatBoldUnderline)
        format_set_align(formatFloatBoldUnderline, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_bottom(formatFloatBoldUnderline, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_num_format(formatFloatBoldUnderline, "0.00;-0.00;")
        
        formatRed = workbook_add_format(workbook)
        format_set_bg_color(formatRed, LXW_COLOR_RED.rawValue)
        format_set_font_color(formatRed, LXW_COLOR_WHITE.rawValue)
        formatRedHatched = workbook_add_format(workbook)
        format_set_bg_color(formatRedHatched, LXW_COLOR_RED.rawValue)
        format_set_fg_color(formatRedHatched, LXW_COLOR_WHITE.rawValue)
        format_set_pattern(formatRedHatched, UInt8(LXW_PATTERN_LIGHT_UP.rawValue))
        formatYellow = workbook_add_format(workbook)
        format_set_bg_color(formatYellow, 0xFFFB00)
        formatGrey = workbook_add_format(workbook)
        format_set_bg_color(formatGrey, 0xc0C0C0C0)

    }
    
    // MARK: - Helper routines
    
    func write(cellType: CellType, worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, content: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        switch cellType {
        case .string:
            write(worksheet: worksheet, row: row, column: column, string: content, format: format)
        case .integer:
            if let integer = Int(content) {
                write(worksheet: worksheet, row: row, column: column, integer: integer, format: format ?? formatInt)
            } else {
                write(worksheet: worksheet, row: row, column: column, string: content, format: format ?? formatInt)
            }
        case .float:
            if let float = Float(content) {
                write(worksheet: worksheet, row: row, column: column, float: float, format: format ?? formatFloat)
            } else {
                write(worksheet: worksheet, row: row, column: column, string: content, format: format ?? formatFloat)
            }
        case .stringFormula:
            write(worksheet: worksheet, row: row, column: column, formula: content, format: format)
        case .integerFormula:
            write(worksheet: worksheet, row: row, column: column, integerFormula: content, format: format ?? formatInt)
        case .floatFormula:
            write(worksheet: worksheet, row: row, column: column, floatFormula: content, format: format ?? formatFloat)
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
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, dynamicFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        worksheet_write_dynamic_formula(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), dynamicFormula, format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, dynamicIntegerFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatInt
        worksheet_write_dynamic_formula(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), dynamicIntegerFormula, format)
    }
    
    func write(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int, dynamicFloatFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let format = format ?? formatFloat
        worksheet_write_dynamic_formula(worksheet, lxw_row_t(Int32(row)), lxw_col_t(Int32(column)), dynamicFloatFormula, format)
    }
    
    fileprivate func writeDynamicReference(rowNumber: Int, columnNumber: Int, content: ()->String, cellType: CellType? = nil, format: UnsafeMutablePointer<lxw_format>? = nil) {
        
        let row = lxw_row_t(Int32(rowNumber))
        let column = lxw_col_t(Int32(columnNumber))
        let maxRow = lxw_row_t(Int32(Settings.current.largestPlayerCount + rowNumber - 1))
        let format = format ?? formatFrom(cellType: cellType)
        worksheet_write_dynamic_array_formula(worksheet, row, column, maxRow, column, content(), format)
    }
    
    func setConditionalFormat(worksheet: UnsafeMutablePointer<lxw_worksheet>?, fromRow: Int, fromColumn: Int, toRow: Int, toColumn: Int, formula: String, checkDuplicates: Bool = false, stopIfTrue: Bool = true, format: UnsafeMutablePointer<lxw_format>, duplicateFormat: UnsafeMutablePointer<lxw_format>? = nil) {
        let formula = formula
        var conditionalFormat = lxw_conditional_format()
        conditionalFormat.type = UInt8(LXW_CONDITIONAL_TYPE_FORMULA.rawValue)
        conditionalFormat.value_string = UnsafeMutablePointer<CChar>(mutating: NSString(string: formula).utf8String)
        conditionalFormat.stop_if_true = stopIfTrue ? 1 : 0
        conditionalFormat.format = format
        
        worksheet_conditional_format_range(worksheet, lxw_row_t(Int32(fromRow)), lxw_col_t(Int32(fromColumn)), lxw_row_t(Int32(toRow)), lxw_col_t(Int32(toColumn)), &conditionalFormat)
        
        if checkDuplicates {
            var dupFormat = lxw_conditional_format()
            dupFormat.type = UInt8(LXW_CONDITIONAL_TYPE_DUPLICATE.rawValue)
            dupFormat.format = duplicateFormat ?? format
            
            worksheet_conditional_format_range(worksheet, lxw_row_t(Int32(fromRow)), lxw_col_t(Int32(fromColumn)), lxw_row_t(Int32(toRow)), lxw_col_t(Int32(toColumn)), &dupFormat)
        }
    }
            
    func setDataValidation(row: Int, column: Int, formula: String) {
        let row = lxw_row_t(Int32(row))
        let column = lxw_col_t(Int32(column))
        var validation = lxw_data_validation()
        validation.validate = UInt8(LXW_VALIDATION_TYPE_LIST_FORMULA.rawValue)
        validation.criteria = UInt8(LXW_VALIDATION_CRITERIA_EQUAL_TO.rawValue)
        validation.value_formula = UnsafeMutablePointer<CChar>(mutating: NSString(string: formula).utf8String)
        worksheet_data_validation_cell(worksheet, row, column, &validation)
    }
    
    func setColumn(worksheet: UnsafeMutablePointer<lxw_worksheet>?, column: Int, toColumn: Int? = nil, width: Float? = nil, hidden: Bool = false, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let toColumn = lxw_col_t(Int32(toColumn ?? column))
        let column = lxw_col_t(Int32(column))
        let width = (width == nil ? LXW_DEF_COL_WIDTH : Double(width!))
        if hidden {
            var options = lxw_row_col_options()
            options.hidden = 1
            worksheet_set_column_opt(worksheet, column, toColumn, width, format, &options)
        } else {
            worksheet_set_column_opt(worksheet, column, toColumn, width, format, nil)
        }
    }
    
    func setRow(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, format: UnsafeMutablePointer<lxw_format>? = nil, height: Float? = nil) {
        let row = lxw_row_t(Int32(row))
        let height = (height == nil ? LXW_DEF_ROW_HEIGHT : Double(height!))
        worksheet_set_row(worksheet, row, height, format)
    }
    
    func freezePanes(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int) {
        let row = lxw_row_t(Int32(row))
        let column = lxw_col_t(Int32(column))
        worksheet_freeze_panes(worksheet, row, column)
    }
    
}
