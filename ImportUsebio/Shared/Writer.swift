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

enum RankCategory: CaseIterable {
    case gold
    case silver
    case bronze
    
    var string: String {
        "\(self)".capitalized
    }
    
    var backgroundColor: lxw_color_t {
        switch self {
        case .gold:
            lxw_color_t(excelGold.rgbValue)
        case .silver:
            lxw_color_t(excelSilver.rgbValue)
        case .bronze:
            lxw_color_t(excelBronze.rgbValue)
        }
    }
    
    var textColor: lxw_color_t {
        switch self {
        case .bronze:
            LXW_COLOR_WHITE.rawValue
        default:
            LXW_COLOR_BLACK.rawValue
        }
    }
}

class Column {
    var title: String
    var content: ((Participant, Int)->String)?
    var playerContent: ((Participant, Player, Int, Int)->String)?
    var calculatedContent: ((Int?, Int) -> String)?
    var referenceContent: ((Int)->Int)?
    var referenceDynamic: (()->String)?
    var aggregateAs: String?
    var playerNumber: Int?
    var cellType: CellType
    var format: UnsafeMutablePointer<lxw_format>?
    var width: Float?
    var roundNumber: Int?
    var sortedDynamic: Bool = false
    
    init(title: String, content: ((Participant, Int)->String)? = nil, playerContent: ((Participant, Player, Int, Int)->String)? = nil, calculatedContent: ((Int?, Int) -> String)? = nil, playerNumber: Int? = nil, referenceContent: ((Int)->Int)? = nil, referenceDynamic: (()->String)? = nil, cellType: CellType = .string, format: UnsafeMutablePointer<lxw_format>? = nil, width: Float? = nil, aggregateAs: String? = nil, roundNumber: Int? = nil, sortedDynamic: Bool = false) {
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
        self.aggregateAs = aggregateAs
        self.roundNumber = roundNumber
        self.sortedDynamic = sortedDynamic
    }
}

class FormattedColumn : Column {
    var referenceColumn: Int = 0
}

class RanksPlusMPsColumn : Column {
    var strataNumber: Int? = nil
    var strataContent: ((Int, Int?, Int) -> String)?
    
    init(title: String, content: ((Participant, Int)->String)? = nil, playerContent: ((Participant, Player, Int, Int)->String)? = nil, calculatedContent: ((Int?, Int) -> String)? = nil, strataContent: ((Int, Int?, Int) -> String)? = nil, playerNumber: Int? = nil, strataNumber: Int? = nil, referenceContent: ((Int)->Int)? = nil, referenceDynamic: (()->String)? = nil, cellType: CellType = .string, format: UnsafeMutablePointer<lxw_format>? = nil, width: Float? = nil, aggregateAs: String? = nil, roundNumber: Int? = nil, sortedDynamic: Bool = false) {
        self.strataNumber = strataNumber
        self.strataContent = strataContent
        super.init(title: title, content: content, playerContent: playerContent, calculatedContent: calculatedContent, playerNumber: playerNumber, referenceDynamic: referenceDynamic, cellType: cellType, format: format, width: width, aggregateAs: aggregateAs, roundNumber: roundNumber, sortedDynamic: sortedDynamic)
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
        if let overrideTeamMembers = scoreData.overrideTeamMembers {
            result = max(event.type?.participantType?.players ?? 0, overrideTeamMembers)
        } else if event.type?.participantType == .team {
            result = max(event.type?.participantType?.players ?? 0, event.participants.map{$0.member.playerList.count}.max() ?? 0)
        } else {
            result = event.type?.participantType?.players ?? 2
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
    var raceExport: RaceExportWriter!
    var raceDetail: RaceDetailWriter!
    var raceFormatted: RaceFormattedWriter!
    var rounds: [Round] = []
    var minRank = 0
    var maxRank = 9999
    var eventCode: String = ""
    var clubCode: String = ""
    var chooseBest: Int = 0
    var eventDescription: String = ""
    var missingNumbers: [String: (NationalId: String, Nbo: String)] = [:]
    var includeInRace = true
    
    var maxPlayers: Int { min(1000, rounds.map{ $0.maxPlayers }.reduce(0, +)) }
    
    init() {
        super.init()
        parameters = ParametersWriter(writer: self)
        summary = SummaryWriter(writer: self)
        csvImport = CsvImportWriter(writer: self)
        consolidated = ConsolidatedWriter(writer: self)
        missing = MissingNumbersWriter(writer: self)
        formatted = FormattedWriter(writer: self)
        raceExport = RaceExportWriter(writer: self)
        raceDetail = RaceDetailWriter(writer: self)
        raceFormatted = RaceFormattedWriter(writer: self)
    }
        
    @discardableResult func add(name: String, shortName: String? = nil, scoreData: ScoreData) -> Round {
        if scoreData.source == .usebio {
            // Need to recalculate the wins/draws in case rounding mode changed
            _ = UsebioParser.calculateWinDraw(scoreData: scoreData)
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
        if includeInRace {
            raceFormatted.prepare(workbook: workbook)
            raceExport.prepare(workbook: workbook)
            raceDetail.prepare(workbook: workbook)
        }
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
        parameters.writeRanks()
        parameters.writeOrientation()
        parameters.writeErrorTypes()
        parameters.writeLastErrorState()
        csvImport.write()
        summary.write()
        missing.write()
        if missingNumbers.count <= 0 {
            worksheet_hide(missing.worksheet)
        }
        formatted.write()
        if includeInRace {
            raceDetail.write()
            raceFormatted.write()
            raceExport.write()
            worksheet_hide(raceDetail.worksheet)
            worksheet_hide(raceExport.worksheet)
        }
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
    
    let ranksFromColumn = 4
    let ranksCategoryColumn = 5
    
    let pageOrientationColumn = 7
    
    let errorTypeColumn = 9
    let errorColorColumn = 10
    
    let lastErrorNameColumn = 12
    let lastErrorValueColumn = 13
   
    let headerRow = 0
    let dataRow = 1
    
    var sortByNameRange: String!
    var sortByAddressRange: String!
    var sortByDirectionRange: String!

    var ranksFromRange: String!
    var ranksCategoryRange: String!
    
    var sortData: [(name: String, column: Int, direction: Int)] = []
    var ranksData: [(from: Int, category: String)] = []
    var errorData: [(type: String, color: Color)] = []
    var lastErrorData: [(name: String, value: Int)] = []
    
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
            write(worksheet: worksheet, row: headerRow, column: sortNameColumn + column, string: header, format: formatBold)
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
    
    func writeRanks() {
        ranksData = [ (  0, RankCategory.bronze.string),
                      (  1, "Non-SBU"),
                      (  2, RankCategory.bronze.string),
                      (165, RankCategory.silver.string),
                      (190, RankCategory.gold.string  )]
        
        for (column, header) in ["From", "Category"].enumerated() {
            write(worksheet: worksheet, row: headerRow, column: ranksFromColumn + column, string: header, format: formatBold)
        }
        
        for (row,element) in ranksData.enumerated() {
            write(worksheet: worksheet, row: dataRow + row, column: ranksFromColumn, integer: element.from, format: formatString)
            write(worksheet: worksheet, row: dataRow + row, column: ranksCategoryColumn, string: element.category)
        }
        
        // Define names
        workbook_define_name(writer.workbook, "RanksFrom", "=\(cell(writer: self, dataRow, rowFixed: true, ranksFromColumn, columnFixed: true)):\(cell(writer: self, dataRow + ranksData.count - 1, rowFixed: true, ranksFromColumn, columnFixed: true))")
        
        workbook_define_name(writer.workbook, "RanksCategory", "=\(cell(writer: self, dataRow, rowFixed: true, ranksCategoryColumn, columnFixed: true)):\(cell(writer: self, dataRow + ranksData.count - 1, rowFixed: true, ranksCategoryColumn, columnFixed: true))")

    }
    
    func writeOrientation() {
        write(worksheet: worksheet, row: headerRow, column: pageOrientationColumn, string: "Orientation", format: formatBold)
        write(worksheet: worksheet, row: dataRow, column: pageOrientationColumn, string: "Portrait")
        write(worksheet: worksheet, row: dataRow + 1, column: pageOrientationColumn, string: "Landscape")
        
        setColumn(worksheet: worksheet, column: pageOrientationColumn, width: 17)
        
        workbook_define_name(writer.workbook, "PageOrientation", "=\(cell(writer: self, dataRow, rowFixed: true, pageOrientationColumn, columnFixed: true)):\(cell(writer: self, dataRow + 1, rowFixed: true, pageOrientationColumn, columnFixed: true))")
    }
    
    func writeErrorTypes() {
        errorData = [ ("Error",                 excelRed),
                      ("Bad ID",                excelGrey),
                      ("Inactive member",       excelNotActive),
                      ("No home club member",   excelNoHomeClub),
                      ("Warning",               excelYellow),
                      ("Not paid member",       excelNotPaid)]
        
        for (column, header) in ["Error Type", "Colour"].enumerated() {
            write(worksheet: worksheet, row: headerRow, column: errorTypeColumn + column, string: header, format: formatBold)
        }
        
        for (row, element) in errorData.enumerated() {
            write(worksheet: worksheet, row: dataRow + row, column: errorTypeColumn, string: element.type)
            write(worksheet: worksheet, row: dataRow + row, column: errorColorColumn, integer: element.color.excelValue)
        }
        
        setColumn(worksheet: worksheet, column: errorTypeColumn, width: 17)
        
        workbook_define_name(writer.workbook, "ErrorTypes", "=\(cell(writer: self, dataRow, rowFixed: true, errorTypeColumn, columnFixed: true)):\(cell(writer: self, dataRow + errorData.count - 1, rowFixed: true, errorColorColumn, columnFixed: true))")
        
    }
    
    func writeLastErrorState() {
        
        lastErrorData = [ ("Type",      0),
                          ("Row",       0),
                          ("StartRow",  0)]
        
        for (column, header) in ["Last error", "Value"].enumerated() {
            write(worksheet: worksheet, row: headerRow, column: lastErrorNameColumn + column, string: header, format: formatBold)
        }
        for (row, element) in lastErrorData.enumerated() {
            write(worksheet: worksheet, row: dataRow + row, column: lastErrorNameColumn, string: element.name)
            write(worksheet: worksheet, row: dataRow + row, column: lastErrorValueColumn, integer: element.value)
            workbook_define_name(writer.workbook, "LastError\(element.name)", "=\(cell(writer: self, dataRow + row, rowFixed: true, lastErrorValueColumn, columnFixed: true))")
        }
        
        setColumn(worksheet: worksheet, column: lastErrorNameColumn, width: 12)
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
    var chooseBestRow: Int?
    var varianceRow: Int?
    
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
        chooseBestRow = totalRow! + 2
        varianceRow = chooseBestRow! + 1
        
        // Setup sheet
        freezePanes(worksheet: worksheet, row: detailRow!, column: 0)
        
        setColumn(worksheet: worksheet, column: descriptionColumn!, width: 30)
        setColumn(worksheet: worksheet, column: localMPsColumn!, width: 12)
        setColumn(worksheet: worksheet, column: nationalMPsColumn!, width: 12)
        setColumn(worksheet: worksheet, column: checksumColumn!, width: 16)
        
        // Write title row
        write(worksheet: worksheet, row: headerRow!, column: descriptionColumn!, string: "Round", format: formatBold)
        write(worksheet: worksheet, row: headerRow!, column: entryColumn!, string: "Entry", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: tablesColumn!, string: "Tables", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: localNationalColumn!, string: "Nat/Local", format: formatBold)
        write(worksheet: worksheet, row: headerRow!, column: localMPsColumn!, string: "Local MPs", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: nationalMPsColumn!, string: "National MPs", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow!, column: checksumColumn!, string: "Checksum", format: formatRightBold)
        
        // Write round totals
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
        
        // Write underlined row
        for column in [entryColumn, tablesColumn, localNationalColumn, localMPsColumn, nationalMPsColumn, checksumColumn] {
            write(worksheet: worksheet, row: detailRow! + writer.rounds.count, column: column!, string: "", format: formatBoldUnderline)
        }
        
        // Write totals of rounds
        write(worksheet: worksheet, row: totalRow!, column: descriptionColumn!, string: "Round totals", format: formatBold)
        
        writeTotal(column: tablesColumn, format: formatInt)
        writeTotal(column: localMPsColumn)
        writeTotal(column: nationalMPsColumn)
        writeTotal(column: checksumColumn)
        
        // Write exported totals
        let csvImport = writer.csvImport!
        write(worksheet: worksheet, row: exportedRow!, column: descriptionColumn!, string: "Exported totals", format: formatBold)
        write(worksheet: worksheet, row: exportedRow!, column: localMPsColumn!, formula: "=\(cell(writer: csvImport, csvImport.localMpsCell!))", format: formatZeroFloat)
        write(worksheet: worksheet, row: exportedRow!, column: nationalMPsColumn!, formula: "=\(cell(writer: csvImport, csvImport.nationalMpsCell!))", format: formatZeroFloat)
        write(worksheet: worksheet, row: exportedRow!, column: checksumColumn!, formula: "=\(cell(writer: csvImport, csvImport.checksumCell!))", format: formatZeroFloat)
        
        // Write discarded results discrepancy (choose best)
        let consolidated = writer.consolidated!
        let chooseBestCell = cell(writer: csvImport, csvImport.chooseBestRow, rowFixed: true, csvImport.valuesColumn, columnFixed: true)
        write(worksheet: worksheet, row : chooseBestRow!, column: descriptionColumn!, dynamicFormula: "=IF(\(chooseBestCell)<>0,\"Discarded results\",\"\")", format: formatBold)
        for (column, label) in [(localMPsColumn!,    "Local"),
                                (nationalMPsColumn!, "National")] {
            var formula = "=ROUND("
            for index in 0..<writer.rounds.count {
                if index > 0 { formula += "+" }
                formula += "IF(\(cell(writer: consolidated, consolidated.localNationalRow!, rowFixed: true, consolidated.dataColumn! + index, columnFixed: true))=\"\(label)\", SUM(\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, consolidated.dataColumn! + index, columnFixed: true)))),0)"
            }
            formula += "-SUM(\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, column, columnFixed: true)))),2)"
            write(worksheet: worksheet, row: chooseBestRow!, column: column, dynamicFormula: formula, format: formatFloatUnderline)
        }
        
        var formula = "=ROUND("
        for index in 0..<writer.rounds.count {
            if index > 0 { formula += "+" }
            formula += "Checksum(\(vstack)(\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, consolidated.dataColumn! + index, columnFixed: true)))),\(vstack)(\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, consolidated.nationalIdColumn!, columnFixed: true)))))"
        }
        formula += "-("
        for (index, column) in [localMPsColumn!, nationalMPsColumn!].enumerated() {
            if index > 0 { formula += "+" }
            formula += "Checksum(\(vstack)(\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, column, columnFixed: true)))),\(vstack)(\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, consolidated.nationalIdColumn!, columnFixed: true)))))"
        }
        formula += "),2)"
        write(worksheet: worksheet, row: chooseBestRow!, column: checksumColumn!, dynamicFormula: formula, format: formatFloatBoldUnderline)
        
        // Write Variance and highlight if non-zero
        write(worksheet: worksheet, row : varianceRow!, column: descriptionColumn!, string: "Variance", format: formatBold)
        for column in [localMPsColumn, nationalMPsColumn, checksumColumn] {
            formula = "=ROUND("
            for (index, row) in [totalRow, exportedRow, chooseBestRow].enumerated() {
                if index != 0 { formula += "-" }
                formula += "\(cell(row!, rowFixed: true, column!))"
            }
            formula += ",2)"
            write(worksheet: worksheet, row: varianceRow!, column: column!, formula: formula, format: formatZeroFloatBold)
        }
        for column in [localMPsColumn, nationalMPsColumn, checksumColumn] {
            highlightNonZeroVariance(row: varianceRow!, column: column!)
        }
    }
    
    private func writeTotal(column: Int?, format: UnsafeMutablePointer<lxw_format>? = nil) {
        write(worksheet: worksheet, row: totalRow!, column: column!, integerFormula: "=ROUND(SUM(\(cell(detailRow!, rowFixed: true, column!)):\(cell(detailRow! + writer.rounds.count - 1, rowFixed: true, column!))),2)", format: format ?? formatZeroFloat)
    }
    
    private func highlightNonZeroVariance(row: Int, column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let field = "\(cell(row, rowFixed: true, column))"
        let formula = "\(field)<>0"
        setConditionalFormat(worksheet: worksheet, fromRow: row, fromColumn: column, toRow: row, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}

// MARK: - Formatted

class FormattedWriter: WriterBase {
    override var name: String { "Formatted" }
    let titleRow = 0
    let dataRow = 1
    var nameColumn: Int?
    var columns: [FormattedColumn] = []
    var leftRightColumns: [String] = []
    var leftColumns: [String] = []
    var directionColumn: Int?
    var hiddenColumns: [Int] = []
    var chooseBestColumns: [Int] = []
    var categoryColumn: Int?
    var fromStratumColumn: Int?
    
    var formatBottom: UnsafeMutablePointer<lxw_format>?
    var formatLeftRight: UnsafeMutablePointer<lxw_format>?
    var formatLeftRightBottom: UnsafeMutablePointer<lxw_format>?
    var formatLeft: UnsafeMutablePointer<lxw_format>?
    var formatLeftBottom: UnsafeMutablePointer<lxw_format>?
    
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
        workbook_define_name(writer.workbook, "FormattedNameArray", "=\(arrayRef)(\(cell(writer: self, dataRow, rowFixed: true, nameColumn!, columnFixed: true)))")
        workbook_define_name(writer.workbook, "FormattedTitleRow", "=\(cell(writer: self, titleRow, rowFixed: true, 0, columnFixed: true)):\(cell(titleRow, rowFixed: true, columns.count - 1, columnFixed: true))")
        workbook_define_name(writer.workbook, "Printing", "=false")
        
        // Setup rows and page format
        worksheet_set_default_row(worksheet, 25, 0)
        setRow(worksheet: worksheet, row: titleRow, height: 50)
        worksheet_fit_to_pages(worksheet, 1, 0)
        freezePanes(worksheet: worksheet, row: dataRow, column: 0)
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
            } else if column.format == formatBodyCenteredString || column.format == formatBodyCenteredInt {
                bannerFormat = formatBannerCenteredString
                bodyFormat = formatBodyCenteredString
            }
            if column.width != nil || column.aggregateAs != nil {
                setColumn(worksheet: worksheet, column: columnNumber, width: column.width, hidden: column.aggregateAs != nil)
            }
            write(worksheet: worksheet, row: titleRow, column: columnNumber, string: column.title, format: bannerFormat)
                        
            if !column.sortedDynamic {
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
        
        if let categoryColumn = categoryColumn {
            let categoryDataAddress = cell(dataRow, rowFixed: false, categoryColumn, columnFixed: true)
            setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: categoryColumn, toRow: dataRow + writer.maxPlayers - 1, toColumn: categoryColumn, formula: "=(\(categoryDataAddress)=\"Non-SBU\")", stopIfTrue: false, format: formatFaint!)
        }
        
        let titleAddress = cell(titleRow, rowFixed: true, 0, columnFixed: false)
        let nameDataAddress = cell(dataRow, rowFixed: false, nameColumn!, columnFixed: true)
        let nextNameDataAddress = cell(dataRow + 1, rowFixed: false, nameColumn!, columnFixed: true)
        
        var leftRightFormula = "AND(\(nameDataAddress)<>\"\",OR("
        for (columnNumber, column) in leftRightColumns.enumerated() {
            if columnNumber != 0 {
                leftRightFormula += ","
            }
            leftRightFormula += "\(titleAddress)=\"\(column)\""
        }
        leftRightFormula += "))"
        
        var leftFormula = "AND(\(nameDataAddress)<>\"\",OR("
        for (columnNumber, column) in leftColumns.enumerated() {
            if columnNumber != 0 {
                leftFormula += ","
            }
            leftFormula += "\(titleAddress)=\"\(column)\""
        }
        leftFormula += "))"
        
        let csvImport = writer.csvImport!
        let ranksPlusMPs = writer.rounds.first!.ranksPlusMps!
        var bottomFormula = "OR(AND(\(nextNameDataAddress)=\"\",\(nameDataAddress)<>\"\"),AND(Printing,MOD(ROW(\(nameDataAddress)),\(cell(writer: csvImport, csvImport.linesPerPageCell!)))=1)"
        if singleEvent && twoWinners {
            let directionAddress = cell(dataRow, rowFixed: false, directionColumn!, columnFixed: true)
            let nextDirectionAddress = cell(dataRow + 1, rowFixed: false, directionColumn!, columnFixed: true)
            bottomFormula += ",AND(\(nameDataAddress)<>\"\",\(directionAddress)<>\(nextDirectionAddress)))"
        } else if singleEvent && !ranksPlusMPs.round.scoreData.strata.isEmpty {
            let stratumAddress = cell(dataRow, rowFixed: false, fromStratumColumn!, columnFixed: true)
            let nextStratumAddress = cell(dataRow + 1, rowFixed: false, fromStratumColumn!, columnFixed: true)
            bottomFormula += ",AND(\(nameDataAddress)<>\"\",\(stratumAddress)<>\(nextStratumAddress)))"
        } else {
            bottomFormula += ")"
        }
        
        for columnNumber in chooseBestColumns {
            let rounds = writer.rounds.count
            let formattedRange = "\(cell(dataRow, chooseBestColumns.first!, columnFixed: true)):\(cell(dataRow, chooseBestColumns.last!, columnFixed: true)))"
            let consolidated = writer.consolidated!
            let localNationalCell = cell(writer: consolidated, consolidated.localNationalRow!, rowFixed: true, consolidated.dataColumn! + columns[columnNumber].referenceColumn)
            let consolidatedRange = "\(cell(writer: consolidated, consolidated.localNationalRow!, rowFixed: true, consolidated.dataColumn!, columnFixed: true)):\(cell(consolidated.localNationalRow!, rowFixed: true, consolidated.dataColumn! + rounds - 1, columnFixed: true))"
            let csvImport = writer.csvImport!
            let chooseBestCell = cell(writer: csvImport, csvImport.chooseBestRow, rowFixed: true, csvImport.valuesColumn, columnFixed: true)
            setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: columnNumber, toRow: dataRow + writer.maxPlayers - 1, toColumn: columnNumber, formula: "NOT(SumMaxIfIncluded(\(vstack)(\(formattedRange),\(vstack)(\(consolidatedRange)),\(localNationalCell),\(chooseBestCell),\(columns[columnNumber].referenceColumn + 1)))", stopIfTrue: false, format: formatFaint!)
        }
        
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: "AND(\(leftRightFormula),\(bottomFormula))", stopIfTrue: true, format: formatLeftRightBottom!)
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: "AND(\(leftFormula),\(bottomFormula))", stopIfTrue: true, format: formatLeftBottom!)
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: leftRightFormula, stopIfTrue: true, format: formatLeftRight!)
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: leftFormula, stopIfTrue: true, format: formatLeft!)
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: 0, toRow: dataRow + writer.maxPlayers - 1, toColumn: columns.count - 1, formula: bottomFormula, stopIfTrue: true, format: formatBottom!)
        for column in hiddenColumns {
            setColumn(worksheet: worksheet, column: column, hidden: true)
        }
    }
    
    private func setupColumns() {
        let consolidated = writer.consolidated!
        let rounds = writer.rounds
        let round = rounds.first!
        let event = round.scoreData.events.first!
        let winDraw = event.type?.requiresWinDraw ?? false
        var local = false
        var national = false
        var aggregateColumns: [String: [Column]] = [:]
        
        if (singleEvent) {
            var nationalIdColumn: [Int] = []
            var rankColumn: [Int] = []
            
            let ranksPlusMPs = round.ranksPlusMps!
            
            for playerNumber in 0..<round.maxParticipantPlayers {
                columns.append(FormattedColumn(title: "National ID (\(playerNumber + 1))", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.nationalIdColumn[playerNumber])))" }, cellType: .numericFormula, width: 9))
                nationalIdColumn.append(columns.count - 1)
                hiddenColumns.append(columns.count - 1)
            }
            
            for playerNumber in 0..<round.maxParticipantPlayers {
                columns.append(FormattedColumn(title: "Rank (\(playerNumber + 1))", referenceDynamic: { [self] in "=\(fnPrefix)IFNA(\(fnPrefix)NUMBERVALUE(\(fnPrefix)XLOOKUP(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn[playerNumber], columnFixed: true))),\(vstack)(\(lookupRange(Settings.current.userDownloadNationalIdColumn))),\(vstack)(\(lookupRange(Settings.current.userDownloadRankColumn))),,0)),1)" }, cellType: .numericFormula, width: 9))
                rankColumn.append(columns.count - 1)
                hiddenColumns.append(columns.count - 1)
            }
            
            var positionWidth: Float? = 9
            if !ranksPlusMPs.scoreData.strata.isEmpty, let ranksFromStratumCodeColumn = ranksPlusMPs.fromStratumCodeColumn.first {
                
                positionWidth = 7
                columns.append(FormattedColumn(title: "Strata", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksFromStratumCodeColumn)))" }, cellType: .numericFormula, format: formatBodyCenteredInt, width: positionWidth))
                fromStratumColumn = columns.count - 1
                leftColumns.append(columns.last!.title)
                
                columns.append(FormattedColumn(title: "Strata\nPosition", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.stratumPositionColumn.first!)))" }, cellType: .numericFormula, width: positionWidth))
            }

            columns.append(FormattedColumn(title: "\((ranksPlusMPs.scoreData.strata.isEmpty ? "Position" : "Overall\nPosition"))", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.positionColumn!)))" }, cellType: .numericFormula, width: positionWidth))
            if ranksPlusMPs.scoreData.strata.isEmpty {
                leftRightColumns.append(columns.last!.title)
            }
            
            if twoWinners {
                columns.append(FormattedColumn(title: "Direction", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.directionColumn!)))" }, cellType: .numericFormula, width: 9))
                directionColumn = columns.count - 1
                leftRightColumns.append(columns.last!.title)
            }
            columns.append(FormattedColumn(title: (round.maxParticipantPlayers == event.type!.participantType!.players ? "Names" : "Names           *Awards for team members will vary by boards played"), referenceDynamic: { [self] in ranksNamesRef(ranksPlusMPs: ranksPlusMPs) }, cellType: .string, width: Float(round.maxParticipantPlayers) * (18.0 - Float(round.maxParticipantPlayers))))
            nameColumn = columns.count - 1
            leftColumns.append(columns.last!.title)
            
            columns.append(FormattedColumn(title: "Category", referenceDynamic: { [self] in categoryNamesRef(rankColumn: rankColumn) }, cellType: .stringFormula, format: formatBodyCenteredString, width: 10))
            categoryColumn = columns.count - 1
            
            columns.append(FormattedColumn(title: "Score", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.scoreColumn!)))" }, cellType: .numericFormula))
            
            if winDraw {
                columns.append(FormattedColumn(title: "Wins / Draws", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.winDrawColumn[0])))" }, cellType: .numericFormula))
            }

            columns.append(FormattedColumn(title: "\(round.scoreData.national ? "National" : "Local") MPs", referenceDynamic: { [self] in "=\(ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: sourceRef(column: ranksPlusMPs.totalMPColumn[0])))" }, cellType: .floatFormula, width: 10))
            leftRightColumns.append(columns.last!.title)
            
        } else {
            let chooseBest = (rounds.filter({$0.scoreData.aggreateAs != nil}).count == 0)
            
            columns.append(FormattedColumn(title: "National ID", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.nationalIdColumn!))" }, cellType: .numericFormula, width: 9, sortedDynamic: true))
            let nationalIdColumn = columns.count - 1
            hiddenColumns.append(columns.count - 1)
            
            columns.append(FormattedColumn(title: "Rank", referenceDynamic: { [self] in "=\(fnPrefix)IFNA(\(fnPrefix)NUMBERVALUE(\(fnPrefix)XLOOKUP(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn, columnFixed: true))),\(vstack)(\(lookupRange(Settings.current.userDownloadNationalIdColumn))),\(vstack)(\(lookupRange(Settings.current.userDownloadRankColumn))),,0)),1)" }, cellType: .numericFormula, width: 9))
            let rankColumn = columns.count - 1
            hiddenColumns.append(columns.count - 1)
            
            columns.append(FormattedColumn(title: "Name", referenceDynamic: { [self] in "CONCATENATE(\(consolidatedArrayRef(column: consolidated.firstNameColumn!)),\" \",\(consolidatedArrayRef(column: consolidated.otherNamesColumn!)))" }, cellType: .stringFormula, width: 30, sortedDynamic: true))
            nameColumn = columns.count - 1
            leftRightColumns.append(columns.last!.title)
            
            columns.append(FormattedColumn(title: "Category", referenceDynamic: { [self] in "=\(categoryNameRef(rankColumn: rankColumn)))" }, cellType: .stringFormula, format: formatBodyCenteredString, width: 10))
            leftRightColumns.append(columns.last!.title)
            categoryColumn = columns.count - 1
            
            for (roundNumber, round) in rounds.enumerated() {
                let column = FormattedColumn(title: round.shortName.replacingOccurrences(of: " ", with: "\n"), referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.dataColumn! + roundNumber))" }, cellType: .floatFormula, width: 10, aggregateAs: round.scoreData.aggreateAs, roundNumber: roundNumber, sortedDynamic: true)
                if chooseBest {
                    // Note relies on fact that won't happen if later are inserting aggregated columns
                    column.referenceColumn = roundNumber
                    chooseBestColumns.append(columns.count)
                }
                columns.append(column)
                if round.scoreData.national {
                    national = true
                } else {
                    local = true
                }
                if let aggregateAs = round.scoreData.aggreateAs {
                    // Add to list of aggregate columns
                    var columns = aggregateColumns[aggregateAs] ?? []
                    columns.append(column)
                    aggregateColumns[aggregateAs] = columns
                }
            }
            
            for (title, sourceColumns) in aggregateColumns {
                if let position = columns.lastIndex(where: {$0.aggregateAs == title}) {
                    var reference = ""
                    for column in sourceColumns {
                        if reference != "" {
                            reference += "+"
                        }
                        reference += "\(consolidatedArrayRef(column: consolidated.dataColumn! + column.roundNumber!))"
                    }
                    
                    columns.insert(FormattedColumn(title: title.replacingOccurrences(of: " ", with: "\n"), referenceDynamic: { reference }, cellType: .floatFormula, width: 10, sortedDynamic: true), at: position + 1)
                }
            }
            
            if local {
                columns.append(FormattedColumn(title: "Local MPs", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.localMPsColumn!))" }, cellType: .floatFormula, width: 10, sortedDynamic: true))
                leftRightColumns.append(columns.last!.title)
            }
            if national {
                columns.append(FormattedColumn(title: "National MPs", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.nationalMPsColumn!))" }, cellType: .floatFormula, width: 10, sortedDynamic: true))
                leftRightColumns.append(columns.last!.title)
                
            }
            if local && national {
                columns.append(FormattedColumn(title: "Total MPs", referenceDynamic: { [self] in "\(consolidatedArrayRef(column: consolidated.localMPsColumn!))+\(consolidatedArrayRef(column: consolidated.nationalMPsColumn!))" }, cellType: .floatFormula, width: 10, sortedDynamic: true))
                leftRightColumns.append(columns.last!.title)
            }
        }
    }
    
    private func consolidatedArrayRef(column: Int) -> String {
        return "\(arrayRef)(\(cell(writer: writer.consolidated, writer.consolidated.dataRow!, rowFixed: true, column, columnFixed: true)))"
    }
    
    private func ranksArrayRef(ranksPlusMPs: RanksPlusMPsWriter, arrayContent: String) -> String {
        
        var result = "\(sortBy)(\(zeroFiltered(arrayContent: arrayContent)), "
        
        if twoWinners {
            let direction = zeroFiltered(arrayContent: sourceRef(column: ranksPlusMPs.directionColumn!))
            result += "\(direction), -1, "
        } else if ranksPlusMPs.scoreData.strata.count > 0 {
            let fromStratum = zeroFiltered(arrayContent: sourceRef(column: ranksPlusMPs.fromStratumNumberColumn[0]))
            result += "\(fromStratum), -1, "
        }
        
        let position = zeroFiltered(arrayContent: sourceRef(column: ranksPlusMPs.positionColumn!))
        result += "\(position), 1)"
        
        return result
    }
    
    private func ranksNamesRef(ranksPlusMPs: RanksPlusMPsWriter ) -> String {
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
        
        return ranksArrayRef(ranksPlusMPs: ranksPlusMPs, arrayContent: arrayContent)
    }
    
    private func categoryNamesRef(rankColumn: [Int]) -> String {
        let round = writer.rounds.first!
        
        var arrayContent = "CombinedCategory(FALSE,"
        for playerNumber in 0..<round.maxParticipantPlayers {
            if playerNumber != 0 {
                arrayContent += ","
            }
            arrayContent += "\(arrayRef)(\(cell(dataRow, rowFixed: true, rankColumn[playerNumber], columnFixed: true)))"
        }
        arrayContent += ")"
        
        return arrayContent
    }
    
    private func categoryNameRef(rankColumn: Int) -> String {
       return "CombinedCategory(FALSE,\(arrayRef)(\(cell(dataRow, rowFixed: true, rankColumn, columnFixed: true)))"
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
        
        formatLeftRight = workbook_add_format(workbook)
        format_set_left(formatLeftRight, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_right(formatLeftRight, UInt8(LXW_BORDER_THIN.rawValue))
        
        formatLeftRightBottom = workbook_add_format(workbook)
        format_set_left(formatLeftRightBottom, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_right(formatLeftRightBottom, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_bottom(formatLeftRightBottom, UInt8(LXW_BORDER_THIN.rawValue))
        
        formatBottom = workbook_add_format(workbook)
        format_set_bottom(formatBottom, UInt8(LXW_BORDER_THIN.rawValue))
        
        formatLeft = workbook_add_format(workbook)
        format_set_left(formatLeft, UInt8(LXW_BORDER_THIN.rawValue))

        formatLeftBottom = workbook_add_format(workbook)
        format_set_left(formatLeftBottom, UInt8(LXW_BORDER_THIN.rawValue))
        format_set_bottom(formatLeftBottom, UInt8(LXW_BORDER_THIN.rawValue))
    }
}

// MARK: - Export CSV

class CsvImportWriter: WriterBase {
    override var name: String { "Import" }
    var localMpsCell: String?
    var nationalMpsCell: String?
    var linesPerPageCell: String?
    var pageOrientationCell: String?
    var checksumCell: String?
    
    let eventDescriptionRow = 0
    let eventCodeRow = 1
    let eventDateRow = 2
    let customFooterRow = 4
    let linesPerPageRow = 5
    let pageOrientationRow = 6
    let minRankRow = 7
    let maxRankRow = 8
    let clubCodeRow = 9
    let chooseBestRow = 10
    let sortByRow = 12
    let awardsRow = 13
    let localMPsRow = 14
    let nationalMPsRow = 15
    let checksumRow = 16
    
    let titleRow = 18
    let dataRow = 19
    
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
    let lookupPaymentStatusColumn = 17
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        // Define ranges
        workbook_define_name(writer.workbook, "ImportDateArray", "=\(arrayRef)(\(cell(writer: self, dataRow, rowFixed: true, eventDateColumn, columnFixed: true)))")
        workbook_define_name(writer.workbook, "ImportTitleRow", "=\(cell(writer: self, titleRow, rowFixed: true, eventDateColumn, columnFixed: true)):\(cell(titleRow, rowFixed: true, nationalMPsColumn, columnFixed: true))")
        workbook_define_name(writer.workbook, "ImportEventDescriptionCell", "=\(cell(writer: self, eventDescriptionRow, rowFixed: true, valuesColumn, columnFixed: true))")
        workbook_define_name(writer.workbook, "ImportCustomFooterCell", "=\(cell(writer: self, customFooterRow, rowFixed: true, valuesColumn, columnFixed: true))")
        workbook_define_name(writer.workbook, "ImportLinesPerPageCell", "=\(cell(writer: self, linesPerPageRow, rowFixed: true, valuesColumn, columnFixed: true))")
        workbook_define_name(writer.workbook, "ImportPageOrientationCell", "=\(cell(writer: self, pageOrientationRow, rowFixed: true, valuesColumn, columnFixed: true))")
        workbook_define_name(writer.workbook, "EventDateCell", "=\(cell(writer: self, eventDateRow, rowFixed: true, valuesColumn, columnFixed: true))")

        freezePanes(worksheet: worksheet, row: dataRow, column: 0)

        // Add macro buttons
        writer.createMacroButton(worksheet: worksheet, title: "Next Error", macro: "NextSheetErrorRow", row: 11, column: 3)
        writer.createMacroButton(worksheet: worksheet, title: "Create CSV", macro: "CreateCSV", row: 11, column: 5)
        writer.createMacroButton(worksheet: worksheet, title: "Create PDF", macro: "PrintFormatted", row: 14, column: 5)
        if self.writer.includeInRace {
            writer.createMacroButton(worksheet: worksheet, title: "Create Race PDF", macro: "PrintRaceFormatted", row: 11, column: 7, width: 70)
            writer.createMacroButton(worksheet: worksheet, title: "Copy Race Data", macro: "CopyRaceExport", row: 14, column: 7, width: 70)
        }
        
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
        setColumn(worksheet: worksheet, column: lookupOtherUnionColumn, width: 24, format: formatZeroBlank)
        setColumn(worksheet: worksheet, column: lookupHomeClubColumn, width: 25)
        setColumn(worksheet: worksheet, column: lookupRankColumn, width: 8)
        setColumn(worksheet: worksheet, column: lookupEmailColumn, width: 25)
        setColumn(worksheet: worksheet, column: lookupStatusColumn, width: 8)
        setColumn(worksheet: worksheet, column: lookupPaymentStatusColumn, width: 25)
        
        // Create hidden group for parameters
        for row in (customFooterRow - 1)...chooseBestRow {
            setRow(worksheet: worksheet, row: row, group: true, hidden: true, collapsed: true)
        }
        
        // Parameters etc
        let parameters = writer.parameters!
        
        write(worksheet: worksheet, row: eventDescriptionRow, column: titleColumn, string: "Event name:", format: formatBold)
        write(worksheet: worksheet, row: eventDescriptionRow, column: valuesColumn, string: writer.eventDescription)
        
        write(worksheet: worksheet, row: customFooterRow, column: titleColumn, string: "PDF Footer:", format: formatBold)
        write(worksheet: worksheet, row: customFooterRow, column: valuesColumn, string: writer.rounds.first!.scoreData.customFooter) // Assumes single round
        
        write(worksheet: worksheet, row: linesPerPageRow, column: titleColumn, string: "Lines/page:", format: formatBold)
        write(worksheet: worksheet, row: linesPerPageRow, column: valuesColumn, integer: Settings.current.linesPerFormattedPage ?? 32, format: formatString)
        linesPerPageCell = cell(linesPerPageRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: worksheet, row: pageOrientationRow, column: titleColumn, string: "Orientation:", format: formatBold)
        write(worksheet: worksheet, row: pageOrientationRow, column: valuesColumn, string: "Portrait")
        pageOrientationCell = cell(linesPerPageRow, rowFixed: true, valuesColumn, columnFixed: true)
        let orientationValidationRange = "=\(cell(writer: parameters, parameters.dataRow, rowFixed: true, parameters.pageOrientationColumn, columnFixed: true)):\(cell(parameters.dataRow + 1, rowFixed: true, parameters.pageOrientationColumn, columnFixed: true))"
        setDataValidation(row: pageOrientationRow, column: valuesColumn, formula: orientationValidationRange)
        
        write(worksheet: worksheet, row: eventCodeRow, column: titleColumn, string: "Event Code:", format: formatBold)
        write(worksheet: worksheet, row: eventCodeRow, column: valuesColumn, string: writer.eventCode)
        
        write(worksheet: worksheet, row: minRankRow, column: titleColumn, string: "Minimum Rank:", format: formatBold)
        write(worksheet: worksheet, row: minRankRow, column: valuesColumn, formula: "=\(writer.minRank)", format: formatString)
        
        write(worksheet: worksheet, row: maxRankRow, column: titleColumn, string: "Maximum Rank:", format: formatBold)
        write(worksheet: worksheet, row: maxRankRow, column: valuesColumn, formula: "=\(writer.maxRank)", format: formatString)
        
        write(worksheet: worksheet, row: eventDateRow, column: titleColumn, string: "Event Date:", format: formatBold)
        write(worksheet: worksheet, row: eventDateRow, column: valuesColumn, floatFormula: "=\(writer.maxEventDate)", format: formatDate)
        
        write(worksheet: worksheet, row: clubCodeRow, column: titleColumn, string: "Club Code:", format: formatBold)
        write(worksheet: worksheet, row: clubCodeRow, column: valuesColumn, string: writer.clubCode, format: formatString)
        
        write(worksheet: worksheet, row: chooseBestRow, column: titleColumn, string: "Choose best:", format: formatBold)
        write(worksheet: worksheet, row: chooseBestRow, column: valuesColumn, integer: writer.chooseBest, format: formatString)
        
        write(worksheet: worksheet, row: sortByRow, column: titleColumn, string: "Sort by:", format: formatBold)
        let sortData = parameters.sortData
        let sortValidationRange = "=\(cell(writer: parameters, parameters.dataRow, rowFixed: true, parameters.sortNameColumn, columnFixed: true)):\(cell(parameters.dataRow + sortData.count - 1, rowFixed: true, parameters.sortNameColumn, columnFixed: true))"
        setDataValidation(row: sortByRow, column: valuesColumn, formula: sortValidationRange)
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
        writeLookup(title: "Payment status", column: lookupPaymentStatusColumn, lookupColumn: Settings.current.userDownloadPaymentStatusColumn)

        highlightBadEventDate()
        
        highlightLookupDifferent(columns: [firstNameColumn, otherNamesColumn, lookupFirstNameColumn, lookupOtherNamesColumn], column: firstNameColumn, offset: 0, format: formatYellow)
        highlightLookupDifferent(columns: [firstNameColumn, otherNamesColumn, lookupFirstNameColumn, lookupOtherNamesColumn], column: otherNamesColumn, offset: 1, format: formatRed)
        highlightLookupError(fromColumn: lookupFirstNameColumn, toColumn: lookupPaymentStatusColumn)
        highlightBadNationalId(column: nationalIdColumn, firstNameColumn: firstNameColumn)
        highlightBadDate(column: eventDateColumn, firstNameColumn: firstNameColumn)
        highlightBadMPs(column: localMPsColumn, firstNameColumn: firstNameColumn)
        highlightBadMPs(column: nationalMPsColumn, firstNameColumn: firstNameColumn)
        highlightNoHomeClub(column: lookupHomeClubColumn)
        highlightBadRank(column: lookupRankColumn)
        highlightBadStatus(column: lookupStatusColumn)
        highlightBadPaymentStatus(column: lookupPaymentStatusColumn)
    }
    
    private func sourceArray(_ column: Int) -> String {
        let consolidated = writer.consolidated!
        return "\(arrayRef)(\(cell(writer: consolidated, consolidated.dataRow!, rowFixed: true, column, columnFixed: true)))"
    }
    
    private func writeLookup(title: String, column: Int, lookupColumn: String, format: UnsafeMutablePointer<lxw_format>? = nil, numeric: Bool = false) {
        write(worksheet: worksheet, row: titleRow, column: column, string: title, format: format ?? formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow, column: column, dynamicFormula: "=\(numeric ? "\(fnPrefix)NUMBERVALUE(" : "")\(fnPrefix)XLOOKUP(\(arrayRef)(\(cell(dataRow, rowFixed: true, nationalIdColumn, columnFixed: true))),\(vstack)(\(lookupRange(Settings.current.userDownloadNationalIdColumn))),\(vstack)(\(lookupRange(lookupColumn))),,0)\(numeric ? ")" : "")")
    }
    
    private func highlightLookupDifferent(columns: [Int], column: Int, offset: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        var cells: [String] = []
        for index in 0..<columns.count {
            cells.append("StripName(\(cell(dataRow,columns[index],columnFixed:true)))")
        }
        let formula = "=(AND(CONCATENATE(\(cells[0]),\(cells[1]))<>CONCATENATE(\(cells[2]),\(cells[3])),\(cells[offset])<>\(cells[offset+2])))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightLookupError(fromColumn: Int, toColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let lookupCell = "\(cell(dataRow, fromColumn, columnFixed: true))"
        let formula = "=ISNA(\(lookupCell))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: fromColumn, toRow: dataRow + writer.maxPlayers - 1, toColumn: toColumn, formula: formula, format: format ?? formatGrey!)
    }
    
    private func highlightNoHomeClub(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let nationalIdCell = cell(dataRow, nationalIdColumn, columnFixed: true)
        let statusCell = cell(dataRow, column, columnFixed: true)
        let formula = "AND(\(nationalIdCell)<>\"\", \(statusCell)=0)"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatNoHomeClub!)
    }
    
    private func highlightBadStatus(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let statusCell = "\(cell(dataRow, column, columnFixed: true))"
        let formula = "=AND(\(statusCell)<>\"\(Settings.current.goodStatus!)\", \(statusCell)<>\"\")"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatNotActive!)
    }

    private func highlightBadPaymentStatus(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let paymentStatusCell = "\(cell(dataRow, column, columnFixed: true))"
        let eventDateCell = cell(eventDateRow, rowFixed: true, valuesColumn, columnFixed: true)
        let formula = "=AND(\(paymentStatusCell)=\"\(Settings.current.notPaidPaymentStatus!)\", \(paymentStatusCell)<>\"\", OR(MONTH(\(eventDateCell))>=\(Settings.current.ignorePaymentTo! + 1),MONTH(\(eventDateCell))<=\(Settings.current.ignorePaymentFrom!-1)))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatNotPaid!)
    }
    
    private func highlightBadNationalId(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let nationalIdCell = cell(dataRow, column, columnFixed: true)
        let firstNameCell = cell(dataRow, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(nationalIdCell)<=0, \(nationalIdCell)>\(Settings.current.maxNationalIdNumber!), NOT(ISNUMBER(\(nationalIdCell)))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, checkDuplicates: true, format: format ?? formatGrey!, duplicateFormat: formatRedHatched!)
    }
    
    private func highlightBadMPs(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let pointsCell = cell(dataRow, column, columnFixed: true)
        let firstNameCell = cell(dataRow, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(pointsCell)>\(Settings.current.maxPoints!), AND(\(pointsCell)<>\"\", NOT(ISNUMBER(\(pointsCell))))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadEventDate(format: UnsafeMutablePointer<lxw_format>? = nil) {
        let eventDateCell = cell(eventDateRow, rowFixed: true, valuesColumn, columnFixed: true)
        let formula = "=OR(INT(\(eventDateCell))>TODAY(), INT(\(eventDateCell))<(TODAY()-8))"
        setConditionalFormat(worksheet: worksheet, fromRow: eventDateRow, fromColumn: valuesColumn, toRow: eventDateRow, toColumn: valuesColumn, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadDate(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let eventDateCell = cell(eventDateRow, rowFixed: true, valuesColumn, columnFixed: true)
        let dateCell = cell(dataRow, column, columnFixed: true)
        let firstNameCell = cell(dataRow, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\",OR(INT(\(dateCell))>TODAY(), INT(\(dateCell))<(TODAY()-8)), \(dateCell)<>\(eventDateCell))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadRank(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let rankCell = cell(dataRow, column, columnFixed: true)
        let formula = "=AND(\(rankCell)<>\"\", OR(\(rankCell)<\(cell(minRankRow, rowFixed: true, valuesColumn, columnFixed: true)), \(rankCell)>\(cell(maxRankRow, rowFixed: true, valuesColumn, columnFixed: true))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow, fromColumn: column, toRow: dataRow + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}
    
// MARK: - RaceExportWriter

class RaceDetailWriter: WriterBase {
    override var name: String { "Race Detail" }
    
    let sectionRow = 0
    let titleRow = 1
    let dataRow = 2
    
    let sectionCategoryColumnOffset = 0
    
    let eventNameColumnOffset = 0
    let categoryColumnOffset = 1
    let nationalIdColumnOffset = 2
    let positionColumnOffset = 3
    let racePositionColumnOffset = 4
    let playerNameColumnOffset = 5
    let playerRankColumnOffset = 6
    let racePointsColumnOffset = 7
    let sectionColumns = 8
    
    var sectionCategoryColumn: [RankCategory:Int] = [:]
    var eventNameColumn: [RankCategory:Int] = [:]
    var categoryColumn: [RankCategory:Int] = [:]
    var nationalIdColumn: [RankCategory:Int] = [:]
    var positionColumn: [RankCategory:Int] = [:]
    var racePositionColumn: [RankCategory:Int] = [:]
    var playerNameColumn: [RankCategory:Int] = [:]
    var playerRankColumn: [RankCategory:Int] = [:]
    var racePointsColumn: [RankCategory:Int] = [:]
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        let rounds = writer.rounds
        let round = rounds.first!
        let event = round.scoreData.events.first!
        
        freezePanes(worksheet: worksheet, row: dataRow, column: 0)
        
        let individual = writer.rounds.first!.individualMPs!
        let indivNationalIdArray = "\(arrayRef)(\(cell(writer: individual, individual.dataRow, rowFixed: true, individual.nationalIdColumn!, columnFixed: true)))"
        let indivPlayerCategoryArray = "\(arrayRef)(\(cell(writer: individual, individual.dataRow, rowFixed: true, individual.playerCategoryColumn!, columnFixed: true)))"
        let indivPlayerRaceCategoryArray = "\(arrayRef)(\(cell(writer: individual, individual.dataRow, rowFixed: true, individual.playerRaceCategoryColumn!, columnFixed: true)))"
        let indivPositionArray = "\(arrayRef)(\(cell(writer: individual, individual.dataRow, rowFixed: true, individual.positionColumn!, columnFixed: true)))"
        let indivFirstNameArray = "\(arrayRef)(\(cell(writer: individual, individual.dataRow, rowFixed: true, individual.firstNameColumn!, columnFixed: true)))"
        let indivOtherNamesArray = "\(arrayRef)(\(cell(writer: individual, individual.dataRow, rowFixed: true, individual.otherNamesColumn!, columnFixed: true)))"
        
        for (index, category) in RankCategory.allCases.enumerated() {
            let startColumn = index * (sectionColumns + 1)
            // Format rows/columns
            setColumn(worksheet: worksheet, column: startColumn + eventNameColumnOffset, width: 30)
            setColumn(worksheet: worksheet, column: startColumn + playerNameColumnOffset, width: 16)
            
            let sectionCell = cell(sectionRow, rowFixed: true, startColumn + sectionCategoryColumnOffset, columnFixed: true)
            let filterLogic = "\(indivPlayerRaceCategoryArray)=\(sectionCell)"
            let sortLogic = "\(filter)(\(indivPositionArray),\(filterLogic),\"\")"
            
            // Section cell
            write(worksheet: worksheet, row: sectionRow, column: startColumn + sectionCategoryColumnOffset, string: category.string, format: formatBold)
            sectionCategoryColumn[category] = startColumn + sectionCategoryColumnOffset
            
            // Data
            let nationalIdArray = "\(arrayRef)(\(cell(dataRow, rowFixed: true, startColumn + nationalIdColumnOffset, columnFixed: true)))"
            
            // Event name
            write(worksheet: worksheet, row: titleRow, column: startColumn + eventNameColumnOffset, string: "Event", format: formatBoldUnderline)
            write(worksheet: worksheet, row: dataRow, column: startColumn + eventNameColumnOffset, dynamicFormula: "=\(fnPrefix)MAKEARRAY(ROWS(\(nationalIdArray)),1,\(lambda)(\(lambdaParam)row,\(lambdaParam)column,ImportEventDescriptionCell))")
            eventNameColumn[category] = startColumn + eventNameColumnOffset
            
            // Category
            write(worksheet: worksheet, row: titleRow, column: startColumn + categoryColumnOffset, string: "\(event.type?.participantType?.string ?? "Team") Category", format: formatBoldUnderline)
            write(worksheet: worksheet, row: dataRow, column: startColumn + categoryColumnOffset, dynamicFormula: "=\(fnPrefix)MAKEARRAY(ROWS(\(nationalIdArray)),1,\(lambda)(\(lambdaParam)row,\(lambdaParam)column,\(sectionCell)))")
            categoryColumn[category] = startColumn + categoryColumnOffset
            
            // National Id
            write(worksheet: worksheet, row: titleRow, column: startColumn + nationalIdColumnOffset, string: "SBU", format: formatBoldUnderline)
            write(worksheet: worksheet, row: dataRow, column: startColumn + nationalIdColumnOffset, dynamicFormula: "=\(sortBy)(\(filter)(\(indivNationalIdArray),\(filterLogic),\"\"), \(sortLogic), 1)")
            nationalIdColumn[category] = startColumn + nationalIdColumnOffset
            
            // Overall Position
            write(worksheet: worksheet, row: titleRow, column: startColumn + positionColumnOffset, string: "Overall position", format: formatBoldUnderline)
            write(worksheet: worksheet, row: dataRow, column: startColumn + positionColumnOffset, dynamicFormula: "=\(sortBy)(\(sortLogic), \(sortLogic), 1)")
            positionColumn[category] = startColumn + positionColumnOffset
            
            // Race Position
            write(worksheet: worksheet, row: titleRow, column: startColumn + racePositionColumnOffset, string: "Race Position", format: formatBoldUnderline)
            let positionArray = "\(arrayRef)(\(cell(dataRow, rowFixed: true, positionColumn[category]!, columnFixed: true)))"
            write(worksheet: worksheet, row: dataRow, column: startColumn + racePositionColumnOffset, dynamicFormula: "=\(fnPrefix)MAKEARRAY(ROWS(\(nationalIdArray)),1,\(lambda)(\(lambdaParam)row,\(lambdaParam)column,IF(RelativeTo(\(nationalIdArray),\(lambdaParam)row,\(lambdaParam)column)=\"\",\"\",\(fnPrefix)RANK.EQ(RelativeTo(\(positionArray),\(lambdaParam)row,\(lambdaParam)column),\(positionArray),1))))")
            racePositionColumn[category] = startColumn + racePositionColumnOffset
            
            // Player Name
            write(worksheet: worksheet, row: titleRow, column: startColumn + playerNameColumnOffset, string: "Name", format: formatBoldUnderline)
            write(worksheet: worksheet, row: dataRow, column: startColumn + playerNameColumnOffset, dynamicFormula: "=\(sortBy)(\(filter)(CONCATENATE(\(indivFirstNameArray),\" \",\(indivOtherNamesArray)),\(filterLogic),\"\"), \(sortLogic), 1)")
            playerNameColumn[category] = startColumn + playerNameColumnOffset
            
            // Actual Rank
            write(worksheet: worksheet, row: titleRow, column: startColumn + playerRankColumnOffset, string: "Actual Category", format: formatBoldUnderline)
            write(worksheet: worksheet, row: dataRow, column: startColumn + playerRankColumnOffset, dynamicFormula: "=\(sortBy)(\(filter)(\(indivPlayerCategoryArray),\(filterLogic),\"\"), \(sortLogic), 1)")
            playerRankColumn[category] = startColumn + playerRankColumnOffset
            
            // Race Points
            write(worksheet: worksheet, row: titleRow, column: startColumn + racePointsColumnOffset, string: "Race Points", format: formatBoldUnderline)
            let racePositionArray = "\(arrayRef)(\(cell(dataRow, rowFixed: true, racePositionColumn[category]!, columnFixed: true)))"
            write(worksheet: worksheet, row: dataRow, column: startColumn + racePointsColumnOffset, dynamicFormula: "=\(fnPrefix)MAKEARRAY(ROWS(\(nationalIdArray)),1,\(lambda)(\(lambdaParam)row,\(lambdaParam)column,IF(RelativeTo(\(nationalIdArray),\(lambdaParam)row,\(lambdaParam)column)=\"\",\"\",MAX(0,11-RelativeTo(\(racePositionArray),\(lambdaParam)row,\(lambdaParam)column)))))")
            racePointsColumn[category] = startColumn + racePointsColumnOffset
        }
    }
}


// MARK: - RaceExportWriter

class RaceExportWriter: WriterBase {
    override var name: String { "Race Export" }
    
    let titleRow = 0
    let dataRow = 1
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
    }
    
    func write() {
        let raceDetail = writer.raceDetail!
        
        workbook_define_name(writer.workbook, "RaceExportFirstArray", "=\(arrayRef)(\(cell(writer: self, dataRow, rowFixed: true, 0, columnFixed: true)))")
        workbook_define_name(writer.workbook, "RaceExportDataRow", "=\(cell(writer: self, dataRow, rowFixed: true, 0, columnFixed: true)):\(cell(dataRow, rowFixed: true, raceDetail.sectionColumns - 1, columnFixed: true))")
        
        freezePanes(worksheet: worksheet, row: dataRow, column: 0)
        
        setRow(worksheet: worksheet, row: titleRow, height: 30)
        setColumn(worksheet: worksheet, column: raceDetail.eventNameColumnOffset, width: 30)
        setColumn(worksheet: worksheet, column: raceDetail.playerNameColumnOffset, width: 16)
        
        for column in 0..<raceDetail.sectionColumns {
            let titleCell = cell(writer: raceDetail, raceDetail.titleRow, rowFixed: true, column, columnFixed: true)
            write(worksheet: worksheet, row: titleRow, column: column, dynamicFormula: "=\(titleCell)", format: formatBoldUnderline)
            
            var logic = "\(vstack)("
            for (index, category) in RankCategory.allCases.enumerated() {
                if index != 0 {
                    logic += ","
                }
                let filterLogic = "\(arrayRef)(\(cell(writer: raceDetail, raceDetail.dataRow, rowFixed: true, raceDetail.racePointsColumn[category]!, columnFixed: true)))<>0"
                let detailCell = cell(writer: raceDetail, raceDetail.dataRow, rowFixed: true, (index * (raceDetail.sectionColumns + 1)) + column, columnFixed: true)
                let dataCell = "\(filter)(\(arrayRef)(\(detailCell)),\(filterLogic))"
                logic += "\(dataCell)"
            }
            logic += ")"
            write(worksheet: worksheet, row: dataRow, column: column, dynamicFormula: logic)
        }
    }
}

// MARK: - Race Formatted

enum FormatType: CaseIterable {
    case largeCenter
    case center
    case left
    
    var fontSize: Double {
        switch self {
        case .largeCenter: 18
        default: 11
        }
    }
    
    var alignment: UInt8 {
        switch self {
        case .left: UInt8(LXW_ALIGN_LEFT.rawValue)
        default: UInt8(LXW_ALIGN_CENTER.rawValue)
        }
    }
}

class RaceFormattedWriter: WriterBase {
    override var name: String { "Race Formatted" }
    var raceDetail: RaceDetailWriter!
    let categoryRow = 0
    let titleRow = 1
    let blankRow = 2
    let dataRow = 3
    let categoryRows = 18
    let blankLeadingColumn = 0
    let racePositionColumn = 1
    let positionColumn = 2
    let playerNameColumn = 3
    let playerRankColumn = 4
    let pointsColumn = 5
    let blankTrailingColumn = 6
    
    var format: [RankCategory:[FormatType:UnsafeMutablePointer<lxw_format>?]] = [:]
    
    init(writer: Writer) {
        super.init(writer: writer)
        self.workbook = workbook
        raceDetail = writer.raceDetail
    }
    
    func write() {
        
        setupFormattedFormats()
        let rounds = writer.rounds
        let round = rounds.first!
        let event = round.scoreData.events.first!
        let raceDetail = writer.raceDetail!
        
        // Format sheet
        worksheet_fit_to_pages(worksheet, 1, 0)
        worksheet_center_horizontally(worksheet)
        worksheet_set_header(worksheet, "&C&20Race to Aviemore - \(writer.eventDescription)")
        
        // Setup Columns
        for column in (blankLeadingColumn...blankTrailingColumn) {
            if column == playerNameColumn {
                setColumn(worksheet: worksheet, column: column, width: 20, format: formatBodyString!)
            } else if column == blankLeadingColumn || column == blankTrailingColumn {
                setColumn(worksheet: worksheet, column: column, width: 2, format: formatBodyString!)
            } else {
                setColumn(worksheet: worksheet, column: column, format: formatBodyCenteredString!)
            }
        }
        
        for (index, category) in RankCategory.allCases.enumerated() {
            
            let startColumn = index * (raceDetail.sectionColumns + 1)
            
                      
            let baseRow = (index * categoryRows)
            
            // Setup Rows
            setRow(worksheet: worksheet, row: baseRow + categoryRow, height: 24)
            setRow(worksheet: worksheet, row: baseRow + titleRow, height: 32)
            setRow(worksheet: worksheet, row: baseRow + blankRow, height: 8)
            
            for column in (blankLeadingColumn...blankTrailingColumn) {
            
                write(worksheet: worksheet, row: baseRow + categoryRow, column: column, string: (column == playerNameColumn ? "\(category.string) \(event.type!.participantType!.string)s" : ""), format: format[category]![.largeCenter]!)
                
                write(worksheet: worksheet, row: baseRow + titleRow, column: column, string: "" , format: format[category]![.center]!)
                
                write(worksheet: worksheet, row: baseRow + blankRow, column: column, string: "" , format: format[category]![.center]!)
            }
            
            
            writeElement(title: "\(category.string) Position", category: category, sourceColumnOffset: startColumn + raceDetail.racePositionColumnOffset, sourcePointsColumnOffset: startColumn + raceDetail.racePointsColumnOffset, titleRow: baseRow + titleRow, dataRow: baseRow + dataRow, column: racePositionColumn, bodyFormat: formatBodyCenteredString)
            
            writeElement(title: "Overall Position", category: category, sourceColumnOffset: startColumn + raceDetail.positionColumnOffset, sourcePointsColumnOffset: startColumn + raceDetail.racePointsColumnOffset, titleRow: baseRow + titleRow, dataRow: baseRow + dataRow, column: positionColumn, bodyFormat: formatBodyCenteredString)
            
            writeElement(title: "Name", category: category, sourceColumnOffset: startColumn + raceDetail.playerNameColumnOffset, sourcePointsColumnOffset: startColumn + raceDetail.racePointsColumnOffset, titleRow: baseRow + titleRow, dataRow: baseRow + dataRow, column: playerNameColumn, bodyFormat: formatBodyString)
            
            writeElement(title: "Actual Rank", category: category, sourceColumnOffset: startColumn + raceDetail.playerRankColumnOffset, sourcePointsColumnOffset: startColumn + raceDetail.racePointsColumnOffset, titleRow: baseRow + titleRow, dataRow: baseRow + dataRow, column: playerRankColumn, bodyFormat: formatBodyCenteredString)
            
            writeElement(title: "Race Points", category: category, sourceColumnOffset: startColumn + raceDetail.racePointsColumnOffset, sourcePointsColumnOffset: startColumn + raceDetail.racePointsColumnOffset, titleRow: baseRow + titleRow, dataRow: baseRow + dataRow, column: pointsColumn, bodyFormat: formatBodyCenteredString)
        }
    }
    
    func writeElement(title: String, category: RankCategory, sourceColumnOffset: Int, sourcePointsColumnOffset: Int, titleRow: Int, dataRow: Int, column: Int, bodyFormat: UnsafeMutablePointer<lxw_format>?) {
        write(worksheet: worksheet, row: titleRow, column: column, string: title, format: format[category]![.center]!)
        let sourceArray = "\(arrayRef)(\(cell(writer: raceDetail, raceDetail.dataRow, rowFixed: true, sourceColumnOffset, columnFixed: true)))"
        let filterLogic = "\(arrayRef)(\(cell(writer: raceDetail, raceDetail.dataRow, rowFixed: true, sourcePointsColumnOffset, columnFixed: true)))>0"
        write(worksheet: worksheet, row: dataRow, column: column, dynamicFormula: "\(filter)(\(sourceArray),\(filterLogic))", format: bodyFormat!)
    }
    
    func setupFormattedFormats() {
        for category in RankCategory.allCases {
            format[category] = [:]
            for formatType in FormatType.allCases {
                format[category]![formatType] = workbook_add_format(workbook)
                format_set_align(format[category]![formatType]!, formatType.alignment)
                format_set_align(format[category]![formatType]!, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
                format_set_bold(format[category]![formatType]!)
                format_set_pattern (format[category]![formatType]!, UInt8(LXW_PATTERN_SOLID.rawValue))
                format_set_bg_color(format[category]![formatType]!, lxw_color_t(category.backgroundColor))
                format_set_font_color(format[category]![formatType]!, lxw_color_t(category.textColor))
                format_set_font_size(format[category]![formatType]!, formatType.fontSize)
                format_set_text_wrap(format[category]![formatType]!)
            }
        }
    }
}

// MARK: - Consolidated
 
class ConsolidatedWriter: WriterBase {
    override var name: String { "Consolidated" }
    var localNationalRow: Int?
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
        localNationalRow = 0
        let totalRow = 1
        let checksumRow = 2
        dataRow = 4
        
        freezePanes(worksheet: worksheet, row: dataRow!, column: dataColumn!)
        
        let localNationalRange = "\(cell(localNationalRow!, rowFixed: true, dataColumn!, columnFixed: true)):\(cell(localNationalRow!, rowFixed: true, dataColumn! + writer.rounds.count - 1, columnFixed: true))"
        
        setRow(worksheet: worksheet, row: titleRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: localNationalRow!, format: formatBoldUnderline)
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
            write(worksheet: worksheet, row: localNationalRow!, column: dataColumn! + column, formula: "=IF(\(cell)=0,\"\",\(cell))", format: formatRightBoldUnderline)
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
            formula += "\(filter)("
            formula += "\(arrayRef)(\(cell(writer: round.individualMPs, 1,rowFixed: true, round.individualMPs.uniqueColumn!, columnFixed: true)))"
            formula += ",\(arrayRef)(\(cell(writer: round.individualMPs, 1,rowFixed: true, round.individualMPs.decimalColumn!, columnFixed: true)))<>0)"
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
        
        let csvImport = writer.csvImport!
        let chooseBestCell = cell(writer: csvImport, csvImport.chooseBestRow, rowFixed: true, csvImport.valuesColumn, columnFixed: true)
        write(worksheet: worksheet, row: dataRow!, column: localMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SumMaxIf(\(vstack)(\(lambdaParam)),\(vstack)(\(localNationalRange)), \"Local\", \(chooseBestCell))))")
        write(worksheet: worksheet, row: dataRow!, column: nationalMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SumMaxIf(\(vstack)(\(lambdaParam)),\(vstack)(\(localNationalRange)), \"National\", \(chooseBestCell))))")
        
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
    internal var scoreData: ScoreData!
    
    override var name: String { "Source" }
    private var columns: [RanksPlusMPsColumn] = []
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
    var frozenRankColumn: [Int] = []
    var strataRankColumn: [Int] = []
    var boardsPlayedColumn: [Int] = []
    var winDrawColumn: [Int] = []
    var bonusMPColumn: [Int] = []
    var winDrawMPColumn: [Int] = []
    var strataMPColumn: [[Int]] = []
    var maxStrataMPColumn: [Int] = []
    var scoreColumn: Int?
    var totalMPColumn: [Int] = []
    var stratumColumn: Int?
    var fromStratumNumberColumn: [Int] = []
    var fromStratumCodeColumn: [Int] = []
    var stratumPositionColumn: [Int] = []
    var frozenRankCategoryColumn: Int?


    var eventDescriptionCell: String?
    var entryCell: String?
    var tablesCell: String?
    var maxAwardCell: String?
    var ewMaxAwardCell: String?
    var minEntryCell: String?
    var awardToCell: String?
    var ewAwardToCell: String?
    var awardPercentCell: String?
    var perWinCell: String?
    var eventDateCell: String?
    var eventIdCell: String?
    var localCell: String?
    var localMPsCell :String?
    var nationalMPsCell :String?
    var checksumCell :String?
    var strataCodeCell: [String] = []
    var strataRankCell: [String] = []
    var strataPercentCell: [String] = []
    var columnWidth: [Int:Float] = [:]
    
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
        setRow(worksheet: worksheet, row: headerTitleRow, height: 30)
        setRow(worksheet: worksheet, row: dataTitleRow, height: 45)
                
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBoldUnderline
            if column.cellType == .integer || column.cellType == .float || column.cellType == .numeric || column.cellType == .integerFormula || column.cellType == .floatFormula || column.cellType == .numericFormula {
                format = formatRightBoldUnderline
            }
            if let width = column.width {
                columnWidth[columnNumber] = max(columnWidth[columnNumber] ?? 0, width)
            }
            write(worksheet: worksheet, row: dataTitleRow, column: columnNumber, string: round.replace(column.title), format: format)
        }
        
        for columnNumber in 0..<max(headerColumns, columns.count) {
            if let width = columnWidth[columnNumber] {
                setColumn(worksheet: worksheet, column: columnNumber, width: width)
            } else {
                setColumn(worksheet: worksheet, column: columnNumber, width: 6.5)
            }
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
                            if let strataNumber = column.strataNumber, let strataContent = column.strataContent?(playerNumber, strataNumber, rowNumber) {
                                write(cellType: (strataContent == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: strataContent)
                            }
                        } else {
                            write(cellType: .string, worksheet: worksheet, row: rowNumber, column: columnNumber, content: "")
                        }
                    }
                }
                
                if column.playerNumber == nil, let strataNumber = column.strataNumber, let strataContent = column.strataContent?(strataNumber, nil, rowNumber) {
                    write(cellType: (strataContent == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: strataContent)
                }
                
                if let calculatedContent = column.calculatedContent?(column.playerNumber, rowNumber) {
                    write(cellType: (calculatedContent == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: calculatedContent)
                }
            }
        }
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
        
        columns.append(RanksPlusMPsColumn(title: "Place", content: { (participant, _) in "\(participant.place ?? 0)" }, cellType: .integer))
        positionColumn = columns.count - 1
        
        if event.winnerType == 2 && event.type?.participantType == .pair {
            columns.append(RanksPlusMPsColumn(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
            directionColumn = columns.count - 1
        }
        
        columns.append(RanksPlusMPsColumn(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer)) ; participantNoColumn = columns.count - 1
        
        columns.append(RanksPlusMPsColumn(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float)); scoreColumn = columns.count - 1
        
        if !self.scoreData.strata.isEmpty {
            // Pre-allocate stratum column so that we can reference it
            // so that we can reference it before we calculate it
            columns.append(RanksPlusMPsColumn(title: "")); stratumColumn = columns.count - 1
        }
        
        if winDraw && playerCount <= event.type?.participantType?.players ?? playerCount {
            columns.append(RanksPlusMPsColumn(title: "Win / Draw", content: { (participant, _) in "\(participant.winDraw ?? 0)" }, cellType: .float))
            winDrawColumn.append(columns.count - 1)
        }
        
        for playerNumber in 0..<playerCount {
            columns.append(RanksPlusMPsColumn(title: "First Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name ?? "Unknown Unknown", element: 0) }, playerNumber: playerNumber, cellType: .string, width: 12))
            firstNameColumn.append(columns.count - 1)
            
            columns.append(RanksPlusMPsColumn(title: "Other Names (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name ?? "Unknown Unknown", element: 1) }, playerNumber: playerNumber, cellType: .string, width: 12))
            otherNamesColumn.append(columns.count - 1)
            
            columns.append(RanksPlusMPsColumn(title: "SBU No (\(playerNumber+1))", playerContent: playerNationalId, playerNumber: playerNumber, cellType: .integer, width: 8))
            nationalIdColumn.append(columns.count - 1)
            
            if self.writer.includeInRace {
                columns.append(RanksPlusMPsColumn(title: "Frozen Rank (\(playerNumber+1))", playerContent: playerFrozenRank, playerNumber: playerNumber, cellType: .integerFormula))
                frozenRankColumn.append(columns.count - 1)
            }
            
            if !self.scoreData.strata.isEmpty {
                columns.append(RanksPlusMPsColumn(title: "Current Rank (\(playerNumber+1))", playerContent: playerCurrentRank, playerNumber: playerNumber, cellType: .integerFormula))
                strataRankColumn.append(columns.count - 1)
                
            }
            
            if !scoreData.manualMPs {
                
                if playerCount > event.type?.participantType?.players ?? playerCount {
                    if event.type != .head_to_head {
                        columns.append(RanksPlusMPsColumn(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in
                            "\(player.boardsPlayed ?? event.boards ?? 1)"
                        }, playerNumber: playerNumber, cellType: .integer))
                        boardsPlayedColumn.append(columns.count - 1)
                    }
                    
                    if winDraw {
                        columns.append(RanksPlusMPsColumn(title: "Win / Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                        winDrawColumn.append(columns.count - 1)
                    }
                    
                    if event.type != .head_to_head {
                        addBonusMPColumns(playerNumber: playerNumber)
                    }
                    
                    if !winDraw {
                        totalMPColumn.append(columns.count - 1)
                    } else {
                        if event.type != .head_to_head {
                            bonusMPColumn.append(columns.count - 1)
                        }
                        columns.append(RanksPlusMPsColumn(title: "Win / Draw MP (\(playerNumber+1))", calculatedContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .floatFormula))
                        winDrawMPColumn.append(columns.count - 1)
                        
                        if event.type != .head_to_head {
                            columns.append(RanksPlusMPsColumn(title: "Total MP (\(playerNumber+1))", calculatedContent: playerTotalAward, playerNumber: playerNumber, cellType: .floatFormula))
                        }
                        totalMPColumn.append(columns.count - 1)
                    }
                }
            }
        }
        
        if self.writer.includeInRace {
            columns.append(RanksPlusMPsColumn(title: "\(event.type?.participantType?.string ?? "Team") Category", calculatedContent: frozenRankCategory, cellType: .stringFormula))
            frozenRankCategoryColumn = columns.count - 1
        }
        
        if scoreData.manualMPs {
            
            columns.append(RanksPlusMPsColumn(title: "Total MPs", content: { (participant, _) in "\(participant.manualMps ?? 0)" }, cellType: .floatFormula))
            for _ in 0..<playerCount {
                totalMPColumn.append(columns.count - 1)
            }
            
        } else {
            
            if !self.scoreData.strata.isEmpty {
                // Note column already allocated
                columns[stratumColumn!] = (RanksPlusMPsColumn(title: "\(event.type?.participantType?.string ?? "Team") Stratum", calculatedContent: stratum, cellType: .integerFormula, format: formatZeroInt))
            }
            
            if playerCount <= event.type?.participantType?.players ?? playerCount {
                
                addBonusMPColumns()
                
                if !winDraw {
                    for _ in 0..<playerCount {
                        totalMPColumn.append(columns.count - 1)
                    }
                } else {
                    bonusMPColumn.append(columns.count - 1)
                    
                    columns.append(RanksPlusMPsColumn(title: "Win / Draw MP", calculatedContent: winDrawAward, cellType: .floatFormula))
                    winDrawMPColumn.append(columns.count - 1)
                    
                    columns.append(RanksPlusMPsColumn(title: "Total MP", calculatedContent: totalAward, cellType: .floatFormula))
                    for _ in 0..<playerCount {
                        totalMPColumn.append(columns.count - 1)
                    }
                }
            }
        }
    }
    
    //
    
    func addBonusMPColumns(playerNumber: Int? = nil) {
        let event = scoreData.events.first!
        let winDraw = event.type?.requiresWinDraw ?? false
        let playerSuffix = (playerNumber == nil ? "" : " (\(playerNumber! + 1))")
        
        if !self.scoreData.strata.isEmpty {
            var playerStrataMPColumn: [Int] = []
            
            for (index, _) in self.scoreData.strata.enumerated() {
                columns.append(RanksPlusMPsColumn(title: "Strata \(index + 1) MP\(playerSuffix)", strataContent: playerStrataBonusAward, playerNumber: playerNumber, strataNumber: index, cellType: .floatFormula))
                playerStrataMPColumn.append(columns.count - 1)
            }
            strataMPColumn.append(playerStrataMPColumn)
            
            columns.append(RanksPlusMPsColumn(title: "Use\nStrata\nNumber", calculatedContent: playerFromStratumNumber, playerNumber: playerNumber, cellType: .integerFormula, format: formatZeroInt))
            fromStratumNumberColumn.append(columns.count - 1)
            
            columns.append(RanksPlusMPsColumn(title: "Use\nStrata\nCode", calculatedContent: playerFromStratumCode, playerNumber: playerNumber, cellType: .integerFormula, format: formatZeroInt))
            fromStratumCodeColumn.append(columns.count - 1)
            
            columns.append(RanksPlusMPsColumn(title: "Stratum\nPosition", calculatedContent: playerStratumPosition, playerNumber: playerNumber, cellType: .integerFormula, format: formatZeroInt))
            stratumPositionColumn.append(columns.count - 1)
            
            columns.append(RanksPlusMPsColumn(title: (winDraw ? "Bonus" : "Total"), calculatedContent: playerBestStratumAward, playerNumber: playerNumber, cellType: .floatFormula))
            maxStrataMPColumn.append(columns.count - 1)
            
        } else {
            columns.append(RanksPlusMPsColumn(title: "\((winDraw ? "Bonus" : "Total")) MP\(playerSuffix)", calculatedContent: playerBonusAward, playerNumber: playerNumber, cellType: .floatFormula))
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
    
    private func playerBestStratumAward(playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "=MAX("
        let playerNumber = playerNumber ?? 0
        for (index, strataMPColumn) in strataMPColumn[playerNumber].enumerated() {
            if index != 0 {
                result += ","
            }
            result += cell(rowNumber, strataMPColumn, columnFixed: true)
        }
        result += ")"
        return result
    }
        
    private func playerFromStratumNumber(playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "=IF(\(cell(rowNumber, positionColumn!, columnFixed: true))=0,0,MIN(\(cell(rowNumber, stratumColumn!, columnFixed: true)),\(self.scoreData.strata.count + 1)-\(fnPrefix)XMATCH("
        let playerNumber = playerNumber ?? 0
        result += cell(rowNumber, maxStrataMPColumn[playerNumber], columnFixed: true) + ",\(hstack)("
        for (index, strataMPColumn) in strataMPColumn[playerNumber].reversed().enumerated() {
            if index != 0 {
                result += ","
            }
            result += cell(rowNumber, strataMPColumn, columnFixed: true)
        }
        result += "))))"
        return result
    }
    
    private func playerFromStratumCode(playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "=IF(\(cell(rowNumber, positionColumn!, columnFixed: true))=0,\"\",\(fnPrefix)CHOOSECOLS(\(hstack)("
        let playerNumber = playerNumber ?? 0
        for (index, strataCodeCell) in strataCodeCell.enumerated() {
            if index != 0 {
                result += ","
            }
            result += strataCodeCell
        }
        result += "),\(cell(rowNumber, fromStratumNumberColumn[playerNumber], columnFixed: true))))"
        return result
    }
    
    private func playerStratumPosition(playerNumber: Int? = nil, rowNumber: Int) -> String {
        let playerNumber = playerNumber ?? 0
        let fromStratumCell = cell(rowNumber, fromStratumNumberColumn[playerNumber], columnFixed: true)
        let allPositionsRange = getAllPositionsRange(positionColumn: positionColumn!, comparisonColumn: stratumColumn!, comparison: "=\(fromStratumCell)")
        let result = "=IF(\(cell(rowNumber, stratumColumn!, columnFixed: true))=0,0,StratumPosition(\(cell(rowNumber, positionColumn!, columnFixed: true)),\(allPositionsRange)))"
        return result
    }
    
    private func bonusAward(_: Int?, rowNumber: Int) -> String {
        return playerStrataBonusAward(rowNumber: rowNumber)
    }
    
    private func playerBonusAward(playerNumber: Int? = nil, rowNumber: Int) -> String {
        return playerStrataBonusAward(playerNumber: playerNumber, rowNumber: rowNumber)
    }
    
    private func playerStrataBonusAward(strataNumber: Int? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = "="
        let event = scoreData.events.first!
        let positionCell = cell(rowNumber, positionColumn!)
        
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
        
        if let strataNumber = strataNumber {
            let stratumCell = cell(rowNumber, stratumColumn!, columnFixed: true)
            let allPositionsRange = getAllPositionsRange(positionColumn: positionColumn!, comparisonColumn: stratumColumn!, comparison: ">=\(strataNumber + 1)")
            result += "IF(\(stratumCell)<\(strataNumber + 1),0,StratumAward(ROUNDUP(\(maxAwardCell!)*\(strataPercentCell[strataNumber]),2),\(positionCell), \(awardPercentCell!), 2, \(allPositionsRange)))"
        } else {
            var useAwardCell: String
            var useAwardToCell: String
            var allPositionsRange = ""
            
            if event.winnerType == 2 && event.type?.participantType == .pair {
                let directionCell = cell(rowNumber, directionColumn!, columnFixed: true)
                useAwardCell = "IF(\(directionCell)=\"\(Direction.ns.string)\",\(maxAwardCell!),\(ewMaxAwardCell!))"
                useAwardToCell = "IF(\(directionCell)=\"\(Direction.ns.string)\",\(awardToCell!),\(ewAwardToCell!))"
                allPositionsRange = getAllPositionsRange(positionColumn: positionColumn!, comparisonColumn: directionColumn!, comparison: "=\(directionCell)")
            } else {
                useAwardCell = maxAwardCell!
                useAwardToCell = awardToCell!
                allPositionsRange = getAllPositionsRange(positionColumn: positionColumn!, comparisonColumn: positionColumn!, comparison: "<>0")
            }
            
            result += "IF(\(positionCell)=0,0,Award(\(useAwardCell), \(positionCell), \(useAwardToCell), 2, \(allPositionsRange)))"
        }
        if playerNumber != nil {
            result += ", 2)"
        }
        return result
    }
    
    private func getAllPositionsRange(positionColumn: Int, comparisonColumn: Int, comparison: String) -> String {
        return  "\(filter)(\(vstack)(\(cell(dataRow, rowFixed: true, positionColumn, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, positionColumn, columnFixed: true))),\(vstack)(\(cell(dataRow, rowFixed: true, comparisonColumn, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, comparisonColumn, columnFixed: true)))\(comparison))"
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
    
    private func playerFrozenRank(_: Participant, player: Player, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let lookupCell = cell(rowNumber, nationalIdColumn[playerNumber ?? 0], columnFixed: true)
        result = "\(fnPrefix)NUMBERVALUE(\(fnPrefix)XLOOKUP(\(lookupCell),\(vstack)(\(lookupFrozenRange(Settings.current.userDownloadNationalIdColumn))),\(vstack)(\(lookupFrozenRange(Settings.current.userDownloadRankColumn))),1,0))"

        return result
    }
    
    private func playerCurrentRank(_: Participant, player: Player, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        var currentRankCode: Int? = nil
        
        // Use current rank if available
        if let nationalId = player.nationalId, let rankCode = MemberViewModel.member(nationalId: nationalId)?.rankCode {
            currentRankCode = rankCode
        }
            
        if let currentRankCode = currentRankCode {
            result = "\(currentRankCode)"
        } else {
            // Couldn't get current rank - look it up in the spreadsheet
            let lookupCell = cell(rowNumber, nationalIdColumn[playerNumber ?? 0], columnFixed: true)
            result = "\(fnPrefix)NUMBERVALUE(\(fnPrefix)XLOOKUP(\(lookupCell),\(vstack)(\(lookupRange(Settings.current.userDownloadNationalIdColumn))),\(vstack)(\(lookupRange(Settings.current.userDownloadRankColumn))),1,0))"
        }
        return result
    }
    
    private func frozenRankCategory(_: Int? = nil, rowNumber: Int) -> String {
        let round = writer.rounds.first!
        var result = "\(fnPrefix)IF(\(cell(rowNumber, frozenRankColumn[0], columnFixed: true))=0,\"\",CombinedCategory(TRUE,"
        for playerNumber in 0..<round.maxParticipantPlayers {
            if playerNumber != 0 {
                result += ","
            }
            result += "\(vstack)(\(cell(rowNumber, frozenRankColumn[playerNumber], columnFixed: true)))"
        }
        result += "))"
        return result
    }

    private func stratum(_: Int? = nil, rowNumber: Int) -> String {
        let round = writer.rounds.first!
        var result = "\(fnPrefix)IF(\(cell(rowNumber, strataRankColumn[0], columnFixed: true))=0,0,Stratum(\(vstack)("
        for playerNumber in 0..<round.maxParticipantPlayers {
            if playerNumber != 0 {
                result += ","
            }
            result += cell(rowNumber, strataRankColumn[playerNumber], columnFixed: true)
        }
        result += "),\(hstack)("
        for (index, _) in self.scoreData.strata.enumerated() {
            if index != 0 {
                result += ","
            }
            result += strataRankCell[index]
        }
        result += ")))"
        return result
    }

    
    // Ranks plus MPs header
    
    func writeheader() {
        let event = scoreData.events.first!
        var column = -1
        var prefix = ""
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        if twoWinners {
            prefix = "NS "
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, string: String, format: UnsafeMutablePointer<lxw_format>? = nil, width: Float? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, string: string, format: format)
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, integer: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, integer: integer, format: format)
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, float: Float, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, float: float, format: format)
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, date: Date, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, date: date, format: format)
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, formula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, floatFormula: formula, format: format)
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, integerFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, integerFormula: integerFormula, format: format)
        }
        
        func writeCell(title: String, titleFormat: UnsafeMutablePointer<lxw_format>? = formatRightBold, floatFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: worksheet, row: headerTitleRow, column: column, string: title, format: titleFormat)
            write(worksheet: worksheet, row: headerDataRow, column: column, floatFormula: floatFormula, format: format)
        }
        
        // Occasionally leave a blank cell to allow column to overflow

        writeCell(title: "Description", titleFormat: formatBold, string: event.description ?? "") ; eventDescriptionCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        columnWidth[column] = 10
        
        if let toe = round.toe {
            writeCell(title: "TOE", integer: toe, format: formatInt)
        }
        
        var entryCells = ""
        var ewEntryRef = ""
        if twoWinners {
            let nsCell = "COUNTIF(\(cell(dataRow, rowFixed: true, directionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, directionColumn!, columnFixed: true)),\"\(Direction.ns.string)\")"
            writeCell(title: "\(prefix)Entry", integerFormula: "=\(nsCell)")
            entryCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
            column += 1
                        
            let ewCell = "COUNTIF(\(cell(dataRow, rowFixed: true, directionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, directionColumn!, columnFixed: true)),\"\(Direction.ew.string)\")"
            writeCell(title: "EW Entry", integerFormula: "=\(ewCell)")
            ewEntryRef = cell(headerDataRow, rowFixed: true, column, columnFixed: true)

            entryCells = "(\(nsCell)+\(ewCell))"
        } else {
            entryCells = "COUNTIF(\(cell(dataRow, rowFixed: true, positionColumn!, columnFixed: true)):\(cell(dataRow + round.fieldSize - 1, rowFixed: true, positionColumn!, columnFixed: true)),\">0\")"
            writeCell(title: "\(prefix)Entry", integerFormula: "=\(entryCells)") ; entryCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        }
        writeCell(title: "Tables", integerFormula: "=ROUNDUP(\(entryCells)*(\(event.type?.participantType?.players ?? 4)/4),0)") ; tablesCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Boards", integer: event.boards ?? 0)
        
        var baseEwMaxAwardCell: String?
        writeCell(title: "\(prefix)Full Award", floatFormula: "=ROUND(\(scoreData.maxAward),2)") ; let baseMaxAwardCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(title: "EW Full Award", floatFormula: "=ROUND(\(scoreData.ewMaxAward ?? scoreData.maxAward),2)") ; baseEwMaxAwardCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(title: "Min Entry", integer: round.scoreData.minEntry, format: formatInt) ; minEntryCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true) ; minEntryCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Entry %", floatFormula: "=IF(\(minEntryCell!)=0,1,MIN(1, ROUNDUP((\(entryCells))/\(minEntryCell!), 4)))", format: formatPercent) ; let entryFactorCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Factor %", float: round.scoreData.reducedTo, format: formatPercent) ; let reduceToCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)

        writeCell(title: "\(prefix)Max Award", floatFormula: "=ROUNDUP(ROUNDUP(\(baseMaxAwardCell)*\(entryFactorCell),2)*\(reduceToCell),2)") ; maxAwardCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(title: "EW Max Award", floatFormula: "=ROUNDUP(ROUNDUP(\(baseEwMaxAwardCell!)*\(entryFactorCell),2)*\(reduceToCell),2)") ; ewMaxAwardCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(title: "Award %", floatFormula: "=ROUND(\(scoreData.awardTo)/100,4)", format: formatPercent) ;  awardPercentCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "\(prefix)Award to", formula: "=ROUNDUP(\(entryCell!)*\(awardPercentCell!),0)") ; awardToCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(title: "EW Award to", formula: "=ROUNDUP(\(ewEntryRef)*\(awardPercentCell!),0)") ; ewAwardToCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(title: "Per Win", floatFormula: "=ROUND(\(scoreData.perWin),2)") ; perWinCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Event Date", titleFormat: formatBold, date: event.date ?? Date.today) ; eventDateCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        columnWidth[column] = 9

        
        writeCell(title: "Event Code", string: event.eventCode ?? "") ; eventIdCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Local/Nat", string: scoreData.national ? "National" : "Local") ; localCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Local MPs", floatFormula: "=IF(\(localCell!)=\"National\",0,ROUND(SUM(\(range(column: totalMPColumn))),2))") ; localMPsCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Nat MPs", floatFormula: "=IF(\(localCell!)<>\"National\",0,ROUND(SUM(\(range(column: totalMPColumn))),2))") ; nationalMPsCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        
        writeCell(title: "Checksum", floatFormula: "=CheckSum(\(vstack)(\(range(column: totalMPColumn))),\(vstack)(\(range(column: nationalIdColumn))))") ; checksumCell = cell(headerDataRow, rowFixed: true, column, columnFixed: true)
        columnWidth[column] = 9
        
        for (index, stratum) in self.scoreData.strata.enumerated() {
            writeCell(title: "Strata \(index + 1) Code", titleFormat: formatCenteredBold, string: stratum.code, format: formatCenteredString) ; strataCodeCell.append(cell(headerDataRow, rowFixed: true, column, columnFixed: true))
            writeCell(title: "Strata \(index + 1) Rank", integer: stratum.rank) ; strataRankCell.append(cell(headerDataRow, rowFixed: true, column, columnFixed: true))
            writeCell(title: "Strata \(index + 1) Award%", float: stratum.percent / 100, format: formatPercent) ; strataPercentCell.append(cell(headerDataRow, rowFixed: true, column, columnFixed: true))
        }
        
        headerColumns = column + 1
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
    var otherNamesColumn: Int?
    var participantCategoryColumn: Int?
    var playerRankColumn: Int?
    var playerCategoryColumn: Int?
    var playerRaceCategoryColumn: Int?
    
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
        setRow(worksheet: worksheet, row: titleRow, height: 30)
        
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBoldUnderline
            if column.cellType == .integer || column.cellType == .float || column.cellType == .numeric || column.cellType == .integerFormula || column.cellType == .floatFormula || column.cellType == .numericFormula {
                format = formatRightBoldUnderline
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
        
        columns.append(Column(title: "", referenceContent: { [self] (playerNumber) in round.ranksPlusMps.otherNamesColumn[playerNumber] }, cellType: .stringFormula, width: 16)) ; otherNamesColumn = columns.count - 1
        
        columns.append(Column(title: "SBU No", referenceContent: { [self] (playerNumber) in round.ranksPlusMps.nationalIdColumn[playerNumber] }, cellType: .integerFormula)) ; nationalIdColumn = columns.count - 1
        
        unique.referenceDynamic = { [self] in "CONCATENATE(\(arrayRef)(\(cell(dataRow, nationalIdColumn!, columnFixed: true))), \"+\", \(arrayRef)(\(cell(dataRow, firstNameColumn!, columnFixed: true))), \"+\", \(arrayRef)(\(cell(dataRow, otherNamesColumn!, columnFixed: true))))" }
        
        if self.writer.includeInRace {
            
            columns.append(Column(title: "\(event.type?.participantType?.string ?? "Team") Category", referenceContent: { [self] (_) in
                round.ranksPlusMps.frozenRankCategoryColumn! }, cellType: .integerFormula)) ; participantCategoryColumn = columns.count - 1

            columns.append(Column(title: "Player Rank", referenceContent: { [self] (playerNumber) in
                round.ranksPlusMps.frozenRankColumn[playerNumber] }, cellType: .integerFormula)) ; playerRankColumn = columns.count - 1
        
            columns.append(Column(title: "Player Category", referenceDynamic: { [self] in "=Category(\(arrayRef)(\(cell(dataRow, playerRankColumn!, columnFixed: true))))" }, cellType: .stringFormula)) ; playerCategoryColumn = columns.count - 1
            
            let logic = "=\(fnPrefix)MAKEARRAY(ROWS(\(arrayRef)(\(cell(dataRow, positionColumn!, columnFixed: true)))),1,\(lambda)(\(lambdaParam)Row,\(lambdaParam)Column,IF(RelativeTo(\(arrayRef)(\(cell(dataRow, playerRankColumn!, columnFixed: true))),\(lambdaParam)Row,1) = \(Settings.current.otherNBORank!), \"Non-SBU\",RelativeTo(\(arrayRef)(\(cell(dataRow, participantCategoryColumn!, columnFixed: true))),\(lambdaParam)Row,1))))"
            columns.append(Column(title: "Race Category", referenceDynamic: { logic }, cellType: .stringFormula)) ; playerRaceCategoryColumn = columns.count - 1
        }
        
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
            result += cell(writer: round.ranksPlusMps, round.ranksPlusMps.dataRow, rowFixed: true, round.ranksPlusMps.positionColumn!)
            result += ":"
            result += cell(round.ranksPlusMps.dataRow + round.fieldSize - 1, rowFixed: true, round.ranksPlusMps.positionColumn!)
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
    var formatCenteredString: UnsafeMutablePointer<lxw_format>?
    var formatZeroBlank: UnsafeMutablePointer<lxw_format>?
    var formatInt: UnsafeMutablePointer<lxw_format>?
    var formatFloat: UnsafeMutablePointer<lxw_format>?
    var formatFloatUnderline: UnsafeMutablePointer<lxw_format>?
    var formatZeroInt: UnsafeMutablePointer<lxw_format>?
    var formatZeroFloat: UnsafeMutablePointer<lxw_format>?
    var formatZeroFloatBold: UnsafeMutablePointer<lxw_format>?
    var formatBold: UnsafeMutablePointer<lxw_format>?
    var formatRightBold: UnsafeMutablePointer<lxw_format>?
    var formatCenteredBold: UnsafeMutablePointer<lxw_format>?
    var formatBoldUnderline: UnsafeMutablePointer<lxw_format>?
    var formatRightBoldUnderline: UnsafeMutablePointer<lxw_format>?
    var formatFloatBoldUnderline: UnsafeMutablePointer<lxw_format>?
    var formatPercent: UnsafeMutablePointer<lxw_format>?
    var formatDate: UnsafeMutablePointer<lxw_format>?
    var formatRed: UnsafeMutablePointer<lxw_format>?
    var formatRedHatched: UnsafeMutablePointer<lxw_format>?
    var formatYellow: UnsafeMutablePointer<lxw_format>?
    var formatGrey: UnsafeMutablePointer<lxw_format>?
    var formatFaint: UnsafeMutablePointer<lxw_format>?
    var formatNotActive: UnsafeMutablePointer<lxw_format>?
    var formatNoHomeClub: UnsafeMutablePointer<lxw_format>?
    var formatNotPaid: UnsafeMutablePointer<lxw_format>?
    var formatBannerString: UnsafeMutablePointer<lxw_format>?
    var formatBannerCenteredString: UnsafeMutablePointer<lxw_format>?
    var formatBannerFloat: UnsafeMutablePointer<lxw_format>?
    var formatBannerNumeric: UnsafeMutablePointer<lxw_format>?
    var formatBodyString: UnsafeMutablePointer<lxw_format>?
    var formatBodyCenteredString: UnsafeMutablePointer<lxw_format>?
    var formatBodyCenteredInt: UnsafeMutablePointer<lxw_format>?
    var formatBodyFloat: UnsafeMutablePointer<lxw_format>?
    var formatBodyNumeric: UnsafeMutablePointer<lxw_format>?
    
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
            self.formatCenteredString = writer.formatCenteredString
            self.formatZeroBlank = writer.formatZeroBlank
            self.formatInt = writer.formatInt
            self.formatFloat = writer.formatFloat
            self.formatFloatUnderline = writer.formatFloatUnderline
            self.formatZeroInt = writer.formatZeroInt
            self.formatZeroFloat = writer.formatZeroFloat
            self.formatZeroFloatBold = writer.formatZeroFloatBold
            self.formatBold = writer.formatBold
            self.formatCenteredBold = writer.formatCenteredBold
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
            self.formatFaint = writer.formatFaint
            self.formatNotActive = writer.formatNotActive
            self.formatNoHomeClub = writer.formatNoHomeClub
            self.formatNotPaid = writer.formatNotPaid
            self.formatBannerString = writer.formatBannerString
            self.formatBannerCenteredString = writer.formatBannerCenteredString
            self.formatBannerFloat = writer.formatBannerFloat
            self.formatBannerNumeric = writer.formatBannerNumeric
            self.formatBodyString = writer.formatBodyString
            self.formatBodyCenteredString = writer.formatBodyCenteredString
            self.formatBodyCenteredInt = writer.formatBodyCenteredInt
            self.formatBodyFloat = writer.formatBodyFloat
            self.formatBodyNumeric = writer.formatBodyNumeric
        }
    }
    
    // MARK: - Utility routines
    
    fileprivate func createMacroButton(worksheet: UnsafeMutablePointer<lxw_worksheet>?, title: String, macro: String, row: Int, column: Int, height: Int = 30, width: Int = 60, xScale: Double = 1.5, yScale: Double = 1.5, xOffset: Int = 2, yOffset: Int = 2) {
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
            var otherNames = names
            otherNames.removeLast()
            return otherNames.joined(separator: " ")
        } else {
            return names.last!
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
    
    internal func lookupRange(_ column: String) -> String {
        return "'\(Settings.current.userDownloadData!)'!$\(column)$\(Settings.current.userDownloadMinRow!):$\(column)$\(Settings.current.userDownloadMaxRow!)"
    }
    
    internal func lookupFrozenRange(_ column: String) -> String {
        return "'\(Settings.current.userDownloadFrozenData!)'!$\(column)$\(Settings.current.userDownloadMinRow!):$\(column)$\(Settings.current.userDownloadMaxRow!)"
    }
    
    func setupFormats() {
        formatString = workbook_add_format(workbook)
        format_set_align(formatString, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatCenteredString = workbook_add_format(workbook)
        format_set_align(formatCenteredString, UInt8(LXW_ALIGN_CENTER.rawValue))
        formatZeroBlank = workbook_add_format(workbook)
        format_set_num_format(formatZeroBlank, "0;-0;")
        format_set_align(formatZeroBlank, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatInt = workbook_add_format(workbook)
        format_set_num_format(formatInt, "0;-0;")
        format_set_align(formatInt, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatFloat = workbook_add_format(workbook)
        format_set_num_format(formatFloat, "0.00;-0.00;")
        format_set_align(formatFloat, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatFloatUnderline = workbook_add_format(workbook)
        format_set_num_format(formatFloatUnderline, "0.00;-0.00;")
        format_set_align(formatFloatUnderline, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_bottom(formatFloatUnderline, UInt8(LXW_BORDER_THIN.rawValue))
        formatZeroInt = workbook_add_format(workbook)
        format_set_num_format(formatZeroInt, "0")
        format_set_align(formatZeroInt, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatZeroFloat = workbook_add_format(workbook)
        format_set_num_format(formatZeroFloat, "0.00")
        format_set_align(formatZeroFloat, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatZeroFloatBold = workbook_add_format(workbook)
        format_set_num_format(formatZeroFloatBold, "0.00")
        format_set_align(formatZeroFloatBold, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_bold(formatZeroFloatBold)
        formatDate = workbook_add_format(workbook)
        format_set_num_format(formatDate, "dd/MM/yyyy")
        format_set_align(formatDate, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatPercent = workbook_add_format(workbook)
        format_set_num_format(formatPercent, "0.00%")
        format_set_align(formatPercent, UInt8(LXW_ALIGN_RIGHT.rawValue))
        formatBold = workbook_add_format(workbook)
        format_set_bold(formatBold)
        format_set_text_wrap(formatBold)
        formatCenteredBold = workbook_add_format(workbook)
        format_set_bold(formatCenteredBold)
        format_set_text_wrap(formatCenteredBold)
        format_set_align(formatCenteredBold, UInt8(LXW_ALIGN_CENTER.rawValue))
        formatRightBold = workbook_add_format(workbook)
        format_set_bold(formatRightBold)
        format_set_align(formatRightBold, UInt8(LXW_ALIGN_RIGHT.rawValue))
        format_set_num_format(formatRightBold, "0;-0;")
        format_set_text_wrap(formatRightBold)
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
        format_set_text_wrap(formatFloatBoldUnderline)
        format_set_num_format(formatFloatBoldUnderline, "0.00;-0.00;")
        
        formatRed = workbook_add_format(workbook)
        format_set_bg_color(formatRed, lxw_color_t(excelRed.rgbValue))
        format_set_font_color(formatRed, LXW_COLOR_WHITE.rawValue)
        formatRedHatched = workbook_add_format(workbook)
        format_set_bg_color(formatRedHatched, lxw_color_t(excelRed.rgbValue))
        format_set_fg_color(formatRedHatched, LXW_COLOR_WHITE.rawValue)
        format_set_pattern(formatRedHatched, UInt8(LXW_PATTERN_LIGHT_UP.rawValue))
        formatYellow = workbook_add_format(workbook)
        format_set_bg_color(formatYellow, lxw_color_t(excelYellow.rgbValue))
        formatGrey = workbook_add_format(workbook)
        format_set_bg_color(formatGrey, lxw_color_t(excelGrey.rgbValue))
        formatFaint = workbook_add_format(workbook)
        format_set_font_color(formatFaint, lxw_color_t(excelFaint.rgbValue))
        formatNotActive = workbook_add_format(workbook)
        format_set_bg_color(formatNotActive, lxw_color_t(excelNotActive.rgbValue))
        format_set_font_color(formatNotActive, LXW_COLOR_WHITE.rawValue)
        formatNoHomeClub = workbook_add_format(workbook)
        format_set_bg_color(formatNoHomeClub, lxw_color_t(excelNoHomeClub.rgbValue))
        format_set_font_color(formatNoHomeClub, LXW_COLOR_WHITE.rawValue)
        format_set_num_format(formatNoHomeClub, "0.00;-0.00;\"No Home Club\"")
        format_set_align(formatNoHomeClub, UInt8(LXW_ALIGN_LEFT.rawValue))
        formatNotPaid = workbook_add_format(workbook)
        format_set_bg_color(formatNotPaid, lxw_color_t(excelNotPaid.rgbValue))
        
        formatBannerString = workbook_add_format(workbook)
        format_set_align(formatBannerString, UInt8(LXW_ALIGN_LEFT.rawValue))
        format_set_align(formatBannerString, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_bold(formatBannerString)
        format_set_pattern (formatBannerString, UInt8(LXW_PATTERN_SOLID.rawValue))
        format_set_bg_color(formatBannerString, lxw_color_t(excelBanner.rgbValue))
        format_set_font_color(formatBannerString, LXW_COLOR_WHITE.rawValue)
        format_set_text_wrap(formatBannerString)
        format_set_indent(formatBannerString, 2)

        formatBannerCenteredString = workbook_add_format(workbook)
        format_set_align(formatBannerCenteredString, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBannerCenteredString, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_bold(formatBannerCenteredString)
        format_set_pattern (formatBannerCenteredString, UInt8(LXW_PATTERN_SOLID.rawValue))
        format_set_bg_color(formatBannerCenteredString, lxw_color_t(excelBanner.rgbValue))
        format_set_font_color(formatBannerCenteredString, LXW_COLOR_WHITE.rawValue)
        format_set_text_wrap(formatBannerCenteredString)
        
        formatBannerFloat = workbook_add_format(workbook)
        format_set_align(formatBannerFloat, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBannerFloat, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_bold(formatBannerFloat)
        format_set_pattern (formatBannerFloat, UInt8(LXW_PATTERN_SOLID.rawValue))
        format_set_bg_color(formatBannerFloat, lxw_color_t(excelBanner.rgbValue))
        format_set_font_color(formatBannerFloat, LXW_COLOR_WHITE.rawValue)
        format_set_text_wrap(formatBannerFloat)
        
        formatBannerNumeric = formatBannerFloat
        
        formatBodyString = workbook_add_format(workbook)
        format_set_align(formatBodyString, UInt8(LXW_ALIGN_LEFT.rawValue))
        format_set_align(formatBodyString, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_indent(formatBodyString, 2)
        
        formatBodyCenteredString = workbook_add_format(workbook)
        format_set_align(formatBodyCenteredString, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBodyCenteredString, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        
        formatBodyCenteredInt = workbook_add_format(workbook)
        format_set_align(formatBodyCenteredInt, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBodyCenteredInt, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_num_format(formatBodyCenteredInt, "0;-0;")

        formatBodyFloat = workbook_add_format(workbook)
        format_set_align(formatBodyFloat, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBodyFloat, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_num_format(formatBodyFloat, "0.00;-0.00;")
        
        formatBodyNumeric = workbook_add_format(workbook)
        format_set_align(formatBodyNumeric, UInt8(LXW_ALIGN_CENTER.rawValue))
        format_set_align(formatBodyNumeric, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
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
    
    func setRow(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, group: Bool = false, hidden: Bool = false, collapsed: Bool = false, format: UnsafeMutablePointer<lxw_format>? = nil, height: Float? = nil) {
        let row = lxw_row_t(Int32(row))
        let height = (height == nil ? LXW_DEF_ROW_HEIGHT : Double(height!))
        if group {
            var options = lxw_row_col_options()
            options.level = 1
            options.hidden = hidden ? 1 : 0
            options.collapsed = collapsed ? 1 : 0
            worksheet_set_row_opt(worksheet, row, height, format, &options)
        } else {
            worksheet_set_row_opt(worksheet, row, height, format, nil)
        }
    }
    
    func freezePanes(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, column: Int) {
        let row = lxw_row_t(Int32(row))
        let column = lxw_col_t(Int32(column))
        worksheet_freeze_panes(worksheet, row, column)
    }
    
}
