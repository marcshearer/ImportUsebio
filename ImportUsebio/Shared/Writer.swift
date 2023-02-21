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
    var referenceDynamic: (()->String)?
    var playerNumber: Int?
    var cellType: CellType
    var width: Float
    
    init(title: String, content: ((Participant, Int)->String)? = nil, playerContent: ((Participant, Player, Int, Int)->String)? = nil, playerNumber: Int? = nil, referenceContent: ((Int)->Int)? = nil, referenceDivisor: Int? = nil, referenceDynamic: (()->String)? = nil, cellType: CellType = .string, width: Float = 10.0) {
        self.title = title
        self.content = content
        self.playerContent = playerContent
        self.playerNumber = playerNumber
        self.referenceContent = referenceContent
        self.referenceDivisor = referenceDivisor
        self.referenceDynamic = referenceDynamic
        self.cellType = cellType
        self.width = width
    }
}

class Sheet {
    var writer: Writer
    var scoreData: ScoreData
    var prefix: String
    var individualMPs: IndividualMPsWriter!
    var ranksPlusMps: RanksPlusMPsWriter!
    var workbook: UnsafeMutablePointer<lxw_workbook>?
    
    init(writer: Writer, prefix: String, scoreData: ScoreData, workbook: UnsafeMutablePointer<lxw_workbook>?) {
        self.writer = writer
        self.prefix = prefix
        self.scoreData = scoreData
        self.individualMPs = IndividualMPsWriter(writer: writer, sheet: self, scoreData: scoreData)
        self.ranksPlusMps = RanksPlusMPsWriter(writer: writer, sheet: self, scoreData: scoreData)
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
        
        return result
    }
    
    var fieldSize: Int { scoreData.events.first!.participants.count }
    
    var maxPlayers: Int { maxParticipantPlayers * fieldSize }
    
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
    fileprivate var summary: SummaryWriter!
    fileprivate var csvExport:CsvExportWriter!
    fileprivate var consolidated: ConsolidatedWriter!
    fileprivate var sheets: [Sheet] = []
    
    var maxPlayers: Int { sheets.map{ $0.maxPlayers }.reduce(0, +) }
    
    init() {
        let name = "New workbook.xlsm"
        super.init(workbook: workbook_new(name))
        summary = SummaryWriter(writer: self)
        csvExport = CsvExportWriter(writer: self)
        consolidated = ConsolidatedWriter(writer: self)
    }
    
    func add(prefix: String, scoreData: ScoreData) {
        sheets.append(Sheet(writer: self, prefix: prefix, scoreData: scoreData, workbook: workbook))
    }
    
    func write() {
        workbook_add_vba_project(workbook, "./Award.bin")
        
        for sheet in sheets {
            sheet.ranksPlusMps.write()
            sheet.individualMPs.write()
        }
        consolidated.write()
        csvExport.write()
        summary.write()
        workbook_close(workbook)
    }
}

// MARK: - Summary
    
fileprivate class SummaryWriter : WriterBase {
    private var writer: Writer!
    private var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    override var name: String { "Summary" }
    var descriptionColumn: Int?
    var nationalLocalColumn: Int?
    var localMPsColumn: Int?
    var nationalMPsColumn: Int?
    var checksumColumn: Int?
    var exportedRow: Int?
    
    init(writer: Writer) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.workbook = workbook
        worksheet = workbook_add_worksheet(workbook, name)
    }
    
    func write() {
        descriptionColumn = 0
        let toeColumn = 1
        let tablesColumn = 2
        nationalLocalColumn = 3
        localMPsColumn = 4
        nationalMPsColumn = 5
        checksumColumn = 6
        let headerRow = 0
        let detailRow = 1
        let totalRow = detailRow + writer.sheets.count + 1
        exportedRow = totalRow + 1
        
        setColumn(worksheet: worksheet, column: descriptionColumn!, width: 30)
        setColumn(worksheet: worksheet, column: localMPsColumn!, width: 12)
        setColumn(worksheet: worksheet, column: nationalMPsColumn!, width: 12)
        setColumn(worksheet: worksheet, column: checksumColumn!, width: 16)
        
        write(worksheet: worksheet, row: headerRow, column: descriptionColumn!, string: "Round", format: formatBold)
        write(worksheet: worksheet, row: headerRow, column: toeColumn, string: "TOE", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow, column: tablesColumn, string: "Tables", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow, column: nationalLocalColumn!, string: "Nat/Local", format: formatBold)
        write(worksheet: worksheet, row: headerRow, column: localMPsColumn!, string: "Local MPs", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow, column: nationalMPsColumn!, string: "National MPs", format: formatRightBold)
        write(worksheet: worksheet, row: headerRow, column: checksumColumn!, string: "Checksum", format: formatRightBold)
        
        for (sheetNumber, sheet) in writer.sheets.enumerated() {
            let row = detailRow + sheetNumber
            write(worksheet: worksheet, row: row, column: descriptionColumn!, formula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.roundCell!)")
            write(worksheet: worksheet, row: row, column: toeColumn, integerFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.toeCell!)")
            write(worksheet: worksheet, row: row, column: tablesColumn, integerFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.tablesCell!)")
            write(worksheet: worksheet, row: row, column: nationalLocalColumn!, formula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.localCell!)")
            write(worksheet: worksheet, row: row, column: localMPsColumn!, floatFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.localMPsCell!)", format: formatZeroFloat)
            write(worksheet: worksheet, row: row, column: nationalMPsColumn!, floatFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.nationalMPsCell!)", format: formatZeroFloat)
            write(worksheet: worksheet, row: row, column: checksumColumn!, floatFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.checksumCell!)", format: formatZeroFloat)
        }
        
        for column in [toeColumn, tablesColumn, nationalLocalColumn, localMPsColumn, nationalMPsColumn, checksumColumn] {
            write(worksheet: worksheet, row: detailRow + writer.sheets.count, column: column!, string: "", format: formatBoldUnderline)
        }
        
        write(worksheet: worksheet, row: totalRow, column: descriptionColumn!, string: "Round totals", format: formatBold)
        
        let tablesColumnRef = columnRef(column: tablesColumn, fixed: true)
        write(worksheet: worksheet, row: totalRow, column: tablesColumn, integerFormula: "=SUM(\(tablesColumnRef)2:\(tablesColumnRef)\(2+writer.sheets.count))", format: formatZeroInt)
        
        let localMPsColumnRef = columnRef(column: localMPsColumn!, fixed: true)
        write(worksheet: worksheet, row: totalRow, column: localMPsColumn!, floatFormula: "=SUM(\(localMPsColumnRef)\(detailRow):\(localMPsColumnRef)\(detailRow + writer.sheets.count))", format: formatZeroFloat)
        
        let nationalMPsColumnRef = columnRef(column: nationalMPsColumn!, fixed: true)
        write(worksheet: worksheet, row: totalRow, column: nationalMPsColumn!, floatFormula: "=SUM(\(nationalMPsColumnRef)\(detailRow):\(nationalMPsColumnRef)\(detailRow + writer.sheets.count))", format: formatZeroFloat)
        
        let checksumColumnRef = columnRef(column: checksumColumn!, fixed: true)
        write(worksheet: worksheet, row: totalRow, column: checksumColumn!, floatFormula: "=SUM(\(checksumColumnRef)\(detailRow):\(checksumColumnRef)\(detailRow + writer.sheets.count))", format: formatZeroFloat)
        
        write(worksheet: worksheet, row: exportedRow!, column: descriptionColumn!, string: "Exported totals", format: formatBold)
        write(worksheet: worksheet, row: exportedRow!, column: localMPsColumn!, formula: "='\(writer.csvExport.name)'!\(writer.csvExport.localMpsCell!)", format: formatZeroFloat)
        write(worksheet: worksheet, row: exportedRow!, column: nationalMPsColumn!, formula: "='\(writer.csvExport.name)'!\(writer.csvExport.nationalMpsCell!)", format: formatZeroFloat)
        write(worksheet: worksheet, row: exportedRow!, column: checksumColumn!, formula: "='\(writer.csvExport.name)'!\(writer.csvExport.checksumCell!)", format: formatZeroFloat)
        
        for column in [localMPsColumn, nationalMPsColumn, checksumColumn] {
            highlightTotalDifferent(row: exportedRow!, compareRow: totalRow, column: column!)
        }
    }
    
    private func highlightTotalDifferent(row: Int, compareRow: Int, column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let field = "\(cell(row, rowFixed: true, column))"
        let matchField = "\(cell(compareRow, rowFixed: true, column))"
        let formula = "\(field)<>\(matchField)"
        setConditionalFormat(worksheet: worksheet, fromRow: row, fromColumn: column, toRow: row, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}

// MARK: - Export CSV

class CsvExportWriter: WriterBase {
    private var writer: Writer!
    var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    override var name: String { "Csv Export" }
    var localMpsCell: String?
    var nationalMpsCell: String?
    var checksumCell: String?
    var dataRow: Int?
    
    init(writer: Writer) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.workbook = workbook
        worksheet = workbook_add_worksheet(workbook, name)
    }
    
    func write() {
        let eventCodeRow = 0
        let eventDescriptionRow = 1
        let licenseNoRow = 2
        let eventDateRow = 3
        let nationalLocalRow = 4
        let localMPsRow = 5
        let nationalMPsRow = 6
        let checksumRow = 7
        let titleRow = 10
        dataRow = 11
        let titleColumn = 0
        let valuesColumn = 1
        let firstNameColumn = 0
        let otherNamesColumn = 1
        let eventDataColumn = 2
        let nationalIdColumn = 3
        let eventCodeColumn = 4
        let clubCodeColumn = 5
        let localMPsColumn = 6
        let nationalMPsColumn = 7
        let lookupFirstNameColumn = 10
        let lookupOtherNamesColumn = 11
        let lookupHomeClubColumn = 12
        let lookupRankColumn = 13
        let lookupEmailColumn = 14
        let lookupStatusColumn = 15
        
        setColumn(worksheet: worksheet, column: titleColumn, width: 12)
        setColumn(worksheet: worksheet, column: valuesColumn, width: 16)
        setColumn(worksheet: worksheet, column: eventDataColumn, width: 10, format: formatDate)
        setColumn(worksheet: worksheet, column: nationalIdColumn, format: formatInt)
        setColumn(worksheet: worksheet, column: localMPsColumn, format: formatFloat)
        setColumn(worksheet: worksheet, column: nationalMPsColumn, format: formatFloat)
        
        setColumn(worksheet: worksheet, column: lookupFirstNameColumn, width: 12)
        setColumn(worksheet: worksheet, column: lookupOtherNamesColumn, width: 12)
        setColumn(worksheet: worksheet, column: lookupHomeClubColumn, width: 25)
        setColumn(worksheet: worksheet, column: lookupRankColumn, width: 8)
        setColumn(worksheet: worksheet, column: lookupEmailColumn, width: 25)
        setColumn(worksheet: worksheet, column: lookupStatusColumn, width: 25)
        
            // Parameters etc
        write(worksheet: worksheet, row: eventCodeRow, column: titleColumn, string: "Event Code:", format: formatBold)
        write(worksheet: worksheet, row: eventCodeRow, column: valuesColumn, formula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.name)'!\(writer.sheets.first!.ranksPlusMps.eventIdCell!)")
        
        write(worksheet: worksheet, row: eventDescriptionRow, column: titleColumn, string: "Event:", format: formatBold)
        write(worksheet: worksheet, row: eventDescriptionRow, column: valuesColumn, formula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.name)'!\(writer.sheets.first!.ranksPlusMps.eventDescriptionCell!)")
        
        write(worksheet: worksheet, row: licenseNoRow, column: titleColumn, string: "License no:", format: formatBold)
        
        write(worksheet: worksheet, row: eventDateRow, column: titleColumn, string: "Date:", format: formatBold)
        write(worksheet: worksheet, row: eventDateRow, column: valuesColumn, floatFormula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.name)'!\(writer.sheets.first!.ranksPlusMps.eventDateCell!)", format: formatDate)
        
        write(worksheet: worksheet, row: nationalLocalRow, column: titleColumn, string: "Nat/Local:", format: formatBold)
        write(worksheet: worksheet, row: nationalLocalRow, column: valuesColumn, formula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.name)'!\(writer.sheets.first!.ranksPlusMps.localCell!)")
        
        write(worksheet: worksheet, row: localMPsRow, column: titleColumn, string: "Local MPs:", format: formatBold)
        write(worksheet: worksheet, row: localMPsRow, column: valuesColumn, dynamicFormula: "=SUM(\(arrayRef)(\(cell(dataRow!, rowFixed: true, localMPsColumn, columnFixed: true))))", format: formatZeroFloat)
        localMpsCell = cell(localMPsRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: worksheet, row: nationalMPsRow, column: titleColumn, string: "National MPs:", format: formatBold)
        write(worksheet: worksheet, row: nationalMPsRow, column: valuesColumn, dynamicFormula: "=SUM(\(arrayRef)(\(cell(dataRow!, rowFixed: true, nationalMPsColumn, columnFixed: true))))", format: formatZeroFloat)
        nationalMpsCell = cell(nationalMPsRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: worksheet, row: checksumRow, column: titleColumn, string: "Checksum:", format: formatBold)
        write(worksheet: worksheet, row: checksumRow, column: valuesColumn, dynamicFormula: "=SUM((\(arrayRef)(\(cell(dataRow!, rowFixed: true, localMPsColumn, columnFixed: true)))+\(arrayRef)(\(cell(dataRow!, rowFixed: true, nationalMPsColumn, columnFixed: true))))*\(arrayRef)(\(cell(dataRow!, rowFixed: true, nationalIdColumn, columnFixed: true))))", format: formatZeroFloat)
        checksumCell = cell(checksumRow, rowFixed: true, valuesColumn, columnFixed: true)
        
            // Data
        write(worksheet: worksheet, row: titleRow, column: firstNameColumn, string: "Names", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: firstNameColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.name)'!\(cell(writer.consolidated.dataRow!, rowFixed: true, writer.consolidated.firstNameColumn!, columnFixed: true)))")
        
        write(worksheet: worksheet, row: titleRow, column: otherNamesColumn, string: "", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: otherNamesColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.name)'!\(cell(writer.consolidated.dataRow!, rowFixed: true, writer.consolidated.otherNamesColumn!, columnFixed: true)))")
        
        write(worksheet: worksheet, row: titleRow, column: eventDataColumn, string: "Event date", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: eventDataColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(cell(dataRow!, rowFixed: true, firstNameColumn, columnFixed: true))), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\", \"\", \(cell(eventDateRow, rowFixed: true, valuesColumn, columnFixed: true)))))", format: formatDate)
        
        write(worksheet: worksheet, row: titleRow, column: nationalIdColumn, string: "MemNo", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: nationalIdColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.name)'!\(cell(writer.consolidated.dataRow!, rowFixed: true, writer.consolidated.nationalIdColumn!, columnFixed: true)))", format: formatInt)
        
        write(worksheet: worksheet, row: titleRow, column: eventCodeColumn, string: "Event Code", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: eventCodeColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(cell(dataRow!, rowFixed: true, firstNameColumn, columnFixed: true))), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\", \"\", \(cell(eventCodeRow, rowFixed: true, valuesColumn, columnFixed: true)))))", format: formatDate)
        
        write(worksheet: worksheet, row: titleRow, column: clubCodeColumn, string: "Club Code", format: formatBoldUnderline)
        
        write(worksheet: worksheet, row: titleRow, column: localMPsColumn, string: "Local", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: localMPsColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.name)'!\(cell(writer.consolidated.dataRow!, rowFixed: true, writer.consolidated.localMPsColumn!, columnFixed: true)))", format: formatFloat)
        
        write(worksheet: worksheet, row: titleRow, column: nationalMPsColumn, string: "National", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: nationalMPsColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.name)'!\(cell(writer.consolidated.dataRow!, rowFixed: true, writer.consolidated.nationalMPsColumn!, columnFixed: true)))", format: formatFloat)
        
            //Lookups
        write(worksheet: worksheet, row: titleRow, column: lookupFirstNameColumn, string: "First Name", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: lookupFirstNameColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(dataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),5,FALSE)")
        
        write(worksheet: worksheet, row: titleRow, column: lookupOtherNamesColumn, string: "Other Names", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: lookupOtherNamesColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(dataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),4,FALSE)")
        
        write(worksheet: worksheet, row: titleRow, column: lookupHomeClubColumn, string: "Home Club", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: lookupHomeClubColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(dataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),19,FALSE)")
        
        write(worksheet: worksheet, row: titleRow, column: lookupRankColumn, string: "Rank", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: lookupRankColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(dataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),6,FALSE)")
        
        write(worksheet: worksheet, row: titleRow, column: lookupEmailColumn, string: "Email", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: lookupEmailColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(dataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),8,FALSE)")
        
        write(worksheet: worksheet, row: titleRow, column: lookupStatusColumn, string: "Status", format: formatBoldUnderline)
        write(worksheet: worksheet, row: dataRow!, column: lookupStatusColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(dataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),25,FALSE)")
        
        highlightLookupDifferent(column: firstNameColumn, lookupColumn: lookupFirstNameColumn, format: formatYellow)
        highlightLookupDifferent(column: otherNamesColumn, lookupColumn: lookupOtherNamesColumn)
        highlightLookupError(fromColumn: lookupFirstNameColumn, toColumn: lookupStatusColumn, format: formatGrey)
        highlightBadNationalId(column: nationalIdColumn, firstNameColumn: firstNameColumn)
        highlightBadDate(column: eventDataColumn, firstNameColumn: firstNameColumn)
        highlightBadMPs(column: localMPsColumn, firstNameColumn: firstNameColumn)
        highlightBadMPs(column: nationalMPsColumn, firstNameColumn: firstNameColumn)
        highlightBadRank(column: lookupRankColumn)
        highlightBadStatus(column: lookupStatusColumn)
        
    }
    
    private func highlightLookupDifferent(column: Int, lookupColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let field = "\(cell(dataRow!, column, columnFixed: true))"
        let lookupfield = "\(cell(dataRow!, lookupColumn, columnFixed: true))"
        let formula = "\(field)<>\(lookupfield)"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: column, toRow: dataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightLookupError(fromColumn: Int, toColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let lookupCell = "\(cell(dataRow!, fromColumn))"
        let formula = "=ISNA(\(lookupCell))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: fromColumn, toRow: dataRow! + writer.maxPlayers - 1, toColumn: toColumn, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadStatus(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let statusCell = "\(cell(dataRow!, column))"
        let formula = "=AND(\(statusCell)<>\"\(goodStatus)\", \(statusCell)<>\"\")"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: column, toRow: dataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadNationalId(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let nationalIdCell = cell(dataRow!, column, columnFixed: true)
        let firstNameCell = cell(dataRow!, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(nationalIdCell)=0, \(nationalIdCell)>\(maxNationalIdNumber), NOT(ISNUMBER(\(nationalIdCell)))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: column, toRow: dataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, checkDuplicates: true, format: format ?? formatRed!, duplicateFormat: formatRedHatched!)
    }
    
    private func highlightBadMPs(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let pointsCell = cell(dataRow!, column, columnFixed: true)
        let firstNameCell = cell(dataRow!, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(pointsCell)>\(maxPoints), AND(\(pointsCell)<>\"\", NOT(ISNUMBER(\(pointsCell))))))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: column, toRow: dataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadDate(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let dateCell = cell(dataRow!, column, columnFixed: true)
        let firstNameCell = cell(dataRow!, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(dateCell)>DATEVALUE(\"\(Date().toString(format: "dd/MM/yyyy"))\"), \(dateCell)<DATEVALUE(\"\(Date().toString(format: "dd/MM/yyyy"))\")-30))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: column, toRow: dataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadRank(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let rankCell = cell(dataRow!, column, columnFixed: true)
        let formula = "=AND(\(rankCell)<>\"\", OR(\(rankCell)<'\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.name)'!\(writer.sheets.first!.ranksPlusMps.minRankCell!), \(rankCell)>'\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.name)'!\(writer.sheets.first!.ranksPlusMps.maxRankCell!)))"
        setConditionalFormat(worksheet: worksheet, fromRow: dataRow!, fromColumn: column, toRow: dataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}
    
// MARK: - Consolidated
 
class ConsolidatedWriter: WriterBase {
    var writer: Writer!
    var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    override var name: String { "Consolidated" }
    var dataRow: Int?
    var firstNameColumn: Int?
    var otherNamesColumn: Int?
    var nationalIdColumn: Int?
    var localMPsColumn: Int?
    var nationalMPsColumn: Int?
    
    init(writer: Writer) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.workbook = workbook
        worksheet = workbook_add_worksheet(workbook, name)
    }
    
    func write() {
        let uniqueColumn = 0
        firstNameColumn = 1
        otherNamesColumn = 2
        nationalIdColumn = 3
        localMPsColumn = 4
        nationalMPsColumn = 5
        let dataColumn = 6
        
        let titleRow = 3
        let nationalLocalRow = 0
        let totalRow = 1
        let checksumRow = 2
        dataRow = 4
        
        let nationalLocalRange = "\(cell(nationalLocalRow, rowFixed: true, dataColumn, columnFixed: true)):\(cell(nationalLocalRow, rowFixed: true, dataColumn + writer.sheets.count - 1, columnFixed: true))"
        
        setRow(worksheet: worksheet, row: titleRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: nationalLocalRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: totalRow, format: formatBoldUnderline)
        setRow(worksheet: worksheet, row: checksumRow, format: formatBoldUnderline)
        
        for column in 0..<(dataColumn + writer.sheets.count) {
            if column == uniqueColumn {
                setColumn(worksheet: worksheet, column: column, hidden: true)
            } else if column == nationalIdColumn {
                setColumn(worksheet: worksheet, column: column, width: 16.5, format: formatInt)
            } else {
                setColumn(worksheet: worksheet, column: column, width: 16.5, format: formatFloat)
            }
        }
        
        // Title row
        write(worksheet: worksheet, row: titleRow, column: firstNameColumn!, string: "First Name", format: formatBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: otherNamesColumn!, string: "Other Names", format: formatBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: nationalIdColumn!, string: "SBU", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: localMPsColumn!, string: "Local MPs", format: formatRightBoldUnderline)
        write(worksheet: worksheet, row: titleRow, column: nationalMPsColumn!, string: "National MPs", format: formatRightBoldUnderline)
        
        for (column, sheet) in writer.sheets.enumerated() {
            // Round titles
            write(worksheet: worksheet, row: titleRow, column: dataColumn + column, string: sheet.prefix, format: formatRightBoldUnderline)
        
            // National/Local row
            let cell = "'\(sheet.prefix) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.localCell!)"
            write(worksheet: worksheet, row: nationalLocalRow, column: dataColumn + column, formula: "=IF(\(cell)=0,\"\",\(cell))", format: formatRightBoldUnderline)
        }
        
        for element in 0...1 {
            // Total row and checksum row
            let row = (element == 0 ? totalRow : checksumRow)
            let totalRange = "\(cell(row, rowFixed: true, dataColumn, columnFixed: true)):\(cell(row, rowFixed: true, dataColumn + writer.sheets.count - 1, columnFixed: true))"
            
            write(worksheet: worksheet, row: row, column: nationalIdColumn!, string: (row == totalRow ? "Total" : "Checksum"), format: formatRightBoldUnderline)
            
            write(worksheet: worksheet, row: row, column: localMPsColumn!, floatFormula: "=SUMIF(\(nationalLocalRange), \"<>National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            write(worksheet: worksheet, row: row, column: nationalMPsColumn!, floatFormula: "=SUMIF(\(nationalLocalRange), \"=National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            for (column, _) in writer.sheets.enumerated() {
                let valueRange = "\(arrayRef)(\(cell(dataRow!, rowFixed: true, dataColumn + column)))"
                let nationalIdRange = "\(arrayRef)(\(cell(dataRow!, rowFixed: true, nationalIdColumn!, columnFixed: true)))"
                
                if row == totalRow {
                    write(worksheet: worksheet, row: row, column: dataColumn + column, floatFormula: "=SUM(\(valueRange))", format: formatFloatBoldUnderline)
                } else {
                    write(worksheet: worksheet, row: row, column: dataColumn + column, floatFormula: "=SUM(\(valueRange)*\(nationalIdRange))", format: formatFloatBoldUnderline)
                }
            }
        }
        
        // Data rows
        let uniqueIdCell = "\(arrayRef)(\(cell(dataRow!, uniqueColumn, columnFixed: true)))"
        
        // Unique ID column
        var formula = "=\(unique)("
        for (sheetNumber, sheet) in writer.sheets.enumerated() {
            if sheetNumber != 0 {
                formula += ","
            }
            formula += "\(vstack)(\(arrayRef)('\(sheet.scoreData.roundName!) \(sheet.individualMPs.name)'!\(cell(1,rowFixed: true, sheet.individualMPs.uniqueColumn!, columnFixed: true)))))"
        }
        write(worksheet: worksheet, row: dataRow!, column: uniqueColumn, dynamicFormula: formula)
        
        // Name columns
        write(worksheet: worksheet, row: dataRow!, column: firstNameColumn!, dynamicFormula: "=IF(\(uniqueIdCell)=\"\",\"\",\(fnPrefix)TEXTBEFORE(\(fnPrefix)TEXTAFTER(\(uniqueIdCell),\"+\"),\"+\"))")
        write(worksheet: worksheet, row: dataRow!, column: otherNamesColumn!, dynamicFormula: "=IF(\(uniqueIdCell)=\"\",\"\",\(fnPrefix)TEXTAFTER(\(uniqueIdCell), \"+\", 2))")
        
        // National ID column
        write(worksheet: worksheet, row: dataRow!, column: nationalIdColumn!, dynamicIntegerFormula: "=IF(\(uniqueIdCell)=\"\",\"\",IFERROR(\(fnPrefix)NUMBERVALUE(\(fnPrefix)TEXTBEFORE(\(uniqueIdCell), \"+\")),\(fnPrefix)TEXTBEFORE(\(uniqueIdCell), \"+\")))")
        
        // Total local/national columns
        let dataRange = "\(arrayRef)(\(cell(dataRow!, dataColumn, columnFixed: true))):\(arrayRef)(\(cell(dataRow!, rowFixed: true, dataColumn + writer.sheets.count - 1, columnFixed: true)))"
        
        write(worksheet: worksheet, row: dataRow!, column: localMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SUMIF(\(nationalLocalRange), \"<>National\", \(lambdaParam))))")
        write(worksheet: worksheet, row: dataRow!, column: nationalMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SUMIF(\(nationalLocalRange), \"=National\", \(lambdaParam))))")
        
        // Lookup data columns
        for (column, sheet) in writer.sheets.enumerated() {
            let sourceDataRange = "\(cell(1, rowFixed: true, sheet.individualMPs.uniqueColumn!, columnFixed: true)):\(cell(writer.maxPlayers, rowFixed: true, sheet.individualMPs.decimalColumn!, columnFixed: true))"
            let sourceOffset = (sheet.individualMPs.sheet.individualMPs.decimalColumn! - sheet.individualMPs.uniqueColumn! + 1)
            write(worksheet: worksheet, row: dataRow!, column: dataColumn + column, dynamicFloatFormula: "=IF(\(uniqueIdCell)=\"\",0,IFERROR(VLOOKUP(\(uniqueIdCell),'\(sheet.prefix) \(sheet.individualMPs.name)'!\(sourceDataRange),\(sourceOffset),FALSE),0))")
        }
    }
}
    
// MARK: - Ranks plus MPs
 
class RanksPlusMPsWriter: WriterBase {
    private var writer: Writer!
    var sheet: Sheet!
    var scoreData: ScoreData!
    var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    
    override var name: String { "Ranks Plus MPs" }
    var columns: [Column] = []
    var positionColumn: Int?
    var directionColumn: Int?
    var participantNoColumn: Int?
    var firstNameColumn: [Int] = []
    var otherNameColumn: [Int] = []
    var nationalIdColumn: [Int] = []
    var boardsPlayedColumn: [Int] = []
    var winDrawColumn: [Int] = []
    var bonusMPColumn: [Int] = []
    var winDrawMPColumn: [Int] = []
    var totalMPColumn: [Int] = []
    
    var roundCell: String?
    var eventDescriptionCell: String?
    var toeCell: String?
    var tablesCell: String?
    var maxAwardCell: String?
    var maxEwAwardCell: String?
    var awardToCell: String?
    var ewAwardToCell: String?
    var perWinCell: String?
    var eventDateCell: String?
    var eventIdCell: String?
    var localCell: String?
    var minRankCell: String?
    var maxRankCell: String?
    var localMPsCell :String?
    var nationalMPsCell :String?
    var checksumCell :String?
    
    let headerRows = 3
    
    init(writer: Writer, sheet: Sheet, scoreData: ScoreData) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.sheet = sheet
        self.scoreData = scoreData
        self.workbook = workbook
        worksheet = workbook_add_worksheet(workbook, "\(scoreData.roundName!) \(name)")
    }
    
    func write() {
        let event = scoreData.events.first!
        let participants = scoreData.events.first!.participants.sorted(by: sortCriteria)
        
        setupColumns()
        writeheader()
        
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: worksheet, column: columnNumber, width: column.width)
            write(worksheet: worksheet, row: headerRows, column: columnNumber, string: sheet.replace(column.title), format: format)
        }
        
        for playerNumber in 0..<sheet.maxParticipantPlayers {
            let nationalIdColumn = nationalIdColumn[playerNumber]
            let firstNameColumn = firstNameColumn[playerNumber]
            let otherNamesColumn = otherNameColumn[playerNumber]
            let firstNameNonBlank = "\(columnRef(column: firstNameColumn, fixed: true))\(headerRows + 2)<>\"\""
            let otherNamesNonBlank = "\(columnRef(column: otherNamesColumn, fixed: true))\(headerRows + 2)<>\"\""
            let nationalIdZero = "\(columnRef(column: nationalIdColumn, fixed: true))\(headerRows + 2)=0"
            let nationalIdLarge = "\(columnRef(column: nationalIdColumn, fixed: true))\(headerRows + 2)>\(maxNationalIdNumber)"
            let formula = "=AND(OR(\(firstNameNonBlank),\(otherNamesNonBlank)), OR(\(nationalIdZero),\(nationalIdLarge)))"
            setConditionalFormat(worksheet: worksheet, fromRow: headerRows + 1, fromColumn: nationalIdColumn, toRow: headerRows + sheet.fieldSize, toColumn: nationalIdColumn, formula: formula, format: formatRed!)
        }
        
        for (rowSequence, participant) in participants.enumerated() {
            
            let rowNumber = rowSequence + headerRows + 1
            
            for (columnNumber, column) in columns.enumerated() {
                
                if let content = column.content?(participant, rowNumber) {
                    write(cellType: (content == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: content)
                }
                
                if let playerNumber = column.playerNumber {
                    let playerList = participant.member.playerList
                    if playerNumber < playerList.count {
                        if let playerContent = column.playerContent?(participant, playerList[playerNumber], playerNumber, rowNumber) {
                            write(cellType: (playerContent == "" ? .string : column.cellType), worksheet: worksheet, row: rowNumber, column: columnNumber, content: playerContent)
                        }
                    } else {
                        write(cellType: .string, worksheet: worksheet, row: rowNumber, column: columnNumber, content: "")
                    }
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
        let playerCount = sheet.maxParticipantPlayers
        let winDraw = event.type?.requiresWinDraw ?? false
        
        columns.append(Column(title: "Place", content: { (participant, _) in "\(participant.place!)" }, cellType: .integer))
        positionColumn = columns.count - 1
        
        if event.winnerType == 2 && event.type?.participantType == .pair {
            columns.append(Column(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
            directionColumn = columns.count - 1
        }
        
        columns.append(Column(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer)) ; participantNoColumn = columns.count - 1
        
        columns.append(Column(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float))
        
        if winDraw && sheet.maxParticipantPlayers <= event.type?.participantType?.players ?? sheet.maxParticipantPlayers {
            columns.append(Column(title: "Win/Draw", content: { (participant, _) in "\(participant.winDraw!)" }, cellType: .float))
            winDrawColumn.append(columns.count - 1)
        }
        
        for playerNumber in 0..<playerCount {
            columns.append(Column(title: "First Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 0) }, playerNumber: playerNumber, cellType: .string))
            firstNameColumn.append(columns.count - 1)
            
            columns.append(Column(title: "Other Names (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 1) }, playerNumber: playerNumber, cellType: .string))
            otherNameColumn.append(columns.count - 1)
            
            columns.append(Column(title: "SBU No (\(playerNumber+1))", playerContent: { (_, player,_, _) in player.nationalId! }, playerNumber: playerNumber, cellType: .integer))
            nationalIdColumn.append(columns.count - 1)
            
            if sheet.maxParticipantPlayers > event.type?.participantType?.players ?? sheet.maxParticipantPlayers {
                
                columns.append(Column(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.boardsPlayed)" }, playerNumber: playerNumber, cellType: .integer))
                boardsPlayedColumn.append(columns.count - 1)
                
                if winDraw {
                    columns.append(Column(title: "Win/Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                    winDrawColumn.append(columns.count - 1)
                }
                
                columns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP (\(playerNumber+1))", playerContent: playerBonusAward, playerNumber: playerNumber, cellType: .floatFormula))
                bonusMPColumn.append(columns.count - 1)
                
                if winDraw {
                    columns.append(Column(title: "Win/Draw MP (\(playerNumber+1))", playerContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .floatFormula))
                    winDrawMPColumn.append(columns.count - 1)
                    
                    columns.append(Column(title: "Total MP (\(playerNumber+1))", playerContent: playerTotalAward, playerNumber: playerNumber, cellType: .floatFormula))
                    totalMPColumn.append(columns.count - 1)
                }
            }
        }
        
        if sheet.maxParticipantPlayers <= event.type?.participantType?.players ?? sheet.maxParticipantPlayers {
            
            columns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP", content: bonusAward, cellType: .floatFormula))
            bonusMPColumn.append(columns.count - 1)
            
            if winDraw {
                
                columns.append(Column(title: "Win/Draw MP", content: winDrawAward, cellType: .integerFormula))
                winDrawMPColumn.append(columns.count - 1)
                
                columns.append(Column(title: "Total MP", content: totalAward, cellType: .integerFormula))
                for _ in 0..<sheet.maxParticipantPlayers {
                    totalMPColumn.append(columns.count - 1)
                }
            }
        }
    }
    
    // Ranks plus MPs Content getters
    
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
        if let nsPairs = sheet.nsPairs {
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
        
        result += "Award(\(useAwardCell!), \(positionCell), \(useAwardToCell!), 2, \(allPositionsRange))"
        
        if playerNumber != nil {
            result += ", 2)"
        }
        
        return result
    }
    
    private func winDrawAward(participant: Participant, rowNumber: Int) -> String {
        return playerWinDrawAward(participant: participant, rowNumber: rowNumber)
    }
    
    private func playerWinDrawAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let winDrawsCell = cell(rowNumber, winDrawColumn[playerNumber ?? 0], columnFixed: true)
        result = "=ROUNDUP(\(winDrawsCell) * \(perWinCell!), 2)"
        
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
    
    // Ranks plus MPs header
    
    func writeheader() {
        let event = scoreData.events.first!
        var column = -1
        var row = 0
        var prefix = ""
        var toe = sheet.fieldSize
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        if twoWinners {
            prefix = "NS "
            toe = sheet.nsPairs!
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
        
        writeCell(string: "Round", format: formatBold)
        writeCell(string: "Description", format: formatBold)
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
        
        writeCell(string: "Min Rank", format: formatBold)
        
        writeCell(string: "Max Rank", format: formatBold)
        
        writeCell(string: "Local MPs", format: formatBold)
        writeCell(string: "National MPs", format: formatBold)
        writeCell(string: "Checksum", format: formatBold)
        
        column = -1
        row = 1
        
        writeCell(string: scoreData.roundName ?? "") ; roundCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.description ?? "") ; eventDescriptionCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: toe) ; toeCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        var toeCells = toeCell!
        var ewToeRef = ""
        if twoWinners {
            writeCell(integer: sheet.fieldSize - sheet.nsPairs!) ; ewToeRef = cell(row, rowFixed: true, column, columnFixed: true)
            toeCells += "+" + ewToeRef
        }
        writeCell(integerFormula: "=ROUNDUP(\(toeCells)*(\(event.type?.participantType?.players ?? 4)/4),0)") ; tablesCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: event.boards ?? 0)
        
        var baseMaxEwAwardCell = ""
        writeCell(float: 10.00) ; let baseMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(floatFormula: "ROUNDUP(\(baseMaxAwardCell)*\(ewToeRef)/\(toeCell!),2)") ; baseMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 1, format: formatPercent) ; let factorCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=ROUNDUP(\(baseMaxAwardCell)*\(factorCell),2)") ; maxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(floatFormula: "=ROUNDUP(\(baseMaxEwAwardCell)*\(factorCell),2)") ; maxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 0.25, format: formatPercent) ; let awardPercentCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(toeCell!)*\(awardPercentCell),0)") ; awardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(ewToeRef)*\(awardPercentCell),0)") ; ewAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 0.25) ; perWinCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(date: event.date ?? Date.today) ; eventDateCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.eventCode ?? "") ; eventIdCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: scoreData.national ? "National" : "Local") ; localCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: scoreData.minRank) ; minRankCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: scoreData.maxRank) ; maxRankCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(localCell!)=\"National\",0,SUM(\(range(column: totalMPColumn))))") ; localMPsCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(localCell!)<>\"National\",0,SUM(\(range(column: totalMPColumn))))") ; nationalMPsCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=SUM(\(vstack)(\(range(column: totalMPColumn)))*\(vstack)(\(range(column: nationalIdColumn))))") ; checksumCell = cell(row, rowFixed: true, column, columnFixed: true)
    }
    
    private func range(column: [Int])->String {
        let event = scoreData.events.first!
        let firstRow = headerRows + 1
        let lastRow = headerRows + event.participants.count
        var result = ""
        for playerNumber in 0..<sheet.maxParticipantPlayers {
            if playerNumber != 0 {
                result += ","
            }
            let from = cell(firstRow, rowFixed: true, column[playerNumber])
            let to = cell(lastRow, rowFixed: true, column[playerNumber])
            result += from + ":" + to
        }
        return result
    }
}
    
// MARK: - Indiviual MPs worksheet

class IndividualMPsWriter: WriterBase {
    private var writer: Writer
    var sheet: Sheet
    var scoreData: ScoreData
    var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    override var name: String { "Individual MPs" }
    var columns: [Column] = []
    var positionColumn: Int?
    var uniqueColumn: Int?
    var decimalColumn: Int?
    var nationalIdColumn: Int?
    var localMPsColumn: Int?
    var nationalMPsColumn: Int?
    var localTotalColumn: Int?
    var nationalTotalColumn: Int?
    var checksumColumn: Int?
    
    init(writer: Writer, sheet: Sheet, scoreData: ScoreData) {
        self.writer = writer
        self.sheet = sheet
        self.scoreData = scoreData
        super.init(workbook: writer.workbook)
        worksheet = workbook_add_worksheet(workbook, "\(scoreData.roundName!) \(name)")
     }
    
    func write() {
        setupIndividualMpsColumns()
        
        for (columnNumber, column) in columns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: worksheet, column: columnNumber, width: column.width)
            write(worksheet: worksheet, row: 0, column: columnNumber, string: sheet.replace(column.title), format: format)
        }
        
        for (columnNumber, column) in columns.enumerated() {
            
            if let referenceContent = column.referenceContent {
                referenceColumn(columnNumber: columnNumber, referencedContent: referenceContent, referenceDivisor: column.referenceDivisor, cellType: column.cellType)
            }
            
            if let referenceDynamicContent = column.referenceDynamic {
                referenceDynamic(columnNumber: columnNumber, content: referenceDynamicContent, cellType: column.cellType)
            }
        }
        
        let localArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, localMPsColumn!, columnFixed: true)))"
        write(worksheet: worksheet, row: 1, column: localTotalColumn!, floatFormula: "=SUM(\(localArrayRef))", format: formatZeroFloat)
        
        let nationalArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, nationalMPsColumn!, columnFixed: true)))"
        write(worksheet: worksheet, row: 1, column: nationalTotalColumn!, floatFormula: "=SUM(\(nationalArrayRef))", format: formatZeroFloat)
        
        let totalArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, decimalColumn!, columnFixed: true)))"
        let nationalIdArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, nationalIdColumn!, columnFixed: true)))"
        write(worksheet: worksheet, row: 1, column: checksumColumn!, floatFormula: "=SUM(\(totalArrayRef)*\(nationalIdArrayRef))", format: formatZeroFloat)
        
        setColumn(worksheet: worksheet, column: uniqueColumn!, hidden: true)
        
    }
    
    private func setupIndividualMpsColumns() {
        let event = scoreData.events.first!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        
        columns.append(Column(title: "Place", referenceContent: { [self] (_) in sheet.ranksPlusMps.positionColumn! }, cellType: .integerFormula)) ; positionColumn = columns.count - 1
        
        if twoWinners {
            columns.append(Column(title: "Direction", referenceContent: { [self] (_) in sheet.ranksPlusMps.directionColumn! }, cellType: .stringFormula))
        }
        
        columns.append(Column(title: "@P no", referenceContent: { [self] (_) in sheet.ranksPlusMps.participantNoColumn! }, cellType: .integerFormula))
        
        columns.append(Column(title: "Unique", cellType: .floatFormula)) ; uniqueColumn = columns.count - 1
        let unique = columns.last!
        
        columns.append(Column(title: "Names", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.firstNameColumn[playerNumber] }, cellType: .stringFormula)) ; let firstNameColumn = columns.count - 1
        
        columns.append(Column(title: "", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.otherNameColumn[playerNumber] }, cellType: .stringFormula)) ; let otherNamesColumn = columns.count - 1
        
        columns.append(Column(title: "SBU No", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.nationalIdColumn[playerNumber] }, cellType: .integerFormula)) ; nationalIdColumn = columns.count - 1
        
        unique.referenceDynamic = { [self] in "CONCATENATE(\(arrayRef)(\(cell(1, nationalIdColumn!, columnFixed: true))), \"+\", \(arrayRef)(\(cell(1, firstNameColumn, columnFixed: true))), \"+\", \(arrayRef)(\(cell(1, otherNamesColumn, columnFixed: true))))" }
        
        columns.append(Column(title: "Total MPs", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.totalMPColumn[playerNumber] }, cellType: .floatFormula)) ; decimalColumn = columns.count - 1
        
        columns.append(Column(title: "Local MPs", referenceDynamic: { [self] in "=\(byRow)(\(arrayRef)(\(cell(1, decimalColumn!, columnFixed: true))),\(lambda)(\(lambdaParam), IF('\(scoreData.roundName!) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.localCell!)<>\"National\",\(lambdaParam),0)))" }, cellType: .floatFormula)) ; localMPsColumn = columns.count - 1
        
        columns.append(Column(title: "National MPs", referenceDynamic: { [self] in "=\(byRow)(\(arrayRef)(\(cell(1, decimalColumn!, columnFixed: true))),\(lambda)(\(lambdaParam), IF('\(scoreData.roundName!) \(sheet.ranksPlusMps.name)'!\(sheet.ranksPlusMps.localCell!)=\"National\",\(lambdaParam),0)))" }, cellType: .floatFormula)) ; nationalMPsColumn = columns.count - 1
        
        columns.append(Column(title: "Total Local", cellType: .floatFormula)) ; localTotalColumn = columns.count - 1
        
        columns.append(Column(title: "Total National", cellType: .floatFormula)) ; nationalTotalColumn = columns.count - 1
        
        columns.append(Column(title: "Checksum", cellType: .floatFormula)) ; checksumColumn = columns.count - 1
    }
    
    private func referenceColumn(columnNumber: Int, referencedContent: (Int)->Int, referenceDivisor: Int? = nil, cellType: CellType? = nil) {
        
        let content = zeroFiltered(referencedContent: referencedContent, divisor: referenceDivisor)
        let position = zeroFiltered(referencedContent: { (_) in positionColumn! })
        
        let result = "=\(sortBy)(\(content), \(position), 1)"
        
        let column = lxw_col_t(Int32(columnNumber))
        worksheet_write_dynamic_array_formula(worksheet, 1, column, 999, column, result, formatFrom(cellType: cellType))
    }
    
    private func zeroFiltered(referencedContent: (Int)->Int ,divisor: Int? = nil) -> String {
        
        var result = "\(filter)(\(vstack)("
        
        for playerNumber in 0..<sheet.maxParticipantPlayers {
            let columnReference = referencedContent(playerNumber)
            if playerNumber != 0 {
                result += ","
            }
            result += "'\(scoreData.roundName!) \(sheet.ranksPlusMps.name)'!" + cell(sheet.ranksPlusMps.headerRows + 1, columnReference)
            result += ":"
            result += cell(sheet.ranksPlusMps.headerRows + sheet.fieldSize, columnReference)
        }
        
        result += ")"
        
        if let divisor = divisor {
            result += "/\(divisor)"
        }
        result += ",\(vstack)("
        
        for playerNumber in 0..<sheet.maxParticipantPlayers {
            if playerNumber != 0 {
                result += ","
            }
            result += "'\(scoreData.roundName!) \(sheet.ranksPlusMps.name)'!" + cell(sheet.ranksPlusMps.headerRows + 1, sheet.ranksPlusMps.totalMPColumn[playerNumber])
            result += ":"
            result += cell(sheet.ranksPlusMps.headerRows + sheet.fieldSize, sheet.ranksPlusMps.totalMPColumn[playerNumber])
        }
        result += ")<>0)"
        
        return result
    }
    
    private func referenceDynamic(columnNumber: Int, content: ()->String, cellType: CellType? = nil) {
        
        let column = lxw_col_t(Int32(columnNumber))
        worksheet_write_dynamic_array_formula(worksheet, 1, column, 999, column, content(), formatFrom(cellType: cellType))
    }
    
}

// MARK - Writer base class

class WriterBase {
    
    let userDownloadData = "user download.csv"
    let userDownloadRange = "$A$2:$AI$6000"
    let maxNationalIdNumber = 30000
    let goodStatus = "Payment Confirmed by SBU"
    let maxPoints: Float = 15.0
    
    let fnPrefix = "_xlfn."
    let dynamicFnPrefix = "_xlfn._xlws."
    let paramPrefix = "_xlpm."
    
    var vstack: String { "\(fnPrefix)VSTACK" }
    var arrayRef: String { "\(fnPrefix)ANCHORARRAY" }
    var filter: String { "\(dynamicFnPrefix)FILTER" }
    var sortBy: String { "\(dynamicFnPrefix)SORTBY" }
    var unique: String { "\(dynamicFnPrefix)UNIQUE" }
    var byRow: String { "\(fnPrefix)BYROW" }
    var lambda: String { "\(fnPrefix)LAMBDA"}
    var lambdaParam: String { "\(paramPrefix)x"}
    
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
    var formatRed: UnsafeMutablePointer<lxw_format>?
    var formatRedHatched: UnsafeMutablePointer<lxw_format>?
    var formatYellow: UnsafeMutablePointer<lxw_format>?
    var formatGrey: UnsafeMutablePointer<lxw_format>?
    
    var name: String { fatalError() }
    
    init(workbook: UnsafeMutablePointer<lxw_workbook>? = nil) {
        if workbook != nil {
            self.workbook = workbook
            setupFormats()
        }
    }
    
    // MARK: - Utility routines
    
    fileprivate func formatFrom(cellType: CellType? = nil) -> UnsafeMutablePointer<lxw_format>? {
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
        let rowRef = rowRef(row: row, fixed: rowFixed)
        let columnRef = columnRef(column: column, fixed: columnFixed)
        return cell(writer: writer, cellRef: "\(columnRef)\(rowRef)")
    }
    
    func cell(writer: WriterBase? = nil, cellRef: String) -> String {
        let sheetRef = (writer == nil ? "" : "'\(writer!.name)'!")
        return "\(sheetRef)\(cellRef)"
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
    
    func setConditionalFormat(worksheet: UnsafeMutablePointer<lxw_worksheet>?, fromRow: Int, fromColumn: Int, toRow: Int, toColumn: Int, formula: String, checkDuplicates: Bool = false, format: UnsafeMutablePointer<lxw_format>, duplicateFormat: UnsafeMutablePointer<lxw_format>? = nil) {
        let formula = formula
        var conditionalFormat = lxw_conditional_format()
        conditionalFormat.type = UInt8(LXW_CONDITIONAL_TYPE_FORMULA.rawValue)
        conditionalFormat.value_string = UnsafeMutablePointer<CChar>(mutating: NSString(string: formula).utf8String)
        conditionalFormat.stop_if_true = 1
        conditionalFormat.format = format
        
        worksheet_conditional_format_range(worksheet, lxw_row_t(Int32(fromRow)), lxw_col_t(Int32(fromColumn)), lxw_row_t(Int32(toRow)), lxw_col_t(Int32(toColumn)), &conditionalFormat)
        
        if checkDuplicates {
            var dupFormat = lxw_conditional_format()
            dupFormat.type = UInt8(LXW_CONDITIONAL_TYPE_DUPLICATE.rawValue)
            dupFormat.format = duplicateFormat ?? format
            
            worksheet_conditional_format_range(worksheet, lxw_row_t(Int32(fromRow)), lxw_col_t(Int32(fromColumn)), lxw_row_t(Int32(toRow)), lxw_col_t(Int32(toColumn)), &dupFormat)
        }
    }
    
    func setColumn(worksheet: UnsafeMutablePointer<lxw_worksheet>?, column: Int, width: Float? = nil, hidden: Bool = false, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let column = lxw_col_t(Int32(column))
        let width = (width == nil ? LXW_DEF_COL_WIDTH : Double(width!))
        if hidden {
            var options = lxw_row_col_options()
            options.hidden = 1
            worksheet_set_column_opt(worksheet, column, column, width, format, &options)
        } else {
            worksheet_set_column_opt(worksheet, column, column, width, format, nil)
        }
    }
    
    func setRow(worksheet: UnsafeMutablePointer<lxw_worksheet>?, row: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let row = lxw_row_t(Int32(row))
        worksheet_set_row(worksheet, row,  LXW_DEF_ROW_HEIGHT, format)
    }
}
