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
    private var summaryWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let summaryName = "Summary"
    var summaryDescriptionColumn: Int?
    var summaryNationalLocalColumn: Int?
    var summaryLocalMPsColumn: Int?
    var summaryNationalMPsColumn: Int?
    var summaryChecksumColumn: Int?
    var summaryExportedRow: Int?
    
    init(writer: Writer) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.workbook = workbook
        summaryWorksheet = workbook_add_worksheet(workbook, summaryName)
    }
    
    func write() {
        summaryDescriptionColumn = 0
        let toeColumn = 1
        let tablesColumn = 2
        summaryNationalLocalColumn = 3
        summaryLocalMPsColumn = 4
        summaryNationalMPsColumn = 5
        summaryChecksumColumn = 6
        let headerRow = 0
        let detailRow = 1
        let totalRow = detailRow + writer.sheets.count + 1
        summaryExportedRow = totalRow + 1
        
        setColumn(worksheet: summaryWorksheet, column: summaryDescriptionColumn!, width: 30)
        setColumn(worksheet: summaryWorksheet, column: summaryLocalMPsColumn!, width: 12)
        setColumn(worksheet: summaryWorksheet, column: summaryNationalMPsColumn!, width: 12)
        setColumn(worksheet: summaryWorksheet, column: summaryChecksumColumn!, width: 16)
        
        write(worksheet: summaryWorksheet, row: headerRow, column: summaryDescriptionColumn!, string: "Round", format: formatBold)
        write(worksheet: summaryWorksheet, row: headerRow, column: toeColumn, string: "TOE", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: headerRow, column: tablesColumn, string: "Tables", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: headerRow, column: summaryNationalLocalColumn!, string: "Nat/Local", format: formatBold)
        write(worksheet: summaryWorksheet, row: headerRow, column: summaryLocalMPsColumn!, string: "Local MPs", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: headerRow, column: summaryNationalMPsColumn!, string: "National MPs", format: formatRightBold)
        write(worksheet: summaryWorksheet, row: headerRow, column: summaryChecksumColumn!, string: "Checksum", format: formatRightBold)
        
        for (sheetNumber, sheet) in writer.sheets.enumerated() {
            let row = detailRow + sheetNumber
            write(worksheet: summaryWorksheet, row: row, column: summaryDescriptionColumn!, formula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsRoundCell!)")
            write(worksheet: summaryWorksheet, row: row, column: toeColumn, integerFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsToeCell!)")
            write(worksheet: summaryWorksheet, row: row, column: tablesColumn, integerFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsTablesCell!)")
            write(worksheet: summaryWorksheet, row: row, column: summaryNationalLocalColumn!, formula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsNationalLocalCell!)")
            write(worksheet: summaryWorksheet, row: row, column: summaryLocalMPsColumn!, floatFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsLocalMPsCell!)", format: formatZeroFloat)
            write(worksheet: summaryWorksheet, row: row, column: summaryNationalMPsColumn!, floatFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsNationalMPsCell!)", format: formatZeroFloat)
            write(worksheet: summaryWorksheet, row: row, column: summaryChecksumColumn!, floatFormula: "='\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsChecksumCell!)", format: formatZeroFloat)
        }
        
        for column in [toeColumn, tablesColumn, summaryNationalLocalColumn, summaryLocalMPsColumn, summaryNationalMPsColumn, summaryChecksumColumn] {
            write(worksheet: summaryWorksheet, row: detailRow + writer.sheets.count, column: column!, string: "", format: formatBoldUnderline)
        }
        
        write(worksheet: summaryWorksheet, row: totalRow, column: summaryDescriptionColumn!, string: "Round totals", format: formatBold)
        
        let tablesColumnRef = columnRef(column: tablesColumn, fixed: true)
        write(worksheet: summaryWorksheet, row: totalRow, column: tablesColumn, integerFormula: "=SUM(\(tablesColumnRef)2:\(tablesColumnRef)\(2+writer.sheets.count))", format: formatZeroInt)
        
        let localMPsColumnRef = columnRef(column: summaryLocalMPsColumn!, fixed: true)
        write(worksheet: summaryWorksheet, row: totalRow, column: summaryLocalMPsColumn!, floatFormula: "=SUM(\(localMPsColumnRef)\(detailRow):\(localMPsColumnRef)\(detailRow + writer.sheets.count))", format: formatZeroFloat)
        
        let nationalMPsColumnRef = columnRef(column: summaryNationalMPsColumn!, fixed: true)
        write(worksheet: summaryWorksheet, row: totalRow, column: summaryNationalMPsColumn!, floatFormula: "=SUM(\(nationalMPsColumnRef)\(detailRow):\(nationalMPsColumnRef)\(detailRow + writer.sheets.count))", format: formatZeroFloat)
        
        let checksumColumnRef = columnRef(column: summaryChecksumColumn!, fixed: true)
        write(worksheet: summaryWorksheet, row: totalRow, column: summaryChecksumColumn!, floatFormula: "=SUM(\(checksumColumnRef)\(detailRow):\(checksumColumnRef)\(detailRow + writer.sheets.count))", format: formatZeroFloat)
        
        write(worksheet: summaryWorksheet, row: summaryExportedRow!, column: summaryDescriptionColumn!, string: "Exported totals", format: formatBold)
        write(worksheet: summaryWorksheet, row: summaryExportedRow!, column: summaryLocalMPsColumn!, formula: "='\(writer.csvExport.csvExportName)'!\(writer.csvExport.csvExportLocalMpsCell!)", format: formatZeroFloat)
        write(worksheet: summaryWorksheet, row: summaryExportedRow!, column: summaryNationalMPsColumn!, formula: "='\(writer.csvExport.csvExportName)'!\(writer.csvExport.csvExportNationalMpsCell!)", format: formatZeroFloat)
        write(worksheet: summaryWorksheet, row: summaryExportedRow!, column: summaryChecksumColumn!, formula: "='\(writer.csvExport.csvExportName)'!\(writer.csvExport.csvExportChecksumCell!)", format: formatZeroFloat)
        
        for column in [summaryLocalMPsColumn, summaryNationalMPsColumn, summaryChecksumColumn] {
            highlightTotalDifferent(row: summaryExportedRow!, compareRow: totalRow, column: column!)
        }
    }
    
    private func highlightTotalDifferent(row: Int, compareRow: Int, column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let field = "\(cell(row, rowFixed: true, column))"
        let matchField = "\(cell(compareRow, rowFixed: true, column))"
        let formula = "\(field)<>\(matchField)"
        setConditionalFormat(worksheet: summaryWorksheet, fromRow: row, fromColumn: column, toRow: row, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}

// MARK: - Export CSV

class CsvExportWriter: WriterBase {
    private var writer: Writer!
    var csvExportWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let csvExportName = "Csv Export"
    var csvExportLocalMpsCell: String?
    var csvExportNationalMpsCell: String?
    var csvExportChecksumCell: String?
    var csvExportDataRow: Int?
    
    init(writer: Writer) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.workbook = workbook
        csvExportWorksheet = workbook_add_worksheet(workbook, csvExportName)
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
        csvExportDataRow = 11
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
        
        setColumn(worksheet: csvExportWorksheet, column: titleColumn, width: 12)
        setColumn(worksheet: csvExportWorksheet, column: valuesColumn, width: 16)
        setColumn(worksheet: csvExportWorksheet, column: eventDataColumn, width: 10, format: formatDate)
        setColumn(worksheet: csvExportWorksheet, column: nationalIdColumn, format: formatInt)
        setColumn(worksheet: csvExportWorksheet, column: localMPsColumn, format: formatFloat)
        setColumn(worksheet: csvExportWorksheet, column: nationalMPsColumn, format: formatFloat)
        
        setColumn(worksheet: csvExportWorksheet, column: lookupFirstNameColumn, width: 12)
        setColumn(worksheet: csvExportWorksheet, column: lookupOtherNamesColumn, width: 12)
        setColumn(worksheet: csvExportWorksheet, column: lookupHomeClubColumn, width: 25)
        setColumn(worksheet: csvExportWorksheet, column: lookupRankColumn, width: 8)
        setColumn(worksheet: csvExportWorksheet, column: lookupEmailColumn, width: 25)
        setColumn(worksheet: csvExportWorksheet, column: lookupStatusColumn, width: 25)
        
            // Parameters etc
        write(worksheet: csvExportWorksheet, row: eventCodeRow, column: titleColumn, string: "Event Code:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: eventCodeRow, column: valuesColumn, formula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.ranksPlusMpsName)'!\(writer.sheets.first!.ranksPlusMps.ranksPlusMpsEventIdCell!)")
        
        write(worksheet: csvExportWorksheet, row: eventDescriptionRow, column: titleColumn, string: "Event:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: eventDescriptionRow, column: valuesColumn, formula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.ranksPlusMpsName)'!\(writer.sheets.first!.ranksPlusMps.ranksPlusMpsEventDescriptionCell!)")
        
        write(worksheet: csvExportWorksheet, row: licenseNoRow, column: titleColumn, string: "License no:", format: formatBold)
        
        write(worksheet: csvExportWorksheet, row: eventDateRow, column: titleColumn, string: "Date:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: eventDateRow, column: valuesColumn, floatFormula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.ranksPlusMpsName)'!\(writer.sheets.first!.ranksPlusMps.ranksPlusMpsEventDateCell!)", format: formatDate)
        
        write(worksheet: csvExportWorksheet, row: nationalLocalRow, column: titleColumn, string: "Nat/Local:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: nationalLocalRow, column: valuesColumn, formula: "='\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.ranksPlusMpsName)'!\(writer.sheets.first!.ranksPlusMps.ranksPlusMpsNationalLocalCell!)")
        
        write(worksheet: csvExportWorksheet, row: localMPsRow, column: titleColumn, string: "Local MPs:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: localMPsRow, column: valuesColumn, dynamicFormula: "=SUM(\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, localMPsColumn, columnFixed: true))))", format: formatZeroFloat)
        csvExportLocalMpsCell = cell(localMPsRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: csvExportWorksheet, row: nationalMPsRow, column: titleColumn, string: "National MPs:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: nationalMPsRow, column: valuesColumn, dynamicFormula: "=SUM(\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, nationalMPsColumn, columnFixed: true))))", format: formatZeroFloat)
        csvExportNationalMpsCell = cell(nationalMPsRow, rowFixed: true, valuesColumn, columnFixed: true)
        
        write(worksheet: csvExportWorksheet, row: checksumRow, column: titleColumn, string: "Checksum:", format: formatBold)
        write(worksheet: csvExportWorksheet, row: checksumRow, column: valuesColumn, dynamicFormula: "=SUM((\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, localMPsColumn, columnFixed: true)))+\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, nationalMPsColumn, columnFixed: true))))*\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, nationalIdColumn, columnFixed: true))))", format: formatZeroFloat)
        csvExportChecksumCell = cell(checksumRow, rowFixed: true, valuesColumn, columnFixed: true)
        
            // Data
        write(worksheet: csvExportWorksheet, row: titleRow, column: firstNameColumn, string: "Names", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: firstNameColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.consolidatedName)'!\(cell(writer.consolidated.consolidatedDataRow!, rowFixed: true, writer.consolidated.consolidatedFirstNameColumn!, columnFixed: true)))")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: otherNamesColumn, string: "", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: otherNamesColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.consolidatedName)'!\(cell(writer.consolidated.consolidatedDataRow!, rowFixed: true, writer.consolidated.consolidatedOtherNamesColumn!, columnFixed: true)))")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: eventDataColumn, string: "Event date", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: eventDataColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, firstNameColumn, columnFixed: true))), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\", \"\", \(cell(eventDateRow, rowFixed: true, valuesColumn, columnFixed: true)))))", format: formatDate)
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: nationalIdColumn, string: "MemNo", format: formatRightBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: nationalIdColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.consolidatedName)'!\(cell(writer.consolidated.consolidatedDataRow!, rowFixed: true, writer.consolidated.consolidatedNationalIdColumn!, columnFixed: true)))", format: formatInt)
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: eventCodeColumn, string: "Event Code", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: eventCodeColumn, dynamicFormula: "=\(byRow)(\(arrayRef)(\(cell(csvExportDataRow!, rowFixed: true, firstNameColumn, columnFixed: true))), \(lambda)(\(lambdaParam), IF(\(lambdaParam)=\"\", \"\", \(cell(eventCodeRow, rowFixed: true, valuesColumn, columnFixed: true)))))", format: formatDate)
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: clubCodeColumn, string: "Club Code", format: formatBoldUnderline)
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: localMPsColumn, string: "Local", format: formatRightBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: localMPsColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.consolidatedName)'!\(cell(writer.consolidated.consolidatedDataRow!, rowFixed: true, writer.consolidated.consolidatedLocalMPsColumn!, columnFixed: true)))", format: formatFloat)
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: nationalMPsColumn, string: "National", format: formatRightBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: nationalMPsColumn, dynamicFormula: "=\(arrayRef)('\(writer.consolidated.consolidatedName)'!\(cell(writer.consolidated.consolidatedDataRow!, rowFixed: true, writer.consolidated.consolidatedNationalMPsColumn!, columnFixed: true)))", format: formatFloat)
        
            //Lookups
        write(worksheet: csvExportWorksheet, row: titleRow, column: lookupFirstNameColumn, string: "First Name", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: lookupFirstNameColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(csvExportDataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),5,FALSE)")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: lookupOtherNamesColumn, string: "Other Names", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: lookupOtherNamesColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(csvExportDataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),4,FALSE)")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: lookupHomeClubColumn, string: "Home Club", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: lookupHomeClubColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(csvExportDataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),19,FALSE)")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: lookupRankColumn, string: "Rank", format: formatRightBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: lookupRankColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(csvExportDataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),6,FALSE)")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: lookupEmailColumn, string: "Email", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: lookupEmailColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(csvExportDataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),8,FALSE)")
        
        write(worksheet: csvExportWorksheet, row: titleRow, column: lookupStatusColumn, string: "Status", format: formatBoldUnderline)
        write(worksheet: csvExportWorksheet, row: csvExportDataRow!, column: lookupStatusColumn, dynamicFormula: "=VLOOKUP(\(arrayRef)(\(cell(csvExportDataRow!, nationalIdColumn, columnFixed: true))),'\(userDownloadData)'!\(userDownloadRange),25,FALSE)")
        
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
        let field = "\(cell(csvExportDataRow!, column, columnFixed: true))"
        let lookupfield = "\(cell(csvExportDataRow!, lookupColumn, columnFixed: true))"
        let formula = "\(field)<>\(lookupfield)"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: column, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightLookupError(fromColumn: Int, toColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let lookupCell = "\(cell(csvExportDataRow!, fromColumn))"
        let formula = "=ISNA(\(lookupCell))"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: fromColumn, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: toColumn, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadStatus(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let statusCell = "\(cell(csvExportDataRow!, column))"
        let formula = "=AND(\(statusCell)<>\"\(goodStatus)\", \(statusCell)<>\"\")"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: column, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadNationalId(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let nationalIdCell = cell(csvExportDataRow!, column, columnFixed: true)
        let firstNameCell = cell(csvExportDataRow!, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(nationalIdCell)=0, \(nationalIdCell)>\(maxNationalIdNumber), NOT(ISNUMBER(\(nationalIdCell)))))"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: column, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, checkDuplicates: true, format: format ?? formatRed!, duplicateFormat: formatRedHatched!)
    }
    
    private func highlightBadMPs(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let pointsCell = cell(csvExportDataRow!, column, columnFixed: true)
        let firstNameCell = cell(csvExportDataRow!, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(pointsCell)>\(maxPoints), AND(\(pointsCell)<>\"\", NOT(ISNUMBER(\(pointsCell))))))"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: column, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadDate(column: Int, firstNameColumn: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let dateCell = cell(csvExportDataRow!, column, columnFixed: true)
        let firstNameCell = cell(csvExportDataRow!, firstNameColumn, columnFixed: true)
        let formula = "=AND(\(firstNameCell)<>\"\", OR(\(dateCell)>DATEVALUE(\"\(Date().toString(format: "dd/MM/yyyy"))\"), \(dateCell)<DATEVALUE(\"\(Date().toString(format: "dd/MM/yyyy"))\")-30))"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: column, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
    
    private func highlightBadRank(column: Int, format: UnsafeMutablePointer<lxw_format>? = nil) {
        let rankCell = cell(csvExportDataRow!, column, columnFixed: true)
        let formula = "=AND(\(rankCell)<>\"\", OR(\(rankCell)<'\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.ranksPlusMpsName)'!\(writer.sheets.first!.ranksPlusMps.ranksPlusMpsMinRankCell!), \(rankCell)>'\(writer.sheets.first!.scoreData.roundName!) \(writer.sheets.first!.ranksPlusMps.ranksPlusMpsName)'!\(writer.sheets.first!.ranksPlusMps.ranksPlusMpsMaxRankCell!)))"
        setConditionalFormat(worksheet: csvExportWorksheet, fromRow: csvExportDataRow!, fromColumn: column, toRow: csvExportDataRow! + writer.maxPlayers - 1, toColumn: column, formula: formula, format: format ?? formatRed!)
    }
}
    
// MARK: - Consolidated
 
class ConsolidatedWriter: WriterBase {
    var writer: Writer!
    var consolidatedWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let consolidatedName = "Consolidated"
    var consolidatedDataRow: Int?
    var consolidatedFirstNameColumn: Int?
    var consolidatedOtherNamesColumn: Int?
    var consolidatedNationalIdColumn: Int?
    var consolidatedLocalMPsColumn: Int?
    var consolidatedNationalMPsColumn: Int?
    
    init(writer: Writer) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.workbook = workbook
        consolidatedWorksheet = workbook_add_worksheet(workbook, consolidatedName)
    }
    
    func write() {
        let uniqueColumn = 0
        consolidatedFirstNameColumn = 1
        consolidatedOtherNamesColumn = 2
        consolidatedNationalIdColumn = 3
        consolidatedLocalMPsColumn = 4
        consolidatedNationalMPsColumn = 5
        let dataColumn = 6
        
        let titleRow = 3
        let nationalLocalRow = 0
        let totalRow = 1
        let checksumRow = 2
        consolidatedDataRow = 4
        
        let nationalLocalRange = "\(cell(nationalLocalRow, rowFixed: true, dataColumn, columnFixed: true)):\(cell(nationalLocalRow, rowFixed: true, dataColumn + writer.sheets.count - 1, columnFixed: true))"
        
        setRow(worksheet: consolidatedWorksheet, row: titleRow, format: formatBoldUnderline)
        setRow(worksheet: consolidatedWorksheet, row: nationalLocalRow, format: formatBoldUnderline)
        setRow(worksheet: consolidatedWorksheet, row: totalRow, format: formatBoldUnderline)
        setRow(worksheet: consolidatedWorksheet, row: checksumRow, format: formatBoldUnderline)
        
        for column in 0..<(dataColumn + writer.sheets.count) {
            if column == uniqueColumn {
                setColumn(worksheet: consolidatedWorksheet, column: column, hidden: true)
            } else if column == consolidatedNationalIdColumn {
                setColumn(worksheet: consolidatedWorksheet, column: column, width: 16.5, format: formatInt)
            } else {
                setColumn(worksheet: consolidatedWorksheet, column: column, width: 16.5, format: formatFloat)
            }
        }
        
        // Title row
        write(worksheet: consolidatedWorksheet, row: titleRow, column: consolidatedFirstNameColumn!, string: "First Name", format: formatBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: consolidatedOtherNamesColumn!, string: "Other Names", format: formatBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: consolidatedNationalIdColumn!, string: "SBU", format: formatRightBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: consolidatedLocalMPsColumn!, string: "Local MPs", format: formatRightBoldUnderline)
        write(worksheet: consolidatedWorksheet, row: titleRow, column: consolidatedNationalMPsColumn!, string: "National MPs", format: formatRightBoldUnderline)
        
        for (column, sheet) in writer.sheets.enumerated() {
            // Round titles
            write(worksheet: consolidatedWorksheet, row: titleRow, column: dataColumn + column, string: sheet.prefix, format: formatRightBoldUnderline)
        
            // National/Local row
            let cell = "'\(sheet.prefix) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsNationalLocalCell!)"
            write(worksheet: consolidatedWorksheet, row: nationalLocalRow, column: dataColumn + column, formula: "=IF(\(cell)=0,\"\",\(cell))", format: formatRightBoldUnderline)
        }
        
        for element in 0...1 {
            // Total row and checksum row
            let row = (element == 0 ? totalRow : checksumRow)
            let totalRange = "\(cell(row, rowFixed: true, dataColumn, columnFixed: true)):\(cell(row, rowFixed: true, dataColumn + writer.sheets.count - 1, columnFixed: true))"
            
            write(worksheet: consolidatedWorksheet, row: row, column: consolidatedNationalIdColumn!, string: (row == totalRow ? "Total" : "Checksum"), format: formatRightBoldUnderline)
            
            write(worksheet: consolidatedWorksheet, row: row, column: consolidatedLocalMPsColumn!, floatFormula: "=SUMIF(\(nationalLocalRange), \"<>National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            write(worksheet: consolidatedWorksheet, row: row, column: consolidatedNationalMPsColumn!, floatFormula: "=SUMIF(\(nationalLocalRange), \"=National\", \(totalRange))", format: formatFloatBoldUnderline)
            
            for (column, _) in writer.sheets.enumerated() {
                let valueRange = "\(arrayRef)(\(cell(consolidatedDataRow!, rowFixed: true, dataColumn + column)))"
                let nationalIdRange = "\(arrayRef)(\(cell(consolidatedDataRow!, rowFixed: true, consolidatedNationalIdColumn!, columnFixed: true)))"
                
                if row == totalRow {
                    write(worksheet: consolidatedWorksheet, row: row, column: dataColumn + column, floatFormula: "=SUM(\(valueRange))", format: formatFloatBoldUnderline)
                } else {
                    write(worksheet: consolidatedWorksheet, row: row, column: dataColumn + column, floatFormula: "=SUM(\(valueRange)*\(nationalIdRange))", format: formatFloatBoldUnderline)
                }
            }
        }
        
        // Data rows
        let uniqueIdCell = "\(arrayRef)(\(cell(consolidatedDataRow!, uniqueColumn, columnFixed: true)))"
        
        // Unique ID column
        var formula = "=\(unique)("
        for (sheetNumber, sheet) in writer.sheets.enumerated() {
            if sheetNumber != 0 {
                formula += ","
            }
            formula += "\(vstack)(\(arrayRef)('\(sheet.scoreData.roundName!) \(sheet.individualMPs.individualMPsName)'!\(cell(1,rowFixed: true, sheet.individualMPs.individualMPsUniqueColumn!, columnFixed: true)))))"
        }
        write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: uniqueColumn, dynamicFormula: formula)
        
        // Name columns
        write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: consolidatedFirstNameColumn!, dynamicFormula: "=IF(\(uniqueIdCell)=\"\",\"\",\(fnPrefix)TEXTBEFORE(\(fnPrefix)TEXTAFTER(\(uniqueIdCell),\"+\"),\"+\"))")
        write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: consolidatedOtherNamesColumn!, dynamicFormula: "=IF(\(uniqueIdCell)=\"\",\"\",\(fnPrefix)TEXTAFTER(\(uniqueIdCell), \"+\", 2))")
        
        // National ID column
        write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: consolidatedNationalIdColumn!, dynamicIntegerFormula: "=IF(\(uniqueIdCell)=\"\",\"\",IFERROR(\(fnPrefix)NUMBERVALUE(\(fnPrefix)TEXTBEFORE(\(uniqueIdCell), \"+\")),\(fnPrefix)TEXTBEFORE(\(uniqueIdCell), \"+\")))")
        
        // Total local/national columns
        let dataRange = "\(arrayRef)(\(cell(consolidatedDataRow!, dataColumn, columnFixed: true))):\(arrayRef)(\(cell(consolidatedDataRow!, rowFixed: true, dataColumn + writer.sheets.count - 1, columnFixed: true)))"
        
        write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: consolidatedLocalMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SUMIF(\(nationalLocalRange), \"<>National\", \(lambdaParam))))")
        write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: consolidatedNationalMPsColumn!, dynamicFloatFormula: "\(byRow)(\(dataRange),\(lambda)(\(lambdaParam),SUMIF(\(nationalLocalRange), \"=National\", \(lambdaParam))))")
        
        // Lookup data columns
        for (column, sheet) in writer.sheets.enumerated() {
            let sourceDataRange = "\(cell(1, rowFixed: true, sheet.individualMPs.individualMPsUniqueColumn!, columnFixed: true)):\(cell(writer.maxPlayers, rowFixed: true, sheet.individualMPs.individualMPsDecimalColumn!, columnFixed: true))"
            let sourceOffset = (sheet.individualMPs.sheet.individualMPs.individualMPsDecimalColumn! - sheet.individualMPs.individualMPsUniqueColumn! + 1)
            write(worksheet: consolidatedWorksheet, row: consolidatedDataRow!, column: dataColumn + column, dynamicFloatFormula: "=IF(\(uniqueIdCell)=\"\",0,IFERROR(VLOOKUP(\(uniqueIdCell),'\(sheet.prefix) \(sheet.individualMPs.individualMPsName)'!\(sourceDataRange),\(sourceOffset),FALSE),0))")
        }
    }
}
    
// MARK: - Ranks plus MPs
 
class RanksPlusMPsWriter: WriterBase {
    private var writer: Writer!
    var sheet: Sheet!
    var scoreData: ScoreData!
    var ranksPlusMpsWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    
    let ranksPlusMpsName = "Ranks Plus MPs"
    var ranksPlusMpsColumns: [Column] = []
    var ranksPlusMpsPositionColumn: Int?
    var ranksPlusMpsDirectionColumn: Int?
    var ranksPlusMpsParticipantNoColumn: Int?
    var ranksPlusMpsfFirstNameColumn: [Int] = []
    var ranksPlusMpsOtherNameColumn: [Int] = []
    var ranksPlusMpsNationalIdColumn: [Int] = []
    var ranksPlusMpsBoardsPlayedColumn: [Int] = []
    var ranksPlusMpsWinDrawColumn: [Int] = []
    var ranksPlusMpsBonusMPColumn: [Int] = []
    var ranksPlusMpsWinDrawMPColumn: [Int] = []
    var ranksPlusMpsTotalMPColumn: [Int] = []
    
    var ranksPlusMpsRoundCell: String?
    var ranksPlusMpsEventDescriptionCell: String?
    var ranksPlusMpsToeCell: String?
    var ranksPlusMpsTablesCell: String?
    var ranksPlusMpsMaxAwardCell: String?
    var ranksPlusMpsMaxEwAwardCell: String?
    var ranksPlusMpsAwardToCell: String?
    var ranksPlusMpsEwAwardToCell: String?
    var ranksPlusMpsPerWinCell: String?
    var ranksPlusMpsEventDateCell: String?
    var ranksPlusMpsEventIdCell: String?
    var ranksPlusMpsNationalLocalCell: String?
    var ranksPlusMpsMinRankCell: String?
    var ranksPlusMpsMaxRankCell: String?
    var ranksPlusMpsLocalMPsCell :String?
    var ranksPlusMpsNationalMPsCell :String?
    var ranksPlusMpsChecksumCell :String?
    
    let ranksPlusMpsHeaderRows = 3
    
    init(writer: Writer, sheet: Sheet, scoreData: ScoreData) {
        super.init(workbook: writer.workbook)
        self.writer = writer
        self.sheet = sheet
        self.scoreData = scoreData
        self.workbook = workbook
        ranksPlusMpsWorksheet = workbook_add_worksheet(workbook, "\(scoreData.roundName!) \(ranksPlusMpsName)")
    }
    
    func write() {
        let event = scoreData.events.first!
        let participants = scoreData.events.first!.participants.sorted(by: sortCriteria)
        
        setupRanksPlusMpsColumns()
        writeRanksPlusMPsHeader()
        
        for (columnNumber, column) in ranksPlusMpsColumns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: ranksPlusMpsWorksheet, column: columnNumber, width: column.width)
            write(worksheet: ranksPlusMpsWorksheet, row: ranksPlusMpsHeaderRows, column: columnNumber, string: sheet.replace(column.title), format: format)
        }
        
        for playerNumber in 0..<sheet.maxParticipantPlayers {
            let nationalIdColumn = ranksPlusMpsNationalIdColumn[playerNumber]
            let firstNameColumn = ranksPlusMpsfFirstNameColumn[playerNumber]
            let otherNamesColumn = ranksPlusMpsOtherNameColumn[playerNumber]
            let firstNameNonBlank = "\(columnRef(column: firstNameColumn, fixed: true))\(ranksPlusMpsHeaderRows + 2)<>\"\""
            let otherNamesNonBlank = "\(columnRef(column: otherNamesColumn, fixed: true))\(ranksPlusMpsHeaderRows + 2)<>\"\""
            let nationalIdZero = "\(columnRef(column: nationalIdColumn, fixed: true))\(ranksPlusMpsHeaderRows + 2)=0"
            let nationalIdLarge = "\(columnRef(column: nationalIdColumn, fixed: true))\(ranksPlusMpsHeaderRows + 2)>\(maxNationalIdNumber)"
            let formula = "=AND(OR(\(firstNameNonBlank),\(otherNamesNonBlank)), OR(\(nationalIdZero),\(nationalIdLarge)))"
            setConditionalFormat(worksheet: ranksPlusMpsWorksheet, fromRow: ranksPlusMpsHeaderRows + 1, fromColumn: nationalIdColumn, toRow: ranksPlusMpsHeaderRows + sheet.fieldSize, toColumn: nationalIdColumn, formula: formula, format: formatRed!)
        }
        
        for (rowSequence, participant) in participants.enumerated() {
            
            let rowNumber = rowSequence + ranksPlusMpsHeaderRows + 1
            
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
    
    private func setupRanksPlusMpsColumns() {
        let event = scoreData.events.first!
        let playerCount = sheet.maxParticipantPlayers
        let winDraw = event.type?.requiresWinDraw ?? false
        
        ranksPlusMpsColumns.append(Column(title: "Place", content: { (participant, _) in "\(participant.place!)" }, cellType: .integer))
        ranksPlusMpsPositionColumn = ranksPlusMpsColumns.count - 1
        
        if event.winnerType == 2 && event.type?.participantType == .pair {
            ranksPlusMpsColumns.append(Column(title: "Direction", content: { (participant, _) in (participant.member as! Pair).direction?.string ?? "NS"} , cellType: .string))
            ranksPlusMpsDirectionColumn = ranksPlusMpsColumns.count - 1
        }
        
        ranksPlusMpsColumns.append(Column(title: "@P number", content: { (participant, _) in participant.member.number!} , cellType: .integer)) ; ranksPlusMpsParticipantNoColumn = ranksPlusMpsColumns.count - 1
        
        ranksPlusMpsColumns.append(Column(title: "Score", content: { (participant, _) in "\(participant.score!)" }, cellType: .float))
        
        if winDraw && sheet.maxParticipantPlayers <= event.type?.participantType?.players ?? sheet.maxParticipantPlayers {
            ranksPlusMpsColumns.append(Column(title: "Win/Draw", content: { (participant, _) in "\(participant.winDraw!)" }, cellType: .float))
            ranksPlusMpsWinDrawColumn.append(ranksPlusMpsColumns.count - 1)
        }
        
        for playerNumber in 0..<playerCount {
            ranksPlusMpsColumns.append(Column(title: "First Name (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 0) }, playerNumber: playerNumber, cellType: .string))
            ranksPlusMpsfFirstNameColumn.append(ranksPlusMpsColumns.count - 1)
            
            ranksPlusMpsColumns.append(Column(title: "Other Names (\(playerNumber+1))", playerContent: { (_, player, _, _) in self.nameColumn(name: player.name!, element: 1) }, playerNumber: playerNumber, cellType: .string))
            ranksPlusMpsOtherNameColumn.append(ranksPlusMpsColumns.count - 1)
            
            ranksPlusMpsColumns.append(Column(title: "SBU No (\(playerNumber+1))", playerContent: { (_, player,_, _) in player.nationalId! }, playerNumber: playerNumber, cellType: .integer))
            ranksPlusMpsNationalIdColumn.append(ranksPlusMpsColumns.count - 1)
            
            if sheet.maxParticipantPlayers > event.type?.participantType?.players ?? sheet.maxParticipantPlayers {
                
                ranksPlusMpsColumns.append(Column(title: "Played (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.boardsPlayed)" }, playerNumber: playerNumber, cellType: .integer))
                ranksPlusMpsBoardsPlayedColumn.append(ranksPlusMpsColumns.count - 1)
                
                if winDraw {
                    ranksPlusMpsColumns.append(Column(title: "Win/Draw (\(playerNumber+1))", playerContent: { (_, player, _, _) in "\(player.winDraw)" }, playerNumber: playerNumber, cellType: .float))
                    ranksPlusMpsWinDrawColumn.append(ranksPlusMpsColumns.count - 1)
                }
                
                ranksPlusMpsColumns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP (\(playerNumber+1))", playerContent: playerBonusAward, playerNumber: playerNumber, cellType: .floatFormula))
                ranksPlusMpsBonusMPColumn.append(ranksPlusMpsColumns.count - 1)
                
                if winDraw {
                    ranksPlusMpsColumns.append(Column(title: "Win/Draw MP (\(playerNumber+1))", playerContent: playerWinDrawAward, playerNumber: playerNumber, cellType: .floatFormula))
                    ranksPlusMpsWinDrawMPColumn.append(ranksPlusMpsColumns.count - 1)
                    
                    ranksPlusMpsColumns.append(Column(title: "Total MP (\(playerNumber+1))", playerContent: playerTotalAward, playerNumber: playerNumber, cellType: .floatFormula))
                    ranksPlusMpsTotalMPColumn.append(ranksPlusMpsColumns.count - 1)
                }
            }
        }
        
        if sheet.maxParticipantPlayers <= event.type?.participantType?.players ?? sheet.maxParticipantPlayers {
            
            ranksPlusMpsColumns.append(Column(title: "\(winDraw ? "Bonus" : "Total") MP", content: bonusAward, cellType: .floatFormula))
            ranksPlusMpsBonusMPColumn.append(ranksPlusMpsColumns.count - 1)
            
            if winDraw {
                
                ranksPlusMpsColumns.append(Column(title: "Win/Draw MP", content: winDrawAward, cellType: .integerFormula))
                ranksPlusMpsWinDrawMPColumn.append(ranksPlusMpsColumns.count - 1)
                
                ranksPlusMpsColumns.append(Column(title: "Total MP", content: totalAward, cellType: .integerFormula))
                for _ in 0..<sheet.maxParticipantPlayers {
                    ranksPlusMpsTotalMPColumn.append(ranksPlusMpsColumns.count - 1)
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
        var useAwardCell = ranksPlusMpsMaxAwardCell
        var useAwardToCell = ranksPlusMpsAwardToCell
        var firstRow = ranksPlusMpsHeaderRows + 1
        var lastRow = ranksPlusMpsHeaderRows + event.participants.count
        if let nsPairs = sheet.nsPairs {
            if event.winnerType == 2 && participant.type == .pair {
                if let pair = participant.member as? Pair {
                    if pair.direction == .ew {
                        useAwardCell = ranksPlusMpsMaxEwAwardCell
                        useAwardToCell = ranksPlusMpsEwAwardToCell
                        firstRow = ranksPlusMpsHeaderRows + nsPairs + 1
                    } else {
                        lastRow = ranksPlusMpsHeaderRows + nsPairs
                    }
                }
            }
        }
        
        if let playerNumber = playerNumber {
                // Need to calculate percentage played
            let totalBoardsPlayed = participant.member.playerList.map{$0.boardsPlayed}.reduce(0, +) / event.type!.participantType!.players
            result += "ROUNDUP((\(cell(rowNumber, ranksPlusMpsBoardsPlayedColumn[playerNumber], columnFixed: true))/\(totalBoardsPlayed))*"
        }
        
        let positionCell = cell(rowNumber, ranksPlusMpsPositionColumn!)
        let allPositionsRange = "\(cell(firstRow, rowFixed: true, ranksPlusMpsPositionColumn!, columnFixed: true)):\(cell(lastRow, rowFixed: true, ranksPlusMpsPositionColumn!, columnFixed: true))"
        
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
        
        let winDrawsCell = cell(rowNumber, ranksPlusMpsWinDrawColumn[playerNumber ?? 0], columnFixed: true)
        result = "=ROUNDUP(\(winDrawsCell) * \(ranksPlusMpsPerWinCell!), 2)"
        
        return result
    }
    
    private func totalAward(participant: Participant, rowNumber: Int) -> String {
        return playerTotalAward(participant: participant, rowNumber: rowNumber)
    }
    
    private func playerTotalAward(participant: Participant, player: Player? = nil, playerNumber: Int? = nil, rowNumber: Int) -> String {
        var result = ""
        
        let bonusMPCell = cell(rowNumber, ranksPlusMpsBonusMPColumn[playerNumber ?? 0], columnFixed: true)
        let winDrawMPCell = cell(rowNumber, ranksPlusMpsWinDrawMPColumn[playerNumber ?? 0], columnFixed: true)
        result = "\(bonusMPCell)+\(winDrawMPCell)"
        
        return result
    }
    
    // Ranks plus MPs header
    
    func writeRanksPlusMPsHeader() {
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
        
        func writeCell(floatFormula: String, format: UnsafeMutablePointer<lxw_format>? = nil) {
            column += 1
            write(worksheet: ranksPlusMpsWorksheet, row: row, column: column, floatFormula: floatFormula, format: format)
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
        
        writeCell(string: scoreData.roundName ?? "") ; ranksPlusMpsRoundCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.description ?? "") ; ranksPlusMpsEventDescriptionCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: toe) ; ranksPlusMpsToeCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        var toeCell = ranksPlusMpsToeCell!
        var ewToeRef = ""
        if twoWinners {
            writeCell(integer: sheet.fieldSize - sheet.nsPairs!) ; ewToeRef = cell(row, rowFixed: true, column, columnFixed: true)
            toeCell += "+" + ewToeRef
        }
        writeCell(integerFormula: "=ROUNDUP(\(toeCell)*(\(event.type?.participantType?.players ?? 4)/4),0)") ; ranksPlusMpsTablesCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: event.boards ?? 0)
        
        var baseMaxEwAwardCell = ""
        writeCell(float: 10.00) ; let baseMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(floatFormula: "ROUNDUP(\(baseMaxAwardCell)*\(ewToeRef)/\(ranksPlusMpsToeCell!),2)") ; baseMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 1, format: formatPercent) ; let factorCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=ROUNDUP(\(baseMaxAwardCell)*\(factorCell),2)") ; ranksPlusMpsMaxAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(floatFormula: "=ROUNDUP(\(baseMaxEwAwardCell)*\(factorCell),2)") ; ranksPlusMpsMaxEwAwardCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 0.25, format: formatPercent) ; let awardPercentCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(formula: "=ROUNDUP(\(ranksPlusMpsToeCell!)*\(awardPercentCell),0)") ; ranksPlusMpsAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        if twoWinners {
            writeCell(formula: "=ROUNDUP(\(ewToeRef)*\(awardPercentCell),0)") ; ranksPlusMpsEwAwardToCell = cell(row, rowFixed: true, column, columnFixed: true)
        }
        
        writeCell(float: 0.25) ; ranksPlusMpsPerWinCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(date: event.date ?? Date.today) ; ranksPlusMpsEventDateCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: event.eventCode ?? "") ; ranksPlusMpsEventIdCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(string: scoreData.national ? "National" : "Local") ; ranksPlusMpsNationalLocalCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: scoreData.minRank) ; ranksPlusMpsMinRankCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(integer: scoreData.maxRank) ; ranksPlusMpsMaxRankCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(ranksPlusMpsNationalLocalCell!)=\"National\",0,SUM(\(range(column: ranksPlusMpsTotalMPColumn))))") ; ranksPlusMpsLocalMPsCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=IF(\(ranksPlusMpsNationalLocalCell!)<>\"National\",0,SUM(\(range(column: ranksPlusMpsTotalMPColumn))))") ; ranksPlusMpsNationalMPsCell = cell(row, rowFixed: true, column, columnFixed: true)
        
        writeCell(floatFormula: "=SUM(\(vstack)(\(range(column: ranksPlusMpsTotalMPColumn)))*\(vstack)(\(range(column: ranksPlusMpsNationalIdColumn))))") ; ranksPlusMpsChecksumCell = cell(row, rowFixed: true, column, columnFixed: true)
    }
    
    private func range(column: [Int])->String {
        let event = scoreData.events.first!
        let firstRow = ranksPlusMpsHeaderRows + 1
        let lastRow = ranksPlusMpsHeaderRows + event.participants.count
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
    var individualMpsWorksheet: UnsafeMutablePointer<lxw_worksheet>?
    let individualMPsName = "Individual MPs"
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
    
    init(writer: Writer, sheet: Sheet, scoreData: ScoreData) {
        self.writer = writer
        self.sheet = sheet
        self.scoreData = scoreData
        super.init(workbook: writer.workbook)
        individualMpsWorksheet = workbook_add_worksheet(workbook, "\(scoreData.roundName!) \(individualMPsName)")
     }
    
    func write() {
        setupIndividualMpsColumns()
        
        for (columnNumber, column) in individualMpsColumns.enumerated() {
            var format = formatBold
            if column.cellType == .integer || column.cellType == .float || column.cellType == .integerFormula || column.cellType == .floatFormula {
                format = formatRightBold
            }
            setColumn(worksheet: individualMpsWorksheet, column: columnNumber, width: column.width)
            write(worksheet: individualMpsWorksheet, row: 0, column: columnNumber, string: sheet.replace(column.title), format: format)
        }
        
        for (columnNumber, column) in individualMpsColumns.enumerated() {
            
            if let referenceContent = column.referenceContent {
                referenceColumn(columnNumber: columnNumber, referencedContent: referenceContent, referenceDivisor: column.referenceDivisor, cellType: column.cellType)
            }
            
            if let referenceDynamicContent = column.referenceDynamic {
                referenceDynamic(columnNumber: columnNumber, content: referenceDynamicContent, cellType: column.cellType)
            }
        }
        
        let localArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, individualMPsLocalMPsColumn!, columnFixed: true)))"
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsLocalTotalColumn!, floatFormula: "=SUM(\(localArrayRef))", format: formatZeroFloat)
        
        let nationalArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, individualMPsNationalMPsColumn!, columnFixed: true)))"
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsNationalTotalColumn!, floatFormula: "=SUM(\(nationalArrayRef))", format: formatZeroFloat)
        
        let totalArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, individualMPsDecimalColumn!, columnFixed: true)))"
        let nationalIdArrayRef = "\(arrayRef)(\(cell(1, rowFixed: true, individualMPsNationalIdColumn!, columnFixed: true)))"
        write(worksheet: individualMpsWorksheet, row: 1, column: individualMPsChecksumColumn!, floatFormula: "=SUM(\(totalArrayRef)*\(nationalIdArrayRef))", format: formatZeroFloat)
        
        setColumn(worksheet: individualMpsWorksheet, column: individualMPsUniqueColumn!, hidden: true)
        
    }
    
    private func setupIndividualMpsColumns() {
        let event = scoreData.events.first!
        let twoWinners = (event.winnerType == 2 && event.type?.participantType == .pair)
        
        individualMpsColumns.append(Column(title: "Place", referenceContent: { [self] (_) in sheet.ranksPlusMps.ranksPlusMpsPositionColumn! }, cellType: .integerFormula)) ; individualMPsPositionColumn = individualMpsColumns.count - 1
        
        if twoWinners {
            individualMpsColumns.append(Column(title: "Direction", referenceContent: { [self] (_) in sheet.ranksPlusMps.ranksPlusMpsDirectionColumn! }, cellType: .stringFormula))
        }
        
        individualMpsColumns.append(Column(title: "@P no", referenceContent: { [self] (_) in sheet.ranksPlusMps.ranksPlusMpsParticipantNoColumn! }, cellType: .integerFormula))
        
        individualMpsColumns.append(Column(title: "Unique", cellType: .floatFormula)) ; individualMPsUniqueColumn = individualMpsColumns.count - 1
        let unique = individualMpsColumns.last!
        
        individualMpsColumns.append(Column(title: "Names", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.ranksPlusMpsfFirstNameColumn[playerNumber] }, cellType: .stringFormula)) ; let firstNameColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.ranksPlusMpsOtherNameColumn[playerNumber] }, cellType: .stringFormula)) ; let otherNamesColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "SBU No", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.ranksPlusMpsNationalIdColumn[playerNumber] }, cellType: .integerFormula)) ; individualMPsNationalIdColumn = individualMpsColumns.count - 1
        
        unique.referenceDynamic = { [self] in "CONCATENATE(\(arrayRef)(\(cell(1, individualMPsNationalIdColumn!, columnFixed: true))), \"+\", \(arrayRef)(\(cell(1, firstNameColumn, columnFixed: true))), \"+\", \(arrayRef)(\(cell(1, otherNamesColumn, columnFixed: true))))" }
        
        individualMpsColumns.append(Column(title: "Total MPs", referenceContent: { [self] (playerNumber) in sheet.ranksPlusMps.ranksPlusMpsTotalMPColumn[playerNumber] }, cellType: .floatFormula)) ; individualMPsDecimalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Local MPs", referenceDynamic: { [self] in "=\(byRow)(\(arrayRef)(\(cell(1, individualMPsDecimalColumn!, columnFixed: true))),\(lambda)(\(lambdaParam), IF('\(scoreData.roundName!) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsNationalLocalCell!)<>\"National\",\(lambdaParam),0)))" }, cellType: .floatFormula)) ; individualMPsLocalMPsColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "National MPs", referenceDynamic: { [self] in "=\(byRow)(\(arrayRef)(\(cell(1, individualMPsDecimalColumn!, columnFixed: true))),\(lambda)(\(lambdaParam), IF('\(scoreData.roundName!) \(sheet.ranksPlusMps.ranksPlusMpsName)'!\(sheet.ranksPlusMps.ranksPlusMpsNationalLocalCell!)=\"National\",\(lambdaParam),0)))" }, cellType: .floatFormula)) ; individualMPsNationalMPsColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Total Local", cellType: .floatFormula)) ; individualMPsLocalTotalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Total National", cellType: .floatFormula)) ; individualMPsNationalTotalColumn = individualMpsColumns.count - 1
        
        individualMpsColumns.append(Column(title: "Checksum", cellType: .floatFormula)) ; individualMPsChecksumColumn = individualMpsColumns.count - 1
    }
    
    private func referenceColumn(columnNumber: Int, referencedContent: (Int)->Int, referenceDivisor: Int? = nil, cellType: CellType? = nil) {
        
        let content = zeroFiltered(referencedContent: referencedContent, divisor: referenceDivisor)
        let position = zeroFiltered(referencedContent: { (_) in individualMPsPositionColumn! })
        
        let result = "=\(sortBy)(\(content), \(position), 1)"
        
        let column = lxw_col_t(Int32(columnNumber))
        worksheet_write_dynamic_array_formula(individualMpsWorksheet, 1, column, 999, column, result, formatFrom(cellType: cellType))
    }
    
    private func zeroFiltered(referencedContent: (Int)->Int ,divisor: Int? = nil) -> String {
        
        var result = "\(filter)(\(vstack)("
        
        for playerNumber in 0..<sheet.maxParticipantPlayers {
            let columnReference = referencedContent(playerNumber)
            if playerNumber != 0 {
                result += ","
            }
            result += "'\(scoreData.roundName!) \(sheet.ranksPlusMps.ranksPlusMpsName)'!" + cell(sheet.ranksPlusMps.ranksPlusMpsHeaderRows + 1, columnReference)
            result += ":"
            result += cell(sheet.ranksPlusMps.ranksPlusMpsHeaderRows + sheet.fieldSize, columnReference)
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
            result += "'\(scoreData.roundName!) \(sheet.ranksPlusMps.ranksPlusMpsName)'!" + cell(sheet.ranksPlusMps.ranksPlusMpsHeaderRows + 1, sheet.ranksPlusMps.ranksPlusMpsTotalMPColumn[playerNumber])
            result += ":"
            result += cell(sheet.ranksPlusMps.ranksPlusMpsHeaderRows + sheet.fieldSize, sheet.ranksPlusMps.ranksPlusMpsTotalMPColumn[playerNumber])
        }
        result += ")<>0)"
        
        return result
    }
    
    private func referenceDynamic(columnNumber: Int, content: ()->String, cellType: CellType? = nil) {
        
        let column = lxw_col_t(Int32(columnNumber))
        worksheet_write_dynamic_array_formula(individualMpsWorksheet, 1, column, 999, column, content(), formatFrom(cellType: cellType))
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
    
    func cell(sheet: String? = nil, _ row: Int, rowFixed: Bool = false, _ column: Int, columnFixed: Bool = false) -> String {
        let sheetRef = (sheet == nil ? "" : "'\(sheet!)'!")
        let rowRef = rowRef(row: row, fixed: rowFixed)
        let columnRef = columnRef(column: column, fixed: columnFixed)
        return "\(sheetRef)\(columnRef)\(rowRef)"
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
