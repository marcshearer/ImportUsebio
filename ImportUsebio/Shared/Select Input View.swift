//
//  Select Input.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

enum Level: Int {
    case local = 0
    case national = 1
}

enum Basis: Int {
    case standard = 0
    case manual = 1
    case headToHead = 2
}

struct SelectInputView: View {
    @State private var inputFilename: String = ""
    @State private var roundName: String = ""
    @State private var eventCode: String = ""
    @State private var basis: Basis = .standard
    @State private var eventDescription: String = ""
    @State private var filterSessionId: String = ""
    @State private var maxTeamMembers: Int = 0
    @State private var manualPointsColumn: String?
    @State private var localNational = Level.local
    @State private var drawsRounded = true
    @State private var winDrawLevel: WinDrawLevel = .board
    @State private var mergeMatches = false
    @State private var vpType: VpType = .discrete
    @State private var minRank: Int = 0
    @State private var maxRank: Int = 999
    @State private var maxAward: Float = 10.0
    @State private var ewMaxAward: Float = 0.0
    @State private var minEntry: Int = 0
    @State private var awardTo: Float = 25
    @State private var perWin: Float = 0.25
    @State private var securityBookmark: Data? = nil
    @State private var refresh = true
    @State private var content: [String] = []
    @State private var writer: Writer? = nil
    @State private var scoreData: ScoreData? = nil
    @State private var roundErrors: [RoundErrorList] = []
    @State private var showErrors = false
    @State private var showSettings = false
    @State private var missingNationalIds = false
    @State private var editSettings = Settings.current.copy()

    var body: some View {
        
        // Just to trigger view refresh
        if refresh { EmptyView() }
        
        StandardView("Select Input") {
            VStack {
                Spacer().frame(height: 10)
                
                HStack {
                    Spacer().frame(width: 30)
                    
                    VStack {
                        VStack {
                            
                            HStack {
                                Input(title: "Event Description:", field: $eventDescription, width: 400, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                                Spacer()
                                VStack {
                                    settingsButton()
                                    Spacer().frame(height:30)
                                }
                                    
                                Spacer().frame(width: 8)
                            }
                            
                            InputTitle(title: "Coding", topSpace: 16)
                            Spacer().frame(height: 8)
                            HStack {
                                Spacer().frame(width: 42)
                                Input(title: "Event code:", field: $eventCode, topSpace: 0, width: 100, inlineTitle: true, inlineTitleWidth: 90, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                                Spacer().frame(width: 40)
                                HStack {
                                    Picker("Calculation basis:", selection: $basis) {
                                        Text("Standard").tag(Basis.standard)
                                        Text("Manual").tag(Basis.manual)
                                        Text("Head-to-head").tag(Basis.headToHead)
                                    }
                                    .help(Text("The method to be used for calculations\nIf this is Head-to-head this means that it is a round were only win/draw awards are applied. There is no place award.\nIf this is Manual then the MPs are not calculated. They are specified in the raw data."))
                                    .onChange(of: basis, initial: false) {
                                        if basis != .standard {
                                            maxAward = 0
                                            ewMaxAward = 0
                                            minEntry = 0
                                            if basis == .headToHead {
                                                awardTo = 100
                                            } else if basis == .manual {
                                                awardTo = 0
                                                perWin = 0
                                            }
                                        }
                                    }
                                    .pickerStyle(.menu)
                                    .focusable(false)
                                    .frame(width: 250)
                                    .font(inputFont)
                                }
                                Spacer()
                            }
                            
                            InputTitle(title: "Ranking restrictions:", topSpace: 16)
                            Spacer().frame(height: 8)
                            HStack {
                                Spacer().frame(width: 30)
                                
                                InputInt(title: "Minimum:", field: $minRank, topSpace: 0, width: 50, inlineTitle: true, inlineTitleWidth: 90)
                                
                                InputInt(title: "Maximum:", field: $maxRank, topSpace: 0, width: 50, inlineTitle: true, inlineTitleWidth: 90)
                                
                                Spacer()
                            }
                            
                            Spacer().frame(height: 16)
                            Separator(thickness: 1)
                            Spacer().frame(height: 16)
                        }
                        VStack {
                            InputTitle(title: "Import Parameters:")
                            HStack {
                                Spacer().frame(width: 42)
                                HStack {
                                    Input(title: "Session ID filter:", field: $filterSessionId, topSpace: 0, width: 100, inlineTitle: true, inlineTitleWidth: 120, isEnabled: true)
                                        .help(Text("Used to restrict the import of a Usebio raw data file to a specific Session Id"))
                                }
                                Spacer().frame(width: 10)
                                HStack {
                                    Picker("Win/Draws from:   ", selection: $winDrawLevel) {
                                        Text("Pairs/Teams").tag(WinDrawLevel.participant)
                                        Text("Matches").tag(WinDrawLevel.match)
                                        Text("Boards").tag(WinDrawLevel.board)
                                    }
                                    .onChange(of: winDrawLevel, initial: false) {
                                        if winDrawLevel == .participant {
                                            mergeMatches = false
                                        }
                                        if winDrawLevel != .board {
                                            vpType = .discrete
                                        }
                                    }
                                    .help(Text("Use Win/Draw data from PARTICIPANT or from MATCH or from MATCH recalculated from BOARD scores - Falls back if data not available at level requested"))
                                    .pickerStyle(.menu)
                                    .focusable(false)
                                    .frame(width: 250)
                                    .font(inputFont)
                                }
                                .frame(height: inputDefaultHeight)
                                Spacer().frame(width: 20)
                                HStack {
                                    Picker("Merge matches:  ", selection: $mergeMatches) {
                                        Text("Merge").tag(true)
                                        Text("Don't Merge").tag(false)
                                    }
                                    .help(Text("Combine matches between the same teams/pairs before calculating wins/draws"))
                                    .pickerStyle(.menu)
                                    .focusable(false)
                                    .frame(width: 230)
                                    .font(inputFont)
                                    .disabled(winDrawLevel == .participant)
                                }
                                .frame(height: inputDefaultHeight)
                                Spacer()
                            }
                            Spacer().frame(height: 10)
                            HStack {
                                Spacer().frame(width: 30)
                                HStack {
                                    InputInt(title: "Max players:", field: $maxTeamMembers, width: 100, inlineTitle: true, inlineTitleWidth: 120)
                                        .help(Text("Used to limit the number of team members considered in the raw data. Primarily used to remove subs from the awards when their impact has not been material"))
                                }
                                Spacer().frame(width: 52)
                                HStack {
                                    Picker("VP type:   ", selection: $vpType) {
                                        Text("Discrete").tag(VpType.discrete)
                                        Text("Continuous").tag(VpType.continuous)
                                    }
                                    .help(Text("When re-calculating from boards what sort of VPs are required?"))
                                    .pickerStyle(.menu)
                                    .focusable(false)
                                    .frame(width: 180)
                                    .font(inputFont)
                                    .disabled(winDrawLevel == .participant)
                                }
                                .frame(height: inputDefaultHeight)
                                Spacer().frame(width: 73)
                                HStack {
                                    Picker("VP Draws:  ", selection: $drawsRounded) {
                                        Text("Exact").tag(false)
                                        Text("Rounded").tag(true)
                                    }
                                    .help(Text("Round Continuous VPs when calculating draws"))
                                    .pickerStyle(.menu)
                                    .focusable(false)
                                    .frame(width: 200)
                                    .font(inputFont)
                                    .disabled(winDrawLevel == .participant || vpType == .discrete)
                                    Spacer()
                                }
                                .frame(height: inputDefaultHeight)
                                Spacer()
                            }
                            HStack {
                                Input(title: "Import filename:", field: $inputFilename, placeHolder: "No import file specified", height: 30, width: 684, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: true, isReadOnly: true).frame(width: 750)
                                
                                VStack() {
                                    Spacer()
                                    self.finderButton()
                                }.frame(height: 62)
                                
                                Spacer()
                            }
                            
                            InputTitle(title: "Other details:", topSpace: 16)
                            Spacer().frame(height: 8)
                            HStack {
                                Spacer().frame(width: 42)
                                HStack {
                                    Input(title: "Round name:", field: $roundName, topSpace: 0, width: 160, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                    Spacer()
                                }
                                .frame(width: 270)
                                Spacer()
                            }
                            
                            HStack {
                                VStack {
                                    InputTitle(title: "Points Award Details:", topSpace: 16)
                                    VStack {
                                        HStack {
                                            Spacer().frame(width: 42)
                                            Picker("Level:             ", selection: $localNational) {
                                                Text("Local").tag(Level.local)
                                                Text("National").tag(Level.national)
                                            }
                                            .pickerStyle(.segmented)
                                            .focusable(false)
                                            .frame(width: 300)
                                            .font(inputFont)
                                            Spacer()
                                        }
                                        
                                        Spacer().frame(height: 10)
                                        
                                        HStack(spacing: 0) {
                                            Spacer().frame(width: 30)
                                            
                                            InputFloat(title: "Max award:", field: $maxAward, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 90)
                                                .disabled(basis != .standard)
                                            
                                            InputFloat(title: "Max E/W award:", field: $ewMaxAward, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 120)
                                                .disabled(basis != .standard)
                                            
                                            Spacer()
                                        }
                                        
                                        Spacer().frame(height: 10)
                                        
                                        HStack(spacing: 0) {
                                            Spacer().frame(width: 30)
                                            
                                            InputInt(title: "Min entry:", field: $minEntry, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 90)
                                                .disabled(basis != .standard)
                                            
                                            InputFloat(title: "Award to %:", field: $awardTo, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 120)
                                                .disabled(basis != .standard)
                                            
                                            InputFloat(title: "Per win:", field: $perWin, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 70)
                                                .disabled(basis == Basis.manual)
                                            
                                            Spacer()
                                        }
                                    }
                                }
                                Spacer()
                            }
                            
                            Spacer().frame(height: 24)
                            Separator(thickness: 1)
                            Spacer().frame(height: 12)
                            HStack {
                                Spacer()
                                addSheetButton()
                                Spacer().frame(width: 50)
                                finishButton()
                                Spacer().frame(width: 50)
                                clearButton()
                                Spacer().frame(width: 50)
                                pasteButton()
                                Spacer()
                            }
                            Spacer().frame(height: 20)
                        }
                    }
                }
            }
            Spacer()
        }
        .sheet(isPresented: $showErrors) {
            ShowErrorsView(roundErrors: roundErrors)
        }
        .sheet(isPresented: $showSettings) {
            SettingsView(settings: editSettings)
        }
    }
    
    private func settingsButton() -> some View {
        return Button {
            showSettings = true
        } label: {
            Image(systemName: "gearshape.fill").font(.largeTitle).foregroundColor(Palette.banner.background)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
    }
    
    private func finderButton() -> some View {
        
        return Button {
            FileSystem.findFile(title: "Select Source File", prompt: "Select", types: ["xml", "csv"]) { (url, bookmarkData) in
                Utility.mainThread {
                    refresh.toggle()
                    securityBookmark = bookmarkData
                    if let data = try? Data(contentsOf: url) {
                        let type = url.pathExtension.lowercased()
                        if type == "xml" {
                            _ = UsebioParser(fileUrl: url, data: data, filterSessionId: filterSessionId == "" ? nil : filterSessionId, overrideEventType: (basis == .headToHead ? .head_to_head: nil), roundContinuousVPDraw: drawsRounded, winDrawLevel: winDrawLevel, mergeMatches: mergeMatches, vpType: vpType, completion: parserComplete)
                        } else if type == "csv" {
                            if  GenericCsvParser(fileUrl: url, data: data, completion: parserComplete) == nil {
                                MessageBox.shared.show("Invalid data")
                            }
                        } else {
                            MessageBox.shared.show("File type \(type) not supported")
                        }
                    } else {
                        MessageBox.shared.show("Unable to read data")
                    }
                }
            }
        } label: {
            Text("Choose file")
                .foregroundColor(Palette.enabledButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.enabledButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
    }
    
    private func addSheetButton() -> some View {
        return Button{
            if let scoreData = scoreData {
                if writer == nil {
                    writer = Writer()
                    writer?.eventDescription = eventDescription
                    writer?.eventCode = eventCode
                    writer?.minRank = minRank
                    writer?.maxRank = maxRank
                }
                scoreData.roundName = roundName
                scoreData.national = (localNational == .national)
                scoreData.maxTeamMembers = (maxTeamMembers == 0 ? nil : maxTeamMembers)
                scoreData.roundContinuousVPDraw = drawsRounded
                scoreData.winDrawLevel = winDrawLevel
                scoreData.mergeMatches = mergeMatches
                scoreData.vpType = vpType
                scoreData.maxAward = maxAward
                scoreData.ewMaxAward = (ewMaxAward != 0 ? ewMaxAward : nil)
                scoreData.minEntry = minEntry
                scoreData.reducedTo = 1
                scoreData.awardTo = awardTo
                scoreData.perWin = perWin
                writer?.add(name: roundName, scoreData: scoreData)
                
                MessageBox.shared.show("Added Successfully", okAction: {
                    self.scoreData = nil
                    self.inputFilename = ""
                    self.roundName = ""
                })
                Utility.executeAfter(delay: 2) {
                    self.scoreData = nil
                    self.inputFilename = ""
                    self.roundName = ""
                    MessageBox.shared.hide()
                }
            }
        } label: {
            Text("Add Sheet")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(scoreData == nil || eventCode == "" || eventDescription == "" || roundName == "" || awardTo <= 0 || (writer?.rounds.contains(where: {$0.shortName == roundName}) ?? false))
    }
    
   
    private func finishButton() -> some View {
        return Button{
            FileSystem.saveFile(title: "Generated Workbook Name", prompt: "Save", filename: "\(eventDescription).xlsm") { (url) in
                Utility.mainThread {
                    if let writer = writer {
                        writer.write(as: url.relativePath)
                        MessageBox.shared.show("Written Successfully", okAction: {
                            self.writer = nil
                        })
                        Utility.executeAfter(delay: 2) {
                            self.writer = nil
                            MessageBox.shared.hide()
                        }
                    }
                }
            }
        } label: {
            Text("Create Workbook")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 150, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(writer == nil)
    }
    
    private func pasteButton() -> some View {
        
        return Button {
            if let data = NSPasteboard.general.string(forType: .string) {
                let dataLines = data.replacingOccurrences(of: "\n", with: "").components(separatedBy: "\r")
                let lines = dataLines.map{$0.components(separatedBy: "\t")}
                ImportRounds.process(lines) { (imported, error, warning) in
                    if let error = error {
                        MessageBox.shared.show(error)
                    } else if let imported = imported {
                        updateFromImport(imported: imported)
                    }
                }
            } else {
                MessageBox.shared.show("Invalid clipboard contents")
            }
        } label: {
            Text("Paste Config")
                .foregroundColor(Palette.enabledButton.text)
                .frame(width: 120, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.enabledButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(writer?.rounds.count ?? 0 > 0)
    }
    
    private func clearButton() -> some View {
        return Button{
            self.writer = nil
        } label: {
            Text("Clear")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(writer == nil)
    }
    
    private func parserComplete(scoreData: ScoreData?, messages: [String]) {
        if let scoreData = scoreData {
            let (roundMissingNationalId, errors, warnings) = scoreData.validate()
            missingNationalIds = roundMissingNationalId
            
            let filename = scoreData.fileUrl?.lastPathComponent.removingPercentEncoding ?? ""
            
            var errorList = RoundErrorList(name: filename, errors: [], warnings: [])
            if let errors = errors {
                errorList.errors = errors
            }
            if let warnings = warnings {
                errorList.warnings = warnings
            }
            if !messages.isEmpty {
                errorList.warnings.append(contentsOf: messages)
            }
            
            roundErrors = []
            if errors != nil || warnings != nil {
                roundErrors.append(errorList)
            }
            if errors != nil {
                self.scoreData = nil
            } else {
                self.scoreData = scoreData
                self.inputFilename = filename
            }
            addMissingNationalIdWarning()
            if !roundErrors.isEmpty {
                showErrors = true
            }
        } else {
            MessageBox.shared.show(messages.last ?? "Unable to parse file \(inputFilename)")
            self.scoreData = nil
            self.inputFilename = ""
            self.roundName = ""
        }
    }
    
    @State private var importInProgress: ImportRounds?
    @State private var importRound: Int = 0
    @State private var sourceDirectory: String?
    
    private func updateFromImport(imported: ImportRounds) {
        FileSystem.findDirectory(prompt: "Select source directory") { (url, bookmarkData) in
            Utility.mainThread {
                securityBookmark = bookmarkData
                importInProgress = imported
                sourceDirectory = url.relativePath
                importRound = 0
                roundErrors = []
                missingNationalIds = false
                processNextRound()
            }
        }
    }
    
    private func processNextRound() {
        if importRound <= (importInProgress?.rounds.count ?? 0) - 1 {
            let round = importInProgress!.rounds[importRound]
            let url = URL(fileURLWithPath: sourceDirectory! + "/" + round.filename!)
            if let data = try? Data(contentsOf: url) {
                Utility.mainThread {
                    round.scoreData?.fileUrl = url
                    let type = url.pathExtension.lowercased()
                    if type == "xml" {
                        _ = UsebioParser(fileUrl: url, data: data, filterSessionId: round.filterSessionId, filterParticipantNumberMin: round.filterParticipantNumberMin, filterParticipantNumberMax: round.filterParticipantNumberMax, overrideEventType: (round.headToHead ? .head_to_head: nil), roundContinuousVPDraw: round.roundContinuousVPDraw, winDrawLevel: round.winDrawLevel, mergeMatches: round.mergeMatches, vpType: round.vpType, completion: importParserComplete)
                    } else if type == "csv" {
                        if GenericCsvParser(fileUrl: url, data: data, manualPointsColumn: round.manualPointsColumn, completion: importParserComplete) == nil {
                            MessageBox.shared.show("Invalid data")
                        }
                    } else {
                        MessageBox.shared.show("File type \(type) not supported")
                    }
                }
            } else {
                MessageBox.shared.show("Unable to access file \(round.filename!)")
                writer = nil
            }
        } else {
            eventDescription = importInProgress!.event!.description!
            eventCode = importInProgress!.event!.code!
            minRank = importInProgress!.event!.minRank ?? 0
            maxRank = ((importInProgress!.event!.maxRank ?? 999) == 0 ? 999 : importInProgress!.event!.maxRank!)
            
            writer = Writer()
            writer!.eventDescription = eventDescription
            writer!.eventCode = eventCode
            writer!.minRank = minRank
            writer!.maxRank = maxRank
            
            for round in importInProgress!.rounds {
                if let scoreData = round.scoreData {
                    roundName = round.name!
                    localNational = (round.localNational ?? importInProgress!.event!.localNational) == .national ? .national : .local
                    maxAward = round.maxAward ?? 0
                    ewMaxAward = round.ewMaxAward ?? maxAward
                    minEntry = round.minEntry ?? 0
                    awardTo = round.awardTo ?? 0
                    perWin = round.perWin ?? 0
                    filterSessionId = round.filterSessionId ?? ""
                    maxTeamMembers = round.maxTeamMembers ?? 0
                    manualPointsColumn = round.manualPointsColumn
                    
                    scoreData.roundName = roundName
                    scoreData.national = localNational == .national
                    scoreData.maxAward = maxAward
                    scoreData.ewMaxAward = ewMaxAward
                    scoreData.reducedTo = round.reducedTo ?? 1
                    scoreData.minEntry = minEntry
                    scoreData.awardTo = awardTo * 100
                    scoreData.perWin = perWin
                    scoreData.filterSessionId = filterSessionId
                    scoreData.aggreateAs = round.aggregateAs
                    scoreData.maxTeamMembers = (maxTeamMembers == 0 ? nil : maxTeamMembers)
                    scoreData.manualPointsColumn = manualPointsColumn
                    if let writerRound = writer?.add(name: round.name!, shortName: round.shortName!, scoreData: round.scoreData!) {
                        writerRound.toe = round.toe
                    }
                }
            }
                    
            importInProgress = nil
            addMissingNationalIdWarning()
            if !roundErrors.isEmpty {
                showErrors = true
            }
        }
    }
    
    private func importParserComplete(scoreData: ScoreData?, messages: [String]) {
        if let scoreData = scoreData {
            let (roundMissingNationalIds, errors, warnings) = scoreData.validate()
            missingNationalIds = missingNationalIds || roundMissingNationalIds
            
            var errorList = RoundErrorList(name: importInProgress!.rounds[importRound].name!, errors: [], warnings: [])
            if let errors = errors {
                errorList.errors =  errors
            }
            if let warnings = warnings {
                errorList.warnings = warnings
            }
            if !messages.isEmpty {
                errorList.warnings.append(contentsOf: messages)
            }
            if !errorList.errors.isEmpty || !errorList.warnings.isEmpty {
                roundErrors.append(errorList)
            }
            
            importInProgress!.rounds[importRound].scoreData = scoreData
            importRound += 1
            processNextRound()
            
        } else {
            MessageBox.shared.show(messages.last ?? "Unable to parse file \(importInProgress!.rounds[importRound].filename!)")
            self.writer = nil
            self.scoreData = nil
            self.inputFilename = ""
            self.roundName = ""
        }
    }
    
    private func addMissingNationalIdWarning() {
        if missingNationalIds {
            roundErrors.append(RoundErrorList(name: "General", errors: [], warnings: ["Some players have missing National Ids"]))
        }
    }
    
}
