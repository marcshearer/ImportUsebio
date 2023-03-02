//
//  SelectInput.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

enum Level: Int {
    case local = 0
    case national = 1
}

struct SelectInputView: View {
    @State private var inputFilename: String = ""
    @State private var roundName: String = ""
    @State private var eventCode: String = ""
    @State private var eventDescription: String = ""
    @State private var localNational = Level.local
    @State private var minRank: Int = 0
    @State private var maxRank: Int = 999
    @State private var maxAward: Float = 10.0
    @State private var minEntry: Int = 0
    @State private var awardTo: Float = 25
    @State private var perWin: Float = 0.25
    @State private var securityBookmark: Data? = nil
    @State private var refresh = true
    @State private var content: [String] = []
    @State private var parser: Parser? = nil
    @State private var writer: Writer? = nil
    @State private var scoreData: ScoreData? = nil

    var body: some View {
        
            // Just to trigger view refresh
        if refresh { EmptyView() }
        
        StandardView("Select Input") {
            HStack {
                
                Spacer().frame(width: 30)
                VStack {
                    Spacer().frame(height: 10)
                    VStack {
                        
                        HStack {
                            Input(title: "Event Description:", field: $eventDescription, width: 400, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                        }
                        
                        HStack {
                            Input(title: "Event code:", field: $eventCode, width: 100, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                            Spacer()
                        }
                        
                        InputTitle(title: "Ranking restrictions:", topSpace: 16)
                        HStack {
                            Spacer().frame(width: 30)
                            
                            InputInt(title: "Minimum:", field: $minRank, topSpace: 0, width: 50, inlineTitle: true, inlineTitleWidth: 90)
                            
                            InputInt(title: "Maximum:", field: $maxRank, topSpace: 0, width: 50, inlineTitle: true, inlineTitleWidth: 90)
                            
                            Spacer()
                        }
                        
                        Spacer().frame(height: 8)
                        Separator(thickness: 1)
                        Spacer().frame(height: 16)
                    }
                    VStack {
                        
                        HStack {
                            Input(title: "Import filename:", field: $inputFilename, placeHolder: "No import file specified", height: 30, width: 700, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: true, isReadOnly: false).frame(width: 750)
                            
                            VStack() {
                                Spacer()
                                self.finderButton()
                            }.frame(height: 62)
                            
                            Spacer()
                        }
                        
                        Spacer().frame(height: 12)
                        
                        HStack {
                            Input(title: "Round name:", field: $roundName, width: 160, isEnabled: true)
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
                                        .frame(width: 300)
                                        .font(inputFont)
                                        Spacer()
                                    }
                                    
                                    Spacer().frame(height: 16)
                                    
                                    HStack(spacing: 0) {
                                        Spacer().frame(width: 30)
                                        
                                        InputFloat(title: "Max award:", field: $maxAward, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 90)
                                        
                                        InputInt(title: "Min entry:", field: $minEntry, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 90)
                                        
                                        InputFloat(title: "Award to %:", field: $awardTo, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 90)
                                        
                                        InputFloat(title: "Per win:", field: $perWin, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 70)
                                        
                                        Spacer()
                                    }
                                }
                            }
                            Spacer()
                        }
                        
                        Spacer().frame(height: 60)
                        
                        HStack {
                            Spacer().frame(width: 16)
                            addSheetButton()
                            Spacer().frame(width: 50)
                            finishButton()
                            Spacer().frame(width: 50)
                            clearButton()
                            Spacer().frame(width: 50)
                            pasteButton()
                            Spacer()
                        }
                        Spacer()
                    }}
                Spacer()
            }
        }
    }
    
    private func finderButton() -> some View {
        
        return Button {
            FileSystem.findFile(types: ["xml"]) { (url, bookmarkData) in
                Utility.mainThread {
                    refresh.toggle()
                    securityBookmark = bookmarkData
                    if let data = try? Data(contentsOf: url) {
                        parser = Parser(fileUrl: url, data: data, completion: parserComplete)
                    } else {
                        // TODO: Handle failure
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
                scoreData.maxAward = maxAward
                scoreData.minEntry = minEntry
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
        .disabled(scoreData == nil || eventCode == "" || eventDescription == "" || roundName == "" || (writer?.rounds.contains(where: {$0.shortName == roundName}) ?? false))
    }
    
   
    private func finishButton() -> some View {
        return Button{
            FileSystem.findDirectory(prompt: "Select target directory") { (url, bookmarkData) in
                Utility.mainThread {
                    if let writer = writer {
                        writer.write(in: url.relativePath)
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
        .disabled(writer == nil)
    }
    
    private func pasteButton() -> some View {
        
        return Button {
            if let data = NSPasteboard.general.string(forType: .string) {
                let dataLines = data.replacingOccurrences(of: "\n", with: "").components(separatedBy: "\r")
                let lines = dataLines.map{$0.components(separatedBy: "\t")}
                Import.process(lines) { (imported, error, warning) in
                    if let error = error {
                        MessageBox.shared.show(error)
                    } else if let imported = imported {
                        updateFromImport(imported: imported)
                    }
                }
            }
        } label: {
            Text("Paste")
                .foregroundColor(Palette.enabledButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.enabledButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
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
        .disabled(writer == nil)
    }
    
    private func parserComplete(scoreData: ScoreData, parseErrors: [String], parseWarnings: [String]) {
        let (errors, warnings) = scoreData.validate()
        if let errors = errors {
            // TODO: Handle errors
            print(parseErrors + errors)
        } else {
            if let warnings = warnings {
                // TODO: Show warnings
                print(parseWarnings + warnings)
            }
            self.scoreData = scoreData
            self.inputFilename = scoreData.fileUrl?.lastPathComponent.removingPercentEncoding ?? ""
        }
    }
    
    @State private var importInProgress: Import?
    @State private var importRound: Int = 0
    @State private var sourceDirectory: String?
    
    private func updateFromImport(imported: Import) {
        FileSystem.findDirectory(prompt: "Select source directory") { (url, bookmarkData) in
            Utility.mainThread {
                securityBookmark = bookmarkData
                importInProgress = imported
                sourceDirectory = url.relativePath
                importRound = 0
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
                    parser = Parser(fileUrl: url, data: data, completion: importParserComplete)
                }
            } else {
                MessageBox.shared.show("Unable to access file \(round.filename!)")
                writer = nil
            }
        } else {
            eventDescription = importInProgress!.event!.description!
            eventCode = importInProgress!.event!.code!
            minRank = importInProgress!.event!.minRank!
            maxRank = (importInProgress!.event!.maxRank! == 0 ? 999 : importInProgress!.event!.maxRank!)
            
            writer = Writer()
            writer!.eventDescription = eventDescription
            writer!.eventCode = eventCode
            writer!.minRank = minRank
            writer!.maxRank = maxRank
            
            for round in importInProgress!.rounds {
                if let scoreData = round.scoreData {
                    roundName = round.name!
                    localNational = (round.localNational ?? importInProgress!.event!.localNational) == .national ? .national : .local
                    maxAward = round.maxAward!
                    minEntry = round.minEntry ?? 0
                    awardTo = round.awardTo!
                    perWin = round.perWin!
                    
                    scoreData.roundName = roundName
                    scoreData.national = localNational == .national
                    scoreData.maxAward = maxAward
                    scoreData.minEntry = minEntry
                    scoreData.awardTo = awardTo * 100
                    scoreData.perWin = perWin
                    if let writerRound = writer?.add(name: round.name!, shortName: round.shortName!, scoreData: round.scoreData!) {
                        writerRound.toe = round.toe
                    }
                }
            }
                    
            importInProgress = nil
        }
    }
    
    private func importParserComplete(scoreData: ScoreData, parseErrors: [String], parseWarnings: [String]) {
        let (errors, warnings) = scoreData.validate()
        if let errors = errors {
            // TODO: Handle errors
            print(parseErrors + errors)
        } else {
            if let warnings = warnings {
                // TODO: Show warnings
                print(parseWarnings + warnings)
            }
            importInProgress!.rounds[importRound].scoreData = scoreData
        }
        importRound += 1
        processNextRound()
    }
    
}
