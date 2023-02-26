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
    @State private var roundName: String = "Pairs"
    @State private var eventCode: String = ""
    @State private var eventDescription: String = ""
    @State private var nationalLocal = Level.local
    @State private var minRank: Int = 0
    @State private var maxRank: Int = 999
    @State private var maxAward: Float = 10.0
    @State private var minField: Int = 0
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
                                        Picker("Level:             ", selection: $nationalLocal) {
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
                                        
                                        InputInt(title: "Min entry:", field: $minField, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 90)
                                        
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
            FileSystem.findFile { (url, bookmarkData) in
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
                scoreData.national = (nationalLocal == .national)
                scoreData.maxAward = maxAward
                scoreData.minField = minField
                scoreData.awardTo = awardTo
                scoreData.perWin = perWin
                writer?.add(prefix: roundName, scoreData: scoreData)
                
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
        .disabled(scoreData == nil || eventCode == "" || eventDescription == "" || roundName == "" || (writer?.rounds.contains(where: {$0.name == roundName}) ?? false))
    }
    
   
    private func finishButton() -> some View {
        return Button{
            if let writer = writer {
                writer.write()
                MessageBox.shared.show("Written Successfully", okAction: {
                })
                Utility.executeAfter(delay: 2) {
                    MessageBox.shared.hide()
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
    
    private func parserComplete(scoreData: ScoreData) {
        let (errors, warnings) = scoreData.validate()
        if let errors = errors {
            // TODO: Handle errors
            print(errors)
        } else {
            if let warnings = warnings {
                // TODO: Show warnings
                print(warnings)
            }
            self.scoreData = scoreData
            self.inputFilename = scoreData.fileUrl?.lastPathComponent.removingPercentEncoding ?? ""
        }
    }
}
