//
//  SelectInput.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

struct SelectInputView: View {
    @State private var inputFilename: String = ""
    @State private var roundName: String = "Pairs"
    @State private var eventCode: String = ""
    @State private var eventDescription: String = ""
    @State private var national: Bool = false
    @State private var minRank: Int = 0
    @State private var maxRank: Int = 999
    @State private var securityBookmark: Data? = nil
    @State private var refresh = true
    @State private var content: [String] = []
    @State private var parser: Parser? = nil
    @State private var scoreData: ScoreData? = nil

    var body: some View {
        
            // Just to trigger view refresh
        if refresh { EmptyView() }
        
        StandardView("Select Input") {
            HStack {
                
                Spacer().frame(width: 30)
                VStack {
                    
                    Spacer().frame(height: 30)
                    
                    HStack {
                        Input(title: "Event Description:", field: $eventDescription, height: 30, width: 400, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                    }
                    
                    HStack {
                        Input(title: "Event code:", field: $eventCode, height: 30, width: 100, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                        Spacer()
                    }
                    
                    VStack {
                        
                        InputInt(title: "Minimum rank:", field: $minRank, inlineTitle: false)
                        
                        InputInt(title: "Maximum rank:", field: $maxRank, inlineTitle: false)
                    }
                    
                    HStack {
                        Input(title: "Import filename:", field: $inputFilename, height: 30, width: 700, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: true, isReadOnly: true).frame(width: 750)
                        
                        VStack() {
                            Spacer()
                            self.finderButton()
                        }.frame(height: 52)
                        
                        Spacer()
                    }
                    
                    HStack {
                        Input(title: "Round name:", field: $roundName, height: 30, width: 100, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                        Spacer()
                    }
                    
                    HStack {
                        VStack {
                            InputTitle(title: "Points Award Level:")
                            HStack {
                                Spacer().frame(width: 32)
                                Toggle(isOn: $national) { Text("National") }.font(inputFont)
                                Spacer()
                            }
                        }
                        Spacer()
                    }
                    
                    Spacer().frame(height: 20)
                    
                    HStack {
                        Spacer().frame(width: 32)
                        Button{
                            if let scoreData = scoreData {
                                let writer = Writer()
                                writer.eventDescription = eventDescription
                                writer.eventCode = eventCode
                                writer.minRank = minRank
                                writer.maxRank = maxRank
                                scoreData.roundName = roundName
                                scoreData.national = national
                                writer.add(prefix: roundName, scoreData: scoreData)
                                writer.write()
                                MessageBox.shared.show("Processed Successfully")
                            }
                        } label: {
                            Text("Process")
                                .font(inputFont)
                                .padding(.all, 10)
                                .frame(width: 100.0, height: 30.0, alignment: .center)
                                .foregroundColor(Palette.bannerButton.text)
                        }
                        .background(Palette.bannerButton.background)
                        .buttonStyle(PlainButtonStyle())
                        .shadow(radius: 2, x: 5, y: 5)
                        
                        Spacer()
                    }
                    Spacer()
                }
                Spacer()
            }
        }
    }
    
    private func finderButton() -> some View {
        
        return Button(action: {
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
        },label: {
            Text("Change...")
        })
        .buttonStyle(DefaultButtonStyle())
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
