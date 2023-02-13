//
//  SelectInput.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

struct SelectInputView: View {
    @State private var inputFilename: String = ""
    @State private var securityBookmark: Data? = nil
    @State private var refresh = true
    @State private var content: [String] = []
    @State private var parser: Parser? = nil
    @State private var scoreData: ScoreData? = nil

    var body: some View {
        
            // Just to trigger view refresh
        if refresh { EmptyView() }
        
        VStack {
            Spacer().frame(height: 30)
            HStack {
                Spacer().frame(width: 30)
                                
                Input(title: "Import filename:", field: $inputFilename, message:nil, height: 30, width: 700, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: true, isReadOnly: true).frame(width: 750)
                
                VStack() {
                    Spacer()
                    self.finderButton()
                }.frame(height: 52)
                Spacer()
            }
            Spacer()
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
            let writer = Writer(scoreData: scoreData)
            writer.write()
        }
    }
}
