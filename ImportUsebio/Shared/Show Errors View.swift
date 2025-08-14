//
//  Show Errors.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 02/03/2023.
//

import SwiftUI

struct RoundErrorList {
    let name: String
    var errors: [String]
    var warnings: [String]
}

struct ShowErrorsView: View {
    @Environment(\.dismiss) private var dismiss
    
    @State var roundErrors: [RoundErrorList]

    @State private var refresh = true
    
    var body: some View {
        if refresh { EmptyView() }
        
        StandardView("Select Input") {
                HStack {
                    Spacer().frame(width: 16)
                    GeometryReader { (geometry) in
                        VStack {
                            Spacer().frame(height: 16)
                            ScrollView {
                                ForEach(roundErrors, id: \.self.name) { roundError in
                                    VStack {
                                        HStack {
                                            Text(roundError.name)
                                                .bold()
                                            Spacer()
                                        }
                                        Spacer().frame(height: 4)
                                        ForEach(roundError.errors, id: \.self) { error in
                                            HStack {
                                                Spacer().frame(width: 20)
                                                HStack {
                                                    Text("Error: ")
                                                    Spacer()
                                                }
                                                .frame(width: 70)
                                                Text(error)
                                                Spacer()
                                            }
                                            .foregroundColor(Palette.background.strongText)
                                        }
                                        ForEach(roundError.warnings, id: \.self) { warning in
                                            HStack {
                                                Spacer().frame(width: 20)
                                                HStack {
                                                    Text("Warning: ")
                                                    Spacer()
                                                }
                                                .frame(width: 70)
                                                Text(warning)
                                                Spacer()
                                            }
                                            .foregroundColor(Palette.background.faintText)
                                        }
                                        Spacer().frame(height: 16)
                                    }
                                }
                                Spacer()
                            }
                            .frame(width: geometry.size.width)
                            Spacer().frame(height: 10)
                        }
                    }
                }
                HStack {
                    Spacer()
                    Button{
                        dismiss()
                    } label: {
                        Text("Continue")
                            .foregroundColor(Palette.highlightButton.text)
                            .frame(width: 100, height: 30)
                            .font(.callout).minimumScaleFactor(0.5)
                            .background(Palette.enabledButton.background)
                            .foregroundColor(Palette.enabledButton.text)
                            .cornerRadius(15)
                    }
                    .buttonStyle(PlainButtonStyle())
                    .focusable(false)
                    Spacer()
            }
            Spacer().frame(height: 10)
        }
        .frame(width: 800, height: 540)
    }
}
