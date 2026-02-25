//
//  Rank Import View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/08/2025.
//

import SwiftUI
import UniformTypeIdentifiers

struct RankImportView: View {
    @Environment(\.dismiss) private var dismiss
    var completion: () -> ()
    @State private var droppedFiles: [(filename: String, contents: String)] = []
    @State private var dropZoneEntered = false
    private let uttypes = [UTType.data]
    
    var body: some View {
        StandardView("RankImport", slideInId: UUID()) {
            dropZone
        }
        .frame(width: 400, height: 450)
    }
    
    var dropZone: some View {
        VStack(spacing: 0) {
            Banner(title: Binding.constant("Import Rank Code Data"), backAction: { completion() ; return true })
            HStack {
                Spacer().frame(width: 50)
                VStack {
                    Spacer().frame(height: 50)
                    ZStack {
                        RoundedRectangle(cornerRadius: 30, style: .continuous)
                            .foregroundColor(dropZoneEntered ? Palette.contrastTile.background : Palette.background.background)
                            .frame(width: 300, height: 300)
                        HStack {
                            Spacer().frame(width: 50)
                            Spacer()
                            VStack {
                                Spacer()
                                Text("Drop Rank Code CSV file here").font(bannerFont)
                                    .multilineTextAlignment(.center)
                                Spacer()
                            }
                            Spacer()
                            Spacer().frame(width: 50)
                        }
                        .frame(width: 300, height: 300)
                        .overlay(RoundedRectangle(cornerRadius: 30)
                            .strokeBorder(style: StrokeStyle(lineWidth: 5, dash: [10, 5]))
                            .foregroundColor(Palette.gridLine))
                    }
                    .onDrop(of: uttypes, delegate: DropFiles(dropZoneEntered: $dropZoneEntered, droppedFiles: $droppedFiles))
                    Spacer().frame(height: 50)
                }
                Spacer().frame(width: 50)
            }
            .onChange(of: droppedFiles.count, initial: false) {
                if !droppedFiles.isEmpty {
                    if droppedFiles.count > 1 {
                        MessageBox.shared.show("Only one file can be dropped")
                    } else {
                        importRankCodes(droppedFiles.first!.contents)
                    }
                    droppedFiles = []
                }
            }
        }
    }
    
    func importRankCodes(_ csvData: String) {
        var ranks: [RankViewModel] = []
        
        let lines = csvData.replacingOccurrences(of: "\r\n", with: "\n").components(separatedBy: "\n")
        if lines.count < 2 {
            MessageBox.shared.show("Invalid CSV data - insufficient lines")
        } else {
            // Check header line
            let headerLine = lines[0].components(separatedBy: ",").map{$0.replacing("\"", with: "")}
            if let rankCodeIndex = headerLine.firstIndex(of: "Code"),
                let rankNameIndex = headerLine.firstIndex(of: "Name") {
                
                let minColumns = max(rankCodeIndex, rankNameIndex) + 1
                for line in lines.dropFirst() {
                    let fields = line.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
                    if fields.count < minColumns {
                        MessageBox.shared.show("Invalid CSV data - line has fewer than required columns")
                        break
                    } else {
                        let rankCode = Int(fields[rankCodeIndex]) ?? 0
                        
                        if rankCode <= 0 {
                            MessageBox.shared.show("Invalid CSV data - non-positive integer rank code found")
                            break
                        }
                        
                        let rankName = fields[rankNameIndex]
                        
                        let rank = RankViewModel(rankCode: rankCode, rankName: rankName)
                        
                        ranks.append(rank)
                    }
                }
                // Clear existing ranks
                for rank in MasterData.shared.ranks.array {
                    rank.remove()
                }
                // Now insert new ranks
                for rank in ranks {
                    rank.insert()
                }
                MessageBox.shared.show("\(ranks.count) rank codes imported successfully", okAction: {
                    dismiss()
                })
            } else {
                MessageBox.shared.show("Invalid CSV data - Mandatory columns not found")
            }
        }
    }
}
