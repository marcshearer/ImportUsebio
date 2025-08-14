//
//  Club Import View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/08/2025.
//

import SwiftUI
import UniformTypeIdentifiers

struct ClubImportView: View {
    @Environment(\.dismiss) private var dismiss
    @State private var droppedFiles: [(filename: String, contents: String)] = []
    @State private var dropZoneEntered = false
    private let uttypes = [UTType.data]
    
    var body: some View {
        StandardView("ClubImport", slideInId: UUID()) {
            dropZone
        }
        .frame(width: 400, height: 450)
    }
    
    var dropZone: some View {
        VStack(spacing: 0) {
            Banner(title: Binding.constant("Import Club Code Data"))
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
                                Text("Drop Club Code CSV file here").font(bannerFont)
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
                        importClubCodes(droppedFiles.first!.contents)
                    }
                    droppedFiles = []
                }
            }
        }
    }
    
    func importClubCodes(_ csvData: String) {
        var clubs: [ClubViewModel] = []
        
        let lines = csvData.replacingOccurrences(of: "\r\n", with: "\n").components(separatedBy: "\n")
        if lines.count < 2 {
            MessageBox.shared.show("Invalid CSV data - insufficient lines")
        } else {
            // Check header line
            let headerLine = lines[0].components(separatedBy: ",").map{$0.replacing("\"", with: "")}
            if let referenceIndex = headerLine.firstIndex(of: "Reference") {
                
                let minColumns = referenceIndex + 1
                for line in lines.dropFirst() {
                    let fields = line.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
                    if fields.count < minColumns {
                        MessageBox.shared.show("Invalid CSV data - line has fewer than required columns")
                        break
                    } else {
                        let bits = fields[referenceIndex].split(separator: " ")
                        let count = bits.count
                        if count < 4 || bits[count - 3] != "(" || bits[count - 1] != ")" {
                            MessageBox.shared.show("Invalid reference column")
                            break
                        }
                        let clubName = bits[0...count - 4].joined(separator: " ")
                        let clubCode = String(bits[count - 2])
                                                
                        let club = ClubViewModel(clubCode: clubCode, clubName: clubName)
                        
                        clubs.append(club)
                    }
                }
                // Clear existing clubs
                for club in MasterData.shared.clubs.array {
                    club.remove()
                }
                // Now insert new clubs
                for club in clubs {
                    club.insert()
                }
                MessageBox.shared.show("\(clubs.count) club codes imported successfully", okAction: {
                    dismiss()
                })
            } else {
                MessageBox.shared.show("Invalid CSV data - Mandatory columns not found")
            }
        }
    }
}
