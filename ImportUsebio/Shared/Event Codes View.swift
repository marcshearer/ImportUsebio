//
//  Event Codes View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/08/2025.
//

import SwiftUI
import UniformTypeIdentifiers

struct EventImportView: View {
    @Environment(\.dismiss) private var dismiss
    @State private var droppedFiles: [(filename: String, contents: String)] = []
    @State private var dropZoneEntered = false
    private let uttypes = [UTType.data]
    
    var body: some View {
        StandardView("EventImport", slideInId: UUID()) {
            dropZone
        }
        .frame(width: 400, height: 450)
    }
    
    var dropZone: some View {
        VStack(spacing: 0) {
            Banner(title: Binding.constant("Import Event Code Data"))
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
                                Text("Drop Event Code CSV file here").font(bannerFont)
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
                        importEventCodes(droppedFiles.first!.contents)
                    }
                    droppedFiles = []
                }
            }
        }
    }
    
    func importEventCodes(_ csvData: String) {
        var events: [EventViewModel] = []
        
        let lines = csvData.replacingOccurrences(of: "\r\n", with: "\n").components(separatedBy: "\n")
        if lines.count < 2 {
            MessageBox.shared.show("Invalid CSV data - insufficient lines")
        } else {
            // Check header line
            let headerLine = lines[0].components(separatedBy: ",").map{$0.replacing("\"", with: "")}
            if let eventCodeIndex = headerLine.firstIndex(of: "Event Code"),
                let eventNameIndex = headerLine.firstIndex(of: "Event Name"),
                let activeIndex = headerLine.firstIndex(of: "Active"),
                let startDateIndex = headerLine.firstIndex(of: "Start Date"),
                let endDateIndex = headerLine.firstIndex(of: "End Date"),
                let validMinRankIndex = headerLine.firstIndex(of: "Valid Min Rank"),
                let validMaxRankIndex = headerLine.firstIndex(of: "Valid Max Rank"),
                let originatingClubCodeIndex = headerLine.firstIndex(of: "Originating Club"),
                let clubCodeOptionalIndex = headerLine.firstIndex(of: "Club Code Optional"),
                let nationalPointsAllowedIndex = headerLine.firstIndex(of: "National Points Allowed"),
                let localPointsAllowedIndex = headerLine.firstIndex(of: "Local Points Allowed") {
                
                let minColumns = max(eventCodeIndex, eventNameIndex, activeIndex, startDateIndex, endDateIndex, validMinRankIndex, validMaxRankIndex, originatingClubCodeIndex, clubCodeOptionalIndex, nationalPointsAllowedIndex, localPointsAllowedIndex) + 1
                for line in lines.dropFirst() {
                    let fields = line.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
                    if fields.count < minColumns {
                        MessageBox.shared.show("Invalid CSV data - line has fewer than require columns")
                        break
                    } else {
                        let eventCode = fields[eventCodeIndex]
                        let eventName = fields[eventNameIndex]
                        let active = (fields[activeIndex] == "Active")
                        let startDate = Utility.dateFromString(fields[startDateIndex])
                        let endDate = Utility.dateFromString(fields[endDateIndex])
                        let validMinRank = RankViewModel.rank(rankName: fields[validMinRankIndex])?.rankCode ?? 0
                        let validMaxRank = RankViewModel.rank(rankName: fields[validMaxRankIndex])?.rankCode ?? 999
                        let originatingClubCode = ClubViewModel.club(clubName: fields[originatingClubCodeIndex])?.clubCode ?? ""
                        let clubCodeOptional = (fields[clubCodeOptionalIndex] == "Yes")
                        let nationalPointsAllowed = (fields[nationalPointsAllowedIndex] == "Yes")
                        let localPointsAllowed = (fields[localPointsAllowedIndex] == "Yes")
                        
                        let event = EventViewModel(eventCode: eventCode, eventName: eventName, active: active, startDate: startDate, endDate: endDate, validMinRank: validMinRank, validMaxRank: validMaxRank, originatingClubCode: originatingClubCode, clubMandatory: !clubCodeOptional, nationalAllowed: nationalPointsAllowed, localAllowed: localPointsAllowed)
                        
                        events.append(event)
                    }
                }
                // Clear existing events
                for event in MasterData.shared.events.array {
                    event.remove()
                }
                // Now insert new events
                for event in events {
                    event.insert()
                }
                MessageBox.shared.show("\(events.count) event codes imported successfully", okAction: {
                    dismiss()
                })
            } else {
                MessageBox.shared.show("Invalid CSV data - Mandatory columns not found")
            }
        }
    }
}
