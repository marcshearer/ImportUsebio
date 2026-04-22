//
//  Member Import View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 17/04/2026.
//

import SwiftUI
import UniformTypeIdentifiers

struct MemberImportView: View {
    @Environment(\.dismiss) private var dismiss
    public var completion: ()->()
    @State private var droppedFiles: [(filename: String, contents: String)] = []
    @State private var dropZoneEntered = false
    @State private var message = "Drop User Download file here"
    @State private var isImporting = false
    private let uttypes = [UTType.data]
    
    var body: some View {
        StandardView("EventImport", slideInId: UUID()) {
            dropZone
        }
        .frame(width: 400, height: 450)
    }
    
    var dropZone: some View {
        return VStack(spacing: 0) {
            Banner(title: Binding.constant("Import Member Data"), backEnabled: {!isImporting})
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
                                Text(message).font(bannerFont)
                                    .multilineTextAlignment(.center)
                                Spacer().frame(height: 20)
                                if isImporting {
                                    ProgressView("")
                                        .progressViewStyle(.linear)
                                }
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
                    message = "Importing\n\nPlease wait..."
                    if droppedFiles.count > 1 {
                        MessageBox.shared.show("Only one file can be dropped")
                        droppedFiles = []
                    } else {
                        let (_, csvData) = droppedFiles.first!
                        Task {
                            isImporting = true
                            await MemberList.shared.importDropped(csvData) { (success, message) in
                                if success {
                                    MessageBox.shared.show("Import complete", okAction: {
                                        dismiss()
                                        completion()
                                    })
                                } else {
                                    MessageBox.shared.show("Import failed.\n(\(message))", okAction: { dismiss() })
                                }
                            }
                            isImporting = false
                            droppedFiles = []
                        }
                    }
                }
            }
        }
    }
}
