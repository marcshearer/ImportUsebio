//
//  Event Codes Search View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 13/08/2025.
//

import SwiftUI

struct EventCodesSearchView: View {
    @Environment(\.dismiss) private var dismiss
    var selectAction: (EventViewModel)->()
    @State private var searchText: String = ""
    @State private var events = (MasterData.shared.events.array as! [EventViewModel]).sorted { $0.eventCode < $1.eventCode }
    @State private var filtered: [EventViewModel] = []
    
    var body: some View {
        VStack(spacing: 0) {
            // Banner
            VStack {
                Spacer()
                HStack {
                    HStack {
                        Spacer().frame(width: 20)
                        Image(systemName: "chevron.left")
                        Spacer().frame(width: 10)
                    }
                    .onTapGesture {
                        dismiss()
                    }
                    HStack {
                        Text("Event Codes")
                        Spacer()
                    }
                }
                .font(toolbarFont)
                Spacer()
            }
            .frame(height: 30)
            .background(Palette.alternateBanner.background)
            .foregroundColor(Palette.alternateBanner.text)
            .bold()
            
            // Search text
            VStack {
                Spacer().frame(height: 10)
                HStack {
                    Spacer().frame(width: 20)
                    Input(title: "Search text", field: $searchText, width: 350, isEnabled: true, onChange: { newText in
                        applyFilter()
                    })
                    Spacer()
                }
                Spacer().frame(height: 20)
            }
                
            // List headings
            VStack {
                Spacer()
                HStack {
                    Spacer().frame(width: 40)
                    HStack {
                        Text("Code")
                        Spacer()
                    }
                    .frame(width: 100)
                    Spacer().frame(width: 20)
                    HStack {
                        Text("Description")
                        Spacer()
                    }
                    .frame(width: 250)
                    Spacer()
                }
                Spacer()
            }
            .frame(height: 30)
            .background(Palette.gridTitle.background)
            .foregroundColor(Palette.gridTitle.text)
            .bold()
            
            // List of results
            ScrollView {
                LazyVStack {
                    ForEach(filtered, id: \.eventCode) { (event) in
                        VStack {
                            Spacer()
                            HStack {
                                Spacer().frame(width: 40)
                                HStack {
                                    Text("\(event.eventCode)")
                                    Spacer()
                                }
                                .frame(width: 100)
                                Spacer().frame(width: 20)
                                HStack {
                                    Text("\(event.eventName)")
                                    Spacer()
                                }
                                .frame(width: 250)
                                Spacer()
                            }
                            Spacer()
                        }
                        .foregroundColor(Palette.background.text)
                        .contentShape(Rectangle())
                        .onTapGesture {
                            selectAction(event)
                            dismiss()
                        }
                        .frame(height: 15)
                    }
                }
            }
        }
        .onAppear {
            applyFilter()
        }
        .ignoresSafeArea(.all)
        .background(Palette.background.background)
        .ignoresSafeArea(.all)
        .frame(width: 450, height: 450)
    }
    
    func applyFilter() {
        filtered = events.filter {
            (searchText == "" || Utility.wordSearch(for: searchText, in: $0.eventCode + " " + $0.eventName)) && $0.active && ($0.startDate ?? Date()) <= Date() && ($0.endDate ?? Date()) >= Date()
        }
    }
}
