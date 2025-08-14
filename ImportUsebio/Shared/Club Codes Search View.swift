//
//  Club Codes Search View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 14/08/2025.
//

import SwiftUI

struct ClubCodesSearchView: View {
    @Environment(\.dismiss) private var dismiss
    var selectAction: (ClubViewModel)->()
    @State private var searchText: String = ""
    @State private var clubs = (MasterData.shared.clubs.array as! [ClubViewModel]).sorted { $0.clubCode < $1.clubCode }
    @State private var filtered: [ClubViewModel] = []
    
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
                        Text("Club Codes")
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
                    ForEach(filtered, id: \.clubCode) { (club) in
                        VStack {
                            Spacer()
                            HStack {
                                Spacer().frame(width: 40)
                                HStack {
                                    Text("\(club.clubCode)")
                                    Spacer()
                                }
                                .frame(width: 100)
                                Spacer().frame(width: 20)
                                HStack {
                                    Text("\(club.clubName)")
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
                            selectAction(club)
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
        filtered = clubs.filter {
            (searchText == "" || Utility.wordSearch(for: searchText, in: $0.clubCode + " " + $0.clubName)) && !$0.clubName.uppercased().contains("CLOSED")
        }
    }
}
