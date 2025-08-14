//
//  Rank Codes Search View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 13/08/2025.
//

import SwiftUI

struct RankCodesSearchView: View {
    @Environment(\.dismiss) private var dismiss
    var showNoMinimum: Bool = false
    var showNoMaximum: Bool = false
    var selectAction: (RankViewModel)->()
    @State private var ranks = MasterData.shared.ranks.array as! [RankViewModel]
    
    var body: some View {
        VStack(spacing: 0) {
            // Combined banner and list headings
            VStack {
                Spacer()
                HStack {
                    HStack {
                        Spacer()
                        Image(systemName: "chevron.left").font(toolbarFont)
                        Spacer()
                    }
                    .frame(width: 12)
                    .onTapGesture {
                        dismiss()
                    }
                    HStack {
                        Spacer()
                        Text("Code")
                    }
                    .frame(width: 50)
                    Spacer().frame(width: 20)
                    HStack {
                        Text("Description")
                        Spacer()
                    }
                    .frame(width: 150)
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
                    ForEach(ranks, id: \.rankCode) { (rank) in
                        VStack {
                            Spacer()
                            HStack {
                                Spacer().frame(width: 40)
                                HStack {
                                    Spacer()
                                    Text("\(rank.rankCode)")
                                }
                                .frame(width: 50)
                                Spacer().frame(width: 20)
                                HStack {
                                    Text("\(rank.rankName)")
                                    Spacer()
                                }
                                .frame(width: 150)
                                Spacer()
                            }
                            .contentShape(Rectangle())
                            .foregroundColor(Palette.background.text)
                            Spacer()
                        }
                        .onTapGesture {
                            selectAction(rank)
                            dismiss()
                        }
                        .frame(height: 15)
                    }
                }
            }
        }
        .background(Palette.background.background)
        .ignoresSafeArea(.all)
        .frame(width: 300, height: 450)
        .onAppear {
            if showNoMinimum {
                ranks.insert(RankViewModel(rankCode: 0, rankName: "No minimum rank"), at: 0)
            }
            if showNoMaximum {
                ranks.append(RankViewModel(rankCode: 999, rankName: "No maximum rank"))
            }
        }
    }
}
