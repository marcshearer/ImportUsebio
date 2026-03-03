//
//  Show Participants View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 27/02/2026.
//

import SwiftUI

enum MemberStatus: Equatable {
    case ok
    case updated
    case memberNotFound(suggested: Int)
    case veryDifferent(suggested: Int)
    case slightlyDifferent
    case triviallyDifferent
    
    var string: String {
        switch self {
        case .ok: return "OK"
        case .updated: return "Updated"
        case .memberNotFound: return "Not found"
        case .veryDifferent: return "Very Different"
        case .slightlyDifferent: return "Slightly Different"
        case .triviallyDifferent: return "Trivially Different"
        }
    }
    
    var matches: String {
        switch self {
        case .memberNotFound(let suggested), .veryDifferent(let suggested):
            "\(suggested) \(suggested == 1 ? "match" : "matches")"
        default:
            ""
        }
    }
    
    var priority: Int {
        switch self {
        case .ok: 0
        case .updated: 0
        case .memberNotFound: 1
        case .veryDifferent: 2
        case .slightlyDifferent: 3
        case .triviallyDifferent: 4
        }
    }
}

@Observable class ParticipantData: Identifiable {
    var id: UUID = UUID()
    var nationalId = ""
    var names = ""
    var memberNationalId: String?
    var memberNames: String?
    var possibleMatches: [MemberViewModel] = []
    var suggested = false
    var player: Player
    var linked: [ParticipantData] = []
    var condition: PlayerCondition = .undefined
    var updated: Bool = false
    
    var status: MemberStatus {
        if memberNationalId == nil || suggested {
            return .memberNotFound(suggested: possibleMatches.count)
        } else if names != memberNames {
            let difference = Utility.levenshteinDistance(names, memberNames ?? "")
            if difference == 0 {
                return .triviallyDifferent
            } else if difference <= 5 {
                return .slightlyDifferent
            } else {
                return .veryDifferent(suggested: possibleMatches.count)
            }
        } else {
            if updated {
                return .updated
            } else {
                return .ok
            }
        }
    }
    
    init(imported: Player) {
        self.player = imported
        self.nationalId = imported.nationalId ?? ""
        self.names = imported.name ?? ""
         
        if let member = MemberViewModel.member(nationalId: nationalId) {
            memberNationalId = member.nationalId
            memberNames = member.names
        }
        switch status {
        case .memberNotFound, .veryDifferent:
            possibleMatches = MemberViewModel.member(names: names).filter( { BlockedViewModel.blocked(nationalId: $0.nationalId) == nil } )
        default:
            break
        }
    }
    
    func copy() -> ParticipantData {
        ParticipantData(imported: player)
    }
    
    func updateFromMember() {
        nationalId = memberNationalId ?? possibleMatches.first?.nationalId ?? nationalId
        names = memberNames ?? possibleMatches.first?.names ?? names
        memberNationalId = nationalId
        memberNames = names
        possibleMatches = []
        updated = true
        suggested = false
        // Write back to imported data
        for item in [self] + self.linked {
            item.player.nationalId = nationalId
            item.player.name = names
            item.player.condition = condition
        }
    }
    
}

struct ParticipantsView: View {
    @State var participants: [ParticipantData] = []
    @State var exit: Bool = true
    @State var showMatches = false
    @State var selected = ParticipantData(imported: Player())
    @State var chooseOnly = false
    
    let tableColumns = [GridItem(.fixed(80),  spacing: 10, alignment: .trailing),
                        GridItem(.fixed(140), spacing: 30, alignment: .leading),
                        GridItem(.fixed(70), spacing: 20, alignment: .center),
                        GridItem(.fixed(80),  spacing: 10, alignment: .center),
                        GridItem(.fixed(140), spacing: 30, alignment: .leading),
                        GridItem(.fixed(120), spacing: 10, alignment: .leading),
                        GridItem(.fixed(80),  spacing: 10, alignment: .leading)]
    
    var body: some View {
        StandardView("Select Input") {
            VStack(spacing: 0) {
                Banner(title: Binding.constant("Check Players Details"), bottomSpace: false, backEnabled: { exit }, backAction: {
                    return true
                })
                ZStack {
                    VStack(spacing: 0) {
                        Spacer().frame(height: 30)
                        Rectangle()
                            .foregroundColor(Palette.background.background)
                    }
                    VStack(spacing: 0) {
                        ScrollView(showsIndicators: true) {
                            VStack(spacing: 0) {
                                LazyVGrid(columns: tableColumns, spacing: 0, pinnedViews: [.sectionHeaders]) {
                                    Section(header: bannerRow()) {
                                        ForEach(self.participants.filter({ if case .ok = $0.status { return false }; return true })) { participant in
                                            gridRow(participant: participant)
                                                .frame(height: 30)
                                        }
                                    }
                                }
                            }
                        }
                        Spacer().frame(height: 5)
                    }
                }
            }
        }
        .sheet(isPresented: $showMatches) {
            ChoosePossibleMatches(participant: $selected, chooseOnly: $chooseOnly)
        }
        .interactiveDismissDisabled(!exit)
        .frame(width: 850, height: 550)
    }
    
    func bannerRow() -> some View {
        ZStack {
            Rectangle()
                .foregroundColor(Palette.tile.background)
                .frame(height: 40)
            VStack(spacing: 0) {
                Spacer()
                LazyVGrid(columns: tableColumns, spacing: 0) {
                    gridRow("Imported", "Imported", "Database","Database", "", "")
                        .frame(height: 15)
                    gridRow("National Id","Names", "National Id", "Names", "Status", "")
                        .frame(height: 15)
                }
                .bold()
                Spacer()
            }
            .frame(height: 40)
        }
    }
    
    func gridRow(_ nationalId: String, _ names: String, _ memberNationalId: String, _ memberNames: String, _ status: String, _ matches: String, participant: ParticipantData? = nil, databaseColor: ThemeTextType = .normal, editAction: ((ParticipantData, Bool)->())? = nil) -> some View {
        
        GridRow {
            TrailingClickableText(nationalId)
            LeadingClickableText(names)
            if let participant = participant {
                actionButtons(participant, { (participant, chooseOnly) in
                    selected = participant
                    self.chooseOnly = chooseOnly
                    showMatches = true
                })
            } else {
                Text("")
            }
            TrailingClickableText(memberNationalId).foregroundColor(Palette.background.textColor(databaseColor))
            LeadingClickableText(memberNames).foregroundColor(Palette.background.textColor(databaseColor))
            LeadingClickableText(status)
            LeadingClickableText(matches)
        }
    }
    
    func gridRow(participant: ParticipantData, editAction: ((ParticipantData, Bool)->())? = nil) -> some View{
        if participant.memberNationalId == nil, let suggest = participant.possibleMatches.first {
            participant.memberNationalId = suggest.nationalId
            participant.memberNames = suggest.otherNames + " " + suggest.lastName
            participant.suggested = true
        }
        return gridRow(participant.nationalId, participant.names, participant.memberNationalId ?? "", participant.memberNames ?? "", participant.status.string, participant.status.matches, participant: participant, databaseColor: (participant.suggested ? .faint : .normal))
    }
    
    func actionButtons(_ participant: ParticipantData, _ editAction: ((ParticipantData, Bool)->())? = nil) -> some View {
        HStack(spacing: 0) {
            if (participant.memberNationalId != nil || !participant.possibleMatches.isEmpty) && (participant.status != .ok && participant.status != .updated) {
                Button(action: {
                    if participant.possibleMatches.count > 1 {
                        editAction?(participant, true)
                    } else {
                        participant.updateFromMember()
                    }
                }) {
                    Text("􁉈").frame(width: 25, height: 25).palette(.highlightButton).cornerRadius(12.5)
                }
                .buttonStyle(PlainButtonStyle())
                .focusable(false)
            }
            Spacer()
            Button(action: {
                editAction?(participant, false)
            }) {
                Text("􀈊").frame(width: 25, height: 25).palette(.highlightButton).cornerRadius(12.5)
            }
            .buttonStyle(PlainButtonStyle())
            .focusable(false)
        }
    }
}

fileprivate enum ViewType {
    case nationalId
    case names
}

struct ChoosePossibleMatches : View {
    @Environment(\.dismiss) var dismiss
    @Binding var participant: ParticipantData
    @Binding var chooseOnly: Bool
    @State var nationalId: String = ""
    @State var names: String = ""
    @State var condition: PlayerCondition = .undefined
    @State var distance = 0
    @FocusState private var focused: ViewType?
    
    let maxDistance = 5
    @State var matches: [MemberViewModel] = []
    @State var notFound: Bool = false
    
    let tableColumns = [GridItem(.fixed(80),  spacing: 0, alignment: .trailing),
                        GridItem(.fixed(140), spacing: 0, alignment: .leading),
                        GridItem(.fixed(140),  spacing: 0, alignment: .leading),
                        GridItem(.fixed(140), spacing: 0, alignment: .leading)]
    
    var body: some View {
        StandardView("Select Input") {
            VStack(spacing: 0) {
                Banner(title: Binding.constant(chooseOnly ? "Choose Player to Update From" : "Update Player Details"), bottomSpace: false, back: false)
                Spacer().frame(height: 30)
                HStack {
                    Spacer().frame(width: 20)
                    VStack {
                        Spacer().frame(height: 10)
                        Text("Possible Matches: ")
                            .frame(width: 120)
                        Spacer()
                    }
                    .frame(height: 120)
                    Spacer().frame(width: 10)
                    VStack {
                        ZStack {
                            VStack(spacing: 0) {
                                Spacer().frame(height: 30)
                                UnevenRoundedRectangle(cornerRadii: .init(bottomLeading: 6, bottomTrailing: 6), style: .continuous)
                                    .foregroundColor(Palette.alternate.background)
                            }
                            ScrollView(showsIndicators: true) {
                                VStack(spacing: 0) {
                                    LazyVGrid(columns: tableColumns, alignment: .center, spacing: 0, pinnedViews: [.sectionHeaders]) {
                                        Section(header: heading()) {
                                            ForEach(matches) { member in
                                                GridRow {
                                                    let rank = RankViewModel.rank(rankCode: member.rankCode)
                                                    TrailingClickableText(member.nationalId)
                                                    LeadingClickableText(member.names)
                                                    LeadingClickableText(member.homeClub)
                                                    LeadingClickableText(rank?.rankName ?? "Unknown")
                                                }
                                                .frame(height: 30)
                                                .padding(.horizontal, 5)
                                                .palette(nationalId == member.nationalId && names == member.names ? .highlightTile : .alternate)
                                                .onTapGesture {
                                                    nationalId = member.nationalId
                                                    names = member.names
                                                    focused = nil
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        .frame(width: 490, height: 120)
                    }
                    Spacer()
                    VStack {
                        Spacer().frame(height: 10)
                        CustomButton.button(title: "Widen Search", width: 120) {
                            widenSearch()
                        }
                        .disabled(distance == maxDistance)
                        Spacer()
                    }
                    .frame(height: 120)
                    Spacer()
                }
                Spacer().frame(height: 30)
                
                if !chooseOnly {
                    HStack {
                        Spacer().frame(width: 20)
                        Input(title: "National Id:", field: $nationalId, width: 140, inlineTitle: true, inlineTitleWidth: 125, onChange: { newValue in
                            if let lookup = MemberViewModel.member(nationalId: newValue) {
                                names = lookup.names
                                focused = nil
                                notFound = false
                                condition = .valid
                            } else {
                                notFound = true
                                condition = (condition == .valid ? .undefined : condition)
                            }
                        })
                        .focused($focused, equals: .nationalId)
                        Spacer().frame(width: 30)
                        let nbos = ["EBU", "WBU", "NIBU", "CBAI", "UNK"]
                        Menu("Add other NBO") {
                            ForEach(nbos, id: \.self) { text in
                                Button(text) {
                                    nationalId = text + "-" + nationalId
                                    focused = nil
                                }
                            }
                        }
                        .focusable(false)
                        Spacer()
                    }
                    
                    Spacer().frame(height: 20)
                    
                    HStack {
                        Spacer().frame(width: 20)
                        Input(title: "Name:", field: $names, width: 140, inlineTitle: true, inlineTitleWidth: 125)
                            .focused($focused, equals: .names)
                        Spacer()
                    }
                    
                    Spacer().frame(height: 20)
                    
                    HStack {
                        Spacer().frame(width: 20)
                        HStack {
                            Text("Treat as:")
                            Spacer()
                        }
                        .frame(width: 115)
                        Picker("", selection: $condition) {
                            Text("Missing").tag(PlayerCondition.missing)
                            Text("Lapsed").tag(PlayerCondition.lapsed)
                        }
                        .pickerStyle(.segmented)
                        .disabled(!notFound)
                        Spacer()
                    }
                }
                Spacer()
                Separator(thickness: 1)
                Spacer().frame(height: 10)
                HStack {
                    CustomButton.button(title: "Cancel") {
                        dismiss()
                    }
                    
                    Spacer().frame(width: 60)
                    
                    CustomButton.button(title: "Update") {
                        participant.memberNationalId = nationalId
                        participant.memberNames = names
                        participant.condition = condition
                        participant.updateFromMember()
                        dismiss()
                    }
                    .disabled((participant.nationalId == nationalId && participant.names == names && participant.condition == condition) || condition == .undefined || nationalId.isEmpty || names.isEmpty)
                }
                Spacer().frame(height: 10)
            }
        }
        .onAppear {
            matches = participant.possibleMatches
            if matches.count == 0 {
                widenSearch()
            }
            nationalId = participant.nationalId
            names = participant.names
            notFound = (MemberViewModel.member(nationalId: nationalId) == nil)
            condition = (notFound && participant.condition == .valid) ? .undefined : participant.condition
        }
        .frame(width: 790, height: chooseOnly ? 290 : 440)
    }
    
    func widenSearch() {
        let originalMatches = matches.count
        repeat {
            distance += 1
            matches = (MasterData.shared.members.array as! [MemberViewModel]).filter({Utility.levenshteinDistance(participant.names, $0.names) <= distance && BlockedViewModel.blocked(nationalId: $0.nationalId) == nil }).sorted(by: { Utility.levenshteinDistance(participant.names, $0.names) < Utility.levenshteinDistance(participant.names, $1.names)})
        } while matches.count == originalMatches && distance <= maxDistance
    }
    
    func heading() -> some View{
        ZStack {
            UnevenRoundedRectangle(cornerRadii: .init(topLeading: 6, topTrailing: 6), style: .continuous)
                .foregroundColor(Palette.tile.background)
                .frame(height: 30)
            VStack(spacing: 0) {
                Spacer()
                LazyVGrid(columns: tableColumns, spacing: 0) {
                    GridRow {
                        Text("National Id")
                        Text("Names")
                        Text("Club")
                        Text("Rank")
                    }
                    .padding(.horizontal, 5)
                }
                .bold()
                Spacer()
            }
            .frame(height: 30)
        }
    }
    
}
