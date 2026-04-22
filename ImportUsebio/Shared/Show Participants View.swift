//
//  Show Participants View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 27/02/2026.
//

import SwiftUI

enum Nbo : String, CaseIterable, Identifiable {
    
    var id: String {
        rawValue
    }
    
    case sbu = ""
    case ebu = "EBU"
    case wbu = "WBU"
    case nibu = "NIBU"
    case cbai = "CBAI"
    case other = "UNK"
    
    init?(nationalId: String) {
        let components = nationalId.uppercased().components(separatedBy: "-")
        if let prefix = components.first {
            self.init(rawValue: prefix)
        } else if Int(nationalId) != nil {
            self = .sbu
        } else {
            self = .other
        }
    }
    
    var string: String {
        switch self {
        case .other:
            "Other"
        default:
            "\(self)".uppercased()
        }
    }
}

indirect enum ParticipantStatus: Equatable {
    case ok
    case updated(original: ParticipantStatus)
    case memberNotFound(suggested: Int)
    case veryDifferent(suggested: Int)
    case memberFoundLocally
    case lapsed
    case slightlyDifferent
    case triviallyDifferent
    
    var string: String {
        switch self {
        case .ok: return "OK"
        case .updated(let original): return "\(original.string) - Updated"
        case .memberNotFound: return "Not found"
        case .memberFoundLocally: return "Found locally"
        case .lapsed: return "Lapsed"
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
        case .ok: 99
        case .updated: 99
        case .memberNotFound: 1
        case .lapsed: 2
        case .veryDifferent: 3
        case .memberFoundLocally: 4
        case .slightlyDifferent: 5
        case .triviallyDifferent: 6
        }
    }
    
    var isUpdated: Bool {
        switch self {
        case .updated:
            return true
        default:
            return false
        }
    }
}

struct ParticipantIdentity: Equatable {
    var nbo: Nbo = .sbu
    var nationalId = ""
    
    init() {
    }
    
    init?(combined: String?) {
        self.init()
        if let combined = combined {
            if !set(combined: combined) {
                return nil
            }
        } else {
            return nil
        }
    }
    
    @discardableResult mutating func set(combined: String) -> Bool {
        nationalId = combined.uppercased()
        let components = nationalId.components(separatedBy: "-")
        if components.count == 2 {
            nbo = Nbo(rawValue: components.first!) ?? .other
            nationalId = components.last!
        } else if components.count == 1 && (components.first!.trim().isEmpty || Int(components.first!) != nil) {
            nbo = .sbu
            nationalId = components.first!
        } else {
            return false
        }
        return true
    }
    
    var isEmpty: Bool {
        combined.trim().isEmpty
    }
    
    var isInvalid: Bool {
        switch nbo {
        case .sbu:
            MemberViewModel.member(nationalId: combined) == nil
        default:
            false
        }
    }
    
    var combined: String {
        (nbo == .sbu ? nationalId : "\(nbo.rawValue)-\(nationalId)")
    }
}

@Observable class ParticipantData: Identifiable {
    var id: UUID = UUID()
    var identity = ParticipantIdentity()
    var names = ""
    var memberIdentity = ParticipantIdentity()
    var memberNames: String?
    var memberStatus: PlayerStatus = .missing
    var originalIdentity = ParticipantIdentity()
    var originalNames = ""
    var originalParticipantStatus = ParticipantStatus.ok
    var possibleMatches: [MemberViewModel] = []
    var suggested = false
    var player: Player
    var linked: [ParticipantData] = []
    var status: PlayerStatus = .missing
    var updated: Bool = false
    
    var participantStatus: ParticipantStatus {
        if updated {
            return .updated(original: originalParticipantStatus)
        } else if memberIdentity.isEmpty || suggested {
            return .memberNotFound(suggested: possibleMatches.count)
        } else if memberStatus == .missing {
            return .memberFoundLocally
        } else if memberStatus != .active {
            return .lapsed
        } else if names != memberNames {
            let difference = Utility.levenshteinDistance(names, memberNames ?? "")
            if difference <= 2 {
                return .triviallyDifferent
            } else if difference <= 5 {
                return .slightlyDifferent
            } else {
                return .veryDifferent(suggested: possibleMatches.count)
            }
        } else {
            return .ok
        }
    }
    
    var otherNames: String { getName(names, last: false)}
    var lastName: String { getName(names, last: true)}
    
    init(imported: Player) {
        self.player = imported
        self.identity = ParticipantIdentity(combined: imported.nationalId) ?? ParticipantIdentity()
        self.names = imported.name ?? ""
        lookupMember()
        self.originalIdentity = self.identity
        self.originalNames = imported.name ?? ""
        self.originalParticipantStatus = self.participantStatus
        switch participantStatus {
        case .memberNotFound, .veryDifferent:
            self.possibleMatches = MemberViewModel.member(names: self.names).filter( { BlockedViewModel.blocked(nationalId: $0.nationalId) == nil } )
        case .triviallyDifferent:
            // Update automatically
            self.names = memberNames ?? self.names
            self.updated = true
        default:
            break
        }
    }
    
    func copy() -> ParticipantData {
        ParticipantData(imported: player)
    }
    
    func revertToOriginal() {
        self.identity = originalIdentity
        self.names = originalNames
        lookupMember()
        self.updated = false
    }
    
    static func sort(_ first: ParticipantData, _ second: ParticipantData) -> Bool {
        // Used to sort by status priority, number of matches and national ID
        var result: Bool
        
        result = sortCheck(first.participantStatus.priority.compare(second.participantStatus.priority)) {
            // First sort key is priority of status
            var comparison: ComparisonResult
            switch first.participantStatus {
            case .veryDifferent, .memberNotFound:
                // Second sort key is number of possible matches (for statuses where that is relevant)
                comparison = matchesSort(first.possibleMatches.count).compare(matchesSort(second.possibleMatches.count))
            default:
                comparison = .orderedSame
            }
            return sortCheck(comparison) {
                // Third sort key is the national Id if it is integer (sorted as number)
                if let firstNumber = Int(first.identity.combined), let secondNumber = Int(second.identity.combined) {
                    comparison = firstNumber.compare(secondNumber)
                } else {
                    comparison = .orderedSame
                }
                return sortCheck(comparison) {
                    // Last check is just the alpha combined identity
                    return sortCheck(first.identity.combined.lowercased().compare(second.identity.combined.lowercased())) {
                        return true
                    }
                }
            }
        }
        return result
    }
    
    static func matchesSort(_ matches: Int) -> Int {
        switch matches {
        case 0:
            0
        case 1:
            2
        default:
            1
        }
    }
                           
    static func sortCheck(_ comparison: ComparisonResult, action: ()->Bool) -> Bool {
        switch comparison {
        case .orderedAscending:
            true
        case .orderedDescending:
            false
        case .orderedSame:
            action()
        }
    }
    
    func lookupMember() {
        if let member = MemberViewModel.member(nationalId: identity.combined) {
            memberIdentity = ParticipantIdentity(combined: member.nationalId) ?? ParticipantIdentity()
            memberNames = member.names
            memberStatus = member.status
        } else {
            if identity.combined != "" {
                if let localMember = LocalMemberViewModel.member(nationalId: identity.combined), names == localMember.names {
                    memberIdentity = ParticipantIdentity(combined: localMember.nationalId)  ?? ParticipantIdentity()
                    memberNames = localMember.names
                    memberStatus = localMember.status
                }
            } else {
                if let localMember = LocalMemberViewModel.member(names: names) {
                    memberIdentity = ParticipantIdentity(combined: localMember.nationalId)  ?? ParticipantIdentity()
                    memberNames = localMember.names
                    memberStatus = localMember.status
                }
            }
        }
    }
    
    func updateFromMember() {
        if !memberIdentity.isEmpty {
            identity = memberIdentity
        } else if let firstMatch = ParticipantIdentity(combined: possibleMatches.first?.nationalId) {
            identity = firstMatch
        }
        names = memberNames ?? names
        status = memberStatus
        memberIdentity = identity
        memberNames = names
        possibleMatches = []
        updated = true
        suggested = false
        // Write back to imported data
        for item in [self] + self.linked {
            item.player.nationalId = identity.combined
            item.player.name = names
            item.player.status = status
        }
    }
    
    private func getName(_ names: String, last: Bool) -> String {
        var otherNames = ""
        var lastName = ""
        var names = names.components(separatedBy: " ")
        if names.last != nil {
            lastName = names.last!
            names.removeLast()
            if !names.isEmpty {
                otherNames = names.joined(separator: " ")
            }
        }
        return (last ? lastName : otherNames)
    }
}

struct ParticipantsView: View {
    @State var participants: [ParticipantData] = []
    @State var exit: Bool = true
    @State var showMatches = false
    @State var selected = ParticipantData(imported: Player())
    @State var chooseOnly = false
    @State var showAll = false
    
    let tableColumns = [GridItem(.fixed(100),  spacing: 10, alignment: .trailing),
                        GridItem(.fixed(140), spacing: 10, alignment: .leading),
                        GridItem(.fixed(20),  spacing: 10, alignment: .center),
                        GridItem(.fixed(100),  spacing: 10, alignment: .center),
                        GridItem(.fixed(140), spacing: 10, alignment: .leading),
                        GridItem(.fixed(200), spacing: 10, alignment: .leading),
                        GridItem(.fixed(80),  spacing: 10, alignment: .leading),
                        GridItem(.fixed(20),  spacing: 10, alignment: .leading)]
    
    var body: some View {
        StandardView("Select Input") {
            VStack(spacing: 0) {
                ZStack {
                    Banner(title: Binding.constant("Check Players Details"), bottomSpace: false, backEnabled: { exit }, backAction: {
                        return true
                    })
                    HStack {
                        Spacer()
                        Picker("Include:", selection: $showAll) {
                            Text("Incorrect").tag(false)
                            Text("All").tag(true)
                        }
                        Spacer().frame(width: 20)
                    }
                }
                ZStack {
                    VStack(spacing: 0) {
                        Spacer().frame(height: 30)
                        Rectangle()
                            .foregroundColor(Palette.background.background)
                    }
                    VStack(spacing: 0) {
                        HStack {
                            ScrollView(showsIndicators: true) {
                                VStack(spacing: 0) {
                                    LazyVGrid(columns: tableColumns, spacing: 0, pinnedViews: [.sectionHeaders]) {
                                        Section(header: bannerRow()) {
                                            ForEach(self.participants.filter({ if case .ok = $0.participantStatus { return showAll }; return true })) { participant in
                                                gridRow(participant: participant)
                                                    .frame(height: 30)
                                                    .onTapGesture {
                                                        selected = participant
                                                        self.chooseOnly = false
                                                        showMatches = true
                                                    }
                                            }
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
        .frame(width: 920, height: 550)
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
            if let participant = participant {
                editButton(participant, { (participant, chooseOnly) in
                    selected = participant
                    self.chooseOnly = chooseOnly
                    showMatches = true
                })
            } else {
                Text("")
            }
        }
    }
    
    func gridRow(participant: ParticipantData, editAction: ((ParticipantData, Bool)->())? = nil) -> some View {
        if participant.memberIdentity.isEmpty, let suggest = participant.possibleMatches.first {
            participant.memberIdentity = ParticipantIdentity(combined: suggest.nationalId) ?? ParticipantIdentity()
            participant.memberNames = suggest.otherNames + " " + suggest.lastName
            participant.suggested = true
        }
        return gridRow(participant.identity.combined, participant.names, participant.memberIdentity.combined, participant.memberNames ?? "", participant.participantStatus.string, participant.participantStatus.matches, participant: participant, databaseColor: (participant.suggested ? .faint : .normal))
    }
    
    func actionButtons(_ participant: ParticipantData, _ editAction: ((ParticipantData, Bool)->())? = nil) -> some View {
        HStack(spacing: 0) {
            Spacer()
            switch participant.participantStatus {
            case .updated:
                Button(action: {
                    participant.revertToOriginal()
                }) {
                    Text("􀄽").frame(width: 25, height: 25).palette(.enabledButton).cornerRadius(12.5)
                }
                .buttonStyle(PlainButtonStyle())
                .focusable(false)
            default:
                if (!participant.memberIdentity.isEmpty || !participant.possibleMatches.isEmpty) && (participant.participantStatus != .ok && !participant.participantStatus.isUpdated) {
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
            }
            Spacer()
        }
    }
    
    func editButton(_ participant: ParticipantData, _ editAction: ((ParticipantData, Bool)->())? = nil) -> some View {
        HStack(spacing: 0) {
            Button(action: {
                editAction?(participant, false)
            }) {
                Text("􀈊").frame(width: 25, height: 25).palette(.highlightButton).cornerRadius(12.5)
            }
            .buttonStyle(PlainButtonStyle())
            .focusable(false)
            Spacer()
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
    @State var identity: ParticipantIdentity = ParticipantIdentity()
    @State var names: String = ""
    @State var status: PlayerStatus = .missing
    @State var distance = 0
    @State var namesChanged = false
    @FocusState private var focused: ViewType?
    var statusDesc: Binding<String> { Binding { status.string } set: { _ in } }
    let maxDistance = 5
    @State var matches: [MemberViewModel] = []
    @State var notFound: Bool = false
    var originalIdentity: Binding<String> { Binding { participant.originalIdentity.combined } set: { _ in } }
    
    let tableColumns = [GridItem(.fixed(100),  spacing: 0, alignment: .trailing),
                        GridItem(.fixed(140), spacing: 0, alignment: .leading),
                        GridItem(.fixed(70), spacing: 0, alignment: .leading),
                        GridItem(.fixed(200), spacing: 0, alignment: .leading),
                        GridItem(.flexible(minimum: 130), spacing: 0, alignment: .leading)]
    
    var body: some View {
        StandardView("Select Input") {
            VStack(spacing: 0) {
                Banner(title: Binding.constant(chooseOnly ? "Choose Player to Update From" : "Update Player Details"), bottomSpace: false, back: false)
                
                Spacer().frame(height: 10)
                HStack {
                    Spacer().frame(width: 20)
                    Input(title: "Imported Id:", field: originalIdentity, placeHolder: "No Id imported", width: 140, inlineTitle: true, inlineTitleWidth: 125, isReadOnly: true)
                    Spacer().frame(width: 30)
                    Input(title: "Imported Name:", field: $participant.originalNames, width: 200, inlineTitle: true, inlineTitleWidth: 120, isReadOnly: true)
                    Spacer()
                }
                Spacer().frame(height: 20)
                
                HStack {
                    Spacer().frame(width: 20)
                    VStack(alignment: .leading) {
                        Spacer().frame(height: 10)
                        Text("Possible Matches: ")
                        Spacer().frame(height: 10)
                        HStack{
                            Spacer()
                            CustomButton.button(title: "Widen", width: 60, height: 20, color: Palette.enabledButton, enabled: { distance != maxDistance } ) {
                                widenSearch()
                            }
                            Spacer()
                        }
                        Spacer()
                    }
                    .frame(width: 120, height: 120)
                    Spacer().frame(width: 10)
                    VStack {
                        ZStack {
                            VStack(spacing: 0) {
                                UnevenRoundedRectangle(cornerRadii: .init(topLeading: 6, topTrailing: 6), style: .continuous)
                                    .foregroundColor(Palette.tile.background)
                                    .frame(height: 24)
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
                                                    LeadingClickableText(member.status.string)
                                                    LeadingClickableText(member.homeClub)
                                                    LeadingClickableText(rank?.rankName ?? "Unknown")
                                                }
                                                .frame(height: 24)
                                                .padding(.horizontal, 5)
                                                .palette(identity.combined == member.nationalId && names == member.names ? .highlightTile : .alternate)
                                                .onTapGesture {
                                                    identity = ParticipantIdentity(combined: member.nationalId) ?? ParticipantIdentity()
                                                    names = member.names
                                                    status = member.status
                                                    focused = nil
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        .frame(width: 610, height: 120)
                        .cornerRadius(8)
                    }
                    Spacer()
                    Spacer()
                }
                Spacer().frame(height: 30)
                
                if !chooseOnly {
                    HStack {
                        Spacer().frame(width: 20)
                        HStack {
                            Text("NBO:")
                            Spacer()
                        }
                        .frame(width: 113)
                        Picker("", selection: $identity.nbo) {
                            ForEach(Nbo.allCases) { nbo in
                                Text(nbo.string).tag(nbo)
                            }
                        }
                        .onChange(of: identity.nbo) {
                            if identity.nbo == .sbu, let lookup = MemberViewModel.member(nationalId: identity.combined) {
                                if names != lookup.names {
                                    names = lookup.names
                                    namesChanged = true
                                }
                                status = lookup.status
                            } else {
                                if names != participant.originalNames {
                                    names = participant.originalNames
                                    namesChanged = true
                                }
                                status = .missing
                            }
                        }
                        .pickerStyle(.segmented)
                        .focusable(false)
                        Spacer()
                    }
                    
                    Spacer().frame(height: 20)
                    
                    HStack {
                        Spacer().frame(width: 20)
                        Input(title: "National Id:", field: $identity.nationalId, width: 140, inlineTitle: true, inlineTitleWidth: 125, onChange: { newValue in
                            if let lookup = MemberViewModel.member(nationalId: identity.combined) {
                                if names != lookup.names {
                                    names = lookup.names
                                    namesChanged =  true
                                }
                                status = lookup.status
                                identity.nbo = .sbu
                                notFound = false
                            } else {
                                if names != participant.originalNames {
                                    names = participant.originalNames
                                    namesChanged = true
                                }
                                notFound = true
                                status = .missing
                            }
                        })
                        .focused($focused, equals: .nationalId)
                        Spacer()
                    }
                    
                    Spacer().frame(height: 20)
                    
                    HStack {
                        Spacer().frame(width: 20)
                        Input(title: "Name:", field: $names, width: 200, inlineTitle: true, inlineTitleWidth: 125, onChange: { newValue in
                            namesChanged = true
                        })
                        .focused($focused, equals: .names)
                        
                        Spacer().frame(width: 20)
                        CustomButton.button(image: "arrow.clockwise", resizeImage: true, title: "Re-Search", width: 110, height: 20, color: Palette.enabledButton) {
                            namesChanged = false
                            widenSearch(newSearch: true)
                        }
                        .disabled(names.trim() == "" || !namesChanged)
                        
                        Spacer()
                    }
                    
                    Spacer().frame(height: 20)
                    
                    HStack {
                        Spacer().frame(width: 20)
                        Input(title: "Member status: ", field: statusDesc, width: 120, inlineTitle: true, inlineTitleWidth: 125, isReadOnly: true)
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
                        // Write back to local members file
                        participant.memberIdentity = identity
                        participant.memberNames = names
                        participant.memberStatus = status
                        participant.updateFromMember()
                        if MemberViewModel.member(nationalId: identity.combined) == nil {
                            saveLocal(participant: participant)
                        }
                        dismiss()
                    }
                    .disabled((participant.identity == identity && participant.names == names && participant.status == status) || identity.isInvalid || names.isEmpty )
                }
                Spacer().frame(height: 10)
            }
        }
        .onAppear {
            matches = participant.possibleMatches
            names = participant.names
            identity = participant.identity
            if let member = MemberViewModel.member(nationalId: identity.combined) {
                status = member.status
                notFound = false
            } else {
                status = .missing
                notFound = true
            }
            if matches.count == 0 {
                widenSearch(newSearch: true)
            }
            Utility.executeAfter(delay: 0.1) {
                namesChanged = false
            }
        }
        .frame(width: 790, height: chooseOnly ? 340 : 490)
    }
    
    func widenSearch(newSearch: Bool = false) {
        if newSearch {
            distance = -1
            matches.removeAll()
        }
        let originalMatches = matches.count
        repeat {
            distance += 1
            let members = MasterData.shared.members.array as! [MemberViewModel] + (MasterData.shared.localMembers.array as! [LocalMemberViewModel]).map{$0.memberViewModel}
            matches = members.filter({Utility.levenshteinDistance(names, $0.names) <= distance && BlockedViewModel.blocked(nationalId: $0.nationalId) == nil }).sorted(by: { Utility.levenshteinDistance(names, $0.names) < Utility.levenshteinDistance(names, $1.names)})
        } while matches.count == originalMatches && distance <= maxDistance
    }
    
    func saveLocal(participant: ParticipantData) {
        if MemberViewModel.member(nationalId: identity.combined) == nil {
            if let localMember = LocalMemberViewModel.member(nationalId: identity.combined) {
                // Already exists - overwrite
                localMember.names = participant.names
                localMember.status = participant.status
                localMember.save()
            } else {
                // New local member - create
                let localMember = LocalMemberViewModel(nationalId: participant.identity.combined, otherNames: participant.otherNames, lastName: participant.lastName, status: participant.status)
                localMember.insert()
            }
        }
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
                        Text("Status")
                        Text("Club")
                        Text("Rank")
                    }
                    .padding(.horizontal, 5)
                }
                .bold()
                Spacer()
            }
            .frame(height: 24)
        }
    }
    
}
