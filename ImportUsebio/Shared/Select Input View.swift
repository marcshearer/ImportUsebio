//
//  Select Input.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

enum Level: Int {
    case local = 0
    case national = 1
}

enum Basis: Int {
    case standard = 0
    case manual = 1
    case headToHead = 2
}

enum ViewField {
    case eventDescription
    case eventCode
    case clubCode
    case minRankCode
    case maxRankCode
}

struct AutoCompleteData : Hashable {
    var index: Int
    var code: String
    var desc: String
}

struct SelectInputView: View {
    @State private var inputFilename: String = ""
    @State private var roundName: String = ""
    @State private var event: EventViewModel? = nil
    @State private var eventCode: String = ""
    @State private var eventMessage: String = ""
    @State private var club: ClubViewModel? = nil
    @State private var clubCode: String = ""
    @State private var chooseBest: Int = 0
    @State private var clubMessage: String = ""
    @State private var minRank: RankViewModel? = nil
    @State private var minRankCode: Int = 0
    @State private var minRankMessage: String = "No minimum rank"
    @State private var maxRank: RankViewModel? = nil
    @State private var maxRankCode: Int = 999
    @State private var maxRankMessage: String = "No maximum rank"
    @State private var basis: Basis = .standard
    @State private var eventDescription: String = ""
    @State private var includeInRace: Bool = false
    @State private var filterSessionId: String = ""
    @State private var overrideTeamMembers: Int = 0
    @State private var manualPointsColumn: String?
    @State private var localNational = Level.local
    @State private var drawsRounded = true
    @State private var winDrawLevel: WinDrawLevel = .board
    @State private var mergeMatches = false
    @State private var vpType: VpType = .discrete
    @State private var maxAward: Float = 10.0
    @State private var ewMaxAward: Float = 0.0
    @State private var minEntry: Int = 0
    @State private var awardTo: Float = 25
    @State private var perWin: Float = 0.25
    @State private var securityBookmark: Data? = nil
    @State private var refresh = true
    @State private var content: [String] = []
    @State private var writer: Writer? = nil
    @State private var scoreData: ScoreData? = nil
    @State private var roundErrors: [RoundErrorList] = []
    @State private var showErrors = false
    @State private var showSettings = false
    @State private var showSettingsMenu = false
    @State private var showImportEvents = false
    @State private var showImportClubs = false
    @State private var showImportRanks = false
    @State private var showEventCodesSearch = false
    @State private var showClubCodesSearch = false
    @State private var showMinRankCodesSearch = false
    @State private var showMaxRankCodesSearch = false
    @State private var missingNationalIds = false
    @State private var showAdvancedParameters = false
    @State private var editSettings = Settings.current.copy()
    @State private var windowHeight: CGFloat = 580
    @State private var expandedWindowHeight: CGFloat = 680
    @FocusState private var focusedField: ViewField?
    @State private var autoCompleteField: ViewField?
    @State private var eventCodeData: [AutoCompleteData] = []
    @State private var clubCodeData: [AutoCompleteData] = []
    @State private var minRankCodeData: [AutoCompleteData] = []
    @State private var maxRankCodeData: [AutoCompleteData] = []
    @State private var selected: Int? = nil
    let detectKeys: Set<KeyEquivalent> = [.upArrow, .downArrow, .return]

    @Namespace private var autoComplete
    
    var body: some View {
        
            // Just to trigger view refresh
        if refresh { EmptyView() }
        
        StandardView("Select Input") {
            ZStack {
                VStack {
                    HStack {
                        Spacer().frame(width: 30)
                        VStack(spacing: 0) {
                            
                            filenameView
                            
                            separatorView
                            
                            eventDescriptionView
                            
                            InputTitle(title: "Coding", topSpace: 16)
                            
                            Spacer().frame(height: 8)
                            
                            HStack {
                                
                                eventCodeView
                                
                                clubCodeView
                                
                                Spacer()
                            }
                            
                            InputTitle(title: "Ranking restrictions", topSpace: 16)
                            Spacer().frame(height: 8)
                            
                            HStack {
                                
                                rankingsView
                                
                            }
                            
                            VStack(spacing: 0) {
                                separatorView
                                InputTitle(title: "Other details")
                                Spacer().frame(height: 8)
                                
                                roundNameView
                                
                                HStack {
                                    VStack {
                                        InputTitle(title: "Points Award Details", topSpace: 10)
                                        pointsAwardView
                                    }
                                    Spacer()
                                }
                            }
                            VStack {
                                Spacer().frame(height: 16)
                                Separator(thickness: 1)
                                Spacer().frame(height: 16)
                                HStack {
                                    VStack {
                                        Spacer()
                                        InputTitle(title: "Advanced Parameters ", fillTrailing: false)
                                        Spacer()
                                    }
                                    VStack {
                                        Spacer()
                                        Spacer().frame(height: 4)
                                        Text(showAdvancedParameters ? "􀄥" : "􀄧")
                                            .foregroundColor(Palette.background.themeText)
                                            .font(inputFont)
                                            .onTapGesture {
                                                showAdvancedParameters.toggle()
                                                windowHeight = showAdvancedParameters ? expandedWindowHeight : windowHeight
                                            }
                                        Spacer()
                                    }
                                    Spacer()
                                }.frame(height: 20)
                                if showAdvancedParameters {
                                    advancedParametersView
                                }
                            }
                            VStack(spacing: 0) {
                                Spacer().frame(height: 24)
                                Separator(thickness: 1)
                                Spacer().frame(height: 12)
                                HStack {
                                    Spacer().frame(width: 40)
                                    addSheetButton()
                                    Spacer().frame(width: 50)
                                    finishButton()
                                    Spacer().frame(width: 50)
                                    clearButton()
                                    Spacer().frame(width: 50)
                                    pasteButton()
                                    Spacer()
                                    settingsButton()
                                        .popover(isPresented: $showSettingsMenu) {
                                            settingsMenuView
                                        }
                                    Spacer().frame(width: 50)
                                }
                                Spacer().frame(height: 20)
                            }
                        }
                    }
                }
                VStack {
                    switch focusedField {
                    case .eventCode:
                        autoCompleteView(field: .eventCode, codeWidth: 80, data: $eventCodeData, valid: event != nil) { (newValue) in
                            eventCode = newValue
                        }
                    case .clubCode:
                        autoCompleteView(field: .clubCode, codeWidth: 80, data: $clubCodeData, valid: club != nil) { (newValue) in
                            clubCode = newValue
                        }
                    case .minRankCode:
                        autoCompleteView(field: .minRankCode, codeWidth: 80, data: $minRankCodeData, valid: minRank != nil) { (newValue) in
                            minRankCode = Int(newValue)!
                        }
                    case .maxRankCode:
                        autoCompleteView(field: .maxRankCode, codeWidth: 80 , data: $maxRankCodeData, valid: maxRank != nil) { (newValue) in
                            maxRankCode = Int(newValue)!
                        }
                    default:
                        EmptyView()
                    }
                }
            }
            Spacer()
        }
        .frame(width: 900, height: windowHeight)
        .sheet(isPresented: $showErrors) {
            ShowErrorsView(roundErrors: roundErrors)
        }
        .sheet(isPresented: $showEventCodesSearch) {
            EventCodesSearchView() { event in
                set(eventCode: event.eventCode)
                focusedField = .eventCode
            }
        }
        .sheet(isPresented: $showClubCodesSearch) {
            ClubCodesSearchView() { club in
                set(clubCode: club.clubCode)
                focusedField = .clubCode
            }
        }
        .sheet(isPresented: $showMinRankCodesSearch) {
            RankCodesSearchView(showNoMinimum: true) { rank in
                set(minRankCode: rank.rankCode)
                focusedField = .minRankCode
            }
        }
        .sheet(isPresented: $showMaxRankCodesSearch) {
            RankCodesSearchView(showNoMaximum: true) { rank in
                set(maxRankCode: rank.rankCode)
                focusedField = .maxRankCode
            }
        }
        .sheet(isPresented: $showSettings) {
            SettingsView(settings: editSettings)
        }
        .sheet(isPresented: $showImportEvents) {
            EventImportView()
        }
        .sheet(isPresented: $showImportClubs) {
            ClubImportView()
        }
        .sheet(isPresented: $showImportRanks) {
            RankImportView()
        }
    }
    
    private func autoCompleteView(field: ViewField, codeWidth: CGFloat, data: Binding<[AutoCompleteData]>, valid: Bool, selectAction: @escaping (String)->()) -> some View {
        VStack(spacing: 0) {
            if data.wrappedValue.count > (valid ? 1 : 0) {
                ScrollView {
                    LazyVStack(spacing: 0) {
                        ForEach(data.wrappedValue, id: \.index) { (element) in
                            VStack(spacing: 0) {
                                HStack(spacing: 0) {
                                    Spacer().frame(width: 12)
                                    HStack(spacing: 0) {
                                        Text(element.code)
                                            .font(inputFont)
                                            .lineLimit(1)
                                            .truncationMode(.tail)
                                            .padding(0)
                                        Spacer()
                                    }
                                    .frame(width: codeWidth - 12)
                                    Text(element.desc)
                                        .font(lookupFont)
                                        .lineLimit(1)
                                        .truncationMode(.tail)
                                        .padding(0)
                                    Spacer()
                                }
                                .contentShape(Rectangle())
                                .onTapGesture {
                                    selectAction(element.code)
                                }
                            }
                            .frame(height: 20)
                            .background(element.index != selected ? Palette.autoComplete.background : Palette.autoCompleteSelected.background)
                            .foregroundColor(element.index != selected ? Palette.autoComplete.text : Palette.autoCompleteSelected.text)
                        }
                    }
                    .scrollTargetLayout()
                    .listStyle(DefaultListStyle())
                }
                .scrollPosition(id: $selected)
            }
        }
        .zIndex(1)
        .clipShape(
            .rect(
                topLeadingRadius: 0,
                bottomLeadingRadius: 8,
                bottomTrailingRadius: 8,
                topTrailingRadius: 0
            )
        )
        .matchedGeometryEffect(
            id: field,
            in: autoComplete,
            properties: .position,
            anchor: .topTrailing,
            isSource: false)
        .frame(width: 270, height: CGFloat(min(6, data.wrappedValue.count) * 20))
    }
    
    private func onKeyPress(_ keyPress: KeyPress, maxSelected: Int, onSelect: ()->()) -> KeyPress.Result {
        switch keyPress.key {
        case .downArrow:
            selected = min(maxSelected - 1, (selected ?? -1) + 1)
            return .handled
        case .upArrow:
            selected = ((selected ?? 0) == 0 ? nil : selected! - 1)
            return .handled
        case .return:
            if selected != nil {
                onSelect()
            }
            return .handled
        default:
            return .ignored
        }
    }
    
    
    
    private func set(eventCode newValue: String) {
        event = EventViewModel.event(eventCode: newValue)
        eventCode = newValue
        eventMessage = event?.eventName ?? "Invalid event code"
        if let event = event {
            if event.localAllowed && !event.nationalAllowed {
                localNational = .local
            } else if !event.localAllowed && event.nationalAllowed {
                localNational = .national
            }
            if event.validMinRank > 0 {
                set(minRankCode: event.validMinRank)
            }
            if event.validMaxRank < 999 {
                set(maxRankCode: event.validMaxRank)
            }
            if event.originatingClubCode != "" {
                set(clubCode: event.originatingClubCode)
            }
        }
    }
    
    private func set(clubCode newValue: String) {
        club = ClubViewModel.club(clubCode: newValue)
        clubCode = newValue
        clubMessage = club?.clubName ?? (clubCode == "" ? "No club code specified" : "Invalid club code")
    }
    
    private func set(minRankCode newValue: Int) {
        minRank = RankViewModel.rank(rankCode: newValue)
        minRankCode = newValue
        minRankMessage = minRank?.rankName ?? (minRankCode == 0 ? "No minimum rank" : "Invalid rank code")
        if minRank != nil && maxRankCode < minRankCode {
            set(maxRankCode: minRankCode)
        }
    }
    
    private func set(maxRankCode newValue: Int) {
        maxRank = RankViewModel.rank(rankCode: newValue)
        maxRankCode = newValue
        if maxRankCode < minRankCode {
            maxRankMessage = "Below minimum rank"
        } else {
            maxRankMessage = maxRank?.rankName ?? (maxRankCode == 999 ? "No maximum rank" : "Invalid rank code")
        }
    }
    
    private var separatorView: some View {
        VStack {
            Spacer().frame(height: 16)
            Separator(thickness: 1)
            Spacer().frame(height: 16)
        }
    }
    
    private var filenameView: some View {
        VStack(spacing: 0) {
            HStack {
                Input(title: "Import filename:", field: $inputFilename, placeHolder: "No import file specified", topSpace: 20, leadingSpace: 30, width: 707, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: true, isReadOnly: true, pickerAction: chooseFile)
                Spacer()
            }
        }
    }
    
    private var eventDescriptionView: some View {
        VStack {
            InputTitle(title: "Event Description")
            HStack {
                Input(field: $eventDescription, leadingSpace: 30, width: 400, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                    .focused($focusedField, equals: .eventDescription)
                Spacer().frame(width: 139)
                Picker("Include in race:     ", selection: $includeInRace) {
                    Text("Include").tag(true)
                    Text("Exclude").tag(false)
                }
                .pickerStyle(.segmented)
                .frame(width: 250)
                .disabled((writer?.rounds.count ?? 0) >= 1)
                Spacer()
            }
        }
    }
    
    private var eventCodeView: some View {
        HStack {
            Spacer().frame(width: 42)
            
            Input(title: "Event code:", field: $eventCode, message: $eventMessage, messageOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true, limitText: 6, pickerAction: { showEventCodesSearch = true }, onKeyPress: eventKeyPress, detectKeys: detectKeys) { (newValue) in
                set(eventCode: newValue)
                eventCodeData = getEventList()
            }
            .focused($focusedField, equals: .eventCode)
            .matchedGeometryEffect(id: ViewField.eventCode, in: autoComplete, anchor: .bottomTrailing)
        }
    }
    
    private func eventKeyPress(_ press: KeyPress) -> KeyPress.Result{
        onKeyPress(press, maxSelected: eventCodeData.count) {
            set(eventCode: eventCodeData[selected!].code)
        }
    }
    
    private var clubCodeView: some View {
        HStack {
            
            Spacer().frame(width: 30)
            
            Input(title: "Club code:", field: $clubCode, message: $clubMessage, messageOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, autoCapitalize: .sentences, autoCorrect: false, isEnabled: event == nil || event!.originatingClubCode == "", limitText: 5, pickerAction: { showClubCodesSearch = true }, onKeyPress: clubKeyPress, detectKeys: detectKeys) { (newValue) in
                set(clubCode: newValue)
                clubCodeData = getClubList()
            }
            .focused($focusedField, equals: .clubCode)
            .matchedGeometryEffect(id: ViewField.clubCode, in: autoComplete, anchor: .bottomTrailing)
        }
    }
    
    private func clubKeyPress(_ press: KeyPress) -> KeyPress.Result{
        onKeyPress(press, maxSelected: clubCodeData.count) {
            set(clubCode: clubCodeData[selected!].code)
        }
    }
    
    private var rankingsView: some View {
        HStack {
            
            Spacer().frame(width: 42)
            
            InputInt(title: "Minimum:", field: $minRankCode, message: $minRankMessage, messageOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, maxValue: 999, isEnabled: event == nil || event!.validMinRank == 0, pickerAction: { showMinRankCodesSearch = true }, onKeyPress: minRankKeyPress, detectKeys: detectKeys) { (newValue) in
                set(minRankCode: newValue)
                minRankCodeData = getMinRankList()
            }
            .focused($focusedField, equals: .minRankCode)
            .matchedGeometryEffect(id: ViewField.minRankCode, in: autoComplete, anchor: .bottomTrailing)
            
            Spacer().frame(width: 30)
            
            InputInt(title: "Maximum:", field: $maxRankCode, message: $maxRankMessage, messageOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, maxValue: 999, isEnabled: event == nil || event!.validMaxRank == 999, pickerAction: { showMaxRankCodesSearch = true }, onKeyPress: maxRankKeyPress, detectKeys: detectKeys) { (newValue) in
                set(maxRankCode: newValue)
                maxRankCodeData = getMaxRankList()
            }
            .focused($focusedField, equals: .maxRankCode)
            .matchedGeometryEffect(id: ViewField.maxRankCode, in: autoComplete, anchor: .bottomTrailing)
            
            Spacer()
        }
    }
    
    private func minRankKeyPress(_ press: KeyPress) -> KeyPress.Result{
        onKeyPress(press, maxSelected: minRankCodeData.count) {
            set(minRankCode: Int(minRankCodeData[selected!].code)!)
        }
    }
    
    private func maxRankKeyPress(_ press: KeyPress) -> KeyPress.Result{
        onKeyPress(press, maxSelected: maxRankCodeData.count) {
            set(maxRankCode: Int(maxRankCodeData[selected!].code)!)
        }
    }
    
    private var roundNameView: some View {
        HStack {
            Spacer().frame(width: 42)
            
            Input(title: "Round name:", field: $roundName, topSpace: 0, width: 160, inlineTitle: true, inlineTitleWidth: 117, isEnabled: true)
            
            Spacer()
        }
    }
    
    private var pointsAwardView: some View {
        VStack {
            HStack {
                
                Spacer().frame(width: 42)
                
                Picker("Level:                    ", selection: $localNational) {
                    Text("Local").tag(Level.local)
                    Text("National").tag(Level.national)
                }
                .disabled(event == nil || !event!.nationalAllowed || !event!.localAllowed)
                .pickerStyle(.segmented)
                .focusable(false)
                .frame(width: 300)
                .font(inputFont)
                Spacer()
            }
            
            Spacer().frame(height: 10)
            
            HStack(spacing: 0) {
                
                Spacer().frame(width: 42)
                
                InputFloat(title: "Max award:", field: $maxAward, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 117)
                    .disabled(basis != .standard)
                
                InputFloat(title: "Max E/W award:", field: $ewMaxAward, topSpace: 0, leadingSpace: 30, width: 60, inlineTitle: true, inlineTitleWidth: 120)
                    .disabled(basis != .standard)
                
                Spacer()
            }
            
            Spacer().frame(height: 10)
            
            HStack(spacing: 0) {
                Spacer().frame(width: 42)
                
                InputInt(title: "Min entry:", field: $minEntry, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 117)
                    .disabled(basis != .standard)
                
                InputFloat(title: "Award to %:", field: $awardTo, topSpace: 0, leadingSpace: 30, width: 60, inlineTitle: true, inlineTitleWidth: 120)
                    .disabled(basis != .standard)
                
                InputFloat(title: "Per win:", field: $perWin, topSpace: 0, leadingSpace: 30, width: 60, inlineTitle: true, inlineTitleWidth: 70)
                    .disabled(basis == Basis.manual)
                
                Spacer()
            }
        }
    }
    
    private var advancedParametersView: some View {
        VStack {
            HStack{
                Spacer().frame(width: 42)
                Picker("Calculation basis:   ", selection: $basis) {
                    Text("Standard").tag(Basis.standard)
                    Text("Manual").tag(Basis.manual)
                    Text("Head-to-head").tag(Basis.headToHead)
                }
                .help(Text("The method to be used for calculations\nIf this is Head-to-head this means that it is a round were only win/draw awards are applied. There is no place award.\nIf this is Manual then the MPs are not calculated. They are specified in the raw data."))
                .onChange(of: basis, initial: false) {
                    if basis != .standard {
                        maxAward = 0
                        ewMaxAward = 0
                        minEntry = 0
                        if basis == .headToHead {
                            awardTo = 100
                        } else if basis == .manual {
                            awardTo = 0
                            perWin = 0
                        }
                    }
                }
                .pickerStyle(.menu)
                .focusable(false)
                .frame(width: 250)
                .font(inputFont)
                Spacer()
            }
            Spacer().frame(height: 15)
            HStack {
                Spacer().frame(width: 42)
                HStack {
                    Input(title: "Session ID filter:", field: $filterSessionId, topSpace: 0, width: 120, inlineTitle: true, inlineTitleWidth: 126, isEnabled: true)
                        .help(Text("Used to restrict the import of a Usebio raw data file to a specific Session Id"))
                }
                Spacer().frame(width: 20)
                HStack {
                    Picker("Win/Draws from:   ", selection: $winDrawLevel) {
                        Text("Pairs/Teams").tag(WinDrawLevel.participant)
                        Text("Matches").tag(WinDrawLevel.match)
                        Text("Boards").tag(WinDrawLevel.board)
                    }
                    .onChange(of: winDrawLevel, initial: false) {
                        if winDrawLevel == .participant {
                            mergeMatches = false
                        }
                        if winDrawLevel != .board {
                            vpType = .discrete
                        }
                    }
                    .help(Text("Use Win/Draw data from PARTICIPANT or from MATCH or from MATCH recalculated from BOARD scores - Falls back if data not available at level requested"))
                    .pickerStyle(.menu)
                    .focusable(false)
                    .frame(width: 250)
                    .font(inputFont)
                }
                .frame(height: inputDefaultHeight)
                Spacer().frame(width: 20)
                HStack {
                    Picker("Merge matches:  ", selection: $mergeMatches) {
                        Text("Merge").tag(true)
                        Text("Don't Merge").tag(false)
                    }
                    .help(Text("Combine matches between the same teams/pairs before calculating wins/draws"))
                    .pickerStyle(.menu)
                    .focusable(false)
                    .frame(width: 230)
                    .font(inputFont)
                    .disabled(winDrawLevel == .participant)
                }
                .frame(height: inputDefaultHeight)
                Spacer()
            }
            Spacer().frame(height: 10)
            HStack {
                Spacer().frame(width: 42)
                HStack {
                    InputInt(title: "Override players:", field: $overrideTeamMembers, width: 127, inlineTitle: true, inlineTitleWidth: 120)
                        .help(Text("Used to override the number of team members considered in the raw data. Primarily used to remove subs from the awards when their impact has not been material, or to increase the number of players to allow manual entry in the spreadsheet"))
                }
                Spacer().frame(width: 72)
                HStack {
                    Picker("VP type:   ", selection: $vpType) {
                        Text("Discrete").tag(VpType.discrete)
                        Text("Continuous").tag(VpType.continuous)
                    }
                    .help(Text("When re-calculating from boards what sort of VPs are required?"))
                    .pickerStyle(.menu)
                    .focusable(false)
                    .frame(width: 196)
                    .font(inputFont)
                    .disabled(winDrawLevel == .participant)
                }
                .frame(height: inputDefaultHeight)
                Spacer().frame(width: 58)
                HStack {
                    Picker("VP Draws:  ", selection: $drawsRounded) {
                        Text("Exact").tag(false)
                        Text("Rounded").tag(true)
                    }
                    .help(Text("Round Continuous VPs when calculating draws"))
                    .pickerStyle(.menu)
                    .focusable(false)
                    .frame(width: 192)
                    .font(inputFont)
                    .disabled(winDrawLevel == .participant || vpType == .discrete)
                    Spacer()
                }
                .frame(height: inputDefaultHeight)
                Spacer()
            }
        }
    }
    
    private var settingsMenuView: some View {
        VStack(alignment: .leading, spacing: 0) {
            Text("Settings").bold()
                .padding([.top], 6)
                .padding([.leading, .trailing], 8)
                .padding([.bottom], 0)
            Button("Edit Settings") { showSettings = true }
                .padding([.leading, .trailing], 28)
                .padding([.top], 6)
                .padding([.bottom], 10)
            Separator()
            Text("Imports").bold()
                .padding([.top], 6)
                .padding([.leading, .trailing], 8)
                .padding([.bottom], 0)
            Button("Import Ranks") { showImportRanks = true }
                .padding([.leading, .trailing], 28)
                .padding([.top, .bottom], 6)
            Button("Import Clubs") { showImportClubs = true}
                .padding([.leading, .trailing], 28)
                .padding([.top, .bottom], 6)
            Button("Import Events") { showImportEvents = true}
                .padding([.leading, .trailing], 28)
                .padding([.top], 6)
                .padding([.bottom], 20)
        }.background(Color.clear)
            .buttonStyle(.borderless)
            .font(inputFont)
            .focusable(false)
    }
    
    private func settingsButton() -> some View {
        return Button {
            showSettingsMenu = true
        } label: {
            Image(systemName: "gearshape.fill").font(.largeTitle).foregroundColor(Palette.banner.background)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
    }
    
    private func chooseFile() {
        FileSystem.findFile(title: "Select Source File", prompt: "Select", types: ["xml", "csv"]) { (url, bookmarkData) in
            Utility.mainThread {
                refresh.toggle()
                securityBookmark = bookmarkData
                if let data = try? Data(contentsOf: url) {
                    let type = url.pathExtension.lowercased()
                    if type == "xml" {
                        _ = UsebioParser(fileUrl: url, data: data, filterSessionId: filterSessionId == "" ? nil : filterSessionId, overrideEventType: (basis == .headToHead ? .head_to_head: nil), roundContinuousVPDraw: drawsRounded, winDrawLevel: winDrawLevel, mergeMatches: mergeMatches, vpType: vpType, completion: parserComplete)
                    } else if type == "csv" {
                        if  GenericCsvParser(fileUrl: url, data: data, completion: parserComplete) == nil {
                            MessageBox.shared.show("Invalid data")
                        }
                    } else {
                        MessageBox.shared.show("File type \(type) not supported")
                    }
                } else {
                    MessageBox.shared.show("Unable to read data")
                }
            }
        }
    }
    
    private func addSheetButton() -> some View {
        return Button{
            if let scoreData = scoreData {
                if writer == nil {
                    writer = Writer()
                    writer?.eventDescription = eventDescription
                    writer?.eventCode = eventCode
                    writer?.clubCode = clubCode
                    writer?.chooseBest = chooseBest
                    writer?.minRank = minRankCode
                    writer?.maxRank = maxRankCode
                }
                writer!.includeInRace = (writer!.rounds.count <= 0 && self.includeInRace)
                scoreData.roundName = roundName
                scoreData.national = (localNational == .national)
                scoreData.overrideTeamMembers = (overrideTeamMembers == 0 ? nil : overrideTeamMembers)
                scoreData.roundContinuousVPDraw = drawsRounded
                scoreData.winDrawLevel = winDrawLevel
                scoreData.mergeMatches = mergeMatches
                scoreData.vpType = vpType
                scoreData.maxAward = maxAward
                scoreData.ewMaxAward = (ewMaxAward != 0 ? ewMaxAward : nil)
                scoreData.minEntry = minEntry
                scoreData.reducedTo = 1
                scoreData.awardTo = awardTo
                scoreData.perWin = perWin
                writer?.add(name: roundName, scoreData: scoreData)
                
                MessageBox.shared.show("Added Successfully", okAction: {
                    self.scoreData = nil
                    self.inputFilename = ""
                    self.roundName = ""
                })
                Utility.executeAfter(delay: 2) {
                    self.scoreData = nil
                    self.inputFilename = ""
                    self.roundName = ""
                    self.includeInRace = false
                    MessageBox.shared.hide()
                }
            }
        } label: {
            Text("Add Sheet")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(scoreData == nil || eventCode == "" || eventDescription == "" || roundName == "" || event == nil || (clubCode != "" && club == nil) || (minRankCode != 0 && minRank == nil) || (maxRankCode != 999 && maxRank == nil) || maxRankCode < minRankCode || (club == nil && event!.clubMandatory) || awardTo <= 0 || (writer?.rounds.contains(where: {$0.shortName == roundName}) ?? false))
    }
    
    
    private func finishButton() -> some View {
        return Button{
            FileSystem.saveFile(title: "Generated Workbook Name", prompt: "Save", filename: "\(eventDescription).xlsm") { (url) in
                Utility.mainThread {
                    if let writer = writer {
                        writer.write(as: url.relativePath)
                        MessageBox.shared.show("Written Successfully", okAction: {
                            self.writer = nil
                        })
                        Utility.executeAfter(delay: 2) {
                            self.writer = nil
                            MessageBox.shared.hide()
                        }
                    }
                }
            }
        } label: {
            Text("Create Workbook")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 150, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(writer == nil)
    }
    
    private func pasteButton() -> some View {
        
        return Button {
            if let data = NSPasteboard.general.string(forType: .string) {
                let dataLines = data.replacingOccurrences(of: "\n", with: "").components(separatedBy: "\r")
                let lines = dataLines.map{$0.components(separatedBy: "\t")}
                ImportRounds.process(lines) { (imported, error, warning) in
                    if let error = error {
                        MessageBox.shared.show(error)
                    } else if let imported = imported {
                        updateFromImport(imported: imported)
                    }
                }
            } else {
                MessageBox.shared.show("Invalid clipboard contents")
            }
        } label: {
            Text("Paste Config")
                .foregroundColor(Palette.enabledButton.text)
                .frame(width: 120, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.enabledButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(writer?.rounds.count ?? 0 > 0)
    }
    
    private func clearButton() -> some View {
        return Button{
            self.writer = nil
        } label: {
            Text("Clear")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(writer == nil)
    }
    
    private func parserComplete(scoreData: ScoreData?, messages: [String]) {
        if let scoreData = scoreData {
            let (roundMissingNationalId, errors, warnings) = scoreData.validate()
            missingNationalIds = roundMissingNationalId
            
            let filename = scoreData.fileUrl?.lastPathComponent.removingPercentEncoding ?? ""
            
            var errorList = RoundErrorList(name: filename, errors: [], warnings: [])
            if let errors = errors {
                errorList.errors = errors
            }
            if let warnings = warnings {
                errorList.warnings = warnings
            }
            if !messages.isEmpty {
                errorList.warnings.append(contentsOf: messages)
            }
            
            roundErrors = []
            if errors != nil || warnings != nil {
                roundErrors.append(errorList)
            }
            if errors != nil {
                self.scoreData = nil
            } else {
                self.scoreData = scoreData
                self.inputFilename = filename
            }
            addMissingNationalIdWarning()
            if !roundErrors.isEmpty {
                showErrors = true
            }
        } else {
            MessageBox.shared.show(messages.last ?? "Unable to parse file \(inputFilename)")
            self.scoreData = nil
            self.inputFilename = ""
            self.roundName = ""
        }
    }
    
    @State private var importInProgress: ImportRounds?
    @State private var importRound: Int = 0
    @State private var sourceDirectory: String?
    
    private func updateFromImport(imported: ImportRounds) {
        FileSystem.findDirectory(prompt: "Select source directory") { (url, bookmarkData) in
            Utility.mainThread {
                securityBookmark = bookmarkData
                importInProgress = imported
                sourceDirectory = url.relativePath
                importRound = 0
                roundErrors = []
                missingNationalIds = false
                processNextRound()
            }
        }
    }
    
    private func processNextRound() {
        if importRound <= (importInProgress?.rounds.count ?? 0) - 1 {
            let round = importInProgress!.rounds[importRound]
            let url = URL(fileURLWithPath: sourceDirectory! + "/" + round.filename!)
            if let data = try? Data(contentsOf: url) {
                Utility.mainThread {
                    round.scoreData?.fileUrl = url
                    let type = url.pathExtension.lowercased()
                    if type == "xml" {
                        _ = UsebioParser(fileUrl: url, data: data, filterSessionId: round.filterSessionId, filterParticipantNumberMin: round.filterParticipantNumberMin, filterParticipantNumberMax: round.filterParticipantNumberMax, overrideEventType: (round.headToHead ? .head_to_head: nil), roundContinuousVPDraw: round.roundContinuousVPDraw, winDrawLevel: round.winDrawLevel, mergeMatches: round.mergeMatches, vpType: round.vpType, completion: importParserComplete)
                    } else if type == "csv" {
                        if GenericCsvParser(fileUrl: url, data: data, manualPointsColumn: round.manualPointsColumn, completion: importParserComplete) == nil {
                            MessageBox.shared.show("Invalid data")
                        }
                    } else {
                        MessageBox.shared.show("File type \(type) not supported")
                    }
                }
            } else {
                MessageBox.shared.show("Unable to access file \(round.filename!)")
                writer = nil
            }
        } else {
            eventDescription = importInProgress!.event!.description!
            eventCode = importInProgress!.event!.code!
            clubCode = importInProgress!.event!.clubCode ?? ""
            chooseBest = importInProgress!.event!.chooseBest ?? 0
            minRankCode = importInProgress!.event!.minRank ?? 0
            maxRankCode = ((importInProgress!.event!.maxRank ?? 999) == 0 ? 999 : importInProgress!.event!.maxRank!)
            
            writer = Writer()
            writer!.eventDescription = eventDescription
            writer!.eventCode = eventCode
            writer!.chooseBest = chooseBest
            writer!.clubCode = clubCode
            writer!.chooseBest = chooseBest
            writer!.minRank = minRankCode
            writer!.maxRank = maxRankCode
            writer!.includeInRace = false
            
            for round in importInProgress!.rounds {
                if let scoreData = round.scoreData {
                    roundName = round.name!
                    localNational = (round.localNational ?? importInProgress!.event!.localNational) == .national ? .national : .local
                    maxAward = round.maxAward ?? 0
                    ewMaxAward = round.ewMaxAward ?? maxAward
                    minEntry = round.minEntry ?? 0
                    awardTo = round.awardTo ?? 0
                    perWin = round.perWin ?? 0
                    filterSessionId = round.filterSessionId ?? ""
                    overrideTeamMembers = round.overrideTeamMembers ?? 0
                    manualPointsColumn = round.manualPointsColumn
                    
                    scoreData.roundName = roundName
                    scoreData.national = localNational == .national
                    scoreData.maxAward = maxAward
                    scoreData.ewMaxAward = ewMaxAward
                    scoreData.reducedTo = round.reducedTo ?? 1
                    scoreData.minEntry = minEntry
                    scoreData.awardTo = awardTo * 100
                    scoreData.perWin = perWin
                    scoreData.filterSessionId = filterSessionId
                    scoreData.aggreateAs = round.aggregateAs
                    scoreData.overrideTeamMembers = (overrideTeamMembers == 0 ? nil : overrideTeamMembers)
                    scoreData.manualPointsColumn = manualPointsColumn
                    if let writerRound = writer?.add(name: round.name!, shortName: round.shortName!, scoreData: round.scoreData!) {
                        writerRound.toe = round.toe
                    }
                }
            }
            
            importInProgress = nil
            addMissingNationalIdWarning()
            if !roundErrors.isEmpty {
                showErrors = true
            }
        }
    }
    
    private func importParserComplete(scoreData: ScoreData?, messages: [String]) {
        if let scoreData = scoreData {
            let (roundMissingNationalIds, errors, warnings) = scoreData.validate()
            missingNationalIds = missingNationalIds || roundMissingNationalIds
            
            var errorList = RoundErrorList(name: importInProgress!.rounds[importRound].name!, errors: [], warnings: [])
            if let errors = errors {
                errorList.errors =  errors
            }
            if let warnings = warnings {
                errorList.warnings = warnings
            }
            if !messages.isEmpty {
                errorList.warnings.append(contentsOf: messages)
            }
            if !errorList.errors.isEmpty || !errorList.warnings.isEmpty {
                roundErrors.append(errorList)
            }
            
            importInProgress!.rounds[importRound].scoreData = scoreData
            importRound += 1
            processNextRound()
            
        } else {
            MessageBox.shared.show(messages.last ?? "Unable to parse file \(importInProgress!.rounds[importRound].filename!)")
            self.writer = nil
            self.scoreData = nil
            self.inputFilename = ""
            self.roundName = ""
        }
    }
    
    private func addMissingNationalIdWarning() {
        if missingNationalIds {
            roundErrors.append(RoundErrorList(name: "General", errors: [], warnings: ["Some players have missing National Ids"]))
        }
    }
    
    private func getEventList() -> [AutoCompleteData] {
        selected = nil
        if eventCode != "" {
            return (MasterData.shared.events.array as! [EventViewModel])
                .filter({$0.eventCode.hasPrefix(eventCode.uppercased()) && $0.active && ($0.startDate ?? Date()) <= Date() && ($0.endDate ?? Date()) >= Date()})
                .sorted(by: {$0.eventCode < $1.eventCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: element.eventCode, desc: element.eventName)})
        } else {
            return []
        }
    }
    
    private func getClubList() -> [AutoCompleteData] {
        selected = nil
        if clubCode != "" {
            return (MasterData.shared.clubs.array as! [ClubViewModel])
                .filter({$0.clubCode.hasPrefix(clubCode.uppercased()) && !$0.clubName.uppercased().contains("CLOSED")})
                .sorted(by: {$0.clubCode < $1.clubCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: element.clubCode, desc: element.clubName)})
        } else {
            return []
        }
    }
    
    private func getMinRankList() -> [AutoCompleteData] {
        getRankList(rankCode: minRankCode, rank: minRank, nullRank: 0)
    }
    
    private func getMaxRankList() -> [AutoCompleteData] {
        getRankList(rankCode: maxRankCode, rank: maxRank, nullRank: 999)
    }
    
    private func getRankList(rankCode: Int, rank: RankViewModel?, nullRank: Int) -> [AutoCompleteData] {
        selected = nil
        if rankCode != nullRank {
            var list = MasterData.shared.ranks.array as! [RankViewModel]
            list.append(RankViewModel(rankCode: nullRank, rankName: "No \(nullRank == 0 ? "minimum" : "maximum") rank"))
            return list.filter({"\($0.rankCode)".hasPrefix("\(rankCode)")})
                .sorted(by: {$0.rankCode < $1.rankCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: "\(element.rankCode)", desc: element.rankName)})
        } else {
            return []
        }
    }
    
}

struct EdgeBorder: Shape {
    var width: CGFloat
    var edges: [Edge]

    func path(in rect: CGRect) -> Path {
        edges.map { edge -> Path in
            switch edge {
            case .top: return Path(.init(x: rect.minX, y: rect.minY, width: rect.width, height: width))
            case .bottom: return Path(.init(x: rect.minX, y: rect.maxY - width, width: rect.width, height: width))
            case .leading: return Path(.init(x: rect.minX, y: rect.minY, width: width, height: rect.height))
            case .trailing: return Path(.init(x: rect.maxX - width, y: rect.minY, width: width, height: rect.height))
            }
        }.reduce(into: Path()) { $0.addPath($1) }
    }
}

extension View {
    func border(width: CGFloat, edges: [Edge], color: Color) -> some View {
        overlay(EdgeBorder(width: width, edges: edges).foregroundColor(color))
    }
}
