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

fileprivate enum ViewField {
    case eventDescription
    case eventCode
    case clubCode
    case minRankText
    case maxRankText
    case strataDefName
}

struct SelectInputView: View {
    @Environment(\.scenePhase) var scenePhase
    @State private var inputFilename: String = ""
    @State private var roundName: String = ""
    @State private var event: EventViewModel? = nil
    @State private var eventCode: String = " " // Intentionally set to space to avoid cursor probs
    @State private var eventCodeDesc: String = ""
    @State private var club: ClubViewModel? = nil
    @State private var clubCodeDesc: String = ""
    @State private var chooseBest: Int = 0
    @State private var clubDesc: String = ""
    @State private var minRank: RankViewModel? = nil
    @State private var minRankCode: Int = 0
    @State private var minRankText: String = ""
    @State private var minRankMessage: String = "No minimum rank"
    @State private var maxRank: RankViewModel? = nil
    @State private var maxRankCode: Int = 999
    @State private var maxRankText: String = ""
    @State private var maxRankDesc: String = "No maximum rank"
    @State private var strataDefName: String = ""
    @State private var strataDef: StrataDefViewModel? = nil
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
    @State private var showBlockedNumbers = false
    @State private var showStratifications = false
    @State private var missingNationalIds = false
    @State private var showAdvancedParameters = false
    @State private var editSettings = Settings.current.copy()
    @State private var windowHeight: CGFloat = 575
    @State private var unexpandedWindowHeight: CGFloat = 575
    @State private var expandedWindowHeight: CGFloat = 675
    @FocusState private var focusedField: ViewField?
    @State private var eventCodeData: [AutoCompleteData] = []
    @State private var clubCodeData: [AutoCompleteData] = []
    @State private var minRankCodeData: [AutoCompleteData] = []
    @State private var maxRankCodeData: [AutoCompleteData] = []
    @State private var strataDefData: [AutoCompleteData] = []
    @State private var selected: Int? = nil
    @State private var downloadingMemberList: Bool = true
    @State private var downloadMemberListMessage: String = ""
    @State var minRankList: [RankViewModel] = []
    @State var maxRankList: [RankViewModel] = []

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
                                
                                HStack {
                                    
                                    roundNameView
                                    
                                    strataDefView
                                    
                                    Spacer()
                                }
                                
                                
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
                                        Spacer()
                                    }
                                    Spacer()
                                }
                                .frame(height: 20)
                                .onTapGesture {
                                    showAdvancedParameters.toggle()
                                    windowHeight = showAdvancedParameters ? expandedWindowHeight : unexpandedWindowHeight
                                }
                                
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
                autoCompleteViews
            }
            Spacer()
        }
        .frame(width: 900, height: windowHeight)
        .onChange(of: scenePhase) { oldPhase, newPhase in
            phaseChange(newPhase)
        }
        .onChange(of: focusedField) { (oldValue, newValue) in
            changeFocus(leaving: oldValue, entering: newValue)
        }
        .sheet(isPresented: $showErrors) {
            ShowErrorsView(roundErrors: roundErrors)
        }
        .sheet(isPresented: $showSettings) {
            SettingsView(settings: editSettings)
        }
        .sheet(isPresented: $showBlockedNumbers) {
            BlockedNumbersView()
        }
        .sheet(isPresented: $showStratifications) {
            StratificationsView()
        }
        .sheet(isPresented: $showImportEvents) {
            EventImportView()
        }
        .sheet(isPresented: $showImportClubs) {
            ClubImportView()
        }
        .sheet(isPresented: $showImportRanks) {
            RankImportView() {
                refreshRankLists()
            }
        }
        .sheet(isPresented: $downloadingMemberList) {
            downloadingMemberListView()
        }
        .onAppear {
            setupFields()
            downloadingMemberList = false
                // downloadMemberList()
            refreshRankLists()
        }
    }
    
    private func phaseChange(_ newPhase: ScenePhase) {
        if newPhase == .active && !downloadingMemberList, let lastDownloaded = MemberList.shared.lastDownloaded {
            // When view becomes active then update memmber list if it is more than 12 hours old
            if Date().timeIntervalSince(lastDownloaded) > 12 * 60 * 60 {
                downloadMemberList(message: "Updating Member List")
            }
        }
    }
    
    private func setupFields() {
        set(eventCode: " ")
        set(clubCode: "")
        set(minRankText: "0")
        set(maxRankText: "999")
        set(strataDefName: "")
    }
    
    private func refreshRankLists() {
        minRankList = ([RankViewModel(rankCode: 0, rankName: "No minimum rank")] + MasterData.shared.ranks.array as! [RankViewModel]).filter({$0.rankCode != 1})
        maxRankList = (MasterData.shared.ranks.array as! [RankViewModel] + [RankViewModel(rankCode: 999, rankName: "No maximum rank")]).filter({$0.rankCode != 1})
    }
    
    private func downloadMemberList(message: String = "Downloading Member List") {
        downloadMemberListMessage = message
        downloadingMemberList = true
        MemberList.shared.download() { (success, errorMessage) in
            downloadingMemberList = false
            if !success {
                MessageBox.shared.show("Failed to download member list \n(\(errorMessage))\n\n\nWill continue using previous version but some ranks may be out of date. This is probably only an issue for stratified events.", okAction: {
                    downloadingMemberList = false
                })
            }
        }
    }
    
    private func downloadingMemberListView() -> some View {
        VStack {
            Spacer()
            Spacer().frame(height: 100)
            HStack {
                Spacer()
                Text(downloadMemberListMessage).font(defaultFont).foregroundColor(.blue)
                Spacer()
            }
            Spacer().frame(height: 10)
            Text("Please wait").font(captionFont)
            Spacer().frame(height: 100)
            Spacer()
        }
    }
    
    private var autoCompleteViews : some View {
        VStack {
            switch focusedField {
            case .eventCode:
                AutoCompleteView(autoComplete: autoComplete, field: focusedField!, selected: $selected, codeWidth: 80, data: $eventCodeData, valid: event != nil) { (newValue) in
                    eventCode = newValue
                }
            case .clubCode:
                AutoCompleteView(autoComplete: autoComplete, field: focusedField!, selected: $selected, codeWidth: 80, data: $clubCodeData, valid: club != nil) { (newValue) in
                    clubCodeDesc = newValue
                }
            case .minRankText:
                AutoCompleteView(autoComplete: autoComplete, field: focusedField!, selected: $selected, codeWidth: 80, data: $minRankCodeData, valid: minRank != nil) { (newValue) in
                    minRankText = newValue
                }
            case .maxRankText:
                AutoCompleteView(autoComplete: autoComplete, field: focusedField!, selected: $selected, codeWidth: 80 , data: $maxRankCodeData, valid: maxRank != nil) { (newValue) in
                    maxRankText = newValue
                }
            case .strataDefName:
                AutoCompleteView(autoComplete: autoComplete, field: focusedField!, selected: $selected, codeWidth: 270, data: $strataDefData, hideList: strataDef != nil, hasDescription: false) { (newValue) in
                    strataDefName = newValue
                }
            default:
                EmptyView()
            }
        }
    }
    
    private func changeFocus(leaving: ViewField?, entering: ViewField?) {
        let listSelected = selected
        switch leaving {
        case .minRankText:
            let list = getMinRankList(text: minRankText)
            if list.count == 1 || (list.count > 0 && listSelected != nil) {
                set(minRankText: "\(list[listSelected ?? 0].code)")
            }
            if Int(minRankText) == 0 {
                set(minRankText: "")
            } else if minRank == nil {
                set(minRankText: "")
            }
        case .maxRankText:
            let list = getMaxRankList(text: maxRankText)
            if list.count == 1 || (list.count > 0 && listSelected != nil) {
                set(maxRankText: "\(list[listSelected ?? 0].code)")
            }
            if Int(maxRankText) == 999 {
                set(maxRankText: "")
            } else if maxRank == nil {
                set(maxRankText: "")
            }
        case .eventCode:
            let list = getEventList()
            if list.count == 1 || (list.count > 0 && listSelected != nil) {
                set(eventCode: (list[listSelected ?? 0].code))
            } else if event == nil {
                set(eventCode: "")
            }
        case .clubCode:
            let list = getClubList()
            if list.count == 1 || (list.count > 0 && listSelected != nil) {
                set(clubCode: list[listSelected ?? 0].code)
            } else if club == nil {
                set(clubCode: "")
            }
        case .strataDefName:
            let list = getStrataDefList()
            if list.count == 1 || (list.count > 0 && listSelected != nil) {
                set(strataDefName: list[listSelected ?? 0].code)
            } else if strataDef == nil {
                set(strataDefName: "")
            }
            if strataDef != nil {
                // Clear 2-winner pairs award
                ewMaxAward = 0
            }
        default:
            break
        }
    }
    
    private func set(eventCode newValue: String) {
        event = EventViewModel.event(eventCode: newValue)
        eventCode = newValue
        eventCodeDesc = event?.eventName ?? "Invalid event code"
        if let event = event {
            if event.localAllowed && !event.nationalAllowed {
                localNational = .local
            } else if !event.localAllowed && event.nationalAllowed {
                localNational = .national
            }
            if event.validMinRank > 0, let minRank = RankViewModel.rank(rankCode: event.validMinRank) {
                set(minRankText: "\(minRank.rankCode)")
            }
            if event.validMaxRank < 999, let maxRank = RankViewModel.rank(rankCode: event.validMaxRank) {
                set(maxRankText: "\(maxRank.rankCode)")
            }
            if event.originatingClubCode != "" {
                set(clubCode: event.originatingClubCode)
            }
        }
    }
    
    private func set(clubCode newValue: String) {
        club = ClubViewModel.club(clubCode: newValue)
        clubCodeDesc = newValue
        clubDesc = club?.clubName ?? (clubCodeDesc == "" ? "No club code specified" : "Invalid club code")
    }
    
    private func set(minRankText newValue: String) {
        if let newRankCode = Int(newValue.trim() == "" ? "0" : newValue) {
            minRank = minRankList.first(where: { $0.rankCode == newRankCode })
        } else {
            minRank = minRankList.first(where: { $0.rankName == newValue })
        }
        minRankText = (newValue.trim() == "" ? "0" : newValue)
        minRankCode = minRank?.rankCode ?? -1
        minRankMessage = (minRank?.rankName ?? "Invalid rank code")
        if minRank != nil && maxRankCode < minRankCode {
            set(maxRankText: minRankText)
        }
    }
    
    private func set(maxRankText newValue: String) {
        if let newRankCode = Int(newValue.trim() == "" ? "999" : newValue) {
            maxRank = maxRankList.first(where: { $0.rankCode == newRankCode })
        } else {
            maxRank = maxRankList.first(where: { $0.rankName == newValue })
        }
        maxRankText = (newValue.trim() == "" ? "999" : newValue)
        maxRankCode = maxRank?.rankCode ?? 0
        if maxRank != nil && maxRankCode < minRankCode {
            maxRankDesc = "Below minimum rank"
        } else {
            maxRankDesc = (maxRank?.rankName ?? "Invalid rank code")
        }
    }
    
    private func set(strataDefName newValue: String) {
        strataDef = StrataDefViewModel.strataDef(name: newValue)
        strataDefName = newValue
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
                Input(title: "Import filename", field: $inputFilename, placeHolder: "No import file specified", topSpace: 16, leadingSpace: 14, width: 365, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: true, isReadOnly: true, pickerAction: chooseFile, onChange: { newValue in
                    // focusedField = .eventDescription
                })
                .help("Select input XML file for a single round event.\nFor a multi-round event use the Paste Config button below to import the details.")
                    
                Spacer()
            }
        }
    }
    
    private var eventDescriptionView: some View {
        VStack(spacing: 5) {
            HStack {
                InputTitle(title: "Event Description", fillTrailing: false)
                VStack(spacing: 0) {
                    Spacer().frame(height: inputTopHeight)
                    Text("(Spreadsheet name)")
                }
                Spacer()
            }
            HStack {
                Input(field: $eventDescription, leadingSpace: 50, width: 365, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true)
                    .focused($focusedField, equals: .eventDescription)
                    .help("Enter the event description. Note that this becomes the default name for the event spreadsheet and the title for outputs.")
                Spacer().frame(width: 30)
                
                Text("Include race:").font(inputFont)
                Picker("", selection: $includeInRace) {
                    Text("Include").tag(true)
                    Text("Exclude").tag(false)
                }
                .pickerStyle(.segmented)
                .help("Include/exclude the event in a multi-event race")
                .disabled((writer?.rounds.count ?? 0) >= 1)
                .focusable(false)
                
                Spacer()
            }
        }
    }
    
    private var eventCodeView: some View {
        HStack {
            Spacer().frame(width: 42)
            
            Input(title: "Event code:", field: $eventCode, desc: $eventCodeDesc, descOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, autoCapitalize: .sentences, autoCorrect: false, isEnabled: true, limitText: 6, onKeyPress: eventKeyPress, detectKeys: AutoComplete.detectKeys) { (newValue) in
                set(eventCode: newValue)
                eventCodeData = getEventList()
            }
            .help("Enter an event code to be used to post the MPs")
            .matchedGeometryEffect(id: ViewField.eventCode, in: autoComplete, anchor: .bottomTrailing)
        }
    }
    
    private func eventKeyPress(_ press: KeyPress) -> KeyPress.Result{
        AutoComplete.onKeyPress(press, selected: $selected, maxSelected: eventCodeData.count) {
            set(eventCode: eventCodeData[selected!].code)
        }
    }
    
    private func getEventList() -> [AutoCompleteData] {
        selected = nil
        if eventCode != "" {
            return (MasterData.shared.events.array as! [EventViewModel])
                .filter({Utility.wordSearch(for: eventCode, in: "\($0.eventCode) \($0.eventName)")
                     && $0.active && ($0.startDate ?? Date()) <= Date() && ($0.endDate ?? Date()) >= Date()})
                .sorted(by: {$0.eventCode < $1.eventCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: element.eventCode, desc: element.eventName)})
        } else {
            return []
        }
    }
    
    private var clubCodeView: some View {
        HStack {
            
            Spacer().frame(width: 30)
            
            Input(title: "Club code:", field: $clubCodeDesc, desc: $clubDesc, descOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, autoCapitalize: .sentences, autoCorrect: false, isEnabled: event == nil || event!.originatingClubCode == "", limitText: 5, onKeyPress: clubKeyPress, detectKeys: AutoComplete.detectKeys) { (newValue) in
                set(clubCode: newValue)
                clubCodeData = getClubList()
            }
            .help("Enter a club code if this is not a licensed event. Leave blank on licensed events.")
            .focused($focusedField, equals: .clubCode)
            .matchedGeometryEffect(id: ViewField.clubCode, in: autoComplete, anchor: .bottomTrailing)
        }
    }
    
    private func clubKeyPress(_ press: KeyPress) -> KeyPress.Result{
        AutoComplete.onKeyPress(press, selected: $selected, maxSelected: clubCodeData.count) {
            set(clubCode: clubCodeData[selected!].code)
        }
    }
    
    private func getClubList() -> [AutoCompleteData] {
        selected = nil
        if clubCodeDesc != "" {
            return (MasterData.shared.clubs.array as! [ClubViewModel])
                .filter({Utility.wordSearch(for: clubCodeDesc, in: "\($0.clubCode) \($0.clubName)") && !$0.clubName.uppercased().contains("CLOSED")})
                .sorted(by: {$0.clubCode < $1.clubCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: element.clubCode, desc: element.clubName)})
        } else {
            return []
        }
    }
    
    private var rankingsView: some View {
        HStack {
            
            Spacer().frame(width: 42)
            
            Input(title: "Minimum:", field: $minRankText, desc: $minRankMessage, descOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, isEnabled: event == nil || event!.validMinRank == 0, onKeyPress: minRankKeyPress, detectKeys: AutoComplete.detectKeys) { (newValue) in
                set(minRankText: newValue)
                minRankCodeData = (newValue == "" ? [] : getMinRankList(text: newValue))
            }
            .help("Enter a minimum rank code for this event or leave blank if you do not want to filter by rank")
            .focused($focusedField, equals: .minRankText)
            .matchedGeometryEffect(id: ViewField.minRankText, in: autoComplete, anchor: .bottomTrailing)
            
            Spacer().frame(width: 30)
            
            Input(title: "Maximum:", field: $maxRankText, desc: $maxRankDesc, descOffset: 80, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, isEnabled: event == nil || event!.validMaxRank == 999, onKeyPress: maxRankKeyPress, detectKeys: AutoComplete.detectKeys) { (newValue) in
                set(maxRankText: newValue)
                maxRankCodeData = (newValue == "" ? [] : getMaxRankList(text: newValue))
            }
            .help("Enter a maximum rank code for this event or leave blank if you do not want to filter by rank")
            .focused($focusedField, equals: .maxRankText)
            .matchedGeometryEffect(id: ViewField.maxRankText, in: autoComplete, anchor: .bottomTrailing)
            
            Spacer()
        }
    }
    
    private func getMinRankList(text: String) -> [AutoCompleteData] {
        getRankList(text: text, minimum: true)
    }
    
    private func getMaxRankList(text: String) -> [AutoCompleteData] {
        getRankList(text: text, minimum: false)
    }
    
    private func getRankList(text: String, minimum: Bool) -> [AutoCompleteData] {
        selected = nil
        return (minimum ? minRankList : maxRankList).filter({Utility.wordSearch(for: text, in: "\($0.rankCode) \($0.rankName)")})
                .sorted(by: {$0.rankCode < $1.rankCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: "\(element.rankCode)", desc: element.rankName)})
    }
    
    private func minRankKeyPress(_ press: KeyPress) -> KeyPress.Result{
        AutoComplete.onKeyPress(press, selected: $selected, maxSelected: minRankCodeData.count) {
            set(minRankText: "\(minRankCodeData[selected!].code)")
        }
    }
    
    private func maxRankKeyPress(_ press: KeyPress) -> KeyPress.Result{
        AutoComplete.onKeyPress(press, selected: $selected, maxSelected: maxRankCodeData.count) {
            set(maxRankText: "\(maxRankCodeData[selected!].code)")
        }
    }
    
    private var roundNameView: some View {
        HStack {
            Spacer().frame(width: 42)
            
            Input(title: "Round name:", field: $roundName, topSpace: 0, width: 160, inlineTitle: true, inlineTitleWidth: 117, isEnabled: true)
                .help("Enter a (short) name for this round of the event. This will be used as a prefix for some tabs in the spreadsheet.")
        }
    }
    
    private var strataDefView: some View {
        HStack {
            Spacer().frame(width: 120)
            
            Input(title: "Stratification:", field: $strataDefName, topSpace: 0, width: 270, inlineTitle: true, inlineTitleWidth: 95, onKeyPress: strataDefKeyPress, detectKeys: AutoComplete.detectKeys) { (newValue) in
                set(strataDefName: newValue)
                strataDefData = getStrataDefList()
            }
            .help("Enter a stratification defintion if this is a stratified event.")
            .focused($focusedField, equals: .strataDefName)
            .matchedGeometryEffect(id: ViewField.strataDefName, in: autoComplete, anchor: .bottomTrailing)
        }
    }
    
    private func strataDefKeyPress(_ press: KeyPress) -> KeyPress.Result{
        AutoComplete.onKeyPress(press, selected: $selected, maxSelected: strataDefData.count) {
            set(strataDefName: strataDefData[selected!].code)
        }
    }
    
    private func getStrataDefList() -> [AutoCompleteData] {
        selected = nil
        if strataDefName != "" {
            return (MasterData.shared.strataDefs.array as! [StrataDefViewModel])
                .filter({Utility.wordSearch(for: strataDefName, in: $0.name)})
                .sorted(by: {StrataDefViewModel.defaultSort($0, $1)}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: element.name, desc: "")})
        } else {
            return []
        }
    }
    
    private var pointsAwardView: some View {
        VStack {
            HStack {
                
                Spacer().frame(width: 24)
                
                Picker("Level:                     ", selection: $localNational) {
                    Text("Local").tag(Level.local)
                    Text("National").tag(Level.national)
                }
                .help("Select the type of MPs to be awarded")
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
                    .help("Enter the maximum MP award. This might be reduced if the minimum entry is not reached.")
                    .disabled(basis != .standard)
                
                InputFloat(title: "Max E/W award:", field: $ewMaxAward, topSpace: 0, leadingSpace: 30, width: 60, inlineTitle: true, inlineTitleWidth: 120)
                    .help("Enter the maximum E/W MP award if this is a 2-winner event with different points for each direction.")
                    .disabled(basis != .standard || strataDef != nil)
                Spacer()
            }
            
            Spacer().frame(height: 10)
            
            HStack(spacing: 0) {
                Spacer().frame(width: 42)
                
                InputInt(title: "Min entry:", field: $minEntry, topSpace: 0, width: 60, inlineTitle: true, inlineTitleWidth: 117)
                    .help("Enter the minimum number of entries for the event to be eligible for the full MP award. If the field is smaller the maximum award will be scaled down pro rata.")
                    .disabled(basis != .standard)
                
                InputFloat(title: "Award to %:", field: $awardTo, topSpace: 0, leadingSpace: 30, width: 60, inlineTitle: true, inlineTitleWidth: 120)
                    .disabled(basis != .standard)
                    .help("Enter the percentage of the field who should receive a bonus award.")
                
                InputFloat(title: "Per win:", field: $perWin, topSpace: 0, leadingSpace: 30, width: 60, inlineTitle: true, inlineTitleWidth: 70)
                    .help("Enter the the number of MPs to be awarded for each win in a Swiss event. Half of this value (rounded up) will be awarded for a draw.")
                    .disabled(basis == Basis.manual)
                
                Spacer()
            }
        }
    }
    
    private var advancedParametersView: some View {
        VStack {
            HStack{
                Spacer().frame(width: 42)
                Picker("Calculation basis:  ", selection: $basis) {
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
                    InputInt(title: "Override players:", field: $overrideTeamMembers, width: 127, inlineTitle: true, inlineTitleWidth: 123)
                        .help(Text("Used to override the number of team members considered in the raw data. Primarily used to remove subs from the awards when their impact has not been material, or to increase the number of players to allow manual entry in the spreadsheet"))
                }
                Spacer().frame(width: 68)
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
                .padding([.top], 10)
                .padding([.leading, .trailing], 10)
                .padding([.bottom], 0)
            Button("Edit Settings") { showSettings = true }
                .padding([.leading, .trailing], 28)
                .padding([.top], 6)
                .padding([.bottom], 10)
            Separator()
            Text("Configuration Data").bold()
                .padding([.top], 10)
                .padding([.leading, .trailing], 10)
                .padding([.bottom], 0)
            Button("Stratifications") { showStratifications = true }
                .padding([.leading, .trailing], 28)
                .padding([.top], 6)
                .padding([.bottom], 10)
            Button("Blocked National Ids") { showBlockedNumbers = true }
                .padding([.leading, .trailing], 28)
                .padding([.top], 6)
                .padding([.bottom], 10)
            Separator()
            Text("Imports").bold()
                .padding([.top], 6)
                .padding([.leading, .trailing], 10)
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
        .help("Click here to change settings, import ranks, clubs or events from Mempad, or to setup other configuration data.")
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
                    writer?.clubCode = clubCodeDesc
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
                scoreData.strata = strataDef?.activeStrata() ?? []
                scoreData.customFooter = strataDef?.customFooter ?? ""
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
            Text("Add Round")
                .help("Click this to add the round defined above to the current event.")
                .foregroundColor(Palette.highlightButton.text)
                .frame(width: 100, height: 30)
                .font(.callout).minimumScaleFactor(0.5)
                .background(Palette.highlightButton.background)
                .cornerRadius(15)
        }
        .buttonStyle(PlainButtonStyle())
        .focusable(false)
        .disabled(scoreData == nil || eventCode == "" || eventDescription == "" || roundName == "" || event == nil || (clubCodeDesc != "" && club == nil) || (minRankCode != 0 && minRank == nil) || (maxRankCode != 999 && maxRank == nil) || maxRankCode < minRankCode || (club == nil && event!.clubMandatory) || (awardTo <= 0 && basis != .manual) || (writer?.rounds.contains(where: {$0.shortName == roundName}) ?? false))
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
            Text("Create Spreadsheet")
                .help("Click this to create the spreadsheet")
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
                .help("Click this to import details of a multi-round event. The clipboard must be in a very specific format, normally copied from a 'rounds' spreadheet.")
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
                .help("Click this to clear the current event. This happens automatically if a spreadsheet is created.")
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
            } else {
                if let usebioDescription = scoreData.events.last?.description, eventDescription == ""  {
                    eventDescription = usebioDescription
                }
                if let usebioType = scoreData.events.last?.type?.participantType?.desc, roundName == ""  {
                    roundName = usebioType
                }
            }
            if let warnings = warnings {
                errorList.warnings = warnings
            }
            if !messages.isEmpty {
                errorList.warnings.append(contentsOf: messages)
            }
            
            roundErrors = []
            if errors != nil || warnings != nil || messages != [] {
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
            clubCodeDesc = importInProgress!.event!.clubCode ?? ""
            chooseBest = importInProgress!.event!.chooseBest ?? 0
            minRankCode = importInProgress!.event!.minRank ?? 0
            maxRankCode = ((importInProgress!.event!.maxRank ?? 999) == 0 ? 999 : importInProgress!.event!.maxRank!)
            
            writer = Writer()
            writer!.eventDescription = eventDescription
            writer!.eventCode = eventCode
            writer!.chooseBest = chooseBest
            writer!.clubCode = clubCodeDesc
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
                writer = nil
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
