//
//  Stratifications View.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 16/02/2026.
//

import SwiftUI

fileprivate enum ViewField {
    case code
    case rank
    case percent
}

struct StratificationsView : View {
    @State var selected = StrataDefViewModel()
    @ObservedObject var edit = StrataDefViewModel()
    @State var exit: Bool = true
    @State var editMode: EditMode? = nil
    var showDetail: Binding<Bool> {
        Binding {
            editMode != nil
        } set: { _ in
        }
    }
    
    var body: some View {
        let strataDefList = (MasterData.shared.strataDefs.array as! [StrataDefViewModel]).sorted(by: { StrataDefViewModel.defaultSort($0, $1)})
        
        StandardView("Select Input") {
            VStack(spacing: 0) {
                Banner(title: Binding.constant("Stratification Definitions"), backEnabled: { exit }, backAction: {
                    return true
                })
                HStack {
                    Spacer().frame(width: 20)
                    VStack {
                        Spacer().frame(height: 20)
                        VStack(spacing: 0) {
                            HStack {
                                Spacer().frame(width: 10)
                                Text("Name")
                                Spacer()
                            }
                            .frame(height: 25)
                            .palette(.contrastTile)
                            ScrollView {
                                ScrollViewReader { scrollViewProxy in
                                    LazyVStack(spacing: 0) {
                                        ForEach(strataDefList, id: \.strataDefId) { (strataDef) in
                                            VStack(spacing: 0) {
                                                HStack {
                                                    Spacer().frame(width: 10)
                                                    Text(strataDef.name)
                                                        .lineLimit(1)
                                                        .truncationMode(.tail)
                                                    Spacer()
                                                }
                                                .frame(height: 25)
                                                .palette(selected.strataDefId == strataDef.strataDefId ? .highlightTile : .tile)
                                                Separator(thickness: 1.0)
                                            }
                                            .id(strataDef.strataDefId)
                                            .onTapGesture {
                                                selected = strataDef
                                                edit.copy(from: strataDef)
                                                editMode = .amend
                                            }
                                        }
                                    }
                                    .onChange(of: selected.strataDefId) {
                                        scrollViewProxy.scrollTo(selected.strataDefId, anchor: nil)
                                    }
                                }
                            }
                        }
                        .background(Palette.tile.background)
                        .cornerRadius(6)
                        Spacer()
                        HStack {
                            Spacer()
                            CustomButton.button(image: "plus", title: EditMode.add.title) {
                                edit.copy(from: StrataDefViewModel())
                                editMode = .add
                            }
                        }
                        Spacer().frame(height: 10)
                    }
                    Spacer().frame(width: 20)
                }
            }
        }
        .sheet(isPresented: showDetail) {
            StratificationsDetailView(edit: edit, editMode: editMode!, completion: { (success) in
                if success {
                    selected = edit
                } else {
                    edit.copy(from: selected)
                }
                editMode = nil
            })
        }
        .interactiveDismissDisabled(!exit)
        .frame(height: 400)
    }
}

struct StratificationsDetailView: View {
    @Environment(\.dismiss) var dismiss
    @StateObject var edit: StrataDefViewModel
    @State var editMode: EditMode
    @State var showStrataView: Bool = false
    @State var editStrata: StrataElement = StrataElement()
    @State var strataIndex: Int = -1
    @State var maxRank: Int = 999
    @State var maxPercent: Float = 100
    @State var strataEditMode: EditMode = .display
    @State var completion: (Bool)->() = { (success) in }
    
    var body: some View {
        
        StandardView("Stratifications") {
            
            VStack(spacing: 0) {
                
                Banner(title: Binding.constant("\(editMode.title) Stratification"), alternateStyle: true, back: false)
                
                Spacer().frame(height: 20)
                
                HStack {
                    Input(title: "Description:", field: $edit.name, message: $edit.nameMessage, leadingSpace: 20, width: 480, isReadOnly: !editMode.enabled)
                    Spacer()
                }
                
                Spacer().frame(height: 30)
                
                let tableColums = [GridItem(.fixed(30), spacing: 10, alignment: .leading),
                                   GridItem(.fixed(80), spacing: 10, alignment: .leading),
                                   GridItem(.fixed(200), spacing: 10, alignment: .leading),
                                   GridItem(.fixed(100), spacing: 10, alignment: .leading)]
                
                HStack {
                    Spacer().frame(width: 10)
                    
                    ZStack {
                        VStack(spacing: 0) {
                            UnevenRoundedRectangle(cornerRadii: .init(topLeading: 6, topTrailing: 6), style: .continuous)
                                .frame(width: 486, height: 30).foregroundColor(Palette.tile.background)
                            UnevenRoundedRectangle(cornerRadii: .init(bottomLeading: 6, bottomTrailing: 6), style: .continuous)
                                .frame(width: 486, height: 80).foregroundColor(Palette.alternate.background)
                        }
                        LazyVGrid(columns: tableColums, spacing: 0) {
                            GridRow {
                                CenteredClickableText(text: "#")
                                CenteredClickableText(text: "Name")
                                LeadingClickableText(text: "Maximum Rank")
                                CenteredClickableText(text: "Percentage")
                            }
                            .bold()
                            .frame(height: 30)
                            .foregroundColor(Palette.tile.text)
                            
                            ForEach(edit.strata.indices, id: \.self) { index in
                                GridRow {
                                    
                                    CenteredClickableText(text: "\(index + 1)").bold()
                                    if index == 0 && edit.strata[index].code.trim() == "" {
                                        CenteredClickableText(text: "Blank")
                                            .foregroundColor(Palette.background.faintText)
                                    } else {
                                        CenteredClickableText(text: "\(edit.strata[index].code)")
                                    }
                                    
                                    if index == 0 {
                                        LeadingClickableText(text: "Any rank")
                                            .foregroundColor(Palette.background.faintText)
                                    } else if edit.strata[index].code.trim() == "" {
                                        LeadingClickableText(text: "")
                                    } else {
                                        let rank = RankViewModel.rank(rankCode: edit.strata[index].rank)
                                        LeadingClickableText(text: "\(rank?.rankName ?? "")")
                                    }
                                    
                                    if index == 0 || edit.strata[index].code.trim() != "" {
                                        CenteredClickableText(text: "\(edit.strata[index].percent.toString(places: 2)) %")
                                    } else {
                                        CenteredClickableText(text: "")
                                    }
                                }
                                .onTapGesture {
                                    if index <= 1 || edit.strata[index - 1].code != "" {
                                        strataIndex = index
                                        editStrata = edit.strata[index]
                                        maxRank = (index == 0 ? 999 : edit.strata[index - 1].rank)
                                        maxPercent = (index == 0 ? 100 : edit.strata[index - 1].percent)
                                        strataEditMode = (editStrata.code.trim() == "" ? .add : .amend)
                                        showStrataView = true
                                    }
                                }
                                .frame(height: 25)
                            }
                        }
                    }
                    
                    Spacer().frame(width: 20)
                }
                
                Spacer()
                Separator(thickness: 1)
                Spacer().frame(height: 10)
                HStack {
                    Spacer()
                    
                    CustomButton.button(title: "Cancel") {
                        completion(false)
                        dismiss()
                    }
                    
                    Spacer().frame(width: 60)
                    
                    CustomButton.button(title: editMode.action) {
                        switch editMode {
                        case .amend:
                            edit.save()
                        case .add:
                            let insert = StrataDefViewModel()
                            insert.copy(from: edit, copyMO: false)
                            insert.insert()
                            edit.copy(from: insert)
                        default:
                            break
                        }
                        completion(true)
                        dismiss()
                    }
                    .disabled($edit.name.wrappedValue.isEmpty)
                    
                    if editMode == .amend {
                        
                        Spacer().frame(width: 60)
                        
                        CustomButton.button(image: "trash", title: "Remove") {
                            MessageBox.shared.show("Are you sure you want to remove this record?", cancelText: "Cancel", okText: "Remove", okAction: {
                                edit.remove()
                                completion(true)
                                dismiss()
                            })
                        }
                    }
                    
                    Spacer()
                }
                Spacer().frame(height: 10)
            }
        }
        .sheet(isPresented: $showStrataView) {
            StratificationDetailStrataView(edit: $editStrata, index: $strataIndex, maxRank: $maxRank, maxPercent: $maxPercent, editMode: $strataEditMode) { (action) in
                switch action {
                case .update:
                    edit.strata[strataIndex] = editStrata
                case .remove:
                    edit.removeStratum(at: strataIndex)
                case .noChange:
                    break
                }
            }
        }
        .frame(width: 600, height: 360)
        .palette(.background)
        .onAppear {
            
        }
    }
}

fileprivate enum Action {
    case update
    case remove
    case noChange
}

struct StratificationDetailStrataView: View {
    @Environment(\.dismiss) var dismiss
    @Binding var edit: StrataElement
    @Binding var index: Int
    @Binding var maxRank: Int
    @Binding var maxPercent: Float
    @Binding var editMode: EditMode
    fileprivate var completion: (Action)->()
    @State private var code: String = ""
    @State private var rank: RankViewModel? = nil
    @State private var rankCode: Int = 999
    @State private var rankText: String = ""
    @State private var rankDesc: String = "No minimum rank"
    @State private var rankMessage: String = ""
    @State private var rankData: [AutoCompleteData] = []
    @State private var percent: Float = 0
    @State private var selected: Int? = nil
    @FocusState private var focusedField: ViewField?
    @State var rankList: [RankViewModel] = []
    
    @Namespace private var autoComplete
    
    var body: some View {
        
        StandardView("Stratum Details") {
            
            ZStack {
                
                VStack(spacing: 0) {
                    
                    Banner(title: Binding.constant("Stratum \(index + 1) Details"), alternateStyle: true, back: false)
                    
                    Spacer().frame(height: 10)
                    
                    HStack {
                        
                        Spacer().frame(width: 30)
                        
                        VStack(spacing: 0) {
                            
                            HStack {
                                Spacer().frame(width: 197)
                                Text(nameMessage(index: index, name: code))
                                    .foregroundColor(Palette.background.strongText)
                                    .font(.caption)
                                Spacer()
                            }
                            .frame(height: 8)
                            HStack {
                                Input(title: "Name: ", field: $code, width: 80, inlineTitle: true, inlineTitleWidth: 190, limitText: 6)
                                    .focused($focusedField, equals: .code)
                                Spacer()
                            }
                            
                            if index != 0 {
                                
                                Spacer().frame(height: 20)
                                
                                HStack {
                                    Spacer().frame(width: 197)
                                    Text(rankMessage)
                                        .foregroundColor(Palette.background.strongText)
                                        .font(.caption)
                                    Spacer()
                                }
                                .frame(height: 8)
                                HStack {
                                    Input(title: "Maximum Rank in Stratum: ", field: $rankText, desc: $rankDesc, descOffset: 80, width: 270, inlineTitle: true, inlineTitleWidth: 185, isEnabled: index != 0, onKeyPress: rankKeyPress, detectKeys: AutoComplete.detectKeys) { (newValue) in
                                        set(rankText: newValue)
                                        rankData = getRankList(text: rankText)
                                    }
                                    .focused($focusedField, equals: .rank)
                                    .matchedGeometryEffect(id: ViewField.rank, in: autoComplete, anchor: .bottomTrailing)
                                    Spacer()
                                }
                            }
                            
                            Spacer().frame(height: 20)
                            
                            HStack {
                                Spacer().frame(width: 197)
                                Text(percentMessage(index: index, percent: percent))
                                    .foregroundColor(Palette.background.strongText)
                                    .font(.caption)
                                Spacer()
                            }
                            .frame(height: 8)
                            Spacer().frame(height: 5)
                            HStack {
                                InputFloat(title: "Percentage of Top Award: ", field: $percent, width: 80, places: 2, inlineTitleWidth: 185)
                                    .focused($focusedField, equals: .percent)
                                Spacer()
                            }
                        }
                    }
                    Spacer()
                    Separator(thickness: 1)
                    Spacer().frame(height: 10)
                    HStack {
                        Spacer()
                        
                        CustomButton.button(title: "Cancel") {
                            completion(.noChange)
                            dismiss()
                        }
                        
                        Spacer().frame(width: 60)
                        
                        CustomButton.button(title: editMode.update) {
                            edit.code = code
                            edit.rank = rank!.rankCode
                            edit.percent = percent
                            completion(.update)
                            dismiss()
                        }
                        .disabled((index != 0 && code.isEmpty) ||
                                  rank == nil ||
                                  percent <= 0 ||
                                  (index == 0 && percent > 100) ||
                                  (index != 0 && percent >= maxPercent))
                        
                        if index != 0 && editMode == .amend {
                            Spacer().frame(width: 60)
                            
                            CustomButton.button(image: "trash", title: "Remove") {
                                completion(.remove)
                                dismiss()
                            }
                        }
                        
                        Spacer()
                    }
                    Spacer().frame(height: 10)
                }
                VStack {
                    switch focusedField {
                    case .rank:
                        AutoComplete.view(autoComplete: autoComplete, field: ViewField.rank, selected: $selected, codeWidth: 80, data: $rankData, valid: rank != nil) { (newValue) in
                            rankText = newValue
                        }
                    default:
                        EmptyView()
                    }
                }
            }
        }
        .onChange(of: focusedField) { (oldValue, _) in
            changeFocus(leaving: oldValue)
        }
        .onAppear {
            rankList = ((MasterData.shared.ranks.array as! [RankViewModel]) + [RankViewModel(rankCode: 999, rankName: "No maximum rank")]).filter({$0.rankCode != 1 && (index == 0 || $0.rankCode < maxRank) })
            code = edit.code
            set(rankText: "\(edit.rank)")
            percent = edit.percent
        }
        .frame(width: 530, height: (index == 0 ? 230 : 280))
        .palette(.background)
    }
    
    private func changeFocus(leaving: ViewField?) {
        switch leaving {
        case .rank:
            let list = getRankList(text: rankText)
            if list.count == 1 {
                set(rankText: "\(list.first!.code)")
            }
            if Int(rankText) == 999 {
                set(rankText: "")
            }
        default:
            break
        }
    }

    func set(rankText newValue: String) {
        if let newRankCode = Int(newValue.trim() == "" ? "999" : newValue) {
            rank = rankList.first(where: { $0.rankCode == newRankCode })
        } else {
            rank = rankList.first(where: { $0.rankName == newValue })
        }
        rankText = newValue
        rankCode = rank?.rankCode ?? -1
        rankDesc = (rank?.rankName ?? "Invalid rank code")
        rankMessage = (rank == nil ? "Invalid rank code" : "")
    }
        
    func getRankList(text: String) -> [AutoCompleteData] {
        selected = nil
        return rankList.filter({Utility.wordSearch(for: text, in: "\($0.rankCode) \($0.rankName)")})
                .sorted(by: {$0.rankCode < $1.rankCode}).enumerated()
                .map({ (index, element) in
                    AutoCompleteData(index: index, code: "\(element.rankCode)", desc: element.rankName)})
    }
    
    private func rankKeyPress(_ press: KeyPress) -> KeyPress.Result{
        AutoComplete.onKeyPress(press, selected: $selected, maxSelected: rankData.count) {
            set(rankText: "\(rankData[selected!].code)")
        }
    }
    
    func nameMessage(index: Int, name: String) -> String {
        return (index == 0 || name != "" ? "" : "Name must not be blank")
    }
    
    func percentMessage(index: Int, percent: Float) -> String {
        if percent <= 0 {
            "Percentage must be > 0"
        } else if percent > 100 {
            "Percentage must be ≤ 100"
        } else {
            ""
        }
    }
}
    

struct LeadingClickableText : View {
    var text: String
    
    var body : some View {
        HStack {
            Text(text)
            Spacer()
        }
        .contentShape(Rectangle())
    }
}

struct CenteredClickableText : View {
    var text: String
    
    var body : some View {
        HStack {
            Spacer()
            Text(text)
            Spacer()
        }
        .contentShape(Rectangle())
    }
}
