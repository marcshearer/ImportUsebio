//
//  Blocked numbers.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 11/02/2026.
//

import SwiftUI

enum EditMode {
    case display
    case amend
    case remove
    case add
    
    var title: String { "\(self)".capitalized }
    var action: String { self == .amend ? "Save" : title }
    var update: String { self == .amend ? "Update" : title }
    var enabled: Bool { self == .amend || self == .add}
}

struct BlockedNumbersView : View {
    @State var selected = BlockedViewModel()
    @ObservedObject var edit = BlockedViewModel()
    @State var exit: Bool = true
    @State var editMode: EditMode? = nil
    var showDetail: Binding<Bool> {
        Binding {
            editMode != nil
        } set: { _ in
        }
    }
    
    var body: some View {
        let blockedList = (MasterData.shared.blocked.array as! [BlockedViewModel]).sorted(by: { BlockedViewModel.defaultSort($0, $1)})
        
        StandardView("Select Input") {
            VStack(spacing: 0) {
                Banner(title: Binding.constant("Blocked National ID Numbers"), backEnabled: { exit }, backAction: {
                    return true
                })
                HStack {
                    Spacer().frame(width: 20)
                    VStack {
                        Spacer().frame(height: 20)
                        VStack(spacing: 0) {
                            HStack {
                                Spacer().frame(width: 10)
                                HStack {
                                    Spacer()
                                    Text("National ID")
                                }
                                .frame(width: 90)
                                Spacer().frame(width: 30)
                                Text("Reason")
                                    .lineLimit(1)
                                    .truncationMode(.tail)
                                Spacer()
                            }
                            .frame(height: 25)
                            .palette(.contrastTile)
                            ScrollView {
                                ScrollViewReader { scrollViewProxy in
                                    LazyVStack(spacing: 0) {
                                        ForEach(blockedList, id: \.nationalId) { (blocked) in
                                            VStack(spacing: 0) {
                                                HStack {
                                                    Spacer().frame(width: 10)
                                                    HStack {
                                                        Spacer()
                                                        Text(blocked.nationalId)
                                                    }
                                                    .frame(width: 90)
                                                    Spacer().frame(width: 30)
                                                    Text(blocked.reason)
                                                        .lineLimit(1)
                                                        .truncationMode(.tail)
                                                    Spacer()
                                                }
                                                .frame(height: 25)
                                                .palette(selected.nationalId == blocked.nationalId ? .highlightTile : .tile)
                                                Separator(thickness: 1.0)
                                            }
                                            .id(blocked.blockedId)
                                            .onTapGesture {
                                                selected = blocked
                                                edit.copy(from: blocked)
                                                editMode = .amend
                                            }
                                        }
                                    }
                                    .onChange(of: selected.nationalId) {
                                        if !selected.nationalId.isEmpty {
                                            scrollViewProxy.scrollTo(selected.blockedId, anchor: nil)
                                        }
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
                                edit.copy(from: BlockedViewModel())
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
            BlockedDetailView(edit: edit, editMode: editMode!, completion: { (success) in
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

struct BlockedDetailView: View {
    @Environment(\.dismiss) var dismiss
    @StateObject var edit: BlockedViewModel
    @State var editMode: EditMode
    @State var completion: (Bool)->() = { (success) in }
    
    var body: some View {
        
        StandardView("Blocked Detail") {
            
            VStack(spacing: 0) {
                
                Banner(title: Binding.constant("\(editMode.title) Blocked National ID"), alternateStyle: true, back: false)
                
                Spacer().frame(height: 20)
                
                HStack {
                    Input(title: "National Id:", field: $edit.nationalId, message: $edit.nationalIdMessage, leadingSpace: 20, width: 450, isReadOnly: !editMode.enabled, limitText: 16)
                    Spacer()
                }
                
                Spacer().frame(height: 30)
                
                HStack {
                    Input(title: "Reason:", field: $edit.reason, message: $edit.reasonMessage, leadingSpace: 20, width: 450, isReadOnly: !editMode.enabled)
                    Spacer()
                }
                
                Spacer().frame(height: 30)
                
                HStack {
                    InputTitle(title: "Warning or Total Block:")
                    Spacer()
                }
                Spacer().frame(height: 10)
                HStack {
                    Picker("", selection: $edit.warnOnly) {
                        Text("Warning Only").tag(true)
                        Text("Total Block").tag(false)
                    }
                    .pickerStyle(.segmented)
                    .frame(width: 300)
                    Spacer()
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
                            let insert = BlockedViewModel()
                            insert.copy(from: edit, copyMO: false)
                            insert.insert()
                            edit.copy(from: insert)
                        default:
                            break
                        }
                        completion(true)
                        dismiss()
                    }
                    .disabled($edit.nationalId.wrappedValue.isEmpty || $edit.reason.wrappedValue.isEmpty)
                    
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
        .frame(width: 600, height: 350)
        .palette(.background)
    }
}

