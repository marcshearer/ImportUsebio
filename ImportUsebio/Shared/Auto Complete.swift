//
//  Auto Complete.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 19/02/2026.
//

import SwiftUI

struct AutoCompleteData : Hashable {
    var index: Int
    var code: String
    var desc: String
}

class AutoComplete {
    
    static let detectKeys: Set<KeyEquivalent> = [.upArrow, .downArrow, .return]
    
    static func view<ID:Hashable>(autoComplete: Namespace.ID, field: ID, selected: Binding<Int?>, codeWidth: CGFloat, data: Binding<[AutoCompleteData]>, valid: Bool, selectAction: @escaping (String) -> ()) -> some View {
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
                            .background(element.index != selected.wrappedValue ? Palette.autoComplete.background : Palette.autoCompleteSelected.background)
                            .foregroundColor(element.index != selected.wrappedValue ? Palette.autoComplete.text : Palette.autoCompleteSelected.text)
                        }
                    }
                    .scrollTargetLayout()
                    .listStyle(DefaultListStyle())
                }
                .scrollPosition(id: selected)
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
    
    static func onKeyPress(_ keyPress: KeyPress, selected: Binding<Int?>, maxSelected: Int, onSelect: ()->()) -> KeyPress.Result {
        switch keyPress.key {
        case .downArrow:
            selected.wrappedValue = min(maxSelected - 1, (selected.wrappedValue ?? -1) + 1)
            return .handled
        case .upArrow:
            selected.wrappedValue = ((selected.wrappedValue ?? 0) == 0 ? nil : selected.wrappedValue! - 1)
            return .handled
        case .return:
            if selected.wrappedValue != nil {
                onSelect()
            }
            return .handled
        default:
            return .ignored
        }
    }
}
