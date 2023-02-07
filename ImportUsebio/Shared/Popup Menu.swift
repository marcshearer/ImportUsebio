//
//  Popup Menu.swift
//  BridgeScore
//
//  Created by Marc Shearer on 01/03/2022.
//

import SwiftUI

struct PopupMenu<Label>: View where Label : View {
    var id: UUID
    var field: Binding<Int?>?
    var selected: Flags?
    let label: Label
    let values: [(text: String, id: AnyHashable)]
    var selectAll: String?
    var top: CGFloat?
    var left: CGFloat?
    var width: CGFloat?
    var animation: ViewAnimation
    var title: String?
    var selectedColor: PaletteColor?
    var hideBackground: Bool
    let onChange: ((Int?)->())?
    @State private var showPopup = false
    
    init(id: UUID, field: Binding<Int?>, values: [String], title: String? = nil, animation: ViewAnimation = .slideLeft, top: CGFloat? = nil, left: CGFloat? = nil, width: CGFloat, selectedColor: PaletteColor? = nil, hideBackground: Bool = true, onChange: ((Int?)->())? = nil, @ViewBuilder label: ()->Label) {
        self.id = id
        self.field = field
        self.values = values.map{($0, $0)}
        self.title = title
        self.top = top
        self.left = left
        self.animation = animation
        self.width = width
        self.selectedColor = selectedColor
        self.onChange = onChange
        self.hideBackground = hideBackground
        self.label = label()
    }
    
    init(id: UUID, selected: Flags, values: [(String, AnyHashable)], title: String? = nil, selectAll: String? = nil, animation: ViewAnimation = .slideLeft, top: CGFloat? = nil, left: CGFloat? = nil, width: CGFloat, selectedColor: PaletteColor? = nil, hideBackground: Bool = true, onChange: ((Int?)->())? = nil, @ViewBuilder label: ()->Label) {
        self.id = id
        self.selected = selected
        self.values = values
        self.title = title
        self.selectAll = selectAll
        self.top = top
        self.left = left
        self.animation = animation
        self.width = width
        self.selectedColor = selectedColor
        self.hideBackground = hideBackground
        self.onChange = onChange
        self.label = label()
    }
    
    public var body: some View {
        label
            .onTapGesture {
                SlideInMenu.shared.show(id: id, title: title, options: values, selected: selected, default: field?.wrappedValue, selectAll: selectAll, animation: animation, top: top, left: left, width: width, selectedColor: selectedColor, hideBackground: hideBackground, completion: { (selected) in
                    if let index = values.firstIndex(where: {$0.text == selected}) {
                        if self.selected != nil {
                            onChange?(index)
                        } else {
                            field?.wrappedValue = index
                            if let value = field?.wrappedValue {
                                onChange?(value)
                            }
                        }
                    } else {
                        onChange?(nil)
                    }
                })
            }
    }
    
}
