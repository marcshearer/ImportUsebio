//
//  Slide In Menu.swift
//  Shortcut Menu
//
//  Created by Marc Shearer on 07/03/2021.
//  Copyright © 2021 Marc Shearer. All rights reserved.
//

import SwiftUI

class SlideInMenu : ObservableObject {
    
    public static let shared = SlideInMenu()
    
    @Published private(set) var title: String? = nil
    @Published public var options: [(text: String, id: AnyHashable)] = []
    @Published public var field: Int?
    @Published private(set) var selected: Flags?
    @Published private(set) var selectAll: String?
    @Published private(set) var top: CGFloat = 0
    @Published private(set) var left: CGFloat? = nil
    @Published fileprivate(set) var width: CGFloat = 0
    @Published private(set) var animation: ViewAnimation = .slideLeft
    @Published private(set) var selectedColor: PaletteColor?
    @Published private(set) var hideBackground: Bool = true
    @Published private(set) var completion: ((String?)->())?
    @Published private(set) var shown: UUID?
    
    public func show(id: UUID, title: String? = nil, options: [(String, AnyHashable)]? = nil, strings: [String] = [], selected: Flags? = nil, default field: Int? = nil, selectAll: String? = nil, animation: ViewAnimation = .slideLeft, top: CGFloat? = nil, left: CGFloat? = nil, width: CGFloat? = nil, selectedColor: PaletteColor? = nil, hideBackground: Bool = true, completion: ((String?)->())? = nil) {
        withAnimation(.none) {
            self.title = title
            self.options = options ?? strings.map{($0, $0)}
            self.field = field
            self.selected = selected
            self.selectAll = selectAll
            self.top = top ?? bannerHeight + 10
            self.left = left
            self.width = width ?? 300
            self.selectedColor = selectedColor
            self.hideBackground = hideBackground
            self.animation = animation
            self.completion = completion
            Utility.mainThread {
                self.shown = id
            }
        }
    }
    
    public func hide() {
        width = 300
        shown = nil
    }
    
    public func isSelected(_ index: Int) -> Bool {
        var result = false
        if let selected = selected {
            let id = options[index].id
            result = selected.value(id)
        } else {
            result = (index == field)
        }
        return result
    }
}

struct SlideInMenuView : View {
    @State var id: UUID
    @ObservedObject var values = SlideInMenu.shared
    @State private var offset: CGFloat = 320
    @State private var refresh = false
    
    var body: some View {
        
        if refresh { EmptyView() }
        
        GeometryReader { (fullGeometry) in
            GeometryReader { (geometry) in
                let (contentHeight, actualTop) = sizeToFit(geometry: geometry)
                ZStack {
                    Rectangle()
                        .foregroundColor(values.shown == id ? (values.hideBackground ? Palette.maskBackground : Palette.clickableBackground) : Color.clear)
                        .onTapGesture {
                            values.hide()
                        }
                        .frame(width: fullGeometry.size.width, height: fullGeometry.size.height + fullGeometry.safeAreaInsets.top + fullGeometry.safeAreaInsets.bottom)
                        .ignoresSafeArea()
                    VStack(spacing: 0) {
                        
                        Spacer().frame(height: actualTop)
                        HStack {
                            Spacer()
                            VStack(spacing: 0) {
                                if let title = values.title {
                                    tile(color: Palette.header, text: title, centered: true, heightFactor: 1.4)
                                    .font(.title)
                                }
                                
                                ScrollView {
                                    VStack(spacing: 0) {
                                        if let selectAll = values.selectAll {
                                            tile(color: Palette.background, text: selectAll,  bottomSeparator: true)
                                            .font(.title2)
                                            .onTapGesture {
                                                values.selected?.clear()
                                                values.completion?(nil)
                                                values.hide()
                                            }
                                        }
                                        ForEach(values.options.indices, id: \.self) { (index) in
                                            let option = values.options[index]
                                            let color = (values.isSelected(index) ? values.selectedColor ?? Palette.background : Palette.background)
                                            let textType: ThemeTextType = (values.isSelected(index) && values.selectedColor == nil ? .theme : .normal)
                                            tile(color: color, text: option.text, textType: textType)
                                            .font(.title2)
                                            .onTapGesture {
                                                if values.selected != nil {
                                                    values.selected?.toggle(option.id)
                                                    values.completion?(option.text)
                                                    refresh.toggle()
                                                } else {
                                                    values.completion?(option.text)
                                                    values.hide()
                                                }
                                            }
                                        }
                                        .listStyle(PlainListStyle())
                                    }
                                }
                                .background(Palette.background.background)
                                .environment(\.defaultMinListRowHeight, slideInMenuRowHeight)
                                .layoutPriority(.greatestFiniteMagnitude)
                                
                                if values.title != nil {
                                    tile(color: Palette.alternate, text: values.selected == nil ? "Cancel" : "Close", centered: true)
                                    .font(Font.title2.bold())
                                    .onTapGesture {
                                        values.hide()
                                    }
                                }
                            }
                            .frame(height: contentHeight)
                            .background(Palette.background.background)
                            .frame(width: values.width)
                            .cornerRadius(20)
                            Spacer().frame(width: 20)
                        }
                        Spacer()
                        
                    }
                    .shadow(color: Palette.maskBackground, radius: 2, x: 4, y: 4)
                    .offset(x: offset)
                }
                
                .if(offset != 0 && values.animation == .fade) { (view) in
                    view.hidden()
                }
                .onChange(of: values.shown, perform: { value in
                    offset = (values.shown == id ? (values.left == nil ? 0 : min(0, (values.left! + values.width - geometry.size.width))) : values.width + 20)
                })
                .animation(values.animation == .none || values.shown != id ? .none : .easeInOut, value: offset)
                .onAppear {
                    SlideInMenu.shared.width = 300
                }
            }
        }
    }
    
    func sizeToFit(geometry: GeometryProxy) -> (CGFloat, CGFloat) {
        let padElements: CGFloat = (values.title == nil ? 0 : 2.4) + (values.selectAll == nil ? 0 : 1)
        let availableHeight = geometry.size.height - bannerHeight - 8
        let maxFit = CGFloat(Int((availableHeight / slideInMenuRowHeight) - padElements)) - 0.4
        let allowedElements = min(maxFit, CGFloat(values.options.count))
        let contentHeight = (allowedElements + padElements) * slideInMenuRowHeight
        let proposedTop = values.top
        let top = min(proposedTop,
                      max(bannerHeight + 8,
                          geometry.size.height - contentHeight))
        return (contentHeight, top)
    }
    
    func tile(color: PaletteColor, text: String, centered: Bool = false, textType: ThemeTextType = .normal, heightFactor: CGFloat = 1, bottomSeparator: Bool = false) -> some View {
        VStack(spacing: 0) {
            Spacer()
            HStack {
                if centered {
                    Spacer()
                } else {
                    Spacer().frame(width: 20)
                }
                Text(text)
                    .foregroundColor(color.textColor(textType))
                Spacer()
            }
            Spacer()
            if bottomSeparator {
                Separator(thickness: 2)
            }
        }
        .background(color.background)
        .frame(height: slideInMenuRowHeight * heightFactor)
    }
    
}
