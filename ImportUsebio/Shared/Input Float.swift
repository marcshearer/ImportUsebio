//  Input Float.swift
//  BridgeScore
//
//  Created by Marc Shearer on 07/02/2022.
//

import SwiftUI

struct InputFloat : View {
    
    var title: String?
    @Binding var field: Float
    var message: Binding<String>?
    var topSpace: CGFloat = 0
    var leadingSpace: CGFloat = 0
    var height: CGFloat = 40
    var width: CGFloat?
    var places: Int = 2
    var inlineTitle: Bool = true
    var inlineTitleWidth: CGFloat = 150
    var onChange: ((Float?)->())?
    
    @State private var refresh = false
    @State private var wrappedText = ""
    var text: Binding<String> {
        Binding {
            wrappedText
        } set: { (newValue) in
            wrappedText = newValue
        }
    }
    
    
    var body: some View {
        
        VStack(spacing: 0) {
            
            // Just to trigger view refresh
            if refresh { EmptyView() }
            
            if title != nil && !inlineTitle {
                HStack {
                    InputTitle(title: title, message: message, topSpace: topSpace)
                }
                Spacer().frame(height: 8)
            } else {
                Spacer().frame(height: topSpace)
            }
            HStack {
                Spacer().frame(width: leadingSpace)
                if title != nil && inlineTitle {
                    HStack {
                        Spacer().frame(width: 6)
                        Text(title!)
                        Spacer()
                    }
                    .frame(width: inlineTitleWidth)
                }

                HStack {
                    Spacer().frame(width: 4)
                    
                    UndoWrapper(text) { text in
                        TextField("", text: text, onEditingChanged: {(editing) in
                            text.wrappedValue = Float(text.wrappedValue)?.toString(places: places) ?? ""
                            field = Float(text.wrappedValue) ?? 0
                        })
                            .onSubmit {
                                text.wrappedValue = Float(text.wrappedValue)?.toString(places: places) ?? ""
                                field = Float(text.wrappedValue) ?? 0
                            }
                            .onChange(of: text.wrappedValue) { newValue in
                                let filtered = newValue.filter { "0123456789 -,.".contains($0) }
                                let oldField = field
                                if filtered != newValue {
                                    text.wrappedValue = filtered
                                }
                                field = Float(text.wrappedValue) ?? 0
                                if oldField != field {
                                    onChange?(field)
                                }
                            }
                            .lineLimit(1)
                            // .padding(.all, 1)
                            .disableAutocorrection(false)
                    }
                }
                .if(width != nil) { (view) in
                    view.frame(width: width)
                }
                .frame(height: height)
                .background(Palette.input.background)
                .cornerRadius(12)
    
                if width == nil {
                    Spacer()
                }
            }
            .font(inputFont)
            .onChange(of: field) { (field) in
                let newValue = field.toString(places: places)
                if newValue != wrappedText {
                    wrappedText = newValue
                }
            }
            .onAppear {
                text.wrappedValue = field.toString(places: places)
            }
        }
        .frame(height: self.height + ((self.inlineTitle ? 0 : self.topSpace) + (title == nil || inlineTitle ? 0 : 30)))
        .if(width != nil) { (view) in
            view.frame(width: width! + leadingSpace + (inlineTitle && title != nil ? inlineTitleWidth : 0))
        }
    }
}
