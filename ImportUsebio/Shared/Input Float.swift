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
    var height: CGFloat = inputDefaultHeight
    var width: CGFloat?
    var places: Int = 2
    var inlineTitle: Bool = true
    var inlineTitleWidth: CGFloat = 150
    var onChange: ((Float?)->())?
    @FocusState private var focus: Bool?

    
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
                        Text(title!)
                            .font(inputFont)
                        Spacer()
                    }
                    .frame(width: inlineTitleWidth)
                }

                HStack {
                    UndoWrapper(text) { text in
                        TextField("", text: text)
                            .onSubmit {
                                text.wrappedValue = Float(text.wrappedValue)?.toString(places: places) ?? ""
                                field = Utility.round(Float(text.wrappedValue) ?? 0, places: places)
                            }
                            .onChange(of: text.wrappedValue, initial: false) { (_, newValue) in
                                let filtered = newValue.filter { "0123456789 -,.".contains($0) }
                                let oldField = field
                                if filtered != newValue {
                                    text.wrappedValue = filtered
                                }
                                field = Utility.round(Float(text.wrappedValue) ?? 0, places: places)
                                if oldField != field {
                                    onChange?(field)
                                }
                            }
                            .onChange(of: focus) {
                                text.wrappedValue = Float(text.wrappedValue)?.toString(places: places) ?? ""
                                field = Utility.round(Float(text.wrappedValue) ?? 0, places: places)
                            }
                            .focused($focus, equals: true)
                            .lineLimit(1)
                            .padding(.all, 10)
                            .disableAutocorrection(false)
                            .textFieldStyle(.plain)
                    }
                }
                .if(width != nil) { (view) in
                    view.frame(width: width)
                }
                .frame(height: height)
                .background(Palette.input.background)
                .cornerRadius(inputCornerRadius)
    
                if width == nil || !inlineTitle {
                    Spacer()
                }
            }
            .font(inputFont)
            .onChange(of: field, initial: false) { (_, field) in
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
    }
}
