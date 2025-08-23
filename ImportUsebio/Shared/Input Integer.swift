//
//  Input Number.swift
//  BridgeScore
//
//  Created by Marc Shearer on 07/02/2022.
//

import SwiftUI

struct InputInt : View {
    
    var title: String?
    @Binding var field: Int
    var message: Binding<String>?
    var messageOffset: CGFloat = 0
    var topSpace: CGFloat = inputTopHeight
    var leadingSpace: CGFloat = 0
    var height: CGFloat = inputDefaultHeight
    var width: CGFloat = 80
    var inlineTitle: Bool = true
    var inlineTitleWidth: CGFloat = 150
    var maxValue: Int?
    var isEnabled: Bool = true
    var isReadOnly: Bool = false
    var pickerAction: (()->())?
    var onKeyPress: ((KeyPress)->(KeyPress.Result))?
    var detectKeys: Set<KeyEquivalent>?
    var onChange: ((Int)->())?
    
    @State private var wrappedText = ""
    var text: Binding<String> {
        Binding {
            wrappedText
        } set: { (newValue) in
            wrappedText = newValue
        }
    }
    
    @State private var refresh = false
    
    var body: some View {
        let pickerWidth: CGFloat = (pickerAction == nil ? 0 : inputDefaultHeight * 0.95)
        VStack(spacing: 0) {
            
                // Just to trigger view refresh
            if refresh { EmptyView() }
            
            if title != nil && !inlineTitle {
                HStack {
                    InputTitle(title: title, message: message, messageOffset: messageOffset, topSpace: topSpace, isEnabled: isEnabled)
                }
                Spacer().frame(height: 8)
            } else {
                Spacer().frame(height: topSpace)
            }
            
            HStack {
                if title != nil && inlineTitle {
                    Spacer().frame(width: leadingSpace)
                    HStack {
                        Text(title!)
                            .font(inputFont)
                        Spacer()
                    }
                    .frame(width: inlineTitleWidth)
                }
                ZStack(alignment: .leading){
                    HStack {
                        Rectangle()
                            .foregroundColor(Palette.input.background)
                            .cornerRadius(inputCornerRadius)
                    }
                    if inlineTitle || title == nil {
                        if let message = message?.wrappedValue {
                            HStack {
                                Spacer().frame(width: messageOffset)
                                HStack {
                                    Text(message).foregroundColor(Palette.input.themeText)
                                        .truncationMode(.tail)
                                    Spacer()
                                }
                                .frame(width: width - messageOffset - pickerWidth)
                            }
                        }
                    }
                    if let pickerAction = pickerAction, isEnabled {
                        HStack {
                            Spacer()
                            pickerButton(width: pickerWidth, height: height, pickerAction: pickerAction)
                                .frame(width: pickerWidth, height: height)
                        }
                    }
                    HStack {
                        UndoWrapper(text) { text in
                            TextField("", text: text, onEditingChanged: { (editing) in
                                valueChanged(oldText: field.toString(), newText: text.wrappedValue)
                            })
                            .onSubmit {
                                valueChanged(oldText: field.toString(), newText: text.wrappedValue)
                            }
                            .onChange(of: text.wrappedValue, initial: false) { (oldValue, newValue) in
                                valueChanged(oldText: oldValue, newText: newValue)

                            }
                            .textFieldStyle(PlainTextFieldStyle())
                            .disabled(!isEnabled || isReadOnly)
                            .foregroundColor(isEnabled ? Palette.input.text : Palette.input.faintText)
                            .lineLimit(1)
                            .padding(.all, 1)
                            .disableAutocorrection(false)
                            .textFieldStyle(.plain)
                            .if(detectKeys != nil) { (view) in
                                view.onKeyPress(keys: detectKeys!) { press in
                                    return (onKeyPress?(press) ?? .ignored)
                                }
                            }
                        }
                    }
                    .frame(width: (messageOffset != 0 ? messageOffset : width - pickerWidth), height: height)
                    .background(Palette.input.background)
                    .cornerRadius(inputCornerRadius)
                    
                    if !inlineTitle {
                        Spacer()
                    }
                }
                .frame(width: width)
            }
            .frame(height: self.height)
        }
        .font(inputFont)
        .onAppear {
            text.wrappedValue = field.toString()
        }
        .onChange(of: field, initial: true) {
            if text.wrappedValue != "" || field != 0 {
                valueChanged(oldText: text.wrappedValue, newText: field.toString())
            }
        }
    }
    
    func valueChanged(oldText: String, newText: String) {
        let newValue = (Int(newText) ?? 0)
        let oldValue = field
        if maxValue != nil && newValue > maxValue! {
            text.wrappedValue = field.toString()
        } else {
            if newText != "" {
                text.wrappedValue = newValue.toString()
            }
            field = newValue
            onChange?(newValue)
        }
    }
    
    func pickerButton(width: CGFloat, height: CGFloat, pickerAction: @escaping ()->()) -> some View {
        VStack {
            let spacing: CGFloat = 2.0
            Spacer().frame(height: spacing * 2)
            HStack {
                Spacer().frame(width: spacing)
                HStack {
                    Spacer()
                    Image(systemName: "chevron.up.chevron.down")
                        .frame(width: width - (2 * spacing), height: height - (2 * spacing))
                        .background(Palette.pickerButton.background)
                        .foregroundColor(Palette.pickerButton.text)
                        .bold()
                        .cornerRadius(inputPickerCornerRadius)
                        .font(pickerFont)
                    Spacer()
                }
                Spacer().frame(width: spacing)
            }
            Spacer().frame(height: spacing * 2)
        }
        .onTapGesture {
            pickerAction()
        }
    }}
