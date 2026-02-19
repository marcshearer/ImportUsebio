    //
    // Input String.swift
    // Bridge Score
    //
    //  Created by Marc Shearer on 10/02/2021.
    //

import SwiftUI

struct Input : View {
        
    var title: String?
    @Binding var field: String
    var message: Binding<String>?
    var messageOffset: CGFloat = 0.0
    var desc: Binding<String>?
    var descOffset: CGFloat = 0.0
    var placeHolder: String = ""
    var secure: Bool = false
    var topSpace: CGFloat = inputTopHeight
    var leadingSpace: CGFloat = 0
    var height: CGFloat = inputDefaultHeight
    var width: CGFloat = 1000
    var inlineTitle: Bool = false
    var inlineTitleWidth: CGFloat = 150
    var keyboardType: KeyboardType = .default
    var autoCapitalize: AutoCapitalization = .sentences
    var autoCorrect: Bool = true
    var isEnabled: Bool = true
    var isReadOnly: Bool = false
    var limitText: Int? = nil
    var pickerAction: (()->())?
    var onKeyPress: ((KeyPress)->(KeyPress.Result))?
    var detectKeys: Set<KeyEquivalent>?
    var onChange: ((String)->())?

    var body: some View {
        let pickerWidth: CGFloat = (pickerAction == nil ? 0 : inputDefaultHeight * 0.95)
        VStack(spacing: 0) {
            if title != nil && !inlineTitle {
                InputTitle(title: title, message: message, messageOffset: messageOffset, topSpace: topSpace, isEnabled: isEnabled)
                    .frame(width: width + (inlineTitle ? inlineTitleWidth : 0) + leadingSpace + 67)
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
                ZStack(alignment: .leading){
                    HStack {
                        Rectangle()
                            .foregroundColor(Palette.input.background)
                            .cornerRadius(inputCornerRadius)
                    }
                    .if(pickerAction != nil && isReadOnly && isEnabled) { (view) in
                        view.onTapGesture {
                            pickerAction?()
                        }
                    }
                    if inlineTitle || title == nil {
                        if let desc = desc?.wrappedValue {
                            HStack {
                                Spacer().frame(width: descOffset)
                                HStack {
                                    Text(desc)
                                        .font(lookupFont)
                                        .foregroundColor(Palette.input.themeText)
                                        .truncationMode(.tail)
                                    Spacer()
                                }
                                .frame(width: width - descOffset - pickerWidth)
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
                        if secure {
                            VStack {
                                Spacer().frame(height: (MyApp.target == .macOS ? 2 :10))
                                SecureField("", text: $field)
                                    .font(inputFont)
                                    .onChange(of: field, initial: false) { (_, value) in
                                        if let limitText = limitText, value.count > limitText {
                                            field = String(value.prefix(limitText))
                                        }
                                        onChange?(value)
                                    }
                                    .textFieldStyle(PlainTextFieldStyle())
                                    .disabled(!isEnabled || isReadOnly)
                                    .foregroundColor(isEnabled ? Palette.input.text : Palette.input.faintText)
                                    .inputStyle(width: width, height: inputDefaultHeight)
                                    .frame(width: descOffset == 0 ? width - pickerWidth : descOffset)
                                    .if(detectKeys != nil) { (view) in
                                        view.onKeyPress(keys: detectKeys!) { press in
                                            return (onKeyPress?(press) ?? .ignored)
                                        }
                                    }
                                Spacer()
                            }
                        } else if height > inputDefaultHeight {
                            VStack {
                                if isEnabled && !isReadOnly {
                                    TextEditor(text: $field)
                                        .font(inputFont)
                                        .onChange(of: field, initial: false)
                                    { (_, value) in
                                        if let limitText = limitText, value.count > limitText {
                                            field = String(value.prefix(limitText))
                                        }
                                        onChange?(field)
                                    }
                                    .disabled(!isEnabled || isReadOnly)
                                    .foregroundColor(isEnabled ? Palette.input.text : Palette.input.faintText)
                                    .inputStyle(width: width, height: height - (MyApp.target == .macOS ? 16 : 0), padding: 5.0)
                                    .myKeyboardType(self.keyboardType)
                                    .myAutocapitalization(autoCapitalize)
                                    .disableAutocorrection(!autoCorrect)
                                    .textCase(.uppercase)
                                    .frame(width: descOffset == 0 ? width - pickerWidth : descOffset)
                                    .if(detectKeys != nil) { (view) in
                                        view.onKeyPress(keys: detectKeys!) { press in
                                            return (onKeyPress?(press) ?? .ignored)
                                        }
                                    }
                                } else {
                                    VStack {
                                        Spacer()
                                        HStack {
                                            Spacer().frame(width: 10)
                                            Text(field)
                                                .foregroundColor(Palette.input.faintText)
                                                .font(inputFont)
                                                .frame(height: height)
                                            Spacer()
                                        }
                                        Spacer()
                                    }
                                    .if(pickerAction != nil && isReadOnly && isEnabled) { (view) in
                                        view.onTapGesture {
                                            pickerAction?()
                                        }
                                    }
                                    .frame(width: descOffset == 0 ? width - pickerWidth : descOffset)
                                }
                            }
                            .font(inputFont)
                            .frame(height: height)
                        } else {
                            TextField("", text: $field)
                                .font(inputFont)
                                .onChange(of: field, initial: false) { (_, value) in
                                    if let limitText = limitText, value.count > limitText {
                                        field = String(value.prefix(limitText))
                                    }
                                    onChange?(value)
                                }
                                .textFieldStyle(PlainTextFieldStyle())
                                .foregroundColor(Palette.input.text)
                                .background(Palette.input.background)
                                .inputStyle(width: width, height: height)
                                .myKeyboardType(self.keyboardType)
                                .myAutocapitalization(autoCapitalize)
                                .disableAutocorrection(!autoCorrect)
                                .disabled(!isEnabled || isReadOnly)
                                .frame(height: height)
                                .if(detectKeys != nil) { (view) in
                                    view.onKeyPress(keys: detectKeys!) { press in
                                        return (onKeyPress?(press) ?? .ignored)
                                    }
                                }
                                .frame(width: descOffset == 0 ? width - pickerWidth : descOffset)
                            Spacer()
                        }
                    }
                    if field.isEmpty {
                        VStack {
                            HStack {
                                Spacer().frame(width: 10)
                                Text(placeHolder)
                                    .font(inputFont)
                                    .foregroundColor(Palette.input.faintText)
                            }
                        }
                        .if(pickerAction != nil && isReadOnly && isEnabled) { (view) in
                            view.onTapGesture {
                                pickerAction?()
                            }
                        }
                    }
                }
                .frame(width: width)
            }
        }
        .frame(height: height + topSpace + (title == nil || inlineTitle ? 0 : 20))
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
    }
}

struct InputViewModifier : ViewModifier {
    @State var width: CGFloat = 0.0
    @State var height: CGFloat = 0.0
    @State var padding: CGFloat = 10.0

    func body(content: Content) -> some View { content
        .frame(height: height)
        .padding([.leading, .trailing], padding)
        .frame(maxWidth: width)
    }
}

extension View {
    fileprivate func inputStyle(width: CGFloat = 0, height: CGFloat = 0, padding: CGFloat = 10.0) -> some View {
        self.modifier(InputViewModifier(width: width, height: height, padding: padding))
    }
}

#if os(macOS)
extension NSTextView {
  open override var frame: CGRect {
    didSet {
        backgroundColor = .clear
        drawsBackground = false
        isRulerVisible = false
    }
  }
}

extension NSTextField {
  open override var frame: CGRect {
    didSet {
        backgroundColor = .clear
        drawsBackground = true
        isBordered = false
        isBezeled = false
        focusRingType = .none
    }
  }
}

extension NSSecureTextField {
  open override var frame: CGRect {
    didSet {
        backgroundColor = .clear
        drawsBackground = true
        isBordered = false
        isBezeled = false
        focusRingType = .none
    }
  }
}
#endif
