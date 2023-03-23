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
    var isEnabled: Bool
    var isReadOnly: Bool = false
    var onChange: ((String)->())?

    var body: some View {
        VStack(spacing: 0) {
            if title != nil && !inlineTitle {
                InputTitle(title: title, message: message, messageOffset: messageOffset, topSpace: topSpace, isEnabled: isEnabled)
                    .frame(width: width + (inlineTitle ? inlineTitleWidth : 0) + leadingSpace + 67)
                Spacer().frame(height: 8)
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
                } else {
                    Spacer().frame(width: 32)
                }
                ZStack(alignment: .leading){
                    HStack {
                        Rectangle()
                            .foregroundColor(Palette.input.background)
                            .cornerRadius(8)
                    }
                    if isEnabled && field.isEmpty {
                        VStack {
                            Spacer().frame(height: 10)
                            HStack {
                                Spacer().frame(width: 10)
                                Text(placeHolder)
                                    .foregroundColor(Palette.input.faintText)
                            }
                            Spacer()
                        }
                    }
                    HStack {
                        if secure {
                            VStack {
                                Spacer().frame(height: (MyApp.target == .macOS ? 2 :10))
                                SecureField("", text: $field)
                                    .font(inputFont)
                                    .onChange(of: field, perform: { (value) in onChange?(value) })
                                    .textFieldStyle(PlainTextFieldStyle())
                                    .disabled(!isEnabled || isReadOnly)
                                    .foregroundColor(isEnabled ? Palette.input.text : Palette.input.faintText)
                                    .inputStyle(width: width, height: inputDefaultHeight)
                                Spacer()
                            }
                        } else if height > inputDefaultHeight || !isEnabled || isReadOnly {
                            VStack {
                                if isEnabled && !isReadOnly {
                                    TextEditor(text: $field)
                                        .font(inputFont)
                                        .onChange(of: field, perform: { (value) in onChange?(value) })
                                        .disabled(!isEnabled || isReadOnly)
                                        .foregroundColor(isEnabled ? Palette.input.text : Palette.input.faintText)
                                        .inputStyle(width: width, height: height - (MyApp.target == .macOS ? 16 : 0), padding: 5.0)
                                        .myKeyboardType(self.keyboardType)
                                        .myAutocapitalization(autoCapitalize)
                                        .disableAutocorrection(!autoCorrect)
                                        .textCase(.uppercase)
                                } else {
                                    VStack {
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
                                }
                            }
                            .font(inputFont)
                            .frame(height: height)
                        } else {
                            TextField("", text: $field)
                                .font(inputFont)
                                .onChange(of: field, perform: { (value) in onChange?(value) })
                                .textFieldStyle(PlainTextFieldStyle())
                                .foregroundColor(Palette.input.text)
                                .background(Palette.input.background)
                                .inputStyle(width: width, height: height)
                                .myKeyboardType(self.keyboardType)
                                .myAutocapitalization(autoCapitalize)
                                .disableAutocorrection(!autoCorrect)
                                .frame(height: height)
                        }
                    }
                }
                .frame(width: width)
            }
            Spacer().frame(height: 8)
        }
        .frame(height: self.height + self.topSpace + (title == nil || inlineTitle ? 0 : 20))
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
