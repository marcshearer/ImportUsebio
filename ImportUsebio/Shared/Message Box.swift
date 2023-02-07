//
//  Miscellanous.swift
// Bridge Score
//
//  Created by Marc Shearer on 02/03/2021.
//

import SwiftUI

class MessageBox : ObservableObject {
    
    public static let shared = MessageBox()
    
    @Published fileprivate var text: String?
    fileprivate var okText: String?
    fileprivate var okDestructive: Bool = false
    fileprivate var cancelText: String?
    fileprivate var showIcon: Bool = false
    fileprivate var showVersion = false
    fileprivate var okAction: (()->())? = nil
    fileprivate var cancelAction: (()->())? = nil

    public var isShown: Bool { MessageBox.shared.text != nil }
    
    public func show(_ text: String, if showMessage: Bool = true,  cancelText: String? = nil, okText: String? = "Close", okDestructive: Bool = true, showIcon: Bool = false, showVersion: Bool = false, cancelAction: (()->())? = nil, okAction: (()->())? = nil) {
        if showMessage {
            MessageBox.shared.text = text
            MessageBox.shared.okText = okText
            MessageBox.shared.okDestructive = okDestructive
            MessageBox.shared.cancelText = cancelText
            MessageBox.shared.showIcon = showIcon
            MessageBox.shared.showVersion = showVersion
            MessageBox.shared.okAction = okAction
            MessageBox.shared.cancelAction = cancelAction
        } else {
            okAction?()
        }
    }
    
    public func hide() {
        MessageBox.shared.text = nil
    }
}

struct MessageBoxView: View {
    @Environment(\.presentationMode) var presentationMode: Binding<PresentationMode>
    @State var suppressIcon: Bool = false
    @ObservedObject var values = MessageBox.shared
    @State var showIcon = true

    var body: some View {
        ZStack {
            Palette.background.background
                .ignoresSafeArea(edges: .all)
            HStack(spacing: 0) {
                if values.showIcon && !suppressIcon {
                    Spacer().frame(width: 30)
                    VStack {
                        Spacer()
                        ZStack {
                            Rectangle().foregroundColor(Palette.alternate.background).frame(width: 80, height: 80).cornerRadius(10)
                            VStack {
                                Spacer()
                                HStack {
                                    Spacer()
                                    Image(appImage).resizable().frame(width: 80, height: 80)
                                    Spacer()
                                }
                                Spacer()
                            }
                        }
                        Spacer()
                    }
                    .frame(width: 80)
                }
                Spacer()
                VStack(alignment: .center) {
                    Spacer()
                    Text(appName).font(.largeTitle).minimumScaleFactor(0.75)
                    if values.showVersion {
                        Text("Version \(Version.current.version) (\(Version.current.build)) \(MyApp.database.name.capitalized)").minimumScaleFactor(0.5)
                    }
                    if let message = $values.text.wrappedValue {
                        Spacer().frame(height: 30)
                        Text(message).multilineTextAlignment(.center).fixedSize(horizontal: false, vertical: true).font(.callout).minimumScaleFactor(0.5)
                    }
                    Spacer().frame(height: 30)
                    HStack {
                        if let okText = values.okText {
                            Button {
                                $values.text.wrappedValue = nil
                                values.okAction?()
                            } label: {
                                Text(okText)
                                    .foregroundColor(values.okDestructive ? Palette.destructiveButton.text : Palette.highlightButton.text)
                                    .font(.callout).minimumScaleFactor(0.5)
                                    .frame(width: 100, height: 30)
                                    .background(values.okDestructive ? Palette.destructiveButton.background : Palette.highlightButton.background)
                                    .cornerRadius(15)
                            }
                        }
                        
                        if let cancelText = values.cancelText {
                            if values.okText != nil {
                                Spacer().frame(width: 30)
                            }
                            Button {
                                $values.text.wrappedValue = nil
                                values.cancelAction?()
                            } label: {
                                Text(cancelText)
                                    .foregroundColor(Palette.highlightButton.text)
                                    .font(.callout).minimumScaleFactor(0.5)
                                    .frame(width: 100, height: 30)
                                    .background(Palette.highlightButton.background)
                                    .cornerRadius(15)
                            }
                        }
                        
                        if values.okText == nil && values.cancelText == nil {
                            Text("").frame(width: 100, height: 30)
                        }
                    }
                    Spacer()
                }
                Spacer()
            }
        }
    }
}
