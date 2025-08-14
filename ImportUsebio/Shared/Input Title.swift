//
//  InputTitle.swift
// Bridge Score
//
//  Created by Marc Shearer on 14/02/2021.
//

import SwiftUI

struct InputTitle : View {
    
    @State var title: String?
    var message: Binding<String>? = nil
    var messageOffset: CGFloat = 0.0
    var topSpace: CGFloat = inputTopHeight
    var fillTrailing: Bool = true
    var buttonImage: AnyView?
    var buttonText: String?
    var buttonAction: (()->())?
    var isEnabled: Bool = true
    
    var body: some View {

        VStack {
            Spacer().frame(height: topSpace)
                
            if let title = self.title {
                HStack(alignment: .center, spacing: nil) {
                    Spacer().frame(width: 16)
                    Text(title).font(.headline).foregroundColor(isEnabled ? Palette.input.text : Palette.input.faintText)
                    
                    if let action = buttonAction {
                        Spacer().frame(width: 16)
                        Button {
                            action()
                        } label: {
                            if let image = buttonImage {
                                image
                            }
                            if let text = buttonText {
                                if buttonImage != nil {
                                    Spacer().frame(width: 8)
                                }
                                Text(text)
                                    .foregroundColor(Palette.background.themeText)
                            }
                        }
                        .menuStyle(DefaultMenuStyle())
                    }
                    if fillTrailing {
                        Spacer()
                    }
                    if let message = message?.wrappedValue {
                        VStack(spacing: 0) {
                            //Spacer().layoutPriority(.leastNonzeroMagnitude)
                            Text(message).foregroundColor(Palette.background.strongText).font(.caption)
                        }
                        Spacer().frame(width: 32 + messageOffset)
                    }
                }
            }
        }
        .frame(height: 12 + topSpace)
    }
}
