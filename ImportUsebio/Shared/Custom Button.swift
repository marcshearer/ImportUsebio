//
//  Custom Button.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/02/2026.
//

import SwiftUI

struct CustomButton {
    
    static func button(image: String? = nil, resizeImage: Bool = false, title: String, help: String? = nil, width: CGFloat = 100, height: CGFloat = 30, cornerRadius: CGFloat = 15, color: PaletteColor = Palette.highlightButton, enabled: ()->Bool = {true}, action: @escaping ()->()) -> some View {
        let color = (enabled() ? color : Palette.disabledButton)
        
        return Button {
            action()
        } label: {
            HStack(spacing: 0) {
                if let image = image {
                    Group {
                        if resizeImage {
                            Image(systemName: image)
                                .resizable()
                                .scaledToFit()
                        } else {
                            Image(systemName: image)
                        }
                    }
                    .padding(EdgeInsets(top: height / 6, leading: 2, bottom: height / 6, trailing: 3))
                }
                Text(title)
                    .padding(EdgeInsets(top: 3, leading: 3, bottom: 3, trailing: 2))
            }
            .foregroundColor(color.text)
            .frame(width: width, height: height)
            .font(inputFont)
            .minimumScaleFactor(0.5)
            .background(color.background)
            .cornerRadius(cornerRadius)
        }
        .help(help ?? "")
        .disabled(!enabled())
        .focusable(false)
        .buttonStyle(PlainButtonStyle())
    }
}

