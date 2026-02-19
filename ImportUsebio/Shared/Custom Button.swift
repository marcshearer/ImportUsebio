//
//  Custom Button.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 12/02/2026.
//

import SwiftUI

struct CustomButton {
    
    static func button(image: String? = nil, title: String, width: CGFloat = 100, height: CGFloat = 30, cornerRadius: CGFloat = 15, enabled: ()->Bool = {true}, action: @escaping ()->()) -> some View {
        
        return Button {
            action()
        } label: {
            HStack(spacing: 0) {
                if let image = image {
                    Image(systemName: image)
                    Spacer().frame(width: 10)
                }
                Text(title)
            }
            .foregroundColor(Palette.highlightButton.text)
            .frame(width: width, height: height)
            .font(inputFont)
            .minimumScaleFactor(0.5)
            .background(Palette.highlightButton.background)
            .cornerRadius(cornerRadius)
        }
        .disabled(!enabled())
        .focusable(false)
        .buttonStyle(PlainButtonStyle())
    }
}
