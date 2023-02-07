//
//  List Tile View.swift
//  BridgeScore
//
//  Created by Marc Shearer on 08/02/2022.
//

import SwiftUI

struct ListTileView <Content> : View where Content : View {
    var color: PaletteColor
    var font: Font
    var content: Content
    var height: CGFloat = 92
    
    init(color: PaletteColor = Palette.tile, font: Font = .largeTitle, height: CGFloat = 92, @ViewBuilder content: ()->Content) {
        self.color = color
        self.font = font
        self.height = height
        self.content = content()
    }

    var body: some View {
        VStack {
            Spacer().frame(height: 4)
            HStack {
                Spacer().frame(width: 16)
                HStack {
                    Spacer().frame(width: 16)
                    content
                    Spacer()
                }
                .frame(height: height - 12)
                .background(color.background)
                .cornerRadius(16)
            Spacer().frame(width: 16)
            }
            Spacer().frame(height: 8)
        }
        .foregroundColor(color.text)
        .font(font)
    }
}
