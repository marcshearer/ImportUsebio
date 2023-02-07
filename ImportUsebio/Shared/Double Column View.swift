//
//  Double Column View.swift
//  BridgeScore
//
//  Created by Marc Shearer on 05/02/2022.
//

import SwiftUI

struct DoubleColumnView <LeftContent, RightContent> : View where LeftContent : View, RightContent : View {
    var leftWidth: CGFloat
    var leftView: LeftContent
    var rightView: RightContent
    var separator: Bool = true
    
    init(leftWidth: CGFloat = 350, separator: Bool = true, @ViewBuilder leftView: ()->LeftContent, @ViewBuilder rightView: ()->RightContent) {
        self.separator =  separator
        self.leftWidth = leftWidth
        self.leftView = leftView()
        self.rightView = rightView()
    }
    var body: some View {
        HStack(spacing: 0) {
            VStack {
                leftView
                Spacer()
            }
            .frame(width: leftWidth)
            if separator {
                Separator(direction: .vertical, thickness: 2.0)
                    .foregroundColor(Palette.banner.background)
            }
            VStack {
                rightView
                Spacer()
            }
        }
    }
}
