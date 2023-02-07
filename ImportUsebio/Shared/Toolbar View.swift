//
//  Toolbar View.swift
//  BridgeScore
//
//  Created by Marc Shearer on 06/02/2022.
//

import SwiftUI

struct ToolbarView : View {
    public var canAdd: (()->(Bool))?
    public var canRemove: (()->(Bool))?
    public var addAction: ()->()
    public var removeAction: ()->()
    
    var body: some View {
        ZStack {
            Rectangle()
                .fill(Palette.alternate.background)
                .ignoresSafeArea()
            HStack {
                Spacer().frame(width: 20)
                Button {
                    addAction()
                } label: {
                    Image(systemName: "plus")
                }
                .opacity(!(canAdd?() ?? true) ? 0.3 : 1.0)
                .disabled(!(canAdd?() ?? true))
                
                Spacer().frame(width: 20)
                Button {
                    removeAction()
                } label: {
                    Image(systemName: "minus")
                }
                .opacity(!(canRemove?() ?? true) ? 0.3 : 1.0)
                .disabled(!(canRemove?() ?? true))
                Spacer()
            }
            .font(.title)
            .foregroundColor(Palette.alternate.text)
        }
        .frame(height: 50)
    }
}
