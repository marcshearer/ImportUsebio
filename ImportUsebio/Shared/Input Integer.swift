//
//  Input Number.swift
//  BridgeScore
//
//  Created by Marc Shearer on 07/02/2022.
//

import SwiftUI

struct InputInt : View {
    
    var title: String?
    var field: Binding<Int>
    var message: Binding<String>?
    var topSpace: CGFloat = inputTopHeight
    var leadingSpace: CGFloat = 32
    var height: CGFloat = inputDefaultHeight
    var width: CGFloat?
    var inlineTitle: Bool = true
    var inlineTitleWidth: CGFloat = 150
    var onChange: ((Int)->())?
    
    @State private var refresh = false
    @State private var text: String = "0"
    
    var body: some View {
        
        VStack(spacing: 0) {
            
            // Just to trigger view refresh
            if refresh { EmptyView() }
            
            if title != nil && !inlineTitle {
                HStack {
                    InputTitle(title: title, message: message, topSpace: topSpace)
                }
                Spacer().frame(height: 8)
            } else {
                Spacer().frame(height: topSpace)
            }
            
            HStack {
                Spacer().frame(width: leadingSpace)
                if title != nil && inlineTitle {
                    HStack {
                        Spacer().frame(width: 6)
                        Text(title!)
                        Spacer()
                    }
                    .frame(width: inlineTitleWidth)
                }
                
                HStack {
                    Spacer().frame(width: 8)
                    UndoWrapper(field) { field in
                        TextField("", value: field, format: .number)
                            .lineLimit(1)
                            .padding(.all, 1)
                            .disableAutocorrection(false)
                    }
                }
                .if(width != nil) { (view) in
                    view.frame(width: width)
                }
                .frame(height: height)
                .background(Palette.input.background)
                .cornerRadius(12)
                
                if width == nil {
                    Spacer()
                }
            }
            .font(inputFont)
        }
        .onAppear {
            text = "\(field.wrappedValue)"
        }
        .onChange(of: field.wrappedValue) { (field) in
            text = "\(field)"
        }
        .frame(height: self.height + ((self.inlineTitle ? 0 : self.topSpace) + (title == nil || inlineTitle ? 0 : 30)))
        .if(width != nil) { (view) in
            view.frame(width: width! + leadingSpace + (inlineTitle ? inlineTitleWidth : 0) + 16)
        }
    }
}
