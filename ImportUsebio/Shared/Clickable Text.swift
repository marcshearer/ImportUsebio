//
//  Clickable Text.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 20/02/2026.
//

import SwiftUI

struct LeadingClickableText : View {
    private var text: String
    
    init(_ text: String) {
        self.text = text
    }
    
    var body : some View {
        HStack {
            Text(text)
            Spacer()
        }
        .contentShape(Rectangle())
    }
}

struct CenteredClickableText : View {
    private var text: String
    
    init(_ text: String) {
        self.text = text
    }
    
    var body : some View {
        HStack {
            Spacer()
            Text(text)
            Spacer()
        }
        .contentShape(Rectangle())
    }
}

struct TrailingClickableText : View {
    private var text: String
    
    init(_ text: String) {
        self.text = text
    }
    
    var body : some View {
        HStack {
            Spacer()
            Text(text)
        }
        .contentShape(Rectangle())
    }
}
