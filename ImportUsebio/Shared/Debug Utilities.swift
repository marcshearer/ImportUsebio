//
//  File.swift
//  Shortcut Menu
//
//  Created by Marc Shearer on 09/03/2021.
//  Copyright © 2021 Marc Shearer. All rights reserved.
//

import SwiftUI

extension View {
    
    func debugAction(_ action: () -> Void) -> Self {
         action()
    
        return self
    }
    
    func debugPrint(_ value: Any) -> Self {
        debugAction { print(value) }
    }
}
