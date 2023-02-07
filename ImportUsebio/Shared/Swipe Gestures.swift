//
//  Swipe Gestures.swift
// Bridge Score
//
//  Created by Marc Shearer on 24/02/2021.
//

import SwiftUI

enum SwipeGestureDirection {
    case right
    case left
    case up
    case down
}

struct SwipeGesture : ViewModifier {
    var minimumDistance: CGFloat = 30
    var action: (SwipeGestureDirection)->()
        
    func body(content: Content) -> some View { content
        .gesture(DragGesture(minimumDistance: minimumDistance, coordinateSpace: .global)
            .onEnded { value in
                let horizontal = value.translation.width
                let vertical = value.translation.height
                
                var swipeDirection: SwipeGestureDirection?
                if abs(horizontal) > 1.5 * abs(vertical) {
                    // Horizontal swipe
                    swipeDirection = (horizontal < 0 ? .left : .right)
                } else if abs(vertical) > 1.5 * abs(horizontal) {
                    // Vertical swipe
                    swipeDirection = (vertical < 0 ? .up : .down)
                }
                if let swipeDirection = swipeDirection {
                    action(swipeDirection)
                }
            })
    }
}

extension View {
    func onSwipe(minimumDistance: CGFloat = 30, action: @escaping (SwipeGestureDirection)->()) -> some View {
        self.modifier(SwipeGesture(minimumDistance: minimumDistance, action: action))
    }
}
