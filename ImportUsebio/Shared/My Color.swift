//
//  My Color.swift
// Bridge Score
//
//  Created by Marc Shearer on 26/02/2021.
//

import SwiftUI

#if canImport(UIKit)
class MyColor : UIColor {
    static override var clear: MyColor { return UIColor.clear as! MyColor}
    static override var black: MyColor { return UIColor.black as! MyColor}
}
#else
class MyColor: NSColor {
    
    convenience init(dynamicProvider: (MyTraitCollection)->MyColor) {
        let color = dynamicProvider(MyTraitCollection())
        self.init(Color(color.cgColor))
    }
    
    static override var clear: MyColor { return NSColor(_colorLiteralRed: 0, green: 0, blue: 0, alpha: 0) as! MyColor}
    static override var black: MyColor { return NSColor(_colorLiteralRed: 0, green: 0, blue: 0, alpha: 1) as! MyColor}
}

class MyTraitCollection {
    enum UserInterfaceStyle {
        case dark
        case light
    }

    public var userInterfaceStyle: UserInterfaceStyle { NSApp?.appearance?.name ?? NSAppearance.Name.aqua == NSAppearance.Name.darkAqua ? .dark : .light }
}

#endif
