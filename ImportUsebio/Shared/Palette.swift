//
//  Palette.swift.background
//  Contract Whist Scorecard
//
//  Created by Marc Shearer on 04/07/2019.
//  Copyright © 2019 Marc Shearer. All rights reserved.
//

import SwiftUI

class Palette {
    
    @BackgroundColor(.clear) static var clear
    @BackgroundColor(.banner) static var banner
    @BackgroundColor(.bannerButton) static var bannerButton
    @BackgroundColor(.bannerShadow) static var bannerShadow
    @BackgroundColor(.alternateBanner) static var alternateBanner
    @BackgroundColor(.alternateBannerButton) static var alternateBannerButton
    @BackgroundColor(.destructiveButton) static var destructiveButton
    @BackgroundColor(.background) static var background
    @BackgroundColor(.alternate) static var alternate
    @BackgroundColor(.tile) static var tile
    @BackgroundColor(.contrastTile) static var contrastTile
    @BackgroundColor(.highlightTile) static var highlightTile
    @BackgroundColor(.inset) static var inset
    @BackgroundColor(.header) static var header
    @BackgroundColor(.subHeader) static var subHeader
    @BackgroundColor(.gridTitle) static var gridTitle
    @BackgroundColor(.gridBoard) static var gridBoard
    @BackgroundColor(.gridBoardSitout) static var gridBoardSitout
    @BackgroundColor(.gridBoardDisabled) static var gridBoardDisabled
    @BackgroundColor(.gridTable) static var gridTable
    @BackgroundColor(.gridTableDisabled) static var gridTableDisabled
    @BackgroundColor(.autoComplete) static var autoComplete
    @BackgroundColor(.contractSelected) static var contractSelected
    @BackgroundColor(.contractUnselected) static var contractUnselected
    @BackgroundColor(.contractDisabled) static var contractDisabled
    @BackgroundColor(.divider) static var divider
    @BackgroundColor(.separator) static var separator
    @BackgroundColor(.listButton) static var listButton
    @BackgroundColor(.menuEntry) static var menuEntry
    @BackgroundColor(.imagePlaceholder) static var imagePlaceholder
    @BackgroundColor(.disabledButton) static var disabledButton
    @BackgroundColor(.enabledButton) static var enabledButton
    @BackgroundColor(.highlightButton) static var highlightButton
    @BackgroundColor(.input) static var input
    @BackgroundColor(.filterTile) static var filterTile
    @BackgroundColor(.filterUsed) static var filterUsed
    @BackgroundColor(.filterUnused) static var filterUnused
    @BackgroundColor(.datePicker) static var datePicker
    
    // Specific colors
    @SpecificColor(.bannerBackButton) static var bannerBackButton
    @SpecificColor(.alternateBannerBackButton) static var alternateBannerBackButton
    @SpecificColor(.maskBackground) static var maskBackground
    @SpecificColor(.clickableBackground) static var clickableBackground
    @SpecificColor(.gridLine) static var gridLine
    @SpecificColor(.clearText) static var clearText
        
    class func colorDetail(color: Color) -> (red: CGFloat, green: CGFloat, blue: CGFloat, alpha: CGFloat) {
        var red: CGFloat = 0
        var green: CGFloat = 0
        var blue: CGFloat = 0
        var alpha: CGFloat = 0
        
        color.getRed(&red, green: &green, blue: &blue, alpha: &alpha)
        
        return (red: red, green: green, blue: blue, alpha: alpha)
    }
    
    class func colorMatch(color: Color) -> String {
        var matches = ""
        
        func addColor(_ color: String) {
            if matches == "" {
                matches = color
            } else {
                matches = matches + ", " + color
            }
        }
        
        for colorName in ThemeBackgroundColorName.allCases {
            if color.cgColor == Themes.currentTheme.background(colorName).cgColor {
                addColor("\(colorName)")
            }
        }
        
        for colorName in ThemeTextColorSetName.allCases {
            for colorType in ThemeTextType.allCases {
                if color.cgColor == Themes.currentTheme.textColor(textColorSetName: colorName, textType: colorType)?.cgColor {
                    addColor("\(colorName)-\(colorType)")
                }
            }
        }
        
        for colorName in ThemeSpecificColorName.allCases {
            if color.cgColor == Themes.currentTheme.specific(colorName).cgColor {
                addColor("\(colorName)")
            }
        }
        
        return matches
    }
}

@propertyWrapper fileprivate final class BackgroundColor {
    var wrappedValue: PaletteColor
    
    fileprivate init(_ colorName: ThemeBackgroundColorName) {
        wrappedValue = PaletteColor(colorName)
    }
}

@propertyWrapper fileprivate final class SpecificColor {
    var wrappedValue: Color
    
    fileprivate init(_ colorName: ThemeSpecificColorName) {
        wrappedValue = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.specific(colorName)}))
    }
}

class PaletteColor: Equatable {
    let background: Color
    let text: Color
    let contrastText: Color
    let strongText: Color
    let faintText: Color
    let themeText: Color
    
    static func == (lhs: PaletteColor, rhs: PaletteColor) -> Bool {
        return (lhs.background == rhs.background && lhs.text == rhs.text && lhs.contrastText == rhs.contrastText && lhs.strongText == rhs.strongText && lhs.faintText == rhs.faintText && lhs.themeText == rhs.themeText)
    }
    
    init(_ colorName: ThemeBackgroundColorName) {
        self.background = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.background(colorName)}))
        self.text = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.text(colorName)}))
        self.contrastText = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.contrastText(colorName)}))
        self.strongText = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.strongText(colorName)}))
        self.faintText = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.faintText(colorName)}))
        self.themeText = Color(MyColor(dynamicProvider: { (_) in Themes.currentTheme.themeText(colorName)}))
    }
    
    public func textColor(_ type: ThemeTextType) -> Color {
        switch type {
        case .normal:
            return self.text
        case .contrast:
            return self.contrastText
        case .strong:
            return self.strongText
        case .faint:
            return self.faintText
        case .theme:
            return self.themeText
        }
    }
}

extension Color {
    func getRed(_ red: inout CGFloat, green: inout CGFloat, blue: inout CGFloat, alpha: inout CGFloat) {

        #if canImport(UIKit)
        typealias NativeColor = MyColor
        #elseif canImport(AppKit)
        typealias NativeColor = NSColor
        #endif

        var r: CGFloat = 0
        var g: CGFloat = 0
        var b: CGFloat = 0
        var o: CGFloat = 0

        NativeColor(self).getRed(&r, green: &g, blue: &b, alpha: &o)
    }
}
