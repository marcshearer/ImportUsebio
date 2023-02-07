//
//  Themes.swift
//  Contract Whist Scorecard
//
//  Created by Marc Shearer on 28/05/2020.
//  Copyright © 2020 Marc Shearer. All rights reserved.
//

import SwiftUI

enum ThemeAppearance: Int {
    case light = 1
    case dark = 2
    case device = 3
    
    #if canImport(UIKit)
    public var userInterfaceStyle: UIUserInterfaceStyle {
        switch self {
        case .light:
            return .light
        case .dark:
            return .dark
        default:
            return .unspecified
        }
    }#else
    public var appearanceStyle: NSAppearance {
        switch self {
        case .light:
            return NSAppearance(named: .aqua)!
        case .dark:
            return NSAppearance(named: .darkAqua)!
        default:
            return NSAppearance(named: .aqua)!
        }
    }
    #endif
}

enum ThemeName: String, CaseIterable {
    case standard = "Default"
    case alternate = "Alternate"
    case red = "Red"
    case blue = "Blue"
    case green = "Green"
    
    public var description: String {
        switch self {
        case .standard:
            return "Default"
        case .alternate:
            return "Alternate"
        case .red:
            return "Red"
        case .blue:
            return "Blue"
        case .green:
            return "Green"
        }
    }
}

enum ThemeTextType: CaseIterable {
    case normal
    case contrast
    case strong
    case faint
    case theme
}

enum ThemeBackgroundColorName: CaseIterable {
    case clear
    case nearlyClear
    case banner
    case bannerButton
    case bannerShadow
    case alternateBanner
    case alternateBannerButton
    case destructiveButton
    case background
    case alternate
    case tile
    case contrastTile
    case highlightTile
    case inset
    case header
    case subHeader
    case gridTitle
    case gridBoard
    case gridBoardSitout
    case gridBoardDisabled
    case gridTable
    case gridTableDisabled
    case autoComplete
    case contractSelected
    case contractUnselected
    case contractDisabled
    case divider
    case separator
    case listButton
    case menuEntry
    case imagePlaceholder
    case disabledButton
    case enabledButton
    case highlightButton
    case input
    case filterTile
    case filterUnused
    case filterUsed
    case datePicker
}

enum ThemeTextColorSetName: CaseIterable {
    case darkBackground
    case midBackground
    case lightBackground
}

enum ThemeSpecificColorName: CaseIterable {
    case bannerBackButton
    case alternateBannerBackButton
    case maskBackground
    case clickableBackground
    case gridLine
    case clearText
}

class Theme {
    private var themeName: ThemeName
    private var textColorConfig: [ThemeTextColorSetName: ThemeTextColor] = [:]
    private var backgroundColor: [ThemeBackgroundColorName : MyColor] = [:]
    private var textColor: [ThemeBackgroundColorName : MyColor] = [:]
    private var contrastTextColor: [ThemeBackgroundColorName : MyColor] = [:]
    private var strongTextColor: [ThemeBackgroundColorName : MyColor] = [:]
    private var faintTextColor: [ThemeBackgroundColorName : MyColor] = [:]
    private var themeTextColor: [ThemeBackgroundColorName : MyColor] = [:]
    private var specificColor: [ThemeSpecificColorName : MyColor] = [:]
    private var _icon: String?
    public var icon: String? { self._icon }
    
    init(themeName: ThemeName) {
        self.themeName = themeName
        if let config = Themes.themes[themeName] {
            self._icon = config.icon
            self.defaultTheme(from: config, all: true)
            
            if let basedOn = config.basedOn,
                let basedOnTheme = Themes.themes[basedOn] {
                self.defaultTheme(from: basedOnTheme)
            }
            if config.basedOn != .standard && themeName != .standard {
                if let defaultTheme = Themes.themes[.standard] {
                    self.defaultTheme(from: defaultTheme)
                }
            }
        }
    }
    
    public func background(_ backgroundColorName: ThemeBackgroundColorName) -> MyColor {
        return self.backgroundColor[backgroundColorName] ?? MyColor.clear
    }
    
    public func text(_ backgroundColorName: ThemeBackgroundColorName, textType: ThemeTextType = .normal) -> MyColor {
        switch textType {
        case .contrast:
            return self.contrastText(backgroundColorName)
        case .strong:
            return self.strongText(backgroundColorName)
        case .faint:
            return self.faintText(backgroundColorName)
        case .theme:
            return self.themeText(backgroundColorName)
        default:
            return self.textColor[backgroundColorName] ?? MyColor.clear
        }
    }
    
    public func contrastText(_ backgroundColorName: ThemeBackgroundColorName) -> MyColor {
        return self.contrastTextColor[backgroundColorName] ?? MyColor.clear
    }
    
    public func strongText(_ backgroundColorName: ThemeBackgroundColorName) -> MyColor {
        return self.strongTextColor[backgroundColorName] ?? MyColor.clear
    }
    
    public func faintText(_ backgroundColorName: ThemeBackgroundColorName) -> MyColor {
        return self.faintTextColor[backgroundColorName] ?? MyColor.clear
    }
    
    public func themeText(_ backgroundColorName: ThemeBackgroundColorName) -> MyColor {
        return self.themeTextColor[backgroundColorName] ?? MyColor.clear
    }
    
    public func textColor(textColorSetName: ThemeTextColorSetName, textType: ThemeTextType) -> MyColor? {
        return self.textColorConfig[textColorSetName]?.color(textType)?.myColor
    }
    
    public func specific(_ specificColorName: ThemeSpecificColorName) -> MyColor {
        return self.specificColor[specificColorName] ?? MyColor.black
    }
    
    private func defaultTheme(from: ThemeConfig, all: Bool = false) {
        // Default in any missing text colors
        for (name, themeTextColor) in from.textColor {
            if all || self.textColorConfig[name] == nil {
                self.textColorConfig[name] = themeTextColor
            }
        }
        
        // Iterate background colors filling in detail
        for (name, themeBackgroundColor) in from.backgroundColor {
            var anyTextColorName: ThemeTextColorSetName
            var darkTextColorName: ThemeTextColorSetName
            if all || self.backgroundColor[name] == nil {
                self.backgroundColor[name] = themeBackgroundColor.backgroundColor.myColor
            }
            anyTextColorName = themeBackgroundColor.anyTextColorName
            darkTextColorName = themeBackgroundColor.darkTextColorName ?? anyTextColorName
            if let anyTextColor = self.textColorConfig[anyTextColorName] {
                let darkTextColor = self.textColorConfig[darkTextColorName]
                if all || self.textColor[name] == nil {
                    self.textColor[name] = self.color(any: anyTextColor, dark: darkTextColor, .normal)
                }
                if all || self.contrastTextColor[name] == nil {
                    self.contrastTextColor[name] = self.color(any: anyTextColor, dark: darkTextColor, .contrast)
                }
                if all || self.strongTextColor[name] == nil {
                    self.strongTextColor[name] = self.color(any: anyTextColor, dark: darkTextColor, .strong)
                }
                if all || self.faintTextColor[name] == nil {
                    self.faintTextColor[name] = self.color(any: anyTextColor, dark: darkTextColor, .faint)
                }
                if all || self.themeTextColor[name] == nil {
                    self.themeTextColor[name] = self.color(any: anyTextColor, dark: darkTextColor, .theme)
                }
            }
        }
        for (name, themeSpecificColor) in from.specificColor {
            if all || self.specificColor[name] == nil {
                self.specificColor[name] = themeSpecificColor.myColor
            }
        }
    }
    
    private func color(any anyTextColor: ThemeTextColor, dark darkTextColor: ThemeTextColor?, _ textType: ThemeTextType) -> MyColor {
        let anyTraitColor = anyTextColor.color(textType) ?? anyTextColor.normal!
        
        if let darkTraitColor = darkTextColor?.color(textType) {
            let darkColor = darkTraitColor.darkColor ?? darkTraitColor.anyColor
            return self.traitColor(anyTraitColor.anyColor, darkColor)
        } else {
            return anyTraitColor.myColor
        }
    }

    private func traitColor(_ anyColor: MyColor, _ darkColor: MyColor?) -> MyColor {
        if let darkColor = darkColor {
            return MyColor(dynamicProvider: { (traitCollection) in
                traitCollection.userInterfaceStyle == .dark ? darkColor : anyColor
                
                })
        } else {
            return anyColor
        }
    }
}

class ThemeTraitColor {
    fileprivate let anyColor: MyColor
    fileprivate let darkColor: MyColor?
    
    init(_ anyColor: MyColor, _ darkColor: MyColor? = nil) {
        self.anyColor = anyColor
        self.darkColor = darkColor
    }
    
    public var myColor: MyColor {
        MyColor(dynamicProvider: { (traitCollection) in
            if traitCollection.userInterfaceStyle == .dark {
                return self.darkColor ?? self.anyColor
            } else {
                return self.anyColor
            }
        })
    }
}

class ThemeColor {
    fileprivate let backgroundColor: ThemeTraitColor
    fileprivate let anyTextColorName: ThemeTextColorSetName
    fileprivate let darkTextColorName: ThemeTextColorSetName?
    
    init(_ anyColor: MyColor, _ darkColor: MyColor? = nil, _ anyTextColorName: ThemeTextColorSetName, _ darkTextColorName: ThemeTextColorSetName? = nil) {
        self.backgroundColor = ThemeTraitColor(anyColor, darkColor)
        self.anyTextColorName = anyTextColorName
        self.darkTextColorName = darkTextColorName
    }
}

class ThemeTextColor {
    fileprivate var normal: ThemeTraitColor?
    fileprivate var contrast: ThemeTraitColor?
    fileprivate var strong: ThemeTraitColor?
    fileprivate var faint: ThemeTraitColor?
    fileprivate var theme: ThemeTraitColor?

    init(normal normalAny: MyColor, _ normalDark: MyColor? = nil, contrast contrastAny: MyColor? = nil, _ contrastDark: MyColor? = nil, strong strongAny: MyColor? = nil, _ strongDark: MyColor? = nil, faint faintAny: MyColor? = nil, _ faintDark: MyColor? = nil, theme themeAny: MyColor? = nil, _ themeDark: MyColor? = nil) {
        self.normal = self.traitColor(normalAny, normalDark)!
        self.contrast = self.traitColor(contrastAny, contrastDark)
        self.strong = self.traitColor(strongAny, strongDark)
        self.faint = self.traitColor(faintAny, faintDark)
        self.theme = self.traitColor(themeAny, themeDark)
    }
    
    fileprivate func color(_ textType: ThemeTextType) -> ThemeTraitColor? {
        switch textType {
        case .normal:
            return self.normal
        case .contrast:
            return self.contrast
        case .strong:
            return self.strong
        case .faint:
            return self.faint
        case .theme:
            return self.theme
        }
    }
    
    private func traitColor(_ anyColor: MyColor?, _ darkColor: MyColor? = nil) -> ThemeTraitColor? {
        if let anyColor = anyColor {
            return ThemeTraitColor(anyColor, darkColor)
        } else {
            return nil
        }
    }
}

class ThemeConfig {
    
    let basedOn: ThemeName?
    let icon: String?
    var backgroundColor: [ThemeBackgroundColorName: ThemeColor]
    var textColor: [ThemeTextColorSetName: ThemeTextColor]
    var specificColor: [ThemeSpecificColorName: ThemeTraitColor]
    
    init(basedOn: ThemeName? = nil, icon: String? = nil, background backgroundColor: [ThemeBackgroundColorName : ThemeColor], text textColor: [ThemeTextColorSetName : ThemeTextColor], specific specificColor: [ThemeSpecificColorName: ThemeTraitColor] = [:]) {
        self.basedOn = basedOn
        self.icon = icon
        self.backgroundColor = backgroundColor
        self.textColor = textColor
        self.specificColor = specificColor
    }
}


class Themes {

    public static var currentTheme: Theme!
    
    fileprivate static let themes: [ThemeName : ThemeConfig] = [
        .standard : ThemeConfig(
            background: [
                .clear                       : ThemeColor(#colorLiteral(red: 0, green: 0, blue: 0, alpha: 0), nil, .lightBackground, .darkBackground),
                .banner                      : ThemeColor(#colorLiteral(red: 0, green: 0.5152708056, blue: 0.7039544028, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .darkBackground, .darkBackground),
                .bannerShadow                : ThemeColor(#colorLiteral(red: 0, green: 0.5529411765, blue: 0.7647058824, alpha: 1), nil, .lightBackground),
                .bannerButton                : ThemeColor(#colorLiteral(red: 0.8374180198, green: 0.8374378085, blue: 0.8374271393, alpha: 1), nil, .lightBackground),
                .alternateBanner             : ThemeColor(#colorLiteral(red: 0.8784313725, green: 0.8784313725, blue: 0.8784313725, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .midBackground, .darkBackground),
                .alternateBannerButton       : ThemeColor(#colorLiteral(red: 0, green: 0.5152708056, blue: 0.7039544028, alpha: 1), #colorLiteral(red: 0, green: 0, blue: 0, alpha: 1), .darkBackground, .darkBackground),
                .destructiveButton           : ThemeColor(#colorLiteral(red: 0.9981788993, green: 0.2295429707, blue: 0.1891850233, alpha: 1), nil, .darkBackground),
                .background                  : ThemeColor(#colorLiteral(red: 0.999904573, green: 1, blue: 0.9998808503, alpha: 1), #colorLiteral(red: 0, green: 0, blue: 0, alpha: 1), .lightBackground, .darkBackground),
                .alternate                   : ThemeColor(#colorLiteral(red: 0.8784313725, green: 0.8784313725, blue: 0.8784313725, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .midBackground, .darkBackground),
                .contrastTile                : ThemeColor(#colorLiteral(red: 0, green: 0.5690457821, blue: 0.5746168494, alpha: 1), #colorLiteral(red: 0.4756349325, green: 0.4756467342, blue: 0.4756404161, alpha: 1), .darkBackground, .darkBackground),
                .tile                        : ThemeColor(#colorLiteral(red: 0.5, green: 0.8086761682, blue: 0.8080882377, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .midBackground, .darkBackground),
                .highlightTile               : ThemeColor(#colorLiteral(red: 0.9981788993, green: 0.2295429707, blue: 0.1891850233, alpha: 1), nil, .midBackground, .darkBackground),
                .inset                       : ThemeColor(#colorLiteral(red: 1, green: 0.9999999404, blue: 1, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .lightBackground, .lightBackground),
                .header                      : ThemeColor(#colorLiteral(red: 0.3921568627, green: 0.5490196078, blue: 0.5490196078, alpha: 1), #colorLiteral(red: 0.6286649108, green: 0.6231410503, blue: 0.6192827821, alpha: 1), .darkBackground,  .lightBackground),
                .subHeader                   : ThemeColor(#colorLiteral(red: 0.4705882353, green: 0.7058823529, blue: 0.7843137255, alpha: 1), #colorLiteral(red: 0.3921568627, green: 0.5490196078, blue: 0.5490196078, alpha: 1), .darkBackground,  .lightBackground),
                .divider                     : ThemeColor(#colorLiteral(red: 0, green: 0, blue: 0, alpha: 1), #colorLiteral(red: 0.6286649108, green: 0.6231410503, blue: 0.6192827821, alpha: 1), .darkBackground,  .lightBackground),
                .gridTitle                   : ThemeColor(#colorLiteral(red: 0.999904573, green: 1, blue: 0.9998808503, alpha: 1), #colorLiteral(red: 0.4756349325, green: 0.4756467342, blue: 0.4756404161, alpha: 1), .lightBackground, .darkBackground),
                .gridBoard                   : ThemeColor(#colorLiteral(red: 0.999904573, green: 1, blue: 0.9998808503, alpha: 1), #colorLiteral(red: 0.4756349325, green: 0.4756467342, blue: 0.4756404161, alpha: 1), .lightBackground, .darkBackground),
                .gridBoardSitout             : ThemeColor(#colorLiteral(red: 0.9254901961, green: 0.9254901961, blue: 0.9254901961, alpha: 1), #colorLiteral(red: 0.4509803922, green: 0.4509803922, blue: 0.4509803922, alpha: 1), .lightBackground, .darkBackground),
                .gridBoardDisabled           : ThemeColor(#colorLiteral(red: 0.9607843137, green: 0.9607843137, blue: 0.9607843137, alpha: 1), #colorLiteral(red: 0.4756349325, green: 0.4756467342, blue: 0.4756404161, alpha: 1), .lightBackground, .darkBackground),
                .gridTable                   : ThemeColor(#colorLiteral(red: 0.8616076436, green: 0.8620930367, blue: 0.8620918949, alpha: 1), #colorLiteral(red: 0.5741485357, green: 0.5741624236, blue: 0.574154973, alpha: 1), .lightBackground, .lightBackground),
                .gridTableDisabled           : ThemeColor(#colorLiteral(red: 0.8616076436, green: 0.8620930367, blue: 0.8620918949, alpha: 1), #colorLiteral(red: 0.5741485357, green: 0.5741624236, blue: 0.574154973, alpha: 1), .lightBackground, .lightBackground),
                .autoComplete                : ThemeColor(#colorLiteral(red: 0.9188528073, green: 0.9193390308, blue: 1, alpha: 1), #colorLiteral(red: 0.5741485357, green: 0.5741624236, blue: 0.574154973, alpha: 1), .lightBackground, .lightBackground),
                .contractSelected            : ThemeColor(#colorLiteral(red: 0.0166248735, green: 0.4766505957, blue: 0.9990670085, alpha: 1), nil, .darkBackground),
                .contractUnselected          : ThemeColor(#colorLiteral(red: 0.4705882353, green: 0.7058823529, blue: 0.7843137255, alpha: 1), nil, .lightBackground),
                .contractDisabled            : ThemeColor(#colorLiteral(red: 0.6096856772, green: 1, blue: 1, alpha: 1), nil, .lightBackground),
                .separator                   : ThemeColor(#colorLiteral(red: 0.6666069031, green: 0.6667050123, blue: 0.6665856242, alpha: 1), nil, .midBackground),
                .listButton                  : ThemeColor(#colorLiteral(red: 0, green: 0.3921568627, blue: 0.7058823529, alpha: 1), nil, .midBackground),
                .menuEntry                   : ThemeColor(#colorLiteral(red: 0.9569241405, green: 0.9567349553, blue: 0.9526277184, alpha: 1), #colorLiteral(red: 0.999904573, green: 1, blue: 0.9998808503, alpha: 1), .darkBackground,  .lightBackground),
                .imagePlaceholder            : ThemeColor(#colorLiteral(red: 0.9215686275, green: 0.9215686275, blue: 0.9215686275, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .lightBackground, .darkBackground),
                .disabledButton              : ThemeColor(#colorLiteral(red: 0.7843137255, green: 0.7843137255, blue: 0.7843137255, alpha: 1), nil, .midBackground),
                .enabledButton               : ThemeColor(#colorLiteral(red: 0.6666069031, green: 0.6667050123, blue: 0.6665856242, alpha: 1), nil, .midBackground),
                .highlightButton             : ThemeColor(#colorLiteral(red: 0, green: 0.5690457821, blue: 0.5746168494, alpha: 1), nil, .darkBackground),
                .input                       : ThemeColor(#colorLiteral(red: 1, green: 0.9999999404, blue: 1, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .lightBackground, .lightBackground),
                .filterTile                  : ThemeColor(#colorLiteral(red: 0.862745098, green: 0.862745098, blue: 0.862745098, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .lightBackground, .darkBackground),
                .filterUnused                : ThemeColor(#colorLiteral(red: 0, green: 0.5690457821, blue: 0.5746168494, alpha: 1), #colorLiteral(red: 0.5704585314, green: 0.5704723597, blue: 0.5704649091, alpha: 1), .darkBackground, .darkBackground),
                .filterUsed                  : ThemeColor(#colorLiteral(red: 0, green: 0.3285208941, blue: 0.5748849511, alpha: 1), #colorLiteral(red: 0.2, green: 0.2, blue: 0.2, alpha: 1), .darkBackground, .darkBackground),
                .datePicker                  : ThemeColor(#colorLiteral(red: 0.9019607843, green: 0.9019607843, blue: 1, alpha: 1), nil, .midBackground)
            ],
            text: [
                .lightBackground             : ThemeTextColor(normal: #colorLiteral(red: 0, green: 0, blue: 0, alpha: 1), contrast: #colorLiteral(red: 0.337254902, green: 0.4509803922, blue: 0.4549019608, alpha: 1), strong: #colorLiteral(red: 0.9981788993, green: 0.2295429707, blue: 0.1891850233, alpha: 1), faint: #colorLiteral(red: 0.7843137255, green: 0.7843137255, blue: 0.7843137255, alpha: 1), theme: #colorLiteral(red: 0.0166248735, green: 0.4766505957, blue: 0.9990670085, alpha: 1)) ,
                .midBackground               : ThemeTextColor(normal: #colorLiteral(red: 0, green: 0, blue: 0, alpha: 1), contrast: #colorLiteral(red: 0.9411764706, green: 0.9411764706, blue: 0.9411764706, alpha: 1), strong: #colorLiteral(red: 0, green: 0, blue: 0, alpha: 1),  faint: #colorLiteral(red: 0.4705882353, green: 0.4705882353, blue: 0.4705882353, alpha: 1), theme: #colorLiteral(red: 0.0166248735, green: 0.4766505957, blue: 0.9990670085, alpha: 1)) ,
                .darkBackground              : ThemeTextColor(normal: #colorLiteral(red: 0.9411764706, green: 0.9411764706, blue: 0.9411764706, alpha: 1), contrast: #colorLiteral(red: 0.6286649108, green: 0.6231410503, blue: 0.6192827821, alpha: 1), strong: #colorLiteral(red: 0.9981788993, green: 0.2295429707, blue: 0.1891850233, alpha: 1), faint: #colorLiteral(red: 0.4705882353, green: 0.4705882353, blue: 0.4705882353, alpha: 1), theme: #colorLiteral(red: 0.0166248735, green: 0.4766505957, blue: 0.9990670085, alpha: 1))
            ],
            specific: [
                .bannerBackButton            : ThemeTraitColor(#colorLiteral(red: 0.9411764706, green: 0.9411764706, blue: 0.9411764706, alpha: 1), #colorLiteral(red: 0.9411764706, green: 0.9411764706, blue: 0.9411764706, alpha: 1)),
                .alternateBannerBackButton   : ThemeTraitColor(#colorLiteral(red: 0.0166248735, green: 0.4766505957, blue: 0.9990670085, alpha: 1), #colorLiteral(red: 0.0166248735, green: 0.4766505957, blue: 0.9990670085, alpha: 1)),
                .maskBackground              : ThemeTraitColor(#colorLiteral(red: 0, green: 0, blue: 0, alpha: 0.3003128759), #colorLiteral(red: 0.999904573, green: 1, blue: 0.9998808503, alpha: 0.2982953811)),
                .clickableBackground         : ThemeTraitColor(#colorLiteral(red: 0, green: 0, blue: 0, alpha: 0.02357007967), #colorLiteral(red: 0.999904573, green: 1, blue: 0.9998808503, alpha: 0.02027367603)),
                .gridLine                    : ThemeTraitColor(#colorLiteral(red: 0, green: 0.5152708056, blue: 0.7039544028, alpha: 1), #colorLiteral(red: 0, green: 0.5152708056, blue: 0.7039544028, alpha: 1)),
                .clearText                   : ThemeTraitColor(#colorLiteral(red: 0.6666666865, green: 0.6666666865, blue: 0.6666666865, alpha: 1), #colorLiteral(red: 0.9411764706, green: 0.9411764706, blue: 0.9411764706, alpha: 1))
            ]
            )
    ]
    
    public static func selectTheme(_ themeName: ThemeName, changeIcon: Bool = false) {
        #if !widget
            let oldIcon = Themes.currentTheme?.icon
        #endif
        Themes.currentTheme = Theme(themeName: themeName)
        #if !widget
            let newIcon = Themes.currentTheme.icon
            #if canImport(UIKit)
                if UIApplication.shared.supportsAlternateIcons && changeIcon && oldIcon != newIcon {
                    Themes.setApplicationIconName(Themes.currentTheme.icon)
                }
            #endif
        #endif
    }
    
    #if !widget
        #if canImport(UIKit)
            private static func setApplicationIconName(_ iconName: String?) {
                if UIApplication.shared.responds(to: #selector(getter: UIApplication.supportsAlternateIcons)) && UIApplication.shared.supportsAlternateIcons {
                    
                    typealias setAlternateIconName = @convention(c) (NSObject, Selector, NSString?, @escaping (NSError) -> ()) -> ()
                    
                    let selectorString = "_setAlternateIconName:completionHandler:"
                    
                    let selector = NSSelectorFromString(selectorString)
                    let imp = UIApplication.shared.method(for: selector)
                    let method = unsafeBitCast(imp, to: setAlternateIconName.self)
                    method(UIApplication.shared, selector, iconName as NSString?, { _ in })
                }
            }
        #endif
    #endif
}
