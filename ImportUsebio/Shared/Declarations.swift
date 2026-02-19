//
//  Declarations.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI
import CoreGraphics

// Parameters

public let appGroup = "group.com.sheareronline.ImportUsebio" // Has to match entitlements
public let widgetKind = "com.sheareronline.ImportUsebio"

// Sizes
public var isLandscape: Bool { false }

var inputTopHeight: CGFloat { 5.0 }
let inputDefaultHeight: CGFloat = 20.0
let inputCornerRadius: CGFloat = 4.0
let inputPickerCornerRadius: CGFloat = 3.0
var inputToggleDefaultHeight: CGFloat { MyApp.format == .tablet ? 30.0 : 16.0 }
var bannerHeight: CGFloat { (MyApp.format == .tablet ? 60.0 : 50.0) }
var alternateBannerHeight: CGFloat { MyApp.format == .tablet ? 50.0 : 35.0 }
var minimumBannerHeight: CGFloat { MyApp.format == .tablet ? 40.0 : 20.0 }
var bannerBottom: CGFloat { (MyApp.format == .tablet ? 30.0 : 20.0) }
var slideInMenuRowHeight: CGFloat { MyApp.target == .iOS ? 50 : 25 }

// Fonts (Font)
var bannerFont: Font { Font.system(size: (MyApp.format == .tablet ? 32.0 : 24.0)) }
var alternateBannerFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 18.0)) }
var defaultFont: Font { Font.system(size: (MyApp.format == .tablet ? 28.0 : 24.0)) }
var toolbarFont: Font { Font.system(size: (MyApp.format == .tablet ? 16.0 : 14.0)) }
var captionFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 18.0)) }
var inputTitleFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 18.0)) }
var inputFont: Font { Font.system(size: 14.0) }
var lookupFont: Font { Font.system(size: 12.0) }
var pickerFont: Font { Font.system(size: 10.0) }
var messageFont: Font { Font.system(size: (MyApp.format == .tablet ? 16.0 : 14.0)) }
var searchFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 16.0)) }

// Excel Colors
let excelRed = Color(red: 255, green: 0, blue: 0)
let excelYellow = Color(red: 255, green: 251,blue: 0)
let excelGrey = Color(red: 192, green: 192, blue: 192)
let excelFaint = Color(red: 217, green: 217, blue: 217)
let excelBanner = Color(red: 31, green: 3, blue: 108)
let excelGold = Color(red: 212, green: 175, blue: 55)
let excelSilver = Color(red: 188, green: 198, blue: 204)
let excelBronze = Color(red: 169, green: 113, blue: 66)
let excelNotActive = Color(red: 192, green: 0, blue: 0)
let excelNoHomeClub = Color(red: 190, green: 0, blue: 0)
let excelNotPaid = Color(red: 255, green: 160, blue: 160)
    
// Backups
let backupDirectoryDateFormat = "yyyy-MM-dd-HH-mm-ss-SSS"
let backupDateFormat = "yyyy-MM-dd HH:mm:ss.SSS Z"

// Other constants
let tagMultiplier = 1000000
let nullUUID = UUID(uuidString: "00000000-0000-0000-0000-000000000000")!

// Localisable names

public let appName = "Import Results"
public let appImage = "ImportUsebio"

public let dateFormat = "EEEE d MMMM yyyy"

public enum UIMode {
    case uiKit
    case appKit
    case unknown
}

#if canImport(UIKit)
public let target: UIMode = .uiKit
#elseif canImport(appKit)
public let target: UIMode = .appKit
#else
public let target: UIMode = .unknown
#endif
