//
//  Declarations.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI
import CoreGraphics

// TODO: - Should be in settings
let maxNationalIdNumber = 30000
let goodStatus = "Payment Confirmed by SBU"
let maxPoints: Float = 15.0
let largestFieldSize = 200
let largestPlayerCount = 600
let defaultWorksheetZoom = 125

let userDownloadData = "user download.csv"
let userDownloadRange = "$A$2:$AI$13001"
let userDownloadMinRow = 1
let userDownloadMaxRow = 13000
let userDownloadNationalIdColumn = 0
let userDownloadFirstNameColumn = 4
let userDownloadOtherNamesColumn = 3
// TODO Sort out the range and colum

// Parameters

public let appGroup = "group.com.sheareronline.ImportUsebio" // Has to match entitlements
public let widgetKind = "com.sheareronline.ImportUsebio"

// Sizes
public var isLandscape: Bool { false }

var inputTopHeight: CGFloat { 5.0 }
let inputDefaultHeight: CGFloat = 30.0
var inputToggleDefaultHeight: CGFloat { MyApp.format == .tablet ? 30.0 : 16.0 }
var bannerHeight: CGFloat { (MyApp.format == .tablet ? 60.0 : 50.0) }
var alternateBannerHeight: CGFloat { MyApp.format == .tablet ? 50.0 : 35.0 }
var minimumBannerHeight: CGFloat { MyApp.format == .tablet ? 40.0 : 20.0 }
var bannerBottom: CGFloat { (MyApp.format == .tablet ? 30.0 : 5.0) }
var slideInMenuRowHeight: CGFloat { MyApp.target == .iOS ? 50 : 25 }

// Fonts (Font)
var bannerFont: Font { Font.system(size: (MyApp.format == .tablet ? 32.0 : 24.0)) }
var alternateBannerFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 18.0)) }
var defaultFont: Font { Font.system(size: (MyApp.format == .tablet ? 28.0 : 24.0)) }
var toolbarFont: Font { Font.system(size: (MyApp.format == .tablet ? 16.0 : 14.0)) }
var captionFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 18.0)) }
var inputTitleFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 18.0)) }
var inputFont: Font { Font.system(size: 14.0) }
var messageFont: Font { Font.system(size: (MyApp.format == .tablet ? 16.0 : 14.0)) }
var searchFont: Font { Font.system(size: (MyApp.format == .tablet ? 20.0 : 16.0)) }

// Backups
let backupDirectoryDateFormat = "yyyy-MM-dd-HH-mm-ss-SSS"
let backupDateFormat = "yyyy-MM-dd HH:mm:ss.SSS Z"

// Other constants
let tagMultiplier = 1000000
let nullUUID = UUID(uuidString: "00000000-0000-0000-0000-000000000000")!

// Localisable names

public let appName = "Import Usebio XML"
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
