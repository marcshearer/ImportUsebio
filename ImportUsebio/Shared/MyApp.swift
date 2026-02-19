//
//  MyApp.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import Foundation

import SwiftUI

class MyApp {
    
    
    enum Target {
        case iOS
        case macOS
    }
   
    enum Format {
        case computer
        case tablet
        case phone
    }

    enum Database: String {
        case development = "Development"
        case production = "Production"
        case unknown = ""
        
        public var name: String {
            return self.rawValue
        }
    }
    
    static let objectModel = Model(eventEntity, clubEntity, rankEntity, memberEntity, blockedEntity, strataDefEntity)
    
    static let shared = MyApp()
    
    static let defaults = UserDefaults(suiteName: appGroup)!
    
    /// Database to use - This  **MUST MUST MUST** match icloud entitlement
    static let expectedDatabase: Database = .production
    
    public static var database: Database = .unknown
    public static var undoManager = UndoManager()
    
    #if targetEnvironment(macCatalyst)
    public static let target: Target = .macOS
    #else
    public static let target: Target = .iOS
    #endif

    public static var format: Format = .tablet
 
    public func start() {
        MasterData.shared.load()
        Themes.selectTheme(.standard)
        self.registerDefaults()
        #if !widget
        Version.current.load()
        // Remove comment (CAREFULLY) if you want to clear the iCloud DB
        // DatabaseUtilities.initialiseAllCloud() {
            // Remove (CAREFULLY) if you want to clear the Core Data DB
            //DatabaseUtilities.initialiseAllCoreData()
            self.setupDatabase()
            // self.setupPreviewData()
        //}
              
        #if canImport(UIKit)
        UITextView.appearance().backgroundColor = .clear
        #endif
        #endif
    }
    
    private func setupDatabase() {
        
        // Get saved database
        MyApp.database = Database(rawValue: UserDefault.database.string) ?? .unknown
        
        #if !widget
            // Check which database we are connected to
        #endif
    }
     
    #if !widget
    private func setupPreviewData() {

    }
    #endif
    
    private func registerDefaults() {
        var initial: [String:Any] = [:]
        for value in UserDefault.allCases {
            initial[value.name] = value.defaultValue ?? ""
        }
        MyApp.defaults.register(defaults: initial)
    }
}

enum BridgeScoreError: Error {
    case invalidData
}
