    //
    //  Settings.swift
    //  ImportUsebio
    //
    //  Created by Marc Shearer on 29/04/2019.
    //  Copyright © 2019 Marc Shearer. All rights reserved.
    //

    import SwiftUI

    class Settings: ObservableObject {
        
        @Published var maxNationalIdNumber: Int!
        @Published var goodStatus: String!
        @Published var maxPoints: Float!
        @Published var largestFieldSize: Int!
        @Published var largestPlayerCount: Int!
        @Published var defaultWorksheetZoom: Float!
        @Published var userDownloadData: String!
        @Published var userDownloadMinRow: Int!
        @Published var userDownloadMaxRow: Int!
        @Published var userDownloadNationalIdColumn: String!
        @Published var userDownloadOtherNamesColumn:String!
        @Published var userDownloadFirstNameColumn: String!
        @Published var userDownloadRankColumn: String!
        @Published var userDownloadEmailColumn: String!
        @Published var userDownloadOtherUnionColumn: String!
        @Published var userDownloadHomeClubColumn: String!
        @Published var userDownloadStatusColumn: String!
        
        static public var current =  Settings(load: true)
        
        init(load: Bool = false) {
            if load {
                self.load()
            }
        }
            
        public func copy() -> Settings {
            let copy = Settings()
            copy.copy(from: self)
            return copy
        }
        
        public func copy(from: Settings) {
            self.maxNationalIdNumber = from.maxNationalIdNumber
            self.goodStatus = from.goodStatus
            self.maxPoints = from.maxPoints
            self.largestFieldSize = from.largestFieldSize
            self.largestPlayerCount = from.largestPlayerCount
            self.defaultWorksheetZoom = from.defaultWorksheetZoom
            self.userDownloadData = from.userDownloadData
            self.userDownloadMinRow = from.userDownloadMinRow
            self.userDownloadMaxRow = from.userDownloadMaxRow
            self.userDownloadNationalIdColumn = from.userDownloadNationalIdColumn
            self.userDownloadOtherNamesColumn = from.userDownloadOtherNamesColumn
            self.userDownloadFirstNameColumn = from.userDownloadFirstNameColumn
            self.userDownloadRankColumn = from.userDownloadRankColumn
            self.userDownloadEmailColumn = from.userDownloadEmailColumn
            self.userDownloadOtherUnionColumn = from.userDownloadOtherUnionColumn
            self.userDownloadHomeClubColumn = from.userDownloadHomeClubColumn
            self.userDownloadStatusColumn = from.userDownloadStatusColumn
        }
        
        public func load() {
            self.maxNationalIdNumber = UserDefault.settingMaxNationalIdNumber.int
            self.goodStatus = UserDefault.settingGoodStatus.string
            self.maxPoints = UserDefault.settingMaxPoints.float
            self.largestFieldSize = UserDefault.settingLargestFieldSize.int
            self.largestPlayerCount = UserDefault.settingLargestPlayerCount.int
            self.defaultWorksheetZoom = UserDefault.settingDefaultWorksheetZoom.float
            self.userDownloadData = UserDefault.settingUserDownloadData.string
            self.userDownloadMinRow = UserDefault.settingUserDownloadMinRow.int
            self.userDownloadMaxRow = UserDefault.settingUserDownloadMaxRow.int
            self.userDownloadNationalIdColumn = UserDefault.settingUserDownloadNationalIdColumn.string
            self.userDownloadOtherNamesColumn = UserDefault.settingUserDownloadOtherNamesColumn.string
            self.userDownloadFirstNameColumn = UserDefault.settingUserDownloadFirtNameColumn.string
            self.userDownloadRankColumn = UserDefault.settingUserDownloadRankColumn.string
            self.userDownloadEmailColumn = UserDefault.settingUserDownloadEmailColumn.string
            self.userDownloadOtherUnionColumn = UserDefault.settingUserDownloadOtherUnionColumn.string
            self.userDownloadHomeClubColumn = UserDefault.settingUserDownloadHomeClubColumn.string
            self.userDownloadStatusColumn = UserDefault.settingUserDownloadStatusColumn.string
        }
        
        public func save() {
            UserDefault.settingMaxNationalIdNumber.set(self.maxNationalIdNumber)
            UserDefault.settingGoodStatus.set(self.goodStatus)
            UserDefault.settingMaxPoints.set(self.maxPoints)
            UserDefault.settingLargestFieldSize.set(self.largestFieldSize)
            UserDefault.settingLargestPlayerCount.set(self.largestPlayerCount)
            UserDefault.settingDefaultWorksheetZoom.set(self.defaultWorksheetZoom)
            UserDefault.settingUserDownloadData.set(self.userDownloadData)
            UserDefault.settingUserDownloadMinRow.set(self.userDownloadMinRow)
            UserDefault.settingUserDownloadMaxRow.set(self.userDownloadMaxRow)
            UserDefault.settingUserDownloadNationalIdColumn.set(self.userDownloadNationalIdColumn)
            UserDefault.settingUserDownloadOtherNamesColumn.set(self.userDownloadOtherNamesColumn)
            UserDefault.settingUserDownloadFirtNameColumn.set(self.userDownloadFirstNameColumn)
            UserDefault.settingUserDownloadRankColumn.set(self.userDownloadRankColumn)
            UserDefault.settingUserDownloadEmailColumn.set(self.userDownloadEmailColumn)
            UserDefault.settingUserDownloadOtherUnionColumn.set(self.userDownloadOtherUnionColumn)
            UserDefault.settingUserDownloadHomeClubColumn.set(self.userDownloadHomeClubColumn)
            UserDefault.settingUserDownloadStatusColumn.set(self.userDownloadStatusColumn)
        }
    }
