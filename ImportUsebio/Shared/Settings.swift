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
        @Published var notPaidPaymentStatus: String!
        @Published var otherNBORank: Int!
        @Published var ignorePaymentFrom: Int!
        @Published var ignorePaymentTo: Int!
        @Published var maxPoints: Float!
        @Published var largestFieldSize: Int!
        @Published var largestPlayerCount: Int!
        @Published var defaultWorksheetZoom: Float!
        @Published var linesPerFormattedPage: Int!
        @Published var userDownloadData: String!
        @Published var userDownloadFrozenData: String!
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
        @Published var userDownloadPaymentStatusColumn: String!

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
            self.notPaidPaymentStatus = from.notPaidPaymentStatus
            self.ignorePaymentFrom = from.ignorePaymentFrom
            self.ignorePaymentTo = from.ignorePaymentTo
            self.otherNBORank = from.otherNBORank
            self.maxPoints = from.maxPoints
            self.largestFieldSize = from.largestFieldSize
            self.largestPlayerCount = from.largestPlayerCount
            self.defaultWorksheetZoom = from.defaultWorksheetZoom
            self.linesPerFormattedPage = from.linesPerFormattedPage
            self.userDownloadData = from.userDownloadData
            self.userDownloadFrozenData = from.userDownloadFrozenData
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
            self.userDownloadPaymentStatusColumn = from.userDownloadPaymentStatusColumn
        }
        
        public func load() {
            self.maxNationalIdNumber = UserDefault.settingMaxNationalIdNumber.int
            self.goodStatus = UserDefault.settingGoodStatus.string
            self.notPaidPaymentStatus = UserDefault.settingnotPaidPaymentStatus.string
            self.otherNBORank = UserDefault.settingNonSBURank.int
            self.ignorePaymentFrom = UserDefault.settingIgnorePaymentFrom.int
            self.ignorePaymentTo = UserDefault.settingIgnorePaymentTo.int
            self.maxPoints = UserDefault.settingMaxPoints.float
            self.largestFieldSize = UserDefault.settingLargestFieldSize.int
            self.largestPlayerCount = UserDefault.settingLargestPlayerCount.int
            self.defaultWorksheetZoom = UserDefault.settingDefaultWorksheetZoom.float
            self.linesPerFormattedPage = UserDefault.settingLinesPerFormattedPage.int
            self.userDownloadData = UserDefault.settingUserDownloadData.string
            self.userDownloadFrozenData = UserDefault.settingUserDownloadFrozenData.string
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
            self.userDownloadPaymentStatusColumn = UserDefault.settingUserDownloadPaymentStatusColumn.string
        }
        
        public func save() {
            UserDefault.settingMaxNationalIdNumber.set(self.maxNationalIdNumber)
            UserDefault.settingGoodStatus.set(self.goodStatus)
            UserDefault.settingnotPaidPaymentStatus.set(self.notPaidPaymentStatus)
            UserDefault.settingNonSBURank.set(self.otherNBORank)
            UserDefault.settingIgnorePaymentFrom.set(self.ignorePaymentFrom)
            UserDefault.settingIgnorePaymentTo.set(self.ignorePaymentTo)
            UserDefault.settingMaxPoints.set(self.maxPoints)
            UserDefault.settingLargestFieldSize.set(self.largestFieldSize)
            UserDefault.settingLargestPlayerCount.set(self.largestPlayerCount)
            UserDefault.settingDefaultWorksheetZoom.set(self.defaultWorksheetZoom)
            UserDefault.settingLinesPerFormattedPage.set(self.linesPerFormattedPage)
            UserDefault.settingUserDownloadFrozenData.set(self.userDownloadFrozenData)
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
            UserDefault.settingUserDownloadPaymentStatusColumn.set(self.userDownloadPaymentStatusColumn)
        }
    }
