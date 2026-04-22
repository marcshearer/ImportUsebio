//
//  Member List.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 03/02/2026.
//

import SwiftUI

fileprivate enum DownloadColumn: String {
    case memberId = "MASTER POINT NUMBER"
    case forename = "FORENAME"
    case surname = "SURNAME"
    case homeClub = "HOME CLUB"
    case postcodeAndRank = "POSTCODE AND RANK"
}

fileprivate enum ImportColumn: String {
    case memberId = "MEMBER"
    case status = "STATUS"
    case forename = "FORENAME"
    case surname = "SURNAME"
    case homeClub = "HOME CLUB"
    case rank = "RANK CODE"
    case postcode = "POST CODE"
}

class MemberList {
    private var memberIdColumn: Int?
    private var forenameColumn: Int?
    private var surnameColumn: Int?
    private var homeClubColumn: Int?
    private var postcodeAndRankColumn: Int?
    private var statusColumn: Int?
    private var rankColumn: Int?
    private var postcodeColumn: Int?
    
    public static let shared = MemberList()
    
    private(set) var lastDownloaded: Date?
    
    func download() async -> (Bool, String) {
        var updated: Int = 0
        var new: Int = 0
        var lapsed: Int = 0
        var importedIds: [UUID:Bool] = [:]
        var result = (true, "")
        
        let downloadedDate = Date()
        
        let url = URL(string: "https://www.mempad.co.uk/sites/default/files/~integration/members.csv")!
        var request = URLRequest(url: url)
        request.cachePolicy = .reloadIgnoringLocalCacheData
        do {
            let (data, response) = try await URLSession.shared.data(for: request)
            var success = true
            if let httpResponse = response as? HTTPURLResponse, !(200...299).contains(httpResponse.statusCode) {
                success = false
            }
            if !success {
                result = (false, "Unable to fetch data")
            } else {
                let csvData = String(data: data, encoding: .utf8)!
                let lines = csvData.replacingOccurrences(of: "\r\n", with: "\n").components(separatedBy: "\n")
                let headerLine = lines.first!
                setupDownloadColumns(from: headerLine)
                if let memberIdColumn = memberIdColumn, let forenameColumn = forenameColumn, let surnameColumn = surnameColumn, let homeClubColumn = homeClubColumn, let postcodeAndRankColumn = postcodeAndRankColumn {
                    for line in lines.dropFirst() {
                        let fields = line.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
                        var member = MemberViewModel.member(nationalId: fields[memberIdColumn])
                        if let member = member {
                            member.nationalId = fields[memberIdColumn]
                            member.otherNames = fields[forenameColumn]
                            member.lastName = fields[surnameColumn]
                            member.status = .active
                            member.homeClub = fields[homeClubColumn]
                            member.postCode = postCode(postCodeAndRank: fields[postcodeAndRankColumn])
                            member.rankCode = rankCode(postCodeAndRank: fields[postcodeAndRankColumn])
                            importedIds[member.memberId] = true
                            if member.changed {
                                member.save()
                                updated += 1
                                Utility.debugMessage("Download", "Member updated")
                            }
                        } else {
                            member = MemberViewModel(nationalId: fields[memberIdColumn], otherNames: fields[forenameColumn], lastName: fields[surnameColumn], status: .active, homeClub: fields[homeClubColumn], postCode: postCode(postCodeAndRank: fields[postcodeAndRankColumn]), rankCode: rankCode(postCodeAndRank: fields[postcodeAndRankColumn]), downloaded: downloadedDate)
                            member!.insert()
                            importedIds[member!.memberId] = false
                            new += 1
                            Utility.debugMessage("Download", "New member added")
                        }
                    }
                    // Mark anything we have not just imported as lapsed
                    let members = MasterData.shared.members.array as! [MemberViewModel]
                    Utility.debugMessage("Import", "Starting lapse set")
                    for member in members.filter({importedIds[$0.memberId] == nil && $0.status == .active}) {
                        member.status = .lapsed
                        member.save()
                        lapsed += 1
                    }
                    Utility.debugMessage("Import", "Ending lapse set")
                    self.lastDownloaded = downloadedDate
                    result =  (true, "")
                } else {
                    result = (false, "Missing mandatory columns in downloaded file")
                }
            }
        } catch {
            return (false, error.localizedDescription)
        }
        return result
    }
    
    func importDropped(_ csvData: String, completion: @escaping (Bool, String)->()) {
        var updated: Int = 0
        var new: Int = 0
        var removed: Int = 0
        
        let importedDate = Date()
        
        let lines = csvData.replacingOccurrences(of: "\r\n", with: "\n").components(separatedBy: "\n")
        let headerLine = lines.first!
        setupImportColumns(from: headerLine)
        if let memberIdColumn = memberIdColumn, let forenameColumn = forenameColumn, let surnameColumn = surnameColumn, let statusColumn = statusColumn, let homeClubColumn = homeClubColumn, let postcodeColumn = postcodeColumn, let rankColumn = rankColumn {
            for line in lines.dropFirst() {
                let fields = line.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
                let nationalId = fields[memberIdColumn]
                let otherNames = fields[forenameColumn]
                let lastName = fields[surnameColumn]
                let status = (fields[statusColumn].uppercased() == "ACTIVE" ? PlayerStatus.active : .lapsed)
                let homeClub = fields[homeClubColumn]
                let postcode = firstSection(of: fields[postcodeColumn])
                let rankCode: Int = Int(fields[rankColumn]) ?? Settings.current.otherNBORank
                
                let member = MemberViewModel(nationalId: nationalId, otherNames: otherNames, lastName: lastName, status: status, homeClub: homeClub, postCode: postcode, rankCode: rankCode, downloaded: importedDate)
                
                if let existingMember = MemberViewModel.member(nationalId: member.nationalId) {
                    existingMember.copy(from: member, copyMO: false)
                    existingMember.save()
                    updated += 1
                } else {
                    member.insert()
                    new += 1
                }
            }
            // Remove anything we have not just imported
            let members = MasterData.shared.members.array as! [MemberViewModel]
            for member in members.filter({$0.downloaded != importedDate}) {
                member.remove()
                removed += 1
            }
            self.lastDownloaded = importedDate
            completion(true, "")
        } else {
            completion(false, "Missing mandatory columns in downloaded file")
        }
    }
    
    private func setupDownloadColumns(from headerLine: String) {
        let fields = headerLine.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
        memberIdColumn = setupDownloadColumn(from: fields, for: .memberId)
        forenameColumn = setupDownloadColumn(from: fields, for: .forename)
        surnameColumn = setupDownloadColumn(from: fields, for: .surname)
        homeClubColumn = setupDownloadColumn(from: fields, for: .homeClub)
        postcodeAndRankColumn = setupDownloadColumn(from: fields, for: .postcodeAndRank)
    }
    
    private func setupDownloadColumn(from fields: [String], for column: DownloadColumn) -> Int? {
        fields.firstIndex(where: {$0.uppercased() == column.rawValue})
    }
    
    private func setupImportColumns(from headerLine: String) {
        let fields = headerLine.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
        memberIdColumn = setupImportColumn(from: fields, for: .memberId)
        forenameColumn = setupImportColumn(from: fields, for: .forename)
        surnameColumn = setupImportColumn(from: fields, for: .surname)
        statusColumn = setupImportColumn(from: fields, for: .status)
        homeClubColumn = setupImportColumn(from: fields, for: .homeClub)
        postcodeColumn = setupImportColumn(from: fields, for: .postcode)
        rankColumn = setupImportColumn(from: fields, for: .rank)
    }
    
    private func setupImportColumn(from fields: [String], for column: ImportColumn) -> Int? {
        fields.firstIndex(where: {$0.uppercased() == column.rawValue})
    }
    
    private func rankCode(postCodeAndRank: String) -> Int {
        if let rankSubString = postCodeAndRank.split(separator: " - ").last, let rankCode = RankViewModel.rank(rankName: String(rankSubString))?.rankCode {
            return rankCode
        } else {
            return -1
        }
    }
    
    private func postCode(postCodeAndRank: String) -> String {
        if let postCode = postCodeAndRank.split(separator: "-").first {
            return String(postCode).replacingOccurrences(of: " ", with: "")
        } else {
            return ""
        }
    }
    
    private func firstSection(of postCode: String) -> String {
        let segments = postCode.components(separatedBy: " ")
        return segments.first ?? ""
    }
}
