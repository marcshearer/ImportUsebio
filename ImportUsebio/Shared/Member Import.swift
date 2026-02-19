//
//  Member List.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 03/02/2026.
//

import SwiftUI

class MemberList {
    
    public static let shared = MemberList()
    
    private(set) var lastDownloaded: Date?
    
    func download(completion: @escaping (Bool, String)->()) {
        let downloaded = Date()
        let url = URL(string: "https://www.mempad.co.uk/sites/default/files/~integration/members.csv")!
        var request = URLRequest(url: url)
        request.cachePolicy = .reloadIgnoringLocalCacheData
        let task = URLSession.shared.dataTask(with: request) {  [self] (data, response, error) in
            var success = (error == nil)
            if success, let httpResponse = response as? HTTPURLResponse, !(200...299).contains(httpResponse.statusCode) {
                success = false
            }
            if !success {
                completion(false, error?.localizedDescription ?? "Unknown Error")
            } else if let data = data {
                let csvData = String(data: data, encoding: .utf8)!
                let lines = csvData.replacingOccurrences(of: "\r\n", with: "\n").components(separatedBy: "\n")
                for line in lines.dropFirst() {
                    let fields = line.components(separatedBy: ",").map{$0.replacing("\"", with: "")}
                    let member = MemberViewModel(nationalId: fields[0], firstName: fields[1], lastName: fields[2], homeClub: fields[3], postCode: postCode(postCodeAndRank: fields[4]), rankCode: rankCode(postCodeAndRank: fields[4]), downloaded: downloaded)
                    if let existingMember = MemberViewModel.member(nationalId: member.nationalId) {
                        existingMember.copy(from: member, copyMO: false)
                        existingMember.updateMO()
                        existingMember.save()
                    } else {
                        member.insert()
                    }
                    
                }
                // Remove anything we have not just updated
                let members = MasterData.shared.members.array as! [MemberViewModel]
                for member in members.filter({$0.downloaded != downloaded}) {
                    member.remove()
                }
                self.lastDownloaded = downloaded
                completion(true, "")
            }
        }
        task.resume()
    }
    
    func rankCode(postCodeAndRank: String) -> Int {
        if let rankSubString = postCodeAndRank.split(separator: " - ").last, let rankCode = RankViewModel.rank(rankName: String(rankSubString))?.rankCode {
            return rankCode
        } else {
            return -1
        }
    }
    
    func postCode(postCodeAndRank: String) -> String {
        if let postCode = postCodeAndRank.split(separator: "-").first {
            return String(postCode).replacingOccurrences(of: " ", with: "")
        } else {
            return ""
        }
    }
}
