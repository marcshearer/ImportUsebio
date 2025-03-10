//
//  Generic Csv Parser.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 25/03/2023.
//

import Foundation

public class GenericCsvParser {
    
    init?(fileUrl: URL, data: Data, manualPointsColumn: String? = nil, completion: @escaping (ScoreData?, [String])->()) {
        if let dataString = String(data: data, encoding: .utf8) {
            let dataLines = dataString.replacingOccurrences(of: "\n", with: "").components(separatedBy: "\r")
            if let firstLine = dataLines.first(where: {!$0.isEmpty}) {
                let upper = firstLine.uppercased()
                if upper.left(10).uppercased() == "PARAMETERS" || upper.left(12) == "INSTRUCTIONS" {
                    let data = dataLines.map{$0.components(separatedBy: ",")}
                    _ = ManualCsvParser(fileUrl: fileUrl, data: data, manualPointsColumn: manualPointsColumn, completion: completion)
                } else {
                    let data = dataLines.map{$0.components(separatedBy: ",")}
                    _ = BridgeWebsCsvParser(fileUrl: fileUrl, data: data, completion: completion)
                }
            } else {
                return nil
            }
        } else {
            return nil
        }
    }
}
