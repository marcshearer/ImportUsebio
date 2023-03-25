//
//  Generic Csv Parser.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 25/03/2023.
//

import Foundation

public class GenericCsvParser {
    
    init?(fileUrl: URL, data: Data, completion: @escaping (ScoreData?, String?)->()) {
        if let dataString = String(data: data, encoding: .utf8) {
            let dataLines = dataString.replacingOccurrences(of: "\n", with: "").components(separatedBy: "\r")
            if let firstLine = dataLines.first(where: {!$0.isEmpty}) {
                if firstLine.left(10).uppercased() == "PARAMETERS" {
                    let data = dataLines.map{$0.components(separatedBy: ",")}
                    _ = ManualCsvParser(fileUrl: fileUrl, data: data, completion: completion)
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
