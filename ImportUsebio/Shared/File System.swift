//
//  File System.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 09/02/2023.
//

import SwiftUI

class FileSystem {
    
    static public func findFile(relativeTo: URL? = nil,completion: @escaping (URL, Data)->()) {
#if canImport(AppKit)
        let openPanel = NSOpenPanel()
        openPanel.allowsMultipleSelection = false
        openPanel.canChooseDirectories = false
        openPanel.canCreateDirectories = false
        openPanel.canChooseFiles = true
        openPanel.prompt = "Select target"
        openPanel.level = .floating
        openPanel.begin { result in
            if result == .OK {
                if !openPanel.urls.isEmpty {
                    let url = openPanel.urls[0]
                    do {
                        let content = try url.bookmarkData(options: .withSecurityScope, includingResourceValuesForKeys: nil, relativeTo: relativeTo)
                        completion(url, content)
                    } catch {
                        // Ignore error
                        print(error.localizedDescription)
                    }
                }
            }
        }
#endif
    }
}
