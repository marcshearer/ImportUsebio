//
//  File System.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 09/02/2023.
//

import SwiftUI
import UniformTypeIdentifiers

class FileSystem {
    
    static public func findFile(relativeTo: URL? = nil, title: String? = nil, prompt: String? = nil, types: [String]? = nil, completion: @escaping (URL, Data)->()) {
#if canImport(AppKit)
        let openPanel = NSOpenPanel()
        openPanel.allowsMultipleSelection = false
        openPanel.canChooseDirectories = false
        openPanel.canCreateDirectories = false
        openPanel.canChooseFiles = true
        if let title = title {
            openPanel.title = title
        }
        openPanel.prompt = prompt ?? "Select target"
        openPanel.level = .floating
        if let types = types {
            openPanel.allowedContentTypes = types.map{UTType(tag: $0, tagClass: UTTagClass.filenameExtension, conformingTo: .text)!}
        }
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
    
    static public func findDirectory(relativeTo: URL? = nil, title: String? = nil, prompt: String? = nil, completion: @escaping (URL, Data)->()) {
#if canImport(AppKit)
        let openPanel = NSOpenPanel()
        openPanel.allowsMultipleSelection = false
        openPanel.canChooseDirectories = true
        openPanel.canCreateDirectories = false
        openPanel.canChooseFiles = false
        if let title = title {
            openPanel.title = title
        }
        openPanel.prompt = prompt ?? "Select directory"
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
    
    static public func saveFile(relativeTo: URL? = nil, title: String? = nil, prompt: String? = nil, filename: String, completion: @escaping (URL, Data)->()) {
#if canImport(AppKit)
        let savePanel = NSSavePanel()
        savePanel.canCreateDirectories = true
        if let title = title {
            savePanel.title = title
        }
        savePanel.prompt = prompt ?? "Select target"
        savePanel.nameFieldStringValue = filename
        savePanel.level = .floating
        savePanel.begin { result in
            if result == .OK {
                if let url = savePanel.url {
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
