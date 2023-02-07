//
//  SelectInput.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

struct SelectInputView: View {
    @State private var inputFilename: String = "/users/shearerm/Documents/Input.xml"
    @State private var securityBookmark: Data? = nil
    @State private var refresh = true

    var body: some View {
        VStack {
  
            // Just to trigger view refresh
            if refresh { EmptyView() }

            Spacer().frame(height: 30)
            HStack {
                Spacer().frame(width: 100)
                VStack {
                    
                    OverlapButton( {
                        Input(title: "Import filename:", field: $inputFilename, message:nil, height: 80, keyboardType: .URL, autoCapitalize: .none, autoCorrect: false, isEnabled: false)
                    }, {
                        self.finderButton()
                    })
                }
                Spacer()
            }
            Spacer()
        }
    }
    
    private func finderButton() -> some View {
        
        return Button(action: {
            SelectInputView.findFile { (url, data) in
                Utility.mainThread {
                    refresh.toggle()
                    securityBookmark = data
                    inputFilename = url.absoluteString
                }
            }
        },label: {
            Image(systemName: "folder.fill")
                .frame(width: 24, height: 24)
                .foregroundColor(Palette.background.themeText)
        })
        .buttonStyle(PlainButtonStyle())
    }
    
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
                        let data = try url.bookmarkData(options: .withSecurityScope, includingResourceValuesForKeys: nil, relativeTo: relativeTo)
                        completion(url, data)
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
