//
//  Import Usebio App.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

@main
struct ImportUsebioApp: App {
    init() {
        MyApp.shared.start()
    }
    
    var body: some Scene {
        MyScene()
    }
}

struct MyScene: Scene {
    @Environment(\.scenePhase) private var scenePhase
    
    var body: some Scene {
        WindowGroup {
                SelectInputView()
                    .navigationTitle("Import XML Results Files")
                    .frame(minWidth: 900, maxWidth: 900,
                           minHeight: 710, maxHeight: 710)
        }
        .windowResizability(.contentSize)
        .onChange(of: scenePhase) { phase in
            if phase == .active {
                
            }
        }
    }
}
