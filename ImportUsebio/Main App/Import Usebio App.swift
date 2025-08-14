//
//  Import Usebio App.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 07/02/2023.
//

import SwiftUI

@main
struct ImportUsebioApp: App {
    
    public let context = PersistenceController.shared.container.viewContext
    
    init() {
        CoreData.context = context
        
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
                    .fixedSize(horizontal: true, vertical: true)
        }
        .windowResizability(.contentSize)
        .onChange(of: scenePhase, initial: false) { (_, phase) in
            if phase == .active {
                
            }
        }
    }
}
