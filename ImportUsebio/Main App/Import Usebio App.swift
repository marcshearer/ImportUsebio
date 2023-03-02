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
            GeometryReader { (geometry) in
                SelectInputView()
                    .frame(minWidth: 900, maxWidth: .infinity, minHeight: 600, maxHeight: .infinity, alignment: .center)
                    .navigationTitle("Import XML Results Files")
                    .onAppear() {
                        MyApp.format = (min(geometry.size.width, geometry.size.height) < 600 ? .phone : .tablet)
                    }
            }
        }
        .windowResizability(.contentSize)
        .onChange(of: scenePhase) { phase in
            if phase == .active {
                
            }
        }
    }
}
