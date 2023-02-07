//
//  Version.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 08/05/2020.
//  Copyright © 2020 Marc Shearer. All rights reserved.
//

import Foundation

class Version {
    
    static private(set) var current = Version()
    
    public var version = "0.0"
    public var build = 0
    public var lastVersion = "0.0"
    public var lastBuild = 0
    public var minVersion = ""
    public var minMessage = ""
    public var infoMessage = ""
    
    public func load() {
        // Dummy routine to trigger creation of singleton
    }
    
    init() {
        // Set up current version and build
        let dictionary = Bundle.main.infoDictionary!
        self.version = dictionary["CFBundleShortVersionString"] as! String? ?? "0.0"
        self.build = Int(dictionary["CFBundleVersion"] as! String) ?? 0
        
        // Get previous version, build etc
        lastVersion = UserDefault.lastVersion.string
        lastBuild = UserDefault.lastBuild.int
        minVersion = UserDefault.minVersion.string
        minMessage = UserDefault.minMessage.string
        infoMessage = UserDefault.infoMessage.string
    }
    
    #if !widget
    public func check(upgrade: Bool = false) {
        if upgrade {
            if Double(self.lastVersion) == 0.0 {
                // New install - just use current version
            } else if compare(self.lastVersion, self.version) == .lessThan {
                // Version has increased - check for upgrade
                MessageBox.shared.show("Upgrading to latest version...", okText: nil, showIcon: true)
                self.upgradeToVersion()
                MessageBox.shared.hide()
            }
            
            // Update last version
            updateLastVersion()
        }
        
        // Check this version is acceptable
        if compare(version, minVersion) == .lessThan {
            MessageBox.shared.show(minMessage, showIcon: true, okAction: {
                exit(1)
            })
        }
        
        // Show info message if it is setup
        if infoMessage != "" {
            MessageBox.shared.show(infoMessage, showIcon: true)
        }
    }
    
    private func upgradeToVersion() {
        if compare(lastVersion, "1.0") == .lessThan {
            // Upgrade to version 1.0
        }
    }
    
    private func updateLastVersion() {
        self.lastVersion = version
        self.lastBuild = build
        UserDefault.lastVersion.set(version)
        UserDefault.lastBuild.set(build)
    }
    
    public func set(minVersion: String, minMessage: String, infoMessage: String) {
        self.minVersion = minVersion
        self.minMessage = minMessage
        self.infoMessage = infoMessage
        UserDefault.minVersion.set(minVersion)
        UserDefault.minMessage.set(minMessage)
        UserDefault.infoMessage.set(infoMessage)
        if !MessageBox.shared.isShown {
            check()
        }
    }
    
    private func compare(_ version1: String, _ version2: String) -> Utility.CompareResult {
        return Utility.compareVersions(version1: version1, version2: version2)
    }
    #endif
}
