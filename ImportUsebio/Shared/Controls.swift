//
//  Controls.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 22/04/2026.
//

import SwiftUI

class Controls: ObservableObject {

    @Published var nextOtherNationalId: Int!

    static public var current =  Controls(load: true)
    
    init(load: Bool = false) {
        if load {
            self.load()
        }
    }
        
    public func copy() -> Controls {
        let copy = Controls()
        copy.copy(from: self)
        return copy
    }
    
    public func copy(from: Controls) {
        self.nextOtherNationalId = from.nextOtherNationalId
    }
    
    public func load() {
        self.nextOtherNationalId = UserDefault.controlNextOtherNationalId.int
    }
    
    public func save() {
        UserDefault.controlNextOtherNationalId.set(self.nextOtherNationalId)
    }
    
    public func getNextOtherNationalId() -> Int {
        let result = self.nextOtherNationalId!
        self.nextOtherNationalId! += 1
        Controls.current.save()
        return result
    }
}
