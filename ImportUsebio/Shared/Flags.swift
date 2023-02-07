//
//  Flags.swift
//  BridgeScore
//
//  Created by Marc Shearer on 30/03/2022.
//

import Foundation

class Flags {
    private var values: [AnyHashable : Bool]
    
    init() {
        self.values = [:]
    }
    
    public var isEmpty: Bool {
        firstValue(equal: true) == nil
    }
    
    public var hasMultiple: Bool {
        self.values.filter({$0.value}).count > 1
    }
    
    public var trueValues: [AnyHashable] {
        values.filter{$0.value}.map{$0.key}
    }
    
    public func toggle(_ id: AnyHashable) {
        if values[id] == nil {
            values[id] = true
        } else {
            values[id]!.toggle()
        }
    }
    
    public func clear(_ id: AnyHashable? = nil) {
        if let id = id {
            values[id] = false
        } else {
            values = [:]
        }
    }
    
    public var isClear: Bool {
        return values.count == 0
    }
    
    public func set(_ id: AnyHashable) {
        clear()
        values[id] = true
    }
    
    public func setArray(_ ids: [AnyHashable], to: Bool = true) {
        clear()
        for id in ids {
            self.values[id] = to
        }
    }
    
    public func value(_ id: AnyHashable) -> Bool {
        if let id = id as? String {
            return values[id] ?? false
        } else if let id = id as? Int {
            return values[id] ?? false
        } else {
            fatalError("Only support integer and string keys")
        }
    }
    
    public func firstValue(equal value: Bool) -> AnyHashable? {
        let first = values.first(where: {$0.value == value})
        return first?.key
    }
}
